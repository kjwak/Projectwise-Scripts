# prepend_qc_on_trigger.ps1
# Polls ProjectWise folder(s) every 30s for documents whose description contains "|QC|".
# For each: runs prepend history PDF workflow, then removes the trigger tag from the description.
#
# DEDICATED MACHINE: use a config file (one folder path per line, # = comment):
#   .\prepend_qc_on_trigger.ps1 -ConfigPath "C:\QC\watch_folders.txt"
#
# Single folder:
#   .\prepend_qc_on_trigger.ps1 -TriggerFolderPath 'pw:\\typsa-us-pw.bentley.com:typsa-us-pw-03\AZDOT 2024\AZFWY1704-FD02-SR202 - I-10 to SR101\CADD\Working\TYPSA\Drainage\JFlint\Prepend Test\incoming'
#   .\prepend_qc_on_trigger.ps1 -WatchFolderPath "path\to\folder" -DatasourceName "..."
# Multiple folders (command line):
#   .\prepend_qc_on_trigger.ps1 -WatchFolderPaths "path1","path2","path3"
# All Sheets folders under a root (discovers project\CADD\Sheets for each project under root):
#   .\prepend_qc_on_trigger.ps1 -WatchUnderRoot "Documents\AZDOT 2024" -SheetsPathFromProject "CADD\Sheets"
# Multiple roots (pipe-separated; use from powershell.exe -File launchers when string[] binding is unreliable):
#   .\prepend_qc_on_trigger.ps1 -WatchUnderRootJoined "Documents\AZDOT 2024|Documents\AZDOT" -SheetsPathFromProject "CADD\Sheets"
# One shot: .\prepend_qc_on_trigger.ps1 -ConfigPath "C:\QC\watch_folders.txt" -RunOnce
# Logging: activity + errors to C:\PW_QC_LOCAL\logs\ (override with -LogDir)
[CmdletBinding(SupportsShouldProcess = $true)]
param(
  [Parameter(Mandatory = $false)]
  [string] $WatchFolderPath,

  [Parameter(Mandatory = $false)]
  [string] $TriggerFolderPath,

  [Parameter(Mandatory = $false)]
  [string[]] $WatchFolderPaths,

  # Additional folder paths to watch (merged with other sources like -WatchUnderRoot discovery).
  [Parameter(Mandatory = $false)]
  [string[]] $ExtraWatchFolderPaths,

  [Parameter(Mandatory = $false)]
  [string] $ConfigPath,

  # Structured watch list (JSON) that can include roots + explicit folders with per-folder oneLevelDeep.
  [Parameter(Mandatory = $false)]
  [string] $WatchListPath,

  [Parameter(Mandatory = $false)]
  [string] $WatchUnderRoot,

  # Pipe-separated watch roots (same pattern as combine_status_set.ps1).
  [Parameter(Mandatory = $false)]
  [string] $WatchUnderRootJoined,

  [Parameter(Mandatory = $false)]
  [string] $SheetsPathFromProject = "CADD\Sheets",

  [Parameter(Mandatory = $false)]
  [string] $DatasourceName = "typsa-us-pw.bentley.com:typsa-us-pw-03",

  [Parameter(Mandatory = $false)]
  [int] $PollIntervalSeconds = 30,

  # When set: also scan the immediate child folders (one level deep) of each watched folder.
  [Parameter(Mandatory = $false)]
  [switch] $OneLevelDeep,

  [Parameter(Mandatory = $false)]
  [switch] $RunOnce,

  [Parameter(Mandatory = $false)]
  [string] $PrependScriptPath,

  [Parameter(Mandatory = $false)]
  [int] $BatchCooldownSeconds = 5,

  [Parameter(Mandatory = $false)]
  [switch] $PromptForCredential,

  [Parameter(Mandatory = $false)]
  [string] $LogDir = "C:\PW_QC_LOCAL\logs",

  # Passed to prepend_qc.ps1 — same as default; override if current-master / work files live elsewhere.
  [Parameter(Mandatory = $false)]
  [string] $LocalRoot = "C:\PW_QC_LOCAL",

  # $true (default): qpdf page 1 -> TEMP --current-master only when OverlaySheetWorkDir:$false. $false: persistent work\ current-master.
  [Parameter(Mandatory = $false)]
  $OverlayOldFromHistoryOnly = $true,

  # $true (default): LocalRoot\work\<sheet>\ split pages + MANIFEST. Passed to prepend_qc.ps1.
  [Parameter(Mandatory = $false)]
  $OverlaySheetWorkDir = $true
)

$ErrorActionPreference = "Stop"
$TriggerTag = "QC_Archivist"

. "$PSScriptRoot\Logging.ps1"
. "$PSScriptRoot\StaMtaRelaunch.ps1"

if ($PSBoundParameters.ContainsKey('OverlayOldFromHistoryOnly')) {
  $OverlayOldFromHistoryOnly = ConvertTo-BoolLoose $PSBoundParameters['OverlayOldFromHistoryOnly']
} else {
  $OverlayOldFromHistoryOnly = $true
}
if ($PSBoundParameters.ContainsKey('OverlaySheetWorkDir')) {
  $OverlaySheetWorkDir = ConvertTo-BoolLoose $PSBoundParameters['OverlaySheetWorkDir']
} else {
  $OverlaySheetWorkDir = $true
}

$WatchRootList = @()
if ($WatchUnderRootJoined -and $WatchUnderRootJoined.Trim()) {
  $WatchRootList = @($WatchUnderRootJoined -split '\|' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
} elseif ($WatchUnderRoot -and $WatchUnderRoot.Trim()) {
  $WatchRootList = @($WatchUnderRoot.Trim())
}
$useWatchUnderRoot = $WatchRootList.Count -gt 0

# ProjectWise credentials: C:\PW_QC_LOCAL\pw_cred.txt
# Format:
#   username=domain\user
#   password=PlainTextPassword
# NOTE: Plain text; use only on a locked-down dedicated machine.
$CredentialPath = 'C:\PW_QC_LOCAL\pw_cred.txt'

function Get-PwCredential {
  if ($PromptForCredential) {
    return Get-Credential -Message "ProjectWise login for $DatasourceName"
  }
  if (-not (Test-Path -LiteralPath $CredentialPath)) {
    throw "Credential file not found: $CredentialPath. Create it with lines: username=..., password=..."
  }
  $lines = Get-Content -LiteralPath $CredentialPath -ErrorAction Stop
  $uLine = $lines | Where-Object { $_ -match '^\s*username\s*=' } | Select-Object -First 1
  $pLine = $lines | Where-Object { $_ -match '^\s*password\s*=' } | Select-Object -First 1
  if (-not $uLine -or -not $pLine) {
    throw "Invalid format in $CredentialPath. Expected lines: username=..., password=..."
  }
  $user = ($uLine -split '=', 2)[1].Trim()
  $pass = ($pLine -split '=', 2)[1].Trim()
  $sec  = ConvertTo-SecureString $pass -AsPlainText -Force
  return [pscredential]::new($user, $sec)
}

$scriptDir = $PSScriptRoot
if (-not $WatchListPath) {
  $candidate = Join-Path $scriptDir "watchlist.json"
  if (Test-Path -LiteralPath $candidate) { $WatchListPath = $candidate }
}

$script:PwpsWarnShown = $false
function Connect-PW([string]$dsName) {
  $cred = Get-PwCredential
  $open = {
    if (-not $script:PwpsWarnShown) {
      Open-PWConnection -DatasourceName $dsName -UserName $cred.UserName -Password $cred.Password | Out-Null
      $script:PwpsWarnShown = $true
    } else {
      Open-PWConnection -DatasourceName $dsName -UserName $cred.UserName -Password $cred.Password -WarningAction SilentlyContinue | Out-Null
    }
  }
  try {
    & $open
  } catch {
    if ($_.Exception.Message -match 'connection is already open') {
      Close-PWConnection -ErrorAction SilentlyContinue
      & $open
    } else { throw }
  }
}

# Normalize one path string to { DatasourceName, FolderPath }. Only first segment is datasource when path starts with pw:\
function Get-NormalizedFolder([string]$path, [string]$defaultDs) {
  $raw = ($path -as [string]).Trim()
  if (-not $raw) { return $null }
  $isFull = $raw -match '^pw:\\?' -or $raw -match '^pw:'
  if ($isFull) {
    $p = $raw -replace '^pw:\\?', '' -replace '^pw:', ''
    $parts = $p -split '\\', 2
    if ($parts.Count -ge 2) {
      $fp = $parts[1].Trim().TrimEnd('\')
      $fp = $fp -replace '^Documents\\', ''
      return @{ DatasourceName = $parts[0].Trim(); FolderPath = $fp }
    }
    $fp = $p.Trim().TrimEnd('\')
    $fp = $fp -replace '^Documents\\', ''
    return @{ DatasourceName = $defaultDs; FolderPath = $fp }
  }
  # Also accept the common "datasource\folder\path" form (without the pw:\ prefix).
  # This prevents DatasourceName from being lost when callers pass a full path that already includes it.
  $parts2 = $raw -split '\\', 2
  if ($parts2.Count -ge 2) {
    $maybeDs = $parts2[0].Trim()
    $rest = $parts2[1].Trim().TrimEnd('\')
    if ($maybeDs -match '[:]' -and $maybeDs -match '[.]') {
      $rest = $rest -replace '^Documents\\', ''
      return @{ DatasourceName = $maybeDs; FolderPath = $rest }
    }
  }
  $fp = $raw.TrimEnd('\')
  $fp = $fp -replace '^Documents\\', ''
  return @{ DatasourceName = $defaultDs; FolderPath = $fp }
}

# De-dupe folder entries and prefer OneLevelDeep=$true when duplicates exist.
function Select-FolderEntriesUniquePreferDeep {
  param([Parameter(Mandatory = $true)] $Entries)
  $map = @{}
  foreach ($e in @($Entries)) {
    if (-not $e -or -not $e.DatasourceName -or -not $e.FolderPath) { continue }
    $key = ($e.DatasourceName + '|' + $e.FolderPath).ToLowerInvariant()
    $deep = $false
    if ($e.ContainsKey('OneLevelDeep')) { $deep = [bool]$e['OneLevelDeep'] }
    if (-not $map.ContainsKey($key)) {
      $map[$key] = $e
      continue
    }
    $existing = $map[$key]
    $existingDeep = $false
    if ($existing.ContainsKey('OneLevelDeep')) { $existingDeep = [bool]$existing['OneLevelDeep'] }
    if ($deep -and -not $existingDeep) {
      $map[$key] = $e
    }
  }
  return @($map.Values)
}

# Loads a JSON watchlist with schema:
# { roots: [ { path, sheetsPathFromProject?, oneLevelDeep? } | "path" ],
#   folders: [ { path, oneLevelDeep? } | "path" ] }
function Get-WatchListFromFile {
  param([Parameter(Mandatory = $true)][string] $Path)
  if (-not (Test-Path -LiteralPath $Path)) { return $null }
  if ($Path -match '\.json$') {
    try {
      $raw = Get-Content -LiteralPath $Path -Encoding UTF8 -Raw -ErrorAction Stop
      $obj = $raw | ConvertFrom-Json -ErrorAction Stop
      return $obj
    } catch {
      Write-Log "WatchListPath JSON load failed ($Path): $_" -Severity WARNING
      return $null
    }
  }
  return $null
}

# Build folder list: ConfigPath > WatchFolderPaths > single WatchFolderPath/TriggerFolderPath. Returns @() when WatchUnderRoot is used (discovery in loop).
function Get-FolderList {
  $list = @()
  if ($ConfigPath -and (Test-Path -LiteralPath $ConfigPath)) {
    $lines = Get-Content -LiteralPath $ConfigPath -Encoding UTF8 -ErrorAction SilentlyContinue
    foreach ($line in $lines) {
      $line = $line.Trim()
      if (-not $line -or $line.StartsWith('#')) { continue }
      $n = Get-NormalizedFolder $line $DatasourceName
      if ($n -and $n.FolderPath) { $list += $n }
    }
  } else {
    if ($WatchFolderPaths -and $WatchFolderPaths.Count -gt 0) {
      foreach ($wp in $WatchFolderPaths) {
        $n = Get-NormalizedFolder $wp $DatasourceName
        if ($n -and $n.FolderPath) { $list += $n }
      }
    } elseif (-not $useWatchUnderRoot) {
      $single = if ($TriggerFolderPath) { $TriggerFolderPath } else { $WatchFolderPath }
      $n = Get-NormalizedFolder $single $DatasourceName
      if ($n -and $n.FolderPath) { $list += $n }
    }
  }

  if ($ExtraWatchFolderPaths -and $ExtraWatchFolderPaths.Count -gt 0) {
    foreach ($wp in $ExtraWatchFolderPaths) {
      $n = Get-NormalizedFolder $wp $DatasourceName
      if ($n -and $n.FolderPath) {
        $n['OneLevelDeep'] = $OneLevelDeep.IsPresent
        $list += $n
      }
    }
  }

  foreach ($it in @($list)) {
    if (-not $it.ContainsKey('OneLevelDeep')) { $it['OneLevelDeep'] = $OneLevelDeep.IsPresent }
  }
  return @($list | Where-Object { $_ -and $_.FolderPath } | Select-Object -Unique)
}

# When WatchUnderRoot is set: connect, list immediate children of root, return list of { DatasourceName, FolderPath } for each child\SheetsPathFromProject.
# Leading "Documents\" is stripped from the path for the API (PW often uses paths without that segment).
function Get-SheetsFoldersUnderRoot {
  param([Parameter(Mandatory = $true)][string] $RootForDiscovery)
  $rootEntry = Get-NormalizedFolder $RootForDiscovery $DatasourceName
  if (-not $rootEntry) { return @() }
  $rootPath = $rootEntry.FolderPath -replace '^Documents\\', ''
  $ds = $rootEntry.DatasourceName
  $childNames = @()
  try {
    $view = Get-PWFolderView -FolderPath $rootPath -ErrorAction Stop
    if ($view.Children) {
      foreach ($c in $view.Children) {
        $name = $c.Name
        if (-not $name -and $c.PSObject.Properties['Name']) { $name = $c.Name }
        if (-not $name -and $c.FolderPath) { $name = [System.IO.Path]::GetFileName($c.FolderPath.TrimEnd('\')) }
        if ($name) { $childNames += $name }
      }
    }
    if ($childNames.Count -eq 0 -and $view.Folders) {
      foreach ($f in $view.Folders) {
        $name = $f.Name; if (-not $name -and $f.PSObject.Properties['Name']) { $name = $f.Name }
        if ($name) { $childNames += $name }
      }
    }
  } catch {
    try {
      $children = Get-PWFoldersImmediateChildren -FolderPath $rootPath -ErrorAction Stop
      foreach ($c in @($children)) {
        $name = $c.Name; if (-not $name -and $c.PSObject.Properties['Name']) { $name = $c.Name }
        if (-not $name -and $c.FolderPath) { $name = [System.IO.Path]::GetFileName($c.FolderPath.TrimEnd('\')) }
        if ($name) { $childNames += $name }
      }
    } catch { Write-Log "Could not list children of $rootPath : $_" -Severity WARNING; return @() }
  }
  $list = @()
  $suffix = $SheetsPathFromProject.Trim().TrimStart('\')
  foreach ($name in $childNames) {
    $folderPath = $rootPath.TrimEnd('\') + '\' + $name + '\' + $suffix
    $list += @{ DatasourceName = $ds; FolderPath = $folderPath }
  }
  return $list
}

$folderList = @(Get-FolderList)
if ($folderList.Count -eq 0 -and -not $useWatchUnderRoot -and -not ($WatchListPath -and (Test-Path -LiteralPath $WatchListPath))) {
  $hint = if ($ConfigPath -and -not (Test-Path -LiteralPath $ConfigPath)) {
    " (-ConfigPath was set but file not found: $ConfigPath; create it on this machine or fix the path)"
  } else { "" }
  throw "No folders to watch.$hint Use run_prepend_qc.ps1 as launcher, or pass -ConfigPath / -WatchFolderPaths / -WatchUnderRoot / -WatchUnderRootJoined / -WatchFolderPath / -TriggerFolderPath."
}

# pwps_dab requires MTA. Re-launch with same bound parameters as strings only (Build-PowerShellExeFileArgs in StaMtaRelaunch.ps1).
# PSBoundParameters can include common parameters (-Verbose, -ErrorAction, -WhatIf, etc.); those are not in our param()
# block and break child -File parsing with "Cannot process argument" if forwarded.
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -eq 'STA') {
  $scriptPath = $PSCommandPath
  if (-not $scriptPath) { $scriptPath = $MyInvocation.MyCommand.Path }
  if (-not $scriptPath) { $scriptPath = Join-Path $PSScriptRoot 'prepend_qc_on_trigger.ps1' }
  if (-not (Test-Path -LiteralPath $scriptPath)) {
    throw "MTA relaunch: could not resolve script path (PSCommandPath / MyInvocation). Tried: $scriptPath"
  }
  $paramNames = @(
    'WatchFolderPath', 'TriggerFolderPath', 'WatchFolderPaths', 'ExtraWatchFolderPaths', 'ConfigPath', 'WatchListPath', 'WatchUnderRoot', 'WatchUnderRootJoined',
    'SheetsPathFromProject', 'DatasourceName', 'PollIntervalSeconds', 'RunOnce', 'PrependScriptPath',
    'BatchCooldownSeconds', 'PromptForCredential', 'LogDir', 'LocalRoot', 'OverlayOldFromHistoryOnly', 'OverlaySheetWorkDir',
    'OneLevelDeep'
  )
  $bp = @{}
  foreach ($n in $paramNames) {
    if ($PSBoundParameters.ContainsKey($n)) { $bp[$n] = $PSBoundParameters[$n] }
  }
  $exeArgs = Build-PowerShellExeFileArgs -ScriptPath $scriptPath -BoundParameters $bp
  & powershell.exe @exeArgs
  exit $LASTEXITCODE
}

if (-not $PrependScriptPath) { $PrependScriptPath = Join-Path $scriptDir "prepend_qc.ps1" }

function Get-ImmediateChildFolderPaths {
  param(
    [Parameter(Mandatory = $true)][string] $FolderPath
  )
  $childFolderPaths = @()
  try {
    $children = @(Get-PWFoldersImmediateChildren -FolderPath $FolderPath -ErrorAction Stop)
    foreach ($c in $children) {
      $fp = $null
      if ($c -and $c.PSObject.Properties['FolderPath']) { $fp = $c.FolderPath }
      if ($fp) {
        $childFolderPaths += $fp.TrimEnd('\')
        continue
      }
      $name = $null
      if ($c -and $c.PSObject.Properties['Name']) { $name = $c.Name }
      if ($name) { $childFolderPaths += (($FolderPath.TrimEnd('\') + '\' + $name.Trim()) -replace '\\{2,}', '\') }
    }
  } catch { }

  $view = $null
  try { $view = Get-PWFolderView -FolderPath $FolderPath -ErrorAction Stop } catch { $view = $null }

  if ($view) {
    if ($view.Children) {
      foreach ($c in @($view.Children)) {
        $fp = $null
        if ($c -and $c.PSObject.Properties['FolderPath']) { $fp = $c.FolderPath }
        $hasDocId = $false
        if ($c.PSObject.Properties['DocumentID'] -and $c.DocumentID) { $hasDocId = $true }
        if ($hasDocId) { continue }
        if ($fp) {
          $childFolderPaths += $fp.TrimEnd('\')
          continue
        }
        $name = $null
        if ($c -and $c.PSObject.Properties['Name']) { $name = $c.Name }
        if ($name) { $childFolderPaths += (($FolderPath.TrimEnd('\') + '\' + $name.Trim()) -replace '\\{2,}', '\') }
      }
    }
    if ($childFolderPaths.Count -eq 0 -and $view.Folders) {
      foreach ($f in @($view.Folders)) {
        $fp = $null
        if ($f -and $f.PSObject.Properties['FolderPath']) { $fp = $f.FolderPath }
        if ($fp) {
          $childFolderPaths += $fp.TrimEnd('\')
          continue
        }
        $name = $null
        if ($f -and $f.PSObject.Properties['Name']) { $name = $f.Name }
        if ($name) { $childFolderPaths += (($FolderPath.TrimEnd('\') + '\' + $name.Trim()) -replace '\\{2,}', '\') }
      }
    }
  }
  return @($childFolderPaths | Where-Object { $_ } | Select-Object -Unique)
}

$folderDesc = if ($WatchListPath) { "WatchList: $WatchListPath" }
  elseif ($ConfigPath) { "Config: $ConfigPath ($($folderList.Count) folders)" }
  elseif ($useWatchUnderRoot) { "Under root(s): $(@($WatchRootList) -join ' | ') -> *\$SheetsPathFromProject (discover each poll)" }
  else { "$($folderList.Count) folder(s)" }
Write-Log "Watching $folderDesc | OneLevelDeep: $OneLevelDeep | Poll: $PollIntervalSeconds s | RunOnce: $RunOnce"

Import-Module pwps_dab -Force

while ($true) {
  $watchListObj = $null
  if ($WatchListPath) { $watchListObj = Get-WatchListFromFile -Path $WatchListPath }

  # If watchlist.json provides roots, prefer them (still allow CLI AddWatchUnderRoot via WatchUnderRootJoined).
  $rootsFromFile = @()
  if ($watchListObj -and $watchListObj.roots) {
    foreach ($r in @($watchListObj.roots)) {
      if ($r -is [string]) {
        $rootsFromFile += @{ path = $r; sheetsPathFromProject = $null; oneLevelDeep = $null }
      } else {
        $rootsFromFile += $r
      }
    }
  }

  if ($rootsFromFile.Count -gt 0) {
    $WatchRootList = @($rootsFromFile | ForEach-Object { ($_.'path' -as [string]).Trim() } | Where-Object { $_ } | Select-Object -Unique)
    $useWatchUnderRoot = $WatchRootList.Count -gt 0
  }

  $folderList = @(Get-FolderList)
  if ($useWatchUnderRoot -and $folderList.Count -eq 0) {
    $merged = @()
    foreach ($root in $WatchRootList) {
      $rootEntry = Get-NormalizedFolder $root $DatasourceName
      if (-not $rootEntry) { continue }
      try {
        Connect-PW $rootEntry.DatasourceName
        $discovered = @(Get-SheetsFoldersUnderRoot -RootForDiscovery $root)
        # Apply root-level overrides from file if present (sheetsPathFromProject / oneLevelDeep)
        $rootCfg = $null
        if ($rootsFromFile.Count -gt 0) {
          $rootCfg = @($rootsFromFile | Where-Object { (($_.'path' -as [string]).Trim()) -eq $root } | Select-Object -First 1)
        }
        $rootDeep = $null
        if ($rootCfg -and $rootCfg.PSObject.Properties['oneLevelDeep']) { $rootDeep = [bool]$rootCfg.oneLevelDeep }
        foreach ($d in @($discovered)) {
          if ($null -ne $rootDeep) { $d['OneLevelDeep'] = $rootDeep } else { $d['OneLevelDeep'] = $OneLevelDeep.IsPresent }
        }
        $merged += $discovered
        Write-Log "Discovered $($discovered.Count) Sheets folders under $($rootEntry.FolderPath)"
      } catch { Write-Log "WatchUnderRoot discovery failed for ${root}: $_" -Severity WARNING }
      Close-PWConnection -ErrorAction SilentlyContinue
    }
    $folderList = @($merged)
  }

  # Merge explicit folders from watchlist.json (each may declare oneLevelDeep).
  if ($watchListObj -and $watchListObj.folders) {
    $extras = @()
    foreach ($f in @($watchListObj.folders)) {
      if ($f -is [string]) {
        $n = Get-NormalizedFolder $f $DatasourceName
        if ($n -and $n.FolderPath) { $n['OneLevelDeep'] = $OneLevelDeep.IsPresent; $extras += $n }
      } else {
        $p = ($f.'path' -as [string]).Trim()
        if (-not $p) { continue }
        $rootPrefix = $null
        if ($f.PSObject.Properties['root']) { $rootPrefix = ($f.'root' -as [string]).Trim() }
        if ($rootPrefix) {
          $p = $rootPrefix.TrimEnd('\') + '\' + $p.TrimStart('\')
        }
        $n = Get-NormalizedFolder $p $DatasourceName
        if (-not $n -or -not $n.FolderPath) { continue }
        $deep = $null
        if ($f.PSObject.Properties['oneLevelDeep']) { $deep = [bool]$f.oneLevelDeep }
        $n['OneLevelDeep'] = if ($null -ne $deep) { $deep } else { $OneLevelDeep.IsPresent }
        $extras += $n
      }
    }
    if ($extras.Count -gt 0) {
      $folderList = @($folderList + $extras | Where-Object { $_ -and $_.DatasourceName -and $_.FolderPath })
    }
  }

  $folderList = @(Select-FolderEntriesUniquePreferDeep -Entries $folderList |
      Sort-Object @{ Expression = { $_.DatasourceName } }, @{ Expression = { $_.FolderPath } })

  Write-Log "Watch set: $($folderList.Count) folder(s) total."
  $cal = @($folderList | Where-Object { $_.FolderPath -like '*Caltrans*' } | Select-Object -First 5)
  foreach ($c in $cal) {
    $d = if ($c.ContainsKey('OneLevelDeep')) { [bool]$c['OneLevelDeep'] } else { $OneLevelDeep.IsPresent }
    Write-Log "Watch Caltrans: [$($c.DatasourceName)] $($c.FolderPath) | OneLevelDeep: $d"
  }

  foreach ($entry in $folderList) {
    $WatchFolderPath = $entry.FolderPath
    $DatasourceName = $entry.DatasourceName
    $entryDeep = $false
    if ($entry.ContainsKey('OneLevelDeep')) { $entryDeep = [bool]$entry['OneLevelDeep'] } else { $entryDeep = $OneLevelDeep.IsPresent }
    Write-Log "[$WatchFolderPath] Scanning folder."
    try {
      #Open-PWConnection -DatasourceName $DatasourceName -BentleyIMS | Out-Null
      Connect-PW $DatasourceName
    } catch {
      Write-Log "Connect failed for $WatchFolderPath : $_" -Severity WARNING
      Close-PWConnection -ErrorAction SilentlyContinue
      continue
    }

    $foldersToScan = @($WatchFolderPath)
    if ($entryDeep) {
      try {
        $children = @(Get-ImmediateChildFolderPaths -FolderPath $WatchFolderPath)
        Write-Log "[$WatchFolderPath] OneLevelDeep: +$($children.Count) child folder(s)"
        if ($children.Count -gt 0) { $foldersToScan += $children }
      } catch {
        Write-Log "[$WatchFolderPath] OneLevelDeep child listing failed: $_" -Severity WARNING
      }
    }
    $foldersToScan = @($foldersToScan | Where-Object { $_ } | Select-Object -Unique)

    foreach ($folderPathToScan in $foldersToScan) {
      # Get all documents in the folder
      $allDocs = @()
      $view = $null
      try { $view = Get-PWFolderView -FolderPath $folderPathToScan -ErrorAction Stop } catch { }
      if ($view -and $view.Documents) { $allDocs = @($view.Documents) }
      elseif ($view -and $view.Children) {
        $allDocs = @($view.Children | Where-Object { $_.DocumentID -or $_.Name })
      }
      if ($allDocs.Count -eq 0) {
        try {
          $allDocs = @(Get-PWDocumentsBySearch -FolderPath $folderPathToScan -JustThisFolder -PopulatePath -ErrorAction Stop)
        } catch {
          $withCols = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $folderPathToScan -JustThisFolder -ReturnColumns @("Description", "Name", "DocumentID") -ErrorAction SilentlyContinue
          if ($withCols) { $allDocs = @($withCols) }
        }
      }
      if ($allDocs.Count -eq 0) {
        Write-Log "[$folderPathToScan] No documents found in folder."
      }

      # Filter to documents whose description contains the trigger
      $triggerDocs = @()
      foreach ($doc in $allDocs) {
        $desc = $null
        if (Get-Member -InputObject $doc -Name Description -MemberType Properties -ErrorAction SilentlyContinue) { $desc = $doc.Description }
        if ($null -eq $desc -and $doc.PSObject.Properties['Description']) { $desc = $doc.Description }
        if ($null -eq $desc) { $desc = "" }
        if ($desc -like "*$TriggerTag*") { $triggerDocs += $doc }
      }

      if ($triggerDocs.Count -eq 0) { continue }

      Close-PWConnection -ErrorAction SilentlyContinue
      Write-Log "[$folderPathToScan] Found $($triggerDocs.Count) document(s) with trigger tag."

      foreach ($doc in $triggerDocs) {
        $docName = $doc.Name
        if (-not $docName -and $doc.PSObject.Properties['Name']) { $docName = $doc.Name }
        if (-not $docName -and $doc.DocumentName) { $docName = $doc.DocumentName }
        if (-not $docName) { $docName = $doc.FullPath; $docName = [System.IO.Path]::GetFileName($docName) }
        $incomingPdf = [System.IO.Path]::GetFileNameWithoutExtension($docName) + ".pdf"

        Write-Log "Processing: $docName (incoming PDF: $incomingPdf)"

        $prependParams = @{
          IncomingFolderPath         = $folderPathToScan
          IncomingDocName            = $incomingPdf
          DatasourceName             = $DatasourceName
          LogDir                     = $LogDir
          LocalRoot                  = $LocalRoot
          OverlayOldFromHistoryOnly  = $OverlayOldFromHistoryOnly
          OverlaySheetWorkDir        = $OverlaySheetWorkDir
        }
        # Resolve bundled qpdf / overlay next to prepend_qc.ps1 (run_prepend_qc often uses a deploy folder without PATH entries).
        $prependRoot = Split-Path -Parent $PrependScriptPath
        foreach ($q in @(
            (Join-Path $prependRoot "tools\qpdf\bin\qpdf.exe"),
            (Join-Path $prependRoot "tools\qpdf\qpdf.exe")
          )) {
          if (Test-Path -LiteralPath $q) {
            $prependParams['QpdfExe'] = $q
            break
          }
        }
        $ov = Join-Path $prependRoot "dist\qc_overlay_prepend\qc_overlay_prepend.exe"
        if (Test-Path -LiteralPath $ov) {
          $prependParams['QcOverlayExe'] = $ov
        }

        try {
          & $PrependScriptPath @prependParams
          if (-not $?) { Write-Log "Prepend failed for $docName" -Severity WARNING; continue }
        } catch {
          Write-Log "Prepend failed for $docName : $_" -Severity WARNING
          continue
        }

        Close-PWConnection -ErrorAction SilentlyContinue
        try { Connect-PW $DatasourceName } catch { }
        $triggerDoc = Get-PWDocumentsBySearch -FolderPath $folderPathToScan -JustThisFolder -DocumentName $docName -PopulatePath
        if (-not $triggerDoc) {
          Write-Log "Could not re-find document to clear tag: $docName" -Severity WARNING
          continue
        }
        $currentDesc = $triggerDoc.Description
        if ($null -eq $currentDesc -and $triggerDoc.PSObject.Properties['Description']) { $currentDesc = $triggerDoc.Description }
        if ($null -eq $currentDesc) { $currentDesc = "" }
        $newDesc = ($currentDesc -replace [regex]::Escape($TriggerTag), "").Trim()

        if ($PSCmdlet.ShouldProcess($triggerDoc.FullPath, "Update description (remove trigger tag)")) {
          try {
            $triggerDoc.Description = $newDesc
            Update-PWDocumentProperties $triggerDoc
            Write-Log "Description updated; |QC| tag removed for $docName"
          } catch {
            Write-Log "Update-PWDocumentProperties failed for $docName : $_" -Severity WARNING
          }
        }
        Close-PWConnection -ErrorAction SilentlyContinue
        if ($BatchCooldownSeconds -gt 0) {
          Start-Sleep -Seconds $BatchCooldownSeconds
        }
      }

      try { Connect-PW $DatasourceName } catch { }
    }
    Close-PWConnection -ErrorAction SilentlyContinue
  }

  if ($RunOnce) {
    Write-Log "Done."
    exit 0
  }
  Start-Sleep -Seconds $PollIntervalSeconds
}