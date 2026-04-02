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

  [Parameter(Mandatory = $false)]
  [string] $ConfigPath,

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

  [Parameter(Mandatory = $false)]
  [switch] $RunOnce,

  [Parameter(Mandatory = $false)]
  [string] $PrependScriptPath,

  [Parameter(Mandatory = $false)]
  [int] $BatchCooldownSeconds = 5,

  [Parameter(Mandatory = $false)]
  [switch] $PromptForCredential,

  [Parameter(Mandatory = $false)]
  [string] $LogDir = "C:\PW_QC_LOCAL\logs"
)

$ErrorActionPreference = "Stop"
$TriggerTag = "QC_Archivist"

$PrependQc_LogDir = $LogDir
. "$PSScriptRoot\Logging.ps1"

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
      return @{ DatasourceName = $parts[0].Trim(); FolderPath = $parts[1].Trim().TrimEnd('\') }
    }
    return @{ DatasourceName = $defaultDs; FolderPath = $p.Trim().TrimEnd('\') }
  }
  return @{ DatasourceName = $defaultDs; FolderPath = $raw.TrimEnd('\') }
}

# Build folder list: ConfigPath > WatchFolderPaths > single WatchFolderPath/TriggerFolderPath. Returns @() when WatchUnderRoot is used (discovery in loop).
function Get-FolderList {
  if ($ConfigPath -and (Test-Path -LiteralPath $ConfigPath)) {
    $lines = Get-Content -LiteralPath $ConfigPath -Encoding UTF8 -ErrorAction SilentlyContinue
    $list = @()
    foreach ($line in $lines) {
      $line = $line.Trim()
      if (-not $line -or $line.StartsWith('#')) { continue }
      $n = Get-NormalizedFolder $line $DatasourceName
      if ($n -and $n.FolderPath) { $list += $n }
    }
    return $list
  }
  if ($WatchFolderPaths -and $WatchFolderPaths.Count -gt 0) {
    $list = @()
    foreach ($wp in $WatchFolderPaths) {
      $n = Get-NormalizedFolder $wp $DatasourceName
      if ($n -and $n.FolderPath) { $list += $n }
    }
    return $list
  }
  if ($useWatchUnderRoot) { return @() }
  $single = if ($TriggerFolderPath) { $TriggerFolderPath } else { $WatchFolderPath }
  $n = Get-NormalizedFolder $single $DatasourceName
  if ($n -and $n.FolderPath) { return @($n) }
  return @()
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
if ($folderList.Count -eq 0 -and -not $useWatchUnderRoot) {
  throw "No folders to watch. Use -ConfigPath, -WatchFolderPaths, -WatchUnderRoot / -WatchUnderRootJoined, or -WatchFolderPath / -TriggerFolderPath."
}

# pwps_dab requires MTA. Re-launch with same params.
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -eq 'STA') {
  $passThrough = @('-DatasourceName', $DatasourceName, '-PollIntervalSeconds', $PollIntervalSeconds)
  if ($ConfigPath) { $passThrough += '-ConfigPath'; $passThrough += $ConfigPath }
  elseif ($useWatchUnderRoot) {
    if ($WatchRootList.Count -gt 1) {
      $passThrough += '-WatchUnderRootJoined'; $passThrough += ($WatchRootList -join '|')
    } else {
      $passThrough += '-WatchUnderRoot'; $passThrough += $WatchRootList[0]
    }
    $passThrough += '-SheetsPathFromProject'; $passThrough += $SheetsPathFromProject
  }
  elseif ($WatchFolderPaths -and $WatchFolderPaths.Count -gt 0) { $passThrough += '-WatchFolderPaths'; $passThrough += $WatchFolderPaths }
  else { $passThrough += '-WatchFolderPath'; $passThrough += $folderList[0].FolderPath }
  if ($RunOnce) { $passThrough += '-RunOnce' }
  if ($PrependScriptPath) { $passThrough += '-PrependScriptPath'; $passThrough += $PrependScriptPath }
  if ($LogDir) { $passThrough += '-LogDir'; $passThrough += $LogDir }
  & powershell.exe -MTA -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath @passThrough
  exit $LASTEXITCODE
}

$scriptDir = $PSScriptRoot
if (-not $PrependScriptPath) { $PrependScriptPath = Join-Path $scriptDir "prepend_qc.ps1" }

$folderDesc = if ($ConfigPath) { "Config: $ConfigPath ($($folderList.Count) folders)" }
  elseif ($useWatchUnderRoot) { "Under root(s): $(@($WatchRootList) -join ' | ') -> *\$SheetsPathFromProject (discover each poll)" }
  else { "$($folderList.Count) folder(s)" }
Write-Log "Watching $folderDesc | Poll: $PollIntervalSeconds s | RunOnce: $RunOnce"

Import-Module pwps_dab -Force

while ($true) {
  $folderList = @(Get-FolderList)
  if ($useWatchUnderRoot -and $folderList.Count -eq 0) {
    $merged = @()
    foreach ($root in $WatchRootList) {
      $rootEntry = Get-NormalizedFolder $root $DatasourceName
      if (-not $rootEntry) { continue }
      try {
        Connect-PW $rootEntry.DatasourceName
        $discovered = @(Get-SheetsFoldersUnderRoot -RootForDiscovery $root)
        $merged += $discovered
        Write-Log "Discovered $($discovered.Count) Sheets folders under $($rootEntry.FolderPath)"
      } catch { Write-Log "WatchUnderRoot discovery failed for ${root}: $_" -Severity WARNING }
      Close-PWConnection -ErrorAction SilentlyContinue
    }
    $folderList = @($merged)
  }
  foreach ($entry in $folderList) {
    $WatchFolderPath = $entry.FolderPath
    $DatasourceName = $entry.DatasourceName
    Write-Log "[$WatchFolderPath] Scanning folder."
    try {
      #Open-PWConnection -DatasourceName $DatasourceName -BentleyIMS | Out-Null
      Connect-PW $DatasourceName
    } catch {
      Write-Log "Connect failed for $WatchFolderPath : $_" -Severity WARNING
      Close-PWConnection -ErrorAction SilentlyContinue
      continue
    }

    # Get all documents in the folder
    $allDocs = @()
    $view = $null
    try { $view = Get-PWFolderView -FolderPath $WatchFolderPath -ErrorAction Stop } catch { }
    if ($view -and $view.Documents) { $allDocs = @($view.Documents) }
    elseif ($view -and $view.Children) {
      $allDocs = @($view.Children | Where-Object { $_.DocumentID -or $_.Name })
    }
    if ($allDocs.Count -eq 0) {
      try {
        $allDocs = @(Get-PWDocumentsBySearch -FolderPath $WatchFolderPath -JustThisFolder -PopulatePath -ErrorAction Stop)
      } catch {
        $withCols = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $WatchFolderPath -JustThisFolder -ReturnColumns @("Description", "Name", "DocumentID") -ErrorAction SilentlyContinue
        if ($withCols) { $allDocs = @($withCols) }
      }
    }
    if ($allDocs.Count -eq 0) {
      Write-Log "[$WatchFolderPath] No documents found in folder."
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

    Close-PWConnection -ErrorAction SilentlyContinue

    if ($triggerDocs.Count -eq 0) { continue }

    Write-Log "[$WatchFolderPath] Found $($triggerDocs.Count) document(s) with trigger tag."

    foreach ($doc in $triggerDocs) {
      $docName = $doc.Name
      if (-not $docName -and $doc.PSObject.Properties['Name']) { $docName = $doc.Name }
      if (-not $docName -and $doc.DocumentName) { $docName = $doc.DocumentName }
      if (-not $docName) { $docName = $doc.FullPath; $docName = [System.IO.Path]::GetFileName($docName) }
      $incomingPdf = [System.IO.Path]::GetFileNameWithoutExtension($docName) + ".pdf"

      Write-Log "Processing: $docName (incoming PDF: $incomingPdf)"

      $prependParams = @{
        IncomingFolderPath = $WatchFolderPath
        IncomingDocName    = $incomingPdf
        DatasourceName     = $DatasourceName
        LogDir             = $LogDir
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
      $triggerDoc = Get-PWDocumentsBySearch -FolderPath $WatchFolderPath -JustThisFolder -DocumentName $docName -PopulatePath
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
  }

  if ($RunOnce) {
    Write-Log "Done."
    exit 0
  }
  Start-Sleep -Seconds $PollIntervalSeconds
}