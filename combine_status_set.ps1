# combine_status_set.ps1
# Purpose: Combine all non-qc PDFs from a ProjectWise CADD/Sheets folder into a single StatusSet PDF.
# Output: _StatusSet.pdf (underscore prefix for top-of-list).
# Runs continuously, monitoring for sheet updates and replacing them in the StatusSet (incremental exchange).
#
# REQUIREMENTS: pwps_dab, qpdf, Logging.ps1
#
# RUN (continuous watch, all Sheets under root):
#   .\combine_status_set.ps1 -WatchUnderRoot @("Documents\AZDOT 2024","Documents\AZDOT") -SheetsPathFromProject "CADD\Sheets" -WriteBackToPW
#   From powershell.exe -File (launcher), use pipe-delimited roots instead (array binding is unreliable):
#   .\combine_status_set.ps1 -WatchUnderRootJoined "Documents\AZDOT 2024|Documents\AZDOT" -SheetsPathFromProject "CADD\Sheets" -WriteBackToPW
#
# RUN (one-shot): add -RunOnce
#
# Manifest JSON stores hashes and pwLastModified. Bump $StatusSetManifestSchemaVersion in this file when those
# semantics change so old manifests are ignored once (no manual delete). Or use -ForceRebuild to re-export all sheets.
#
[CmdletBinding(SupportsShouldProcess = $true)]
param(
  [Parameter(Mandatory = $false)]
  [string] $SheetsFolderPath,

  [Parameter(Mandatory = $false)]
  [string[]] $WatchUnderRoot,

  # Pipe-separated watch roots for subprocess launches (powershell.exe -File does not bind string[] reliably).
  [Parameter(Mandatory = $false)]
  [string] $WatchUnderRootJoined,

  [Parameter(Mandatory = $false)]
  [string] $SheetsPathFromProject = "CADD\Sheets",

  [Parameter(Mandatory = $false)]
  [string] $DatasourceName = "typsa-us-pw.bentley.com:typsa-us-pw-03",

  [Parameter(Mandatory = $false)]
  [string] $LocalRoot = "C:\PW_QC_LOCAL",

  [Parameter(Mandatory = $false)]
  [string] $LogDir = "",

  [Parameter(Mandatory = $false)]
  [switch] $WriteBackToPW,

  [Parameter(Mandatory = $false)]
  [switch] $ForceRebuild,

  [Parameter(Mandatory = $false)]
  # 0 = no delay between watch cycles (tight loop). Use e.g. 30 to pause between full scans.
  [int] $PollIntervalSeconds = 0,

  [Parameter(Mandatory = $false)]
  [switch] $RunOnce,

  [Parameter(Mandatory = $false)]
  [switch] $TestColumns,

  [Parameter(Mandatory = $false)]
  [string] $QpdfExe = "qpdf",

  [Parameter(Mandatory = $false)]
  [switch] $PromptForCredential
)

# pwps_dab requires MTA; re-launch with same params (same pattern as prepend_qc_on_trigger)
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -eq 'STA') {
  $passThrough = @('-DatasourceName', $DatasourceName)
  if ($WatchUnderRootJoined) {
    $passThrough += '-WatchUnderRootJoined'; $passThrough += $WatchUnderRootJoined
    $passThrough += '-SheetsPathFromProject'; $passThrough += $SheetsPathFromProject
  } elseif ($WatchUnderRoot) {
    $joined = @($WatchUnderRoot) -join '|'
    $passThrough += '-WatchUnderRootJoined'; $passThrough += $joined
    $passThrough += '-SheetsPathFromProject'; $passThrough += $SheetsPathFromProject
  } elseif ($SheetsFolderPath) {
    $passThrough += '-SheetsFolderPath'; $passThrough += $SheetsFolderPath
  }
  if ($LocalRoot -ne 'C:\PW_QC_LOCAL') { $passThrough += '-LocalRoot'; $passThrough += $LocalRoot }
  if ($LogDir) { $passThrough += '-LogDir'; $passThrough += $LogDir }
  if ($WriteBackToPW) { $passThrough += '-WriteBackToPW' }
  if ($ForceRebuild) { $passThrough += '-ForceRebuild' }
  $passThrough += '-PollIntervalSeconds'; $passThrough += $PollIntervalSeconds
  if ($RunOnce) { $passThrough += '-RunOnce' }
  if ($TestColumns) { $passThrough += '-TestColumns' }
  if ($QpdfExe -ne 'qpdf') { $passThrough += '-QpdfExe'; $passThrough += $QpdfExe }
  if ($PromptForCredential) { $passThrough += '-PromptForCredential' }
  & powershell.exe -MTA -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath @passThrough
  exit $LASTEXITCODE
}

$ErrorActionPreference = "Stop"
if (-not $LogDir) { $LogDir = Join-Path $LocalRoot "logs" }

$CredentialPath = 'C:\PW_QC_LOCAL\pw_cred.txt'
. "$PSScriptRoot\Logging.ps1"

trap {
  if (Get-Command Write-Log -ErrorAction SilentlyContinue) {
    Write-Log "Unhandled error: $_" -Severity ERROR
  }
  throw
}

function Ensure-Dir([string]$path) {
  if (-not (Test-Path $path)) {
    New-Item -ItemType Directory -Force -Path $path | Out-Null
  }
}

function Assert-Command([string]$exeName) {
  $cmd = Get-Command $exeName -ErrorAction SilentlyContinue
  if (-not $cmd) {
    throw "Required executable not found on PATH: '$exeName'. Install it or set -QpdfExe to the full path."
  }
}

# Prefer local tools\qpdf if present
if ($QpdfExe -eq "qpdf") {
  $localQpdf = Join-Path $PSScriptRoot "tools\qpdf\qpdf.exe"
  $localQpdfBin = Join-Path $PSScriptRoot "tools\qpdf\bin\qpdf.exe"
  if (Test-Path $localQpdf) { $QpdfExe = $localQpdf }
  elseif (Test-Path $localQpdfBin) { $QpdfExe = $localQpdfBin }
}

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
  $sec = ConvertTo-SecureString $pass -AsPlainText -Force
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

# Normalize path: strip leading Documents\ for PW API
function Get-PwFolderPath([string]$path) {
  $p = ($path -as [string]).Trim().TrimEnd('\')
  $p = $p -replace '^Documents\\', ''
  return $p
}

# Folder paths to try for document search. pwps_dab usually resolves folders by path *without* "Documents\"
# (Explorer shows Documents\...). Try stripped path first to avoid spurious "not found" warnings.
function Get-PwSheetsTryPaths([string]$sheetsFolderPath) {
  $trimmed = ($sheetsFolderPath -as [string]).Trim().TrimEnd('\')
  if (-not $trimmed) { return @() }
  $pw = Get-PwFolderPath $trimmed
  $withDoc = "Documents\$pw"
  $ordered = @()
  if ($pw) { $ordered += $pw }
  if ($withDoc -and $ordered -notcontains $withDoc) { $ordered += $withDoc }
  if ($trimmed -match '^Documents\\' -and $trimmed -ne $withDoc -and $ordered -notcontains $trimmed) {
    $ordered += $trimmed
  }
  return $ordered
}

# Normalize path for Get-NormalizedFolder (pw:\ or plain path)
function Get-NormalizedFolder([string]$path, [string]$defaultDs) {
  $raw = ($path -as [string]).Trim()
  if (-not $raw) { return $null }
  if ($raw -match '^pw:\\?' -or $raw -match '^pw:') {
    $p = $raw -replace '^pw:\\?', '' -replace '^pw:', ''
    $parts = $p -split '\\', 2
    if ($parts.Count -ge 2) {
      return @{ DatasourceName = $parts[0].Trim(); FolderPath = $parts[1].Trim().TrimEnd('\') }
    }
    return @{ DatasourceName = $defaultDs; FolderPath = $p.Trim().TrimEnd('\') }
  }
  return @{ DatasourceName = $defaultDs; FolderPath = $raw.TrimEnd('\') }
}

# Discover Sheets folders under root (like prepend_qc_on_trigger)
function Get-SheetsFoldersUnderRoot {
  param([string]$rootPath, [string]$sheetsSuffix, [string]$ds)
  $rootPathRaw = $rootPath.Trim().TrimEnd('\')
  $hadDocuments = $rootPathRaw -match '^Documents\\'
  $rootPath = $rootPathRaw -replace '^Documents\\', ''
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
        $name = $f.Name
        if (-not $name -and $f.PSObject.Properties['Name']) { $name = $f.Name }
        if ($name) { $childNames += $name }
      }
    }
  } catch {
    try {
      $children = Get-PWFoldersImmediateChildren -FolderPath $rootPath -ErrorAction Stop
      foreach ($c in @($children)) {
        $name = $c.Name
        if (-not $name -and $c.PSObject.Properties['Name']) { $name = $c.Name }
        if (-not $name -and $c.FolderPath) { $name = [System.IO.Path]::GetFileName($c.FolderPath.TrimEnd('\')) }
        if ($name) { $childNames += $name }
      }
    } catch { return @() }
  }
  $suffix = $sheetsSuffix.Trim().TrimStart('\')
  $list = @()
  foreach ($name in $childNames) {
    $folderPath = if ($hadDocuments) {
      "Documents\$rootPath\$name\$suffix"
    } else {
      $rootPath.TrimEnd('\') + '\' + $name + '\' + $suffix
    }
    $list += @{ DatasourceName = $ds; FolderPath = $folderPath }
  }
  return $list
}

# Sort-Object -Unique on @{...} hashtables can collapse every row to one (property binding fails). Use explicit paths.
function ConvertTo-FolderListEntry([object]$o) {
  $ds = $null
  $fp = $null
  if ($o -is [hashtable]) {
    $ds = [string]$o['DatasourceName']
    $fp = [string]$o['FolderPath']
  } else {
    $ds = [string]$o.DatasourceName
    $fp = [string]$o.FolderPath
  }
  [PSCustomObject]@{ DatasourceName = $ds; FolderPath = $fp }
}

# Sanitize path for manifest filename (replace \ and : with _)
function Get-ManifestPath([string]$folderPath, [string]$localRoot) {
  $safe = ($folderPath -replace '[\\/:]', '_').Trim()
  if (-not $safe) { $safe = "default" }
  return Join-Path $localRoot "status_set_manifest_$safe.json"
}

# Coerce PW/COM date fields to DateTime (avoid [DateTime]::TryParse - can fail binding ref on some PS/.NET combos)
function ConvertTo-DateTimeFromPwValue([object]$v) {
  if ($null -eq $v) { return $null }
  if ($v -is [DateTime]) { if ($v.Year -gt 1) { return $v }; return $null }
  if ($v -is [DateTimeOffset]) {
    $utc = $v.UtcDateTime
    if ($utc.Year -gt 1) { return $utc }
    return $null
  }
  if ($v -is [double] -or $v -is [float] -or $v -is [decimal]) {
    try {
      $ole = [DateTime]::FromOADate([double]$v)
      if ($ole.Year -gt 1) { return $ole }
    } catch { }
  }
  $str = $null
  try {
    $str = ($v.ToString()).Trim()
  } catch {
    return $null
  }
  if ([string]::IsNullOrWhiteSpace($str)) { return $null }
  try {
    $dt = [DateTime]::Parse($str, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind)
    if ($dt.Year -gt 1) { return $dt }
  } catch { }
  try {
    $dt = [DateTime]::Parse($str, [System.Globalization.CultureInfo]::CurrentCulture)
    if ($dt.Year -gt 1) { return $dt }
  } catch { }
  try {
    $dt = [System.Management.Automation.LanguagePrimitives]::ConvertTo($str, [datetime])
    if ($dt.Year -gt 1) { return $dt }
  } catch { }
  return $null
}

# Last-modified time from a PW document row (export skip + manifest).
# Prefer File Updated (FileUpdatedDate / FileUpdateDate) - matches ProjectWise document Properties "File Updated".
function Get-DocLastModified([psobject]$doc) {
  foreach ($prop in @("FileUpdatedDate", "FileUpdateDate", "DocumentUpdateDate", "VersionModifiedDate", "Version Modified Date")) {
    $v = $null
    if ($doc.PSObject.Properties[$prop]) { $v = $doc.PSObject.Properties[$prop].Value }
    if (-not $v -and $doc.$prop) { $v = $doc.$prop }
    if ($null -eq $v) { continue }
    $dt = ConvertTo-DateTimeFromPwValue $v
    if ($dt) { return $dt }
  }
  return $null
}

function Get-StatusSetDocumentForSheets {
  param(
    [string]$OutputName,
    [string]$SheetsFolderPath,
    [string[]]$DateCols
  )
  $paths = Get-PwSheetsTryPaths $SheetsFolderPath
  foreach ($p in $paths) {
    try {
      $d = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $p -JustThisFolder -DocumentName $OutputName -ColumnsToReturn $DateCols -ErrorAction SilentlyContinue | Select-Object -First 1
      if ($d) { return $d }
    } catch { }
  }
  foreach ($p in $paths) {
    try {
      $d = Get-PWDocumentsBySearch -FolderPath $p -JustThisFolder -DocumentName $OutputName -PopulatePath -ErrorAction SilentlyContinue | Select-Object -First 1
      if ($d) { return $d }
    } catch { }
  }
  return $null
}

function Parse-IsoDateTime([object]$s) {
  return ConvertTo-DateTimeFromPwValue $s
}

# Compute SHA256 hash of file
function Get-FileHashSha256([string]$path) {
  if (-not (Test-Path -LiteralPath $path)) { return $null }
  $bytes = [System.IO.File]::ReadAllBytes($path)
  $sha = [System.Security.Cryptography.SHA256]::Create()
  $hash = $sha.ComputeHash($bytes)
  $sha.Dispose()
  return [BitConverter]::ToString($hash) -replace '-', ''
}

# Get page count via qpdf
function Get-PdfPageCount([string]$path) {
  try {
    $info = & $QpdfExe --show-npages $path 2>&1
    $line = ($info | Select-Object -First 1) -as [string]
    if ($line -match '^\d+$') { return [int]$line }
  } catch { }
  return 1
}

# Combine PDFs via qpdf (batched: Windows command-line length limit; default batch 100)
function Merge-Pdfs([string[]]$pdfPaths, [string]$outPath) {
  if ($pdfPaths.Count -eq 0) {
    throw "No PDFs to merge."
  }
  if ($pdfPaths.Count -eq 1) {
    Copy-Item -LiteralPath $pdfPaths[0] -Destination $outPath -Force
    return
  }
  # ~32K command-line limit on Windows; tune down if qpdf fails with "filename too long"
  $maxPerBatch = 100
  $chunkDir = Split-Path -Parent $outPath
  if (-not $chunkDir) { $chunkDir = $env:TEMP }
  Ensure-Dir $chunkDir
  if ($pdfPaths.Count -le $maxPerBatch) {
    $allArgs = @('--empty', '--pages') + @($pdfPaths) + @('--', $outPath)
    & $QpdfExe @allArgs | Out-Null
    if (-not (Test-Path -LiteralPath $outPath)) {
      throw "qpdf failed to create output: $outPath"
    }
    return
  }
  $chunkTemp = @()
  try {
    for ($i = 0; $i -lt $pdfPaths.Count; $i += $maxPerBatch) {
      $end = [Math]::Min($i + $maxPerBatch - 1, $pdfPaths.Count - 1)
      $batch = @($pdfPaths[$i..$end])
      $chunkOut = Join-Path $chunkDir ("_pw_status_chunk_{0}.pdf" -f [guid]::NewGuid().ToString('N'))
      $chunkTemp += $chunkOut
      Merge-Pdfs -pdfPaths $batch -outPath $chunkOut
    }
    Merge-Pdfs -pdfPaths $chunkTemp -outPath $outPath
  } finally {
    foreach ($t in $chunkTemp) {
      Remove-Item -LiteralPath $t -Force -ErrorAction SilentlyContinue
    }
  }
  if (-not (Test-Path -LiteralPath $outPath)) {
    throw "qpdf failed to create output: $outPath"
  }
}

# Replace pages pageStart..pageEnd in combined PDF with new PDF content (Option B: incremental exchange)
function Replace-PdfPages([string]$combinedPath, [string]$newPdfPath, [int]$pageStart, [int]$pageEnd, [string]$outPath) {
  $totalPages = Get-PdfPageCount $combinedPath
  if ($pageStart -lt 1 -or $pageEnd -lt $pageStart) {
    throw "Replace-PdfPages: invalid range $pageStart-$pageEnd."
  }
  if ($pageEnd -gt $totalPages -or $pageStart -gt $totalPages) {
    throw "Replace-PdfPages: range $pageStart-$pageEnd is outside combined PDF ($totalPages page(s))."
  }
  $pageArgs = @()
  if ($pageStart -gt 1) {
    $pageArgs += $combinedPath
    $pageArgs += "1-$($pageStart - 1)"
  }
  $pageArgs += $newPdfPath
  if ($pageEnd -lt $totalPages) {
    $pageArgs += $combinedPath
    $pageArgs += "$($pageEnd + 1)-z"
  }
  $allArgs = @($combinedPath, '--pages') + $pageArgs + @('--', $outPath)
  & $QpdfExe @allArgs | Out-Null
  if (-not (Test-Path -LiteralPath $outPath)) {
    throw "qpdf failed to replace pages: $outPath"
  }
}

# Validate: need SheetsFolderPath or WatchUnderRoot
# Merge pipe-delimited roots (powershell.exe -File does not bind string[] reliably; use -WatchUnderRootJoined from launchers).
if ($WatchUnderRootJoined -and $WatchUnderRootJoined.Trim()) {
  $fromPipe = @($WatchUnderRootJoined -split '\|' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
  $WatchUnderRoot = @(@($WatchUnderRoot | Where-Object { $_ -and $_.Trim() }) + $fromPipe)
}
$watchRoots = @($WatchUnderRoot | Where-Object { $_ -and $_.Trim() })
$useWatchUnderRoot = ($watchRoots.Count -gt 0)
if (-not $SheetsFolderPath -and -not $useWatchUnderRoot) {
  throw "Specify -SheetsFolderPath or -WatchUnderRoot / -WatchUnderRootJoined (e.g. -WatchUnderRootJoined 'Documents\AZDOT 2024|Documents\AZDOT' -SheetsPathFromProject 'CADD\Sheets')."
}

if ($PollIntervalSeconds -gt 0) {
  Write-Log "Starting combine_status_set (watch mode, pause $PollIntervalSeconds s between full scans)."
} else {
  Write-Log 'Starting combine_status_set (watch mode, continuous loop - no pause between scans).'
}
Write-Log "LocalRoot: $LocalRoot"
Write-Log "WriteBackToPW: $WriteBackToPW"
Write-Log "ForceRebuild: $ForceRebuild"
Write-Log "RunOnce: $RunOnce"

Assert-Command $QpdfExe

if ($TestColumns) {
  $rootEntries = if ($useWatchUnderRoot) {
    @($watchRoots | ForEach-Object { Get-NormalizedFolder $_ $DatasourceName } | Where-Object { $_ })
  } else {
    @(@{ DatasourceName = $DatasourceName; FolderPath = $SheetsFolderPath })
  }
  if ($useWatchUnderRoot -and $rootEntries.Count -eq 0) { throw "Invalid path for TestColumns" }
  $folders = if ($useWatchUnderRoot) {
    $all = @()
    foreach ($rootEntry in $rootEntries) {
      Connect-PW $rootEntry.DatasourceName
      $all += @(Get-SheetsFoldersUnderRoot -rootPath $rootEntry.FolderPath -sheetsSuffix $SheetsPathFromProject -ds $rootEntry.DatasourceName)
      Close-PWConnection -ErrorAction SilentlyContinue
    }
    $all
  } else {
    @(@{ DatasourceName = $DatasourceName; FolderPath = $SheetsFolderPath })
  }
  if ($folders.Count -eq 0) { throw "No Sheets folders found" }
  $dateCols = @("Name", "DocumentID", "FileUpdatedDate", "FileUpdateDate", "DocumentUpdateDate", "VersionModifiedDate", "Version Modified Date")
  $allDocs = $null
  $docSearchPath = $null
  foreach ($entry in $folders) {
    Connect-PW $entry.DatasourceName
    $pwSheetsPath = Get-PwFolderPath $entry.FolderPath
    foreach ($tryPath in (Get-PwSheetsTryPaths $entry.FolderPath)) {
      try {
        $withCols = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $tryPath -JustThisFolder -ColumnsToReturn $dateCols -PopulatePath -ErrorAction Stop
        if ($withCols -and $withCols.Count -gt 0) {
          $pdfs = $withCols | Where-Object { $n = if ($_.Name) { $_.Name } elseif ($_.DocumentName) { $_.DocumentName } else { [System.IO.Path]::GetFileName($_.FullPath) }; $n -match '\.pdf$' -and $n -notmatch '-qc\.pdf$' }
          if ($pdfs.Count -gt 0) {
            $allDocs = @($pdfs)
            $docSearchPath = $tryPath
            break
          }
        }
      } catch { }
      try {
        $plain = Get-PWDocumentsBySearch -FolderPath $tryPath -JustThisFolder -PopulatePath -ErrorAction Stop
        if ($plain -and $plain.Count -gt 0) {
          $pdfs = $plain | Where-Object { $n = if ($_.Name) { $_.Name } elseif ($_.DocumentName) { $_.DocumentName } else { [System.IO.Path]::GetFileName($_.FullPath) }; $n -match '\.pdf$' -and $n -notmatch '-qc\.pdf$' }
          if ($pdfs.Count -gt 0) {
            $allDocs = @($plain)
            $docSearchPath = $tryPath
            $req = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $tryPath -JustThisFolder -DocumentName $pdfs[0].Name -ColumnsToReturn $dateCols -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($req) { $allDocs = @($req) + ($plain | Where-Object { $_.Name -ne $pdfs[0].Name }) }
            break
          }
        }
      } catch { }
    }
    if ($allDocs -and $allDocs.Count -gt 0) { break }
    try {
      foreach ($tp in (Get-PwSheetsTryPaths $entry.FolderPath)) {
        $folder = Get-PWFolders -FolderPath $tp -JustOne -ErrorAction SilentlyContinue
        if ($folder) {
          $view = $folder | Get-PWFolderView -ErrorAction SilentlyContinue
          if ($view -and $view.Documents) { $allDocs = @($view.Documents); $docSearchPath = $tp; break }
          elseif ($view -and $view.Children) {
            $allDocs = @($view.Children | Where-Object { $_.DocumentID })
            if ($allDocs.Count -gt 0) { $docSearchPath = $tp; break }
          }
          if ($allDocs -and $allDocs.Count -gt 0) {
            $pdfs = $allDocs | Where-Object { $n = if ($_.Name) { $_.Name } elseif ($_.DocumentName) { $_.DocumentName } else { [System.IO.Path]::GetFileName($_.FullPath) }; $n -match '\.pdf$' -and $n -notmatch '-qc\.pdf$' }
            if ($pdfs.Count -gt 0) {
              $docName = if ($pdfs[0].Name) { $pdfs[0].Name } elseif ($pdfs[0].DocumentName) { $pdfs[0].DocumentName } else { [System.IO.Path]::GetFileName($pdfs[0].FullPath) }
              $req = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $docSearchPath -JustThisFolder -DocumentName $docName -ColumnsToReturn $dateCols -ErrorAction SilentlyContinue | Select-Object -First 1
              if ($req) { $allDocs = @($req) }
              break
            }
          }
        }
      }
    } catch { }
  }
  if (-not $allDocs -or $allDocs.Count -eq 0) {
    Write-Log "No PDFs in any Sheets folder" -Severity WARNING
    exit 1
  }
  $d = ($allDocs | Where-Object { $n = if ($_.Name) { $_.Name } elseif ($_.DocumentName) { $_.DocumentName } else { [System.IO.Path]::GetFileName($_.FullPath) }; $n -match '\.pdf$' -and $n -notmatch '-qc\.pdf$' } | Select-Object -First 1)
  if (-not $d) { $d = $allDocs[0] }
  Write-Log "Document: $($d.Name)"
  Write-Log "Properties:"
  $d.PSObject.Properties | ForEach-Object { Write-Log "  $($_.Name) = $($_.Value)" }
  $foundCol = $null
  foreach ($p in @("DocumentUpdateDate", "VersionModifiedDate", "Version Modified Date", "FileUpdatedDate", "Date Last Saved", "File Updated")) {
    $v = $null
    if ($d.PSObject.Properties[$p]) { $v = $d.PSObject.Properties[$p].Value }
    if (-not $v) { $v = $d.$p }
    if ($v) { $foundCol = $p; Write-Log "  DATE COLUMN: $p = $v" }
  }
  if (-not $foundCol) {
    $d.PSObject.Properties | Where-Object { $_.Value -and $_.Value -is [DateTime] } | ForEach-Object { Write-Log "  POSSIBLE DATE: $($_.Name) = $($_.Value)" }
  }
  Close-PWConnection -ErrorAction SilentlyContinue
  exit 0
}

$exportDir = Join-Path $LocalRoot "export_status_set"
$tempWorkDir = Join-Path $env:TEMP "PW_QC_StatusSet"
Ensure-Dir $exportDir
Ensure-Dir $tempWorkDir

Import-Module pwps_dab -Force

# Bump when manifest fields or date rules change (older JSON files are treated as stale for one run).
$StatusSetManifestSchemaVersion = 2

while ($true) {
  # Build folder list (re-discover each poll when using WatchUnderRoot)
  $folderList = @()
  if ($useWatchUnderRoot) {
    $rootEntries = @($watchRoots | ForEach-Object { Get-NormalizedFolder $_ $DatasourceName } | Where-Object { $_ })
    if ($rootEntries.Count -eq 0) { throw "Invalid WatchUnderRoot: $($watchRoots -join ', ')" }
    try {
      foreach ($rootEntry in $rootEntries) {
        Connect-PW $rootEntry.DatasourceName
        Write-Log "  Watch root (list child projects, then append $($SheetsPathFromProject)): $($rootEntry.FolderPath) [datasource: $($rootEntry.DatasourceName)]"
        $discovered = @(Get-SheetsFoldersUnderRoot -rootPath $rootEntry.FolderPath -sheetsSuffix $SheetsPathFromProject -ds $rootEntry.DatasourceName)
        Write-Log "Discovered $($discovered.Count) Sheets folders under $($rootEntry.FolderPath)"
        # Force array concat (+= with a single hashtable can flatten incorrectly)
        $folderList = @($folderList) + @($discovered)
        Close-PWConnection -ErrorAction SilentlyContinue
      }
    } catch { Write-Log "WatchUnderRoot discovery failed: $_" -Severity WARNING }
    if ($folderList.Count -gt 0) {
      $folderList = @(
        $folderList | ForEach-Object { ConvertTo-FolderListEntry $_ } |
          Sort-Object -Property DatasourceName, FolderPath -Unique
      )
      Write-Log "Queued $($folderList.Count) unique sheet folder(s) to process this poll (sorted by path)."
    }
  } else {
    $folderList = @(ConvertTo-FolderListEntry @{ DatasourceName = $DatasourceName; FolderPath = $SheetsFolderPath })
  }

  if ($folderList.Count -eq 0 -and $useWatchUnderRoot) {
    Write-Log "No Sheets folders found this poll."
  } elseif ($folderList.Count -gt 0) {
    # Process each folder (defensive @() so a single hashtable is not enumerated as key/value pairs)
    foreach ($entry in @($folderList)) {
  $SheetsFolderPath = $entry.FolderPath
  $DatasourceName = $entry.DatasourceName
  Write-Log "--- Processing sheet folder: $SheetsFolderPath ---"

  try {
    Close-PWConnection -ErrorAction SilentlyContinue
    Connect-PW $DatasourceName
  } catch {
    Write-Log "Connect failed for $SheetsFolderPath : $_" -Severity WARNING
    continue
  }

  try {
$pwSheetsPath = Get-PwFolderPath $SheetsFolderPath
$outputName = "_StatusSet.pdf"

$sheetTryPaths = @(Get-PwSheetsTryPaths $SheetsFolderPath)
Write-Log "PW paths searched for documents (in order; first match wins):"
foreach ($tp in $sheetTryPaths) {
  Write-Log "  $tp"
}

# List documents in folder only (no subfolders); include File Updated columns (same as PW Properties)
$dateCols = @("Name", "DocumentID", "FileUpdatedDate", "FileUpdateDate", "DocumentUpdateDate", "VersionModifiedDate", "Version Modified Date")
$allDocs = @()
$docSearchPath = $pwSheetsPath
foreach ($tryPath in $sheetTryPaths) {
  try {
    # Primary: Get-PWDocumentsBySearchWithReturnColumns with VersionModifiedDate for last-saved comparison
    $withCols = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $tryPath -JustThisFolder -ColumnsToReturn $dateCols -PopulatePath -ErrorAction Stop
    if ($withCols -and $withCols.Count -gt 0) {
      $allDocs = @($withCols)
      $docSearchPath = $tryPath
      break
    }
  } catch { }
  try {
    $allDocs = @(Get-PWDocumentsBySearch -FolderPath $tryPath -JustThisFolder -PopulatePath -ErrorAction Stop)
    if ($allDocs.Count -gt 0) {
      $docSearchPath = $tryPath
      break
    }
  } catch { }
  # Fallback: Get-PWFolderView (may have limits)
  try {
    $folder = Get-PWFolders -FolderPath $tryPath -JustOne -ErrorAction SilentlyContinue
    if ($folder) {
      $view = $folder | Get-PWFolderView -ErrorAction SilentlyContinue
      if ($view -and $view.Documents) {
        $allDocs = @($view.Documents)
        $docSearchPath = $tryPath
        break
      }
      if ($view -and $view.Children) {
        $allDocs = @($view.Children | Where-Object { $_.DocumentID })
        if ($allDocs.Count -gt 0) { $docSearchPath = $tryPath; break }
      }
    }
  } catch { }
}

if ($allDocs.Count -gt 0) {
  Write-Log "Documents loaded from PW folder path: $docSearchPath ($($allDocs.Count) item(s) in folder listing)."
} else {
  Write-Log "No documents returned from any path listed above for this sheet folder." -Severity WARNING
}

# Base names (lowercase) of docs that have .dgn, .dwg, or DGN/DWG without extension in PW
$hasDgnOrDwg = @{}
$nonCadExt = '\.(pdf|xlsx|xls|doc|docx|txt|zip|jpg|jpeg|png|gif|bmp|tif|tiff|log|xml|json|csv)$'
foreach ($doc in $allDocs) {
  $name = $doc.Name
  if (-not $name -and $doc.PSObject.Properties['Name']) { $name = $doc.Name }
  if (-not $name -and $doc.DocumentName) { $name = $doc.DocumentName }
  if (-not $name) { $name = [System.IO.Path]::GetFileName($doc.FullPath) }
  if ($name -match '\.pdf$') { continue }
  if ($name -match $nonCadExt) { continue }
  if ($name -match '\.(dgn|dwg)$') {
    $base = ($name -replace '\.(dgn|dwg)$', '').ToLowerInvariant()
  } else {
    $base = $name.ToLowerInvariant()
  }
  $hasDgnOrDwg[$base] = $true
}

$pdfDocs = @()
foreach ($doc in $allDocs) {
  $name = $doc.Name
  if (-not $name -and $doc.PSObject.Properties['Name']) { $name = $doc.Name }
  if (-not $name -and $doc.DocumentName) { $name = $doc.DocumentName }
  if (-not $name) { $name = [System.IO.Path]::GetFileName($doc.FullPath) }
  if ($name -notmatch '\.pdf$') { continue }
  if ($name -match '-qc\.pdf$') { continue }
  $base = ($name -replace '\.pdf$', '').ToLowerInvariant()
  if (-not $hasDgnOrDwg[$base]) { continue }
  $pdfDocs += $doc
}

# Sort alphabetically by filename
function Get-DocName($doc) {
  if ($doc.Name) { return $doc.Name }
  if ($doc.DocumentName) { return $doc.DocumentName }
  return [System.IO.Path]::GetFileName($doc.FullPath)
}
$pdfDocs = $pdfDocs | Sort-Object { Get-DocName $_ }

if ($pdfDocs.Count -eq 0) {
  $pdfCount = ($allDocs | Where-Object { $n = if ($_.Name) { $_.Name } elseif ($_.DocumentName) { $_.DocumentName } else { [System.IO.Path]::GetFileName($_.FullPath) }; $n -match '\.pdf$' -and $n -notmatch '-qc\.pdf$' }).Count
  Write-Log "Skip (no matching sheet PDFs): $SheetsFolderPath - total docs in folder: $($allDocs.Count), non-qc PDFs: $pdfCount (need paired DGN/DWG base name)." -Severity WARNING
  continue
}

Write-Log "Found $($pdfDocs.Count) PDF(s) with DGN/DWG to combine (total docs in folder: $($allDocs.Count))."

# 3. Load manifest early for change detection; use persistent cache
$manifestPath = Get-ManifestPath -folderPath $pwSheetsPath -localRoot $LocalRoot
$manifest = $null
if (Test-Path $manifestPath) {
  try {
    $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
  } catch { }
}
$manifestSchemaStale = $false
if ($manifest) {
  $mv = $manifest.manifestSchemaVersion
  if ($null -eq $mv -or [int]$mv -lt $StatusSetManifestSchemaVersion) {
    $manifestSchemaStale = $true
    Write-Log "Manifest schema is v$(if ($null -ne $mv) { $mv } else { 'unset' }) (need v$StatusSetManifestSchemaVersion); ignoring file for skip logic so sheets re-export with current date rules." -Severity WARNING
    $manifest = $null
  }
}
if ($ForceRebuild) {
  Write-Log "ForceRebuild: ignoring manifest for export skip (re-export all sheet PDFs from PW)."
  $manifest = $null
}
Write-Log "Manifest file (hashes/cutoff): $manifestPath (exists: $(Test-Path -LiteralPath $manifestPath))"

$safeForCache = ($pwSheetsPath -replace '[\\/:]', '_').Trim()
if (-not $safeForCache) { $safeForCache = "default" }
$cacheSubDir = Join-Path (Join-Path $LocalRoot "status_set_cache") $safeForCache
Ensure-Dir $cacheSubDir
Write-Log "Local export cache directory: $cacheSubDir"

# StatusSet doc: try multiple folder spellings; extra columns help when PW omits default date fields
$statusSetDateCols = @($dateCols + @("FileUpdatedDate", "FileUpdateDate")) | Select-Object -Unique
$statusSetLastModified = $null
try {
  $statusSetDoc = Get-StatusSetDocumentForSheets -OutputName $outputName -SheetsFolderPath $SheetsFolderPath -DateCols $statusSetDateCols
  if ($statusSetDoc) {
    $statusSetLastModified = Get-DocLastModified $statusSetDoc
    Write-Log "Found existing $outputName in PW (searched same paths as document list); using its dates for export cutoff when available."
  } else {
    Write-Log "No $outputName document in PW at those paths yet (or no date on file)."
  }
} catch { }

$manifestPinnedCutoff = $null
if ($manifest -and $manifest.pinnedStatusSetLastModified) {
  $manifestPinnedCutoff = Parse-IsoDateTime ([string]$manifest.pinnedStatusSetLastModified)
}
$exportCutoff = $statusSetLastModified
if (-not $exportCutoff -and $manifestPinnedCutoff) {
  $exportCutoff = $manifestPinnedCutoff
}

if ($statusSetLastModified) {
  Write-Log "StatusSet last modified in PW: $statusSetLastModified"
} elseif ($exportCutoff -and $manifestPinnedCutoff) {
  Write-Log "StatusSet date not returned from PW; using manifest pinned cutoff for export skip: $exportCutoff"
} elseif (-not $exportCutoff) {
  $hasManifestSources = $manifest -and $manifest.sources -and (@($manifest.sources).Count -gt 0)
  if (-not $hasManifestSources) {
    Write-Log 'No _StatusSet.pdf in PW yet (or no date) and no prior manifest for this folder: first run will export all matching sheet PDFs to the local cache, then build _StatusSet.pdf.'
  } else {
    Write-Log 'No StatusSet cutoff from PW or manifest pinned date; per-file manifest/hash rules will skip unchanged sheets where possible.'
  }
}

# Export only PDFs newer than cutoff; use manifest per-file + cache when PW omits StatusSet dates
$localPdfPaths = @()
$sourcePwLastModByName = @{}
foreach ($doc in $pdfDocs) {
  $docName = Get-DocName $doc
  $cachedPath = Join-Path $cacheSubDir $docName
  $docLastMod = Get-DocLastModified $doc
  $sourcePwLastModByName[$docName] = $docLastMod
  $shouldExport = $true
  if ($exportCutoff -and $docLastMod) {
    if ($docLastMod -le $exportCutoff -and (Test-Path $cachedPath)) {
      $shouldExport = $false
    }
  }
  if ($shouldExport -and $manifest -and $manifest.sources) {
    $oldSrc = @($manifest.sources) | Where-Object { $_.name -eq $docName } | Select-Object -First 1
    if ($oldSrc) {
      if ($oldSrc.pwLastModified) {
        $stored = Parse-IsoDateTime ([string]$oldSrc.pwLastModified)
        if ($stored -and $docLastMod -and ($docLastMod -le $stored) -and (Test-Path $cachedPath)) {
          $shouldExport = $false
        }
      }
      # Cached bytes still match manifest: skip re-export on restart unless PW says the sheet is newer than our last snapshot.
      # (Previously we only skipped when docLastMod was null, so any PW date re-exported everything.)
      if ($shouldExport -and $oldSrc.hash -and (Test-Path $cachedPath)) {
        $h = Get-FileHashSha256 $cachedPath
        if ($h -eq $oldSrc.hash) {
          $stored = $null
          if ($oldSrc.pwLastModified) { $stored = Parse-IsoDateTime ([string]$oldSrc.pwLastModified) }
          if ($docLastMod -and $stored -and ($docLastMod -gt $stored)) {
            # Metadata newer than last build snapshot - refresh from PW
          } else {
            $shouldExport = $false
          }
        }
      }
    }
  }
  if ($shouldExport) {
    Write-Log "Exporting $docName..."
    Export-PWDocumentsSimple -InputDocuments $doc -TargetFolder $cacheSubDir | Out-Null
    Start-Sleep -Milliseconds 400
  } else {
    Write-Log "Using cache for $docName (unchanged vs cutoff or manifest)."
  }
  $localPath = $cachedPath
  if (-not (Test-Path $localPath) -and $doc.CopiedOutLocalFileName -and (Test-Path $doc.CopiedOutLocalFileName)) {
    Copy-Item -LiteralPath $doc.CopiedOutLocalFileName -Destination $localPath -Force -ErrorAction SilentlyContinue
  }
  if (-not (Test-Path $localPath)) {
    $found = Get-ChildItem $cacheSubDir -File | Where-Object { $_.Name -ieq $docName } | Select-Object -First 1
    if ($found) {
      $localPath = $found.FullName
      if ($localPath -ne $cachedPath) { Copy-Item -LiteralPath $localPath -Destination $cachedPath -Force -ErrorAction SilentlyContinue; $localPath = $cachedPath }
    }
  }
  if (Test-Path $localPath) {
    $localPdfPaths += $localPath
  } else {
    Write-Log "Export failed for $docName" -Severity WARNING
  }
}

if ($localPdfPaths.Count -eq 0) {
  throw "No PDFs were exported successfully."
}

# Total pages if we merged current exports (incremental replace must match existing _StatusSet.pdf or qpdf ranges go out of range)
$expectedTotalPages = 0
foreach ($lp in $localPdfPaths) {
  $expectedTotalPages += Get-PdfPageCount $lp
}

# 4. Check manifest for changes; determine full rebuild vs incremental exchange
$outPdf = Join-Path $tempWorkDir $outputName
$needsFullRebuild = $ForceRebuild
$changedIndices = @()

if (-not $needsFullRebuild -and (Test-Path $manifestPath) -and -not $manifestSchemaStale) {
  try {
    $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $manifestSources = @($manifest.sources)
    if ($manifestSources.Count -ne $localPdfPaths.Count) {
      $needsFullRebuild = $true
      Write-Log "Manifest source count mismatch (add/remove); full rebuild."
    } else {
      $pageEnd = 0
      for ($i = 0; $i -lt $localPdfPaths.Count; $i++) {
        $localPath = $localPdfPaths[$i]
        $name = [System.IO.Path]::GetFileName($localPath)
        $hash = Get-FileHashSha256 $localPath
        $pageCount = Get-PdfPageCount $localPath
        $pageStart = $pageEnd + 1
        $pageEnd = $pageStart + $pageCount - 1

        $old = $manifestSources | Where-Object { $_.name -eq $name } | Select-Object -First 1
        if (-not $old -or $old.hash -ne $hash) {
          $changedIndices += @{ Index = $i; Name = $name; PageStart = $pageStart; PageEnd = $pageEnd; LocalPath = $localPath }
          Write-Log "Source changed: $name"
        }
      }
    }
  } catch {
    Write-Log "Manifest read failed; full rebuild: $_" -Severity WARNING
    $needsFullRebuild = $true
  }
} else {
  $needsFullRebuild = $true
  if ($manifestSchemaStale) {
    Write-Log "Manifest schema out of date; full rebuild."
  } elseif (-not (Test-Path $manifestPath)) {
    Write-Log "No manifest found; full rebuild."
  } elseif ($ForceRebuild) {
    Write-Log "ForceRebuild; full rebuild."
  }
}

# 5. Build or update combined PDF (Option B: incremental exchange when possible)
if ($needsFullRebuild) {
  Write-Log "Full rebuild: combining $($localPdfPaths.Count) PDF(s) into $outputName..."
  Merge-Pdfs -pdfPaths $localPdfPaths -outPath $outPdf
  Write-Log "Combined PDF created: $outPdf"
} elseif ($changedIndices.Count -gt 0 -and (Test-Path $outPdf)) {
  $actualPages = Get-PdfPageCount $outPdf
  if ($actualPages -ne $expectedTotalPages) {
    Write-Log "Existing combined PDF has $actualPages page(s) on disk; merging all $($localPdfPaths.Count) current source PDF(s) would yield $expectedTotalPages page(s). Full rebuild (stale or partial StatusSet vs folder - not limited to changed sources)." -Severity WARNING
    Merge-Pdfs -pdfPaths $localPdfPaths -outPath $outPdf
    Write-Log "Combined PDF created: $outPdf"
    $needsFullRebuild = $true
  } else {
    Write-Log "Incremental exchange: replacing $($changedIndices.Count) changed PDF(s) in place."
    # Process from highest page range first so earlier ranges stay valid
    $sorted = $changedIndices | Sort-Object { $_.PageStart } -Descending
    $workPdf = $outPdf
    foreach ($ch in $sorted) {
      $tempOut = Join-Path $tempWorkDir "status_set_replace_$([guid]::NewGuid().ToString('N').Substring(0,8)).pdf"
      Replace-PdfPages -combinedPath $workPdf -newPdfPath $ch.LocalPath -pageStart $ch.PageStart -pageEnd $ch.PageEnd -outPath $tempOut
      if ($workPdf -ne $outPdf) { Remove-Item -LiteralPath $workPdf -Force -ErrorAction SilentlyContinue }
      $workPdf = $tempOut
    }
    if ($workPdf -ne $outPdf) {
      Move-Item -LiteralPath $workPdf -Destination $outPdf -Force
    }
    Write-Log "Combined PDF updated: $outPdf"
  }
} else {
  Write-Log "No changes detected; using existing combined PDF."
  if (-not (Test-Path $outPdf)) {
    Write-Log "Existing output missing; full rebuild." -Severity WARNING
    Merge-Pdfs -pdfPaths $localPdfPaths -outPath $outPdf
  }
}

# 6. Update manifest (after full rebuild or incremental exchange)
if ($needsFullRebuild -or $changedIndices.Count -gt 0) {
  $sources = @()
  $pageEnd = 0
  foreach ($lp in $localPdfPaths) {
    $name = [System.IO.Path]::GetFileName($lp)
    $hash = Get-FileHashSha256 $lp
    $pageCount = Get-PdfPageCount $lp
    $pageStart = $pageEnd + 1
    $pageEnd = $pageStart + $pageCount - 1
    $row = @{ name = $name; hash = $hash; pageStart = $pageStart; pageEnd = $pageEnd }
    $dlm = $sourcePwLastModByName[$name]
    if ($dlm) {
      $row.pwLastModified = $dlm.ToUniversalTime().ToString('o')
    }
    $sources += $row
  }
  $pinned = $null
  if ($statusSetLastModified) {
    $pinned = $statusSetLastModified.ToUniversalTime().ToString('o')
  } elseif ($manifest -and $manifest.pinnedStatusSetLastModified) {
    $pinned = [string]$manifest.pinnedStatusSetLastModified
  }
  $manifestObj = @{
    manifestSchemaVersion       = $StatusSetManifestSchemaVersion
    folderPath                  = $pwSheetsPath
    pinnedStatusSetLastModified = $pinned
    sources                     = $sources
  }
  Ensure-Dir (Split-Path $manifestPath -Parent)
  $manifestObj | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
  Write-Log "Manifest updated: $manifestPath"
}

# 7. Write back to PW if requested (when we made changes, or when StatusSet is missing in PW)
$existingDoc = Get-PWDocumentsBySearch -FolderPath $docSearchPath -JustThisFolder -DocumentName $outputName -PopulatePath -ErrorAction SilentlyContinue
$shouldWriteBack = $WriteBackToPW -and (Test-Path $outPdf) -and (($needsFullRebuild -or $changedIndices.Count -gt 0) -or -not $existingDoc)
if ($shouldWriteBack) {
  if ($PSCmdlet.ShouldProcess("$docSearchPath\$outputName", "Create or update StatusSet document")) {
    if ($existingDoc) {
      Write-Log "Updating existing StatusSet document in PW..."
      $updateCmd = Get-Command Update-PWDocumentFile
      $fileParamName = $updateCmd.Parameters.Keys | Where-Object {
        $_ -match '^(LocalPath|SourcePath|FilePath|Path|SourceFile|File)$' -and $_ -notin @('Verbose', 'Debug', 'ErrorAction', 'WarningAction', 'InformationAction', 'ErrorVariable', 'WarningVariable', 'OutVariable', 'OutBuffer', 'PipelineVariable')
      } | Select-Object -First 1
      if (-not $fileParamName) {
        $fileParamName = $updateCmd.Parameters.Keys | Where-Object { $_ -match 'path|file' -and $_ -ne 'InputDocument' } | Select-Object -First 1
      }
      if (-not $fileParamName) { throw "Update-PWDocumentFile: could not find file path parameter." }
      $updateArgs = @{ InputDocument = $existingDoc; $fileParamName = $outPdf }
      Update-PWDocumentFile @updateArgs | Out-Null
      Write-Log "Updated StatusSet document in PW."
    } else {
      Write-Log "Creating new StatusSet document in PW..."
      New-PWDocument -FolderPath $docSearchPath -FilePath $outPdf -DocumentName $outputName | Out-Null
      Write-Log "Created StatusSet document in PW."
    }
    try {
      if (Test-Path $manifestPath) {
        $sd2 = Get-StatusSetDocumentForSheets -OutputName $outputName -SheetsFolderPath $SheetsFolderPath -DateCols $statusSetDateCols
        if ($sd2) {
          $dt2 = Get-DocLastModified $sd2
          if ($dt2) {
            $m2 = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
            $m2.pinnedStatusSetLastModified = $dt2.ToUniversalTime().ToString('o')
            $m2 | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
            Write-Log "Manifest pinned StatusSet date refreshed from PW after write-back."
          }
        }
      }
    } catch { }
  }
  }

  Write-Log "Done. Output: $outPdf"
  } catch {
    Write-Log "Failed for $SheetsFolderPath : $_" -Severity ERROR
  }
  Close-PWConnection -ErrorAction SilentlyContinue
  }
  }

  if ($RunOnce) {
    Write-Log "Done (RunOnce)."
    exit 0
  }
  if ($PollIntervalSeconds -gt 0) {
    Write-Log "Sleeping $PollIntervalSeconds s until next poll..."
    Start-Sleep -Seconds $PollIntervalSeconds
  }
}
