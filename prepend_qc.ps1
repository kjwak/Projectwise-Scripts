# prepend_qc.ps1
# Purpose: prepend incoming PDF into a QC history PDF in the same PW folder (create/update (filename)-qc.pdf).
# Notes:
# - Uses Bentley IMS login via Open-PWConnection.
# - Avoids wildcard searches (your environment doesn't handle them reliably).
# - History document is (incoming filename base)-qc.pdf, saved in the same PW folder as the incoming file.
# - If the history doc doesn't exist, it creates it from the incoming PDF (base case).
#   - If it exists, it exports both and prepends (with overlay layers when available), then updates the history doc in PW.
# - PDF tools (qpdf / qc_overlay_prepend) need local temp files: export from PW → process → upload merged PDF back.
#   "Old" vs "new" for the overlay always comes from PW content (exported history + exported incoming).
#   -OverlaySheetWorkDir:$true (default): LocalRoot\work\<historyBase>\ splits each history page to <stem>-NN.pdf, MANIFEST,
#   and overlay artifacts; Old is extracted from page 1 of the full exported history (same as test; splits are audit/debug).
#   -OverlayOldFromHistoryOnly:$true with -OverlaySheetWorkDir:$false: qpdf page 1 -> TEMP --current-master.
#   -OverlayOldFromHistoryOnly:$false: persistent work\<sheet>_current_master.pdf.
#
# REQUIREMENTS:
#   - pwps_dab module installed
#   - qpdf installed and on PATH (or set $QpdfExe to full path) - used when overlay exe not found, and to seed LocalRoot\work\<historyBase>_current_master.pdf on first overlay run (same role as test\run_f0548dv206_qc_two_step.ps1)
#   - dist\qc_overlay_prepend\qc_overlay_prepend.exe (onedir build) or dist\qc_overlay_prepend.exe (onefile) - optional layered overlay
#
# RUN (standalone example):
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\path\prepend_qc.ps1" -IncomingDocName "input1.pdf"
#   (Script auto re-launches with -MTA when needed for pwps_dab.)
#
# Optional:
#   -WhatIf (dry-run, no writes)
#
[CmdletBinding(SupportsShouldProcess=$true)]
param(
  [Parameter(Mandatory=$false)]
  [string] $DatasourceName = "typsa-us-pw.bentley.com:typsa-us-pw-03",

  [Parameter(Mandatory=$false)]
  [string] $IncomingFolderPath = "AZDOT 2024\AZFWY1704-FD02-SR202 - I-10 to SR101\CADD\Working\TYPSA\Drainage\JFlint\Prepend Test\incoming",

  [Parameter(Mandatory=$false)]
  [string] $IncomingDocName = "input1.pdf",

  [Parameter(Mandatory=$false)]
  [string] $HistoryDocName = "",  # default: (IncomingDocName base)-qc.pdf

  [Parameter(Mandatory=$false)]
  [string] $LocalRoot = "C:\PW_QC_LOCAL",

  [Parameter(Mandatory=$false)]
  [string] $QpdfExe = "qpdf",

  [Parameter(Mandatory=$false)]
  [string] $QcOverlayExe = "",  # default: first existing of dist\qc_overlay_prepend\qc_overlay_prepend.exe, dist\qc_overlay_prepend.exe

  # Optional: when -OverlayOldFromHistoryOnly:$false, path for persistent current-master (default: LocalRoot\work\<sheet>_current_master.pdf).
  [Parameter(Mandatory=$false)]
  [string] $OverlayCurrentMasterPath = "",

  # $true: no persistent work\ current-master file; each run qpdf slices page 1 of exported *-qc.pdf to TEMP and passes --current-master (matches PW source; avoids Python page-1 extract). $false: persistent current-master under work\.
  [Parameter(Mandatory=$false)]
  $OverlayOldFromHistoryOnly = $false,

  # $true (default): work\<historyBase>\ per-sheet splits + MANIFEST (--sheet-work-dir on exe). $false: no split folder; use current-master paths below.
  [Parameter(Mandatory=$false)]
  $OverlaySheetWorkDir = $true,

  [Parameter(Mandatory=$false)]
  [switch] $NoOverlayLayers,  # if set, use qpdf only (no layered overlay)

  [Parameter(Mandatory=$false)]
  [switch] $PromptForCredential,

  [Parameter(Mandatory=$false)]
  [string] $LogDir = ""
)

if (-not $HistoryDocName) {
  $HistoryDocName = [System.IO.Path]::GetFileNameWithoutExtension($IncomingDocName) + "-qc.pdf"
}

# Prefer local tools\qpdf if present (e.g. unpacked qpdf MSVC64 package)
if ($QpdfExe -eq "qpdf") {
  $localQpdf = Join-Path $PSScriptRoot "tools\qpdf\qpdf.exe"
  $localQpdfBin = Join-Path $PSScriptRoot "tools\qpdf\bin\qpdf.exe"
  if (Test-Path $localQpdf) { $QpdfExe = $localQpdf }
  elseif (Test-Path $localQpdfBin) { $QpdfExe = $localQpdfBin }
}

. "$PSScriptRoot\StaMtaRelaunch.ps1"

if ($PSBoundParameters.ContainsKey('OverlayOldFromHistoryOnly')) {
  $OverlayOldFromHistoryOnly = ConvertTo-BoolLoose $PSBoundParameters['OverlayOldFromHistoryOnly']
}
if ($PSBoundParameters.ContainsKey('OverlaySheetWorkDir')) {
  $OverlaySheetWorkDir = ConvertTo-BoolLoose $PSBoundParameters['OverlaySheetWorkDir']
} else {
  $OverlaySheetWorkDir = $true
}

# pwps_dab requires MTA; Cursor/VS Code terminals often use STA. Re-launch in MTA to avoid ThreadOptions error.
# Do not splat full $PSBoundParameters to powershell.exe (common params -Verbose/-WhatIf/etc. break child -File; bool/switch types break native argv).
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -eq 'STA') {
  $scriptPath = $PSCommandPath
  if (-not $scriptPath) { $scriptPath = $MyInvocation.MyCommand.Path }
  if (-not $scriptPath) { $scriptPath = Join-Path $PSScriptRoot 'prepend_qc.ps1' }
  if (-not (Test-Path -LiteralPath $scriptPath)) {
    throw "MTA relaunch: could not resolve script path (PSCommandPath / MyInvocation). Tried: $scriptPath"
  }
  $paramNames = @(
    'DatasourceName', 'IncomingFolderPath', 'IncomingDocName', 'HistoryDocName', 'LocalRoot', 'QpdfExe',
    'QcOverlayExe', 'OverlayCurrentMasterPath', 'OverlayOldFromHistoryOnly', 'OverlaySheetWorkDir', 'NoOverlayLayers',
    'PromptForCredential', 'LogDir'
  )
  $bp = @{}
  foreach ($n in $paramNames) {
    if ($PSBoundParameters.ContainsKey($n)) { $bp[$n] = $PSBoundParameters[$n] }
  }
  $exeArgs = Build-PowerShellExeFileArgs -ScriptPath $scriptPath -BoundParameters $bp
  & powershell.exe @exeArgs
  exit $LASTEXITCODE
}

$ErrorActionPreference = "Stop"
if (-not $LogDir) { $LogDir = Join-Path $LocalRoot "logs" }

# ProjectWise credentials: C:\PW_QC_LOCAL\pw_cred.txt
# Format: username=domain\user and password=... on separate lines.
$CredentialPath = 'C:\PW_QC_LOCAL\pw_cred.txt'

. "$PSScriptRoot\Logging.ps1"
. "$PSScriptRoot\Resolve-OverlayExe.ps1"

# Default overlay exe: prefer onedir from build_overlay_exe.ps1, else onefile
if (-not $QcOverlayExe) {
  $found = Select-ExistingOverlayExePath $PSScriptRoot
  if ($found) {
    $QcOverlayExe = $found
  } else {
    $QcOverlayExe = (Join-Path $PSScriptRoot "dist\qc_overlay_prepend\qc_overlay_prepend.exe")
  }
}

trap {
  if (Get-Command Write-Log -ErrorAction SilentlyContinue) {
    Write-Log "Unhandled error: $_" -Severity ERROR
  }
  throw
}

function Initialize-Directory([string]$path) {
  if (-not (Test-Path $path)) {
    New-Item -ItemType Directory -Force -Path $path | Out-Null
  }
}

function Test-CommandExists([string]$exeName) {
  $cmd = Get-Command $exeName -ErrorAction SilentlyContinue
  if (-not $cmd) {
    throw "Required executable not found on PATH: '$exeName'. Install it or set -QpdfExe to the full path."
  }
}

function Test-ExecutableAvailable([string]$path) {
  if (-not $path) { return $false }
  if ($path -match '\\' -or $path -match '^[A-Za-z]:') {
    return (Test-Path -LiteralPath $path)
  }
  return [bool](Get-Command $path -ErrorAction SilentlyContinue)
}

# Load PW credential from C:\PW_QC_LOCAL\pw_cred.txt
# Format: username=domain\user and password=... on separate lines.
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

function Connect-PW([string]$dsName) {
  $cred = Get-PwCredential
  try {
    Open-PWConnection -DatasourceName $dsName -UserName $cred.UserName -Password $cred.Password | Out-Null
  } catch {
    if ($_.Exception.Message -match 'connection is already open') {
      Close-PWConnection -ErrorAction SilentlyContinue
      Open-PWConnection -DatasourceName $dsName -UserName $cred.UserName -Password $cred.Password | Out-Null
    } else { throw }
  }
}

# Retry a scriptblock on transient "Access denied" / "in use" (e.g. antivirus, PW). Throws after last attempt.
function Invoke-RetryOnAccessDenied([scriptblock]$sb, [int]$maxAttempts = 4) {
  $attempt = 0
  while ($true) {
    try {
      return & $sb
    } catch {
      $attempt++
      $msg = $_.Exception.Message
      $isAccess = ($msg -match 'Access to the path is denied' -or $msg -match 'being used by another process' -or $msg -match 'The process cannot access the file')
      if (-not $isAccess -or $attempt -ge $maxAttempts) { throw }
      Start-Sleep -Milliseconds (300 * $attempt)
    }
  }
}

# Delete file without failing when antivirus blocks Remove-Item (tries .NET delete, retries on access denied, then logs and continues)
function Remove-ItemWithRetry([string]$path) {
  if (-not $path -or -not (Test-Path -LiteralPath $path)) { return }
  try {
    Invoke-RetryOnAccessDenied {
      Remove-Item -LiteralPath $path -Force -ErrorAction Stop
    }
  } catch {
    try {
      Invoke-RetryOnAccessDenied { [System.IO.File]::Delete($path) }
    } catch {
      Write-Log "Could not delete file (antivirus may block): $path" -Severity WARNING
    }
  }
}

function Invoke-PdfPrependMerge([string]$newPdf, [string]$historyPdf, [string]$outPdf) {
  if (-not (Test-Path $historyPdf)) {
    Copy-Item -Path $newPdf -Destination $outPdf -Force
    return
  }

  # qpdf: merge with new first (prepend), then history. Order of --pages = output order.
  & $QpdfExe --empty --pages $newPdf $historyPdf -- $outPdf | Out-Null

  if (-not (Test-Path $outPdf)) {
    throw "qpdf failed to create output: $outPdf"
  }
}

function Convert-OverlayExeOutputLine($obj) {
  if ($null -eq $obj) { return '' }
  return [string]$obj
}

function Invoke-PdfPrependOverlay(
  [string]$incomingPdf,
  [string]$historyPdf,
  [string]$outPdf,
  [string]$overlayExe,
  [string]$currentMasterPath = "",
  [string]$sheetWorkDir = ""
) {
  if (-not (Test-Path $historyPdf)) {
    Copy-Item -Path $incomingPdf -Destination $outPdf -Force
    return
  }

  $exeToRun = Resolve-OverlayExePath $overlayExe

  # qc_overlay_prepend: page1 of history = Old/red, incoming = New/green + Current/black, prepended to history
  # With $ErrorActionPreference = Stop, stderr from the exe (Python tracebacks) becomes terminating on the first line; use Continue for this call only.
  $prevEap = $ErrorActionPreference
  $overlayExit = $null
  $overlayOut = $null
  try {
    $ErrorActionPreference = 'Continue'
    $exeArgs = @($incomingPdf, $historyPdf, '-o', $outPdf)
    if ($currentMasterPath -and (Test-Path -LiteralPath $currentMasterPath)) {
      $exeArgs += @('--current-master', $currentMasterPath)
    }
    if ($sheetWorkDir -and $sheetWorkDir.Trim()) {
      $exeArgs += @('--sheet-work-dir', $sheetWorkDir)
    }
    $overlayOut = & $exeToRun @exeArgs 2>&1
    $overlayExit = $LASTEXITCODE
  } catch {
    $m = $_.Exception.Message
    if ($m -match 'corrupted and unreadable') {
      throw "Windows could not load qc_overlay_prepend.exe ($exeToRun). Often: (1) exe on a slow/network/cloud path or blocked by AV - copy to a local folder (e.g. C:\Tools\) and set -QcOverlayExe; (2) incomplete or wrong-architecture copy - redeploy a Windows-built exe; (3) PyInstaller onedir - deploy the full folder including _internal. Original error: $m"
    }
    throw
  } finally {
    $ErrorActionPreference = $prevEap
  }
  $sev = if ($null -ne $overlayExit -and $overlayExit -ne 0) { 'ERROR' } else { 'INFO' }
  foreach ($line in $overlayOut) {
    $s = Convert-OverlayExeOutputLine $line
    if ($s) { Write-Log $s -Severity $sev }
  }
  if ($null -ne $overlayExit -and $overlayExit -ne 0) {
    throw "qc_overlay_prepend exited with code $overlayExit"
  }
  if (-not (Test-Path $outPdf)) {
    throw "qc_overlay_prepend failed to create output: $outPdf"
  }
}

Write-Log "Starting prepend test..."
Write-Log "Datasource: $DatasourceName"
Write-Log "Folder (incoming + history): $IncomingFolderPath"
Write-Log "Incoming doc:    $IncomingDocName"
Write-Log "History doc:     $HistoryDocName"
Write-Log "LocalRoot:       $LocalRoot"
Write-Log "qpdf:            $QpdfExe"

# Check for overlay exe (creates layered PDF when available)
$haveOverlay = $false
if (-not $NoOverlayLayers -and $QcOverlayExe -and (Test-Path -LiteralPath $QcOverlayExe)) {
  $haveOverlay = $true
  Write-Log "Overlay exe:     $QcOverlayExe (will create layered Old/New/Current)"
} else {
  Write-Log "Overlay exe:     not used (NoOverlayLayers=$NoOverlayLayers or exe not found)"
}

# Local folders (exports); copies and merge use %TEMP% to avoid AV/lock on LocalRoot
$exportDir   = Join-Path $LocalRoot "export_test"
$workDir     = Join-Path $LocalRoot "work"
$tempWorkDir = Join-Path $env:TEMP "PW_QC"
Initialize-Directory $exportDir
Initialize-Directory $workDir
Initialize-Directory $tempWorkDir

# Validate qpdf presence (merge fallback + seeding current-master). Prefer Test-Path for full paths — Get-Command is unreliable for tools\qpdf\... on some hosts.
$haveQpdf = Test-ExecutableAvailable $QpdfExe
if (-not $haveQpdf) {
  Write-Log "qpdf not available at '$QpdfExe' (merge fallback and current-master seed disabled). Install qpdf or add tools\qpdf under the script folder." -Severity WARNING
}

# Load module + connect IMS
Import-Module pwps_dab -Force
#Write-Log "Connecting via IMS..."
#Open-PWConnection -DatasourceName $DatasourceName -BentleyIMS | Out-Null
Write-Log "Connecting with stored credential..."
Connect-PW $DatasourceName


# Resolve incoming doc (exact name search)
Write-Log "Resolving incoming document..."
$incomingDoc = Get-PWDocumentsBySearch -FolderPath $IncomingFolderPath -JustThisFolder -DocumentName $IncomingDocName -PopulatePath
if (-not $incomingDoc) {
  throw "Incoming document not found: '$IncomingDocName' in '$IncomingFolderPath'"
}
Write-Log ("Incoming resolved: DocumentID={0}, FullPath={1}" -f $incomingDoc.DocumentID, $incomingDoc.FullPath)

# Export incoming doc to local
Write-Log "Exporting incoming document to local..."
Export-PWDocumentsSimple -InputDocuments $incomingDoc -TargetFolder $exportDir | Out-Null
Start-Sleep -Milliseconds 600   # let PW/antivirus release file handle (Error 100 can leave file locked briefly)

$localIncoming = Join-Path $exportDir $IncomingDocName
if (-not (Test-Path $localIncoming)) {
  # Some exports may rename; fall back to CopiedOutLocalFileName if set
  if ($incomingDoc.CopiedOutLocalFileName -and (Test-Path $incomingDoc.CopiedOutLocalFileName)) {
    $localIncoming = $incomingDoc.CopiedOutLocalFileName
  } else {
    $found = Get-ChildItem $exportDir -File | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($found) { $localIncoming = $found.FullName }
  }
}

if (-not (Test-Path $localIncoming)) {
  throw "Incoming export failed; no local file found in $exportDir"
}
# Copy to %TEMP%; on "Access denied" retry with a different filename (suffix)
$incomingBase = "incoming_" + [System.IO.Path]::GetFileNameWithoutExtension($IncomingDocName) + "_" + (Get-Date -Format "yyyyMMdd_HHmmss")
$localIncomingTmp = $null
foreach ($suffix in @('', '_2', '_3', '_4', '_5')) {
  $dest = Join-Path $tempWorkDir ($incomingBase + $suffix + ".pdf")
  try {
    Invoke-RetryOnAccessDenied { Copy-Item -LiteralPath $localIncoming -Destination $dest -Force }
    $localIncomingTmp = $dest
    break
  } catch {
    if ($_.Exception.Message -match 'Access to the path .* is denied' -and $suffix -ne '_5') { continue }
    throw
  }
}
Remove-ItemWithRetry $localIncoming   # best effort; PW may still have lock
$localIncoming = $localIncomingTmp
$fi = Get-Item -LiteralPath $localIncoming
if ($fi.Length -eq 0) {
  Remove-ItemWithRetry $localIncoming
  throw "Incoming export failed (zero file size - unmanaged copy?). Document may have no file in PW."
}
$header = [System.IO.File]::ReadAllBytes($localIncoming) | Select-Object -First 4
$isPdf = ($header.Count -ge 4) -and ([char]$header[0] -eq '%' -and [char]$header[1] -eq 'P' -and [char]$header[2] -eq 'D' -and [char]$header[3] -eq 'F')
if (-not $isPdf) {
  Remove-ItemWithRetry $localIncoming
  throw "Incoming export failed (file not a valid PDF - Error 100 unmanaged copy?). Path: $localIncoming"
}
Write-Log "Local incoming file: $localIncoming"

# Resolve history doc (exact name search)
Write-Log "Checking for existing history document..."
$historyDoc = Get-PWDocumentsBySearch -FolderPath $IncomingFolderPath -JustThisFolder -DocumentName $HistoryDocName -PopulatePath

$localHistory = Join-Path $tempWorkDir $HistoryDocName   # history export will go to TEMP
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($HistoryDocName)
$localMerged = Join-Path $tempWorkDir ("${baseName}_MERGED_$stamp.pdf")

if (-not $historyDoc) {
  Write-Log "History document does not exist yet."

  if ($PSCmdlet.ShouldProcess("$IncomingFolderPath\$HistoryDocName", "Create history document from incoming")) {
    Write-Log "Creating $HistoryDocName in same folder as incoming (base case = incoming becomes history)..."
    New-PWDocument -FolderPath $IncomingFolderPath -FilePath $localIncoming -DocumentName $HistoryDocName | Out-Null
    Write-Log "Created history document."
    if (Test-Path $localIncoming) { Remove-ItemWithRetry $localIncoming }
  } else {
    Write-Log "WhatIf: would create history document."
  }

  Write-Log "Done."
  Close-PWConnection -ErrorAction SilentlyContinue
  exit 0
}

Write-Log ("History resolved: DocumentID={0}, FullPath={1}" -f $historyDoc.DocumentID, $historyDoc.FullPath)

# Export history to %TEMP% (avoids AV/lock on LocalRoot)
Write-Log "Exporting existing history document from PW to local..."
Export-PWDocumentsSimple -InputDocuments $historyDoc -TargetFolder $tempWorkDir | Out-Null
Start-Sleep -Milliseconds 600   # let PW/antivirus release file handle (Error 100 can leave file locked briefly)

# Prefer path reported by export; else expected path; else newest matching name
if ($historyDoc.CopiedOutLocalFileName -and (Test-Path $historyDoc.CopiedOutLocalFileName)) {
  $localHistory = $historyDoc.CopiedOutLocalFileName
} elseif (Test-Path $localHistory) {
  # use $localHistory as set
} else {
  $foundHist = Get-ChildItem $tempWorkDir -File | Where-Object { $_.Name -ieq $HistoryDocName } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if ($foundHist) { $localHistory = $foundHist.FullName }
}

if (-not (Test-Path $localHistory)) {
  throw "History export failed; expected local file not found: $localHistory"
}
# Copy to unique file in TEMP; on "Access denied" retry with a different filename (suffix)
$localHistoryExport = $localHistory
$histBase = [System.IO.Path]::GetFileNameWithoutExtension($HistoryDocName) + "_hist_" + (Get-Date -Format "yyyyMMdd_HHmmss")
$localHistoryTmp = $null
foreach ($suffix in @('', '_2', '_3', '_4', '_5')) {
  $dest = Join-Path $tempWorkDir ($histBase + $suffix + ".pdf")
  try {
    Invoke-RetryOnAccessDenied { Copy-Item -LiteralPath $localHistoryExport -Destination $dest -Force }
    $localHistoryTmp = $dest
    break
  } catch {
    if ($_.Exception.Message -match 'Access to the path .* is denied' -and $suffix -ne '_5') { continue }
    throw
  }
}
Remove-ItemWithRetry $localHistoryExport   # best effort; PW may still have lock on export
$localHistory = $localHistoryTmp
$fh = Get-Item -LiteralPath $localHistory
if ($fh.Length -eq 0) {
  Remove-ItemWithRetry $localHistory
  throw "History export failed (zero file size - unmanaged copy?). Document may have no file in PW."
}
$header = [System.IO.File]::ReadAllBytes($localHistory) | Select-Object -First 4
$isPdf = ($header.Count -ge 4) -and ([char]$header[0] -eq '%' -and [char]$header[1] -eq 'P' -and [char]$header[2] -eq 'D' -and [char]$header[3] -eq 'F')
if (-not $isPdf) {
  Remove-ItemWithRetry $localHistory
  throw "History export failed (file not a valid PDF - Error 100 unmanaged copy?). Document may have no file in PW."
}
Write-Log "Local history file: $localHistory"

# Ephemeral qpdf page-1 slice when OverlayOldFromHistoryOnly (deleted after PW update; not written over by exe).
$overlayEphemeralPage1Master = $null
$overlaySheetWorkDirForExe = ""

# Optional --current-master: persistent work\ file, or temp page-1 slice when OverlayOldFromHistoryOnly + qpdf.
# --sheet-work-dir: per-sheet split pages under work\<baseName>\ (default $true); skips qpdf temp current-master.
$overlayCurrentMasterForExe = ""
if ($haveOverlay -and $OverlaySheetWorkDir) {
  $sheetWorkDirPath = Join-Path $workDir $baseName
  Initialize-Directory $sheetWorkDirPath
  $overlaySheetWorkDirForExe = $sheetWorkDirPath
  Write-Log "Overlay: per-sheet work dir (split pages + MANIFEST): $sheetWorkDirPath"
} elseif ($haveOverlay -and -not $OverlayOldFromHistoryOnly) {
  $resolvedOverlayCurrentMaster = if ($OverlayCurrentMasterPath) {
    $OverlayCurrentMasterPath
  } else {
    Join-Path $workDir ("${baseName}_current_master.pdf")
  }
  Initialize-Directory (Split-Path -Parent $resolvedOverlayCurrentMaster)
  if (-not (Test-Path -LiteralPath $resolvedOverlayCurrentMaster)) {
    if ($haveQpdf) {
      Write-Log "Seeding overlay current-master from history page 1 (first run for this sheet): $resolvedOverlayCurrentMaster"
      & $QpdfExe --empty --pages $localHistory 1-1 -- $resolvedOverlayCurrentMaster | Out-Null
      if (-not (Test-Path -LiteralPath $resolvedOverlayCurrentMaster)) {
        Write-Log "qpdf did not create current-master seed; Old layer may be empty until this file exists." -Severity WARNING
      }
    } else {
      Write-Log "No current-master file and qpdf not available to seed it. Install qpdf or pass -OverlayCurrentMasterPath to a baseline PDF (see test\run_f0548dv206_qc_two_step.ps1). Overlay may show an empty Old layer." -Severity WARNING
    }
  }
  if (Test-Path -LiteralPath $resolvedOverlayCurrentMaster) {
    $overlayCurrentMasterForExe = $resolvedOverlayCurrentMaster
    Write-Log "Overlay current-master: $overlayCurrentMasterForExe"
  }
} elseif ($haveOverlay -and $OverlayOldFromHistoryOnly) {
  if ($haveQpdf) {
    $overlayEphemeralPage1Master = Join-Path $tempWorkDir ("${baseName}_qc_page1_$stamp.pdf")
    Write-Log "Overlay: qpdf page 1 of exported *-qc.pdf -> temp --current-master (PW source; exe may overwrite temp with incoming before delete)"
    & $QpdfExe --empty --pages $localHistory 1-1 -- $overlayEphemeralPage1Master | Out-Null
    if (Test-Path -LiteralPath $overlayEphemeralPage1Master) {
      $overlayCurrentMasterForExe = $overlayEphemeralPage1Master
    } else {
      Write-Log "qpdf failed to create page-1 slice; overlay will use Python extract from history only." -Severity WARNING
    }
  } else {
    Write-Log "OverlayOldFromHistoryOnly but qpdf missing: Python page-1 extract only (Old may be empty for some Civil PDFs)." -Severity WARNING
  }
}

# Prepend merge locally (use overlay when available for layered Old/New/Current)
Write-Log "Merging (prepend) incoming -> history..."
if ($haveOverlay) {
  Write-Log "Using qc_overlay_prepend (layered output)..."
  Invoke-PdfPrependOverlay -incomingPdf $localIncoming -historyPdf $localHistory -outPdf $localMerged -overlayExe $QcOverlayExe -currentMasterPath $overlayCurrentMasterForExe -sheetWorkDir $overlaySheetWorkDirForExe
} elseif ($haveQpdf) {
  Write-Log "Using qpdf (simple merge, no layers)..."
  Invoke-PdfPrependMerge -newPdf $localIncoming -historyPdf $localHistory -outPdf $localMerged
} else {
  throw "Neither overlay exe nor qpdf found. Run .\overlay\build_overlay_exe.ps1 (dist\qc_overlay_prepend\...) or place dist\qc_overlay_prepend.exe, or install qpdf."
}
Write-Log "Merged file created: $localMerged"

# Upload / replace history in PW (parameter name varies by pwps_dab version: LocalPath, SourcePath, etc.)
Write-Log "Updating history document content in PW..."
if ($PSCmdlet.ShouldProcess($historyDoc.FullPath, "Update document file content from merged PDF")) {

  $updateCmd = Get-Command Update-PWDocumentFile
  $fileParamName = $updateCmd.Parameters.Keys | Where-Object {
    $_ -match '^(LocalPath|SourcePath|FilePath|Path|SourceFile|File)$' -and $_ -notin @('Verbose','Debug','ErrorAction','WarningAction','InformationAction','ErrorVariable','WarningVariable','OutVariable','OutBuffer','PipelineVariable')
  } | Select-Object -First 1
  if (-not $fileParamName) {
    $fileParamName = $updateCmd.Parameters.Keys | Where-Object { $_ -match 'path|file' -and $_ -ne 'InputDocument' } | Select-Object -First 1
  }
  if (-not $fileParamName) { throw "Update-PWDocumentFile: could not find file path parameter. Parameters: $($updateCmd.Parameters.Keys -join ', ')" }
  $pwUpdateFileParams = @{ InputDocument = $historyDoc; $fileParamName = $localMerged }
  Update-PWDocumentFile @pwUpdateFileParams | Out-Null

  Write-Log "Updated history document."
  # Clear working files after successful PW update (Remove-ItemWithRetry continues if antivirus blocks)
  @($localIncoming, $localHistory, $localMerged, $overlayEphemeralPage1Master) | Where-Object { $_ } | ForEach-Object { Remove-ItemWithRetry $_ }
} else {
  Write-Log "WhatIf: would update history document file content."
}

Write-Log "Done."
Close-PWConnection -ErrorAction SilentlyContinue
