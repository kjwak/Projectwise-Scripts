# prepend_qc.ps1
# Purpose: prepend incoming PDF into a QC history PDF in the same PW folder (create/update (filename)-qc.pdf).
# Notes:
# - Uses Bentley IMS login via Open-PWConnection.
# - Avoids wildcard searches (your environment doesn't handle them reliably).
# - History document is (incoming filename base)-qc.pdf, saved in the same PW folder as the incoming file.
# - If the history doc doesn't exist, it creates it from the incoming PDF (base case).
#   - If it exists, it exports both and prepends (with overlay layers when available), then updates the history doc in PW.
#
# REQUIREMENTS:
#   - pwps_dab module installed
#   - qpdf installed and on PATH (or set $QpdfExe to full path) - used when overlay exe not found
#   - dist\qc_overlay_prepend.exe (optional) - when present, creates layered overlay (Old red, New green, Current black)
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
  [string] $QcOverlayExe = "",  # default: dist\qc_overlay_prepend.exe next to script

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

# Default overlay exe path (standalone, no Python needed)
if (-not $QcOverlayExe) {
  $QcOverlayExe = Join-Path $PSScriptRoot "dist\qc_overlay_prepend.exe"
}

# pwps_dab requires MTA; Cursor/VS Code terminals often use STA. Re-launch in MTA to avoid ThreadOptions error.
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -eq 'STA') {
  & powershell.exe -MTA -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath @args
  exit $LASTEXITCODE
}

$ErrorActionPreference = "Stop"
if (-not $LogDir) { $LogDir = Join-Path $LocalRoot "logs" }

# ProjectWise credentials: C:\PW_QC_LOCAL\pw_cred.txt
# Format: username=domain\user and password=... on separate lines.
$CredentialPath = 'C:\PW_QC_LOCAL\pw_cred.txt'

$PrependQc_LogDir = $LogDir
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
function Safe-RemoveFile([string]$path) {
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

function Prepend-Pdf([string]$newPdf, [string]$historyPdf, [string]$outPdf) {
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

# Avoids cryptic "corrupted and unreadable" from the Windows loader when the path is an LFS pointer, truncated file, or non-PE.
function Get-HexPrefix([byte[]]$bytes, [int]$count) {
  $n = [Math]::Min($count, $bytes.Length)
  $parts = for ($i = 0; $i -lt $n; $i++) { '{0:X2}' -f $bytes[$i] }
  return ($parts -join ' ')
}

function Read-OverlayExeHeaderSample([string]$path) {
  $len = (Get-Item -LiteralPath $path).Length
  $fs = [System.IO.File]::OpenRead($path)
  try {
    $toRead = [Math]::Min(4096, [int]$len)
    if ($toRead -lt 2) {
      return @{ Ok = $false; Err = "Overlay exe too small ($len bytes): $path" }
    }
    $buf = New-Object byte[] $toRead
    $read = $fs.Read($buf, 0, $toRead)
    if ($read -lt 2) {
      return @{ Ok = $false; Err = "Could only read $read byte(s) from overlay exe (file locked or incomplete). Path: $path" }
    }
    return @{ Ok = $true; Buf = $buf; Read = $read; Len = $len }
  } finally {
    $fs.Dispose()
  }
}

function Test-OverlayHeaderBytes([byte[]]$buf, [int]$read, [long]$len, [string]$path) {
  $hex = Get-HexPrefix $buf 16
  if ($buf[0] -eq 0x4D -and $buf[1] -eq 0x5A) {
    return @{ Ok = $true }
  }
  $head = [System.Text.Encoding]::ASCII.GetString($buf[0..([Math]::Min(199, $read - 1))])
  if ($head -match 'git-lfs|oid sha256') {
    return @{ Ok = $false; Err = "Overlay exe is a Git LFS pointer, not the real binary. Run 'git lfs pull' or copy qc_overlay_prepend.exe from a build machine. Path: $path" }
  }
  if ($read -ge 4 -and $buf[0] -eq 0x7F -and $buf[1] -eq 0x45 -and $buf[2] -eq 0x4C -and $buf[3] -eq 0x46) {
    return @{ Ok = $false; Err = "Overlay exe is ELF (Linux), not Windows. Run overlay\build_overlay_exe.ps1 on a Windows machine and copy that qc_overlay_prepend.exe. First bytes: $hex Path: $path" }
  }
  if ($read -ge 4 -and $buf[0] -eq 0xCF -and $buf[1] -eq 0xFA -and $buf[2] -eq 0xED -and $buf[3] -eq 0xFE) {
    return @{ Ok = $false; Err = "Overlay exe is Mach-O (macOS), not Windows. Build with PyInstaller on Windows and deploy that .exe. First bytes: $hex Path: $path" }
  }
  if ($read -ge 4 -and $buf[0] -eq 0xFE -and $buf[1] -eq 0xED -and $buf[2] -eq 0xFA -and $buf[3] -eq 0xCE) {
    return @{ Ok = $false; Err = "Overlay exe is Mach-O (macOS), not Windows. Build with PyInstaller on Windows and deploy that .exe. First bytes: $hex Path: $path" }
  }
  $sampleEnd = [Math]::Min(512, $read) - 1
  $nonZero = 0
  for ($i = 0; $i -le $sampleEnd; $i++) { if ($buf[$i] -ne 0) { $nonZero++ } }
  if ($nonZero -eq 0) {
    return @{ Ok = $false; AllZeros = $true; Hex = $hex; Len = $len }
  }
  return @{ Ok = $false; Err = "Overlay exe does not look like a Windows PE (expected MZ at start). First bytes: $hex size=$len bytes Path: $path Rebuild with .\overlay\build_overlay_exe.ps1 on Windows, or copy a Windows-built qc_overlay_prepend.exe." }
}

# Returns path to invoke (original or TEMP copy when first read returned zeros; Copy-Item often yields real bytes).
function Resolve-OverlayExePath([string]$path) {
  if (-not (Test-Path -LiteralPath $path)) {
    throw "Overlay exe not found: $path"
  }
  $s = Read-OverlayExeHeaderSample $path
  if (-not $s.Ok) { throw $s.Err }
  $v = Test-OverlayHeaderBytes $s.Buf $s.Read $s.Len $path
  if ($v.Ok) {
    return $path
  }
  if ($v.Err) {
    throw $v.Err
  }
  if ($v.AllZeros) {
    $tempDir = Join-Path $env:TEMP "PW_QC"
    if (-not (Test-Path -LiteralPath $tempDir)) {
      New-Item -ItemType Directory -Force -Path $tempDir | Out-Null
    }
    $tempExe = Join-Path $tempDir "qc_overlay_prepend.exe"
    Copy-Item -LiteralPath $path -Destination $tempExe -Force
    $s2 = Read-OverlayExeHeaderSample $tempExe
    if (-not $s2.Ok) { throw $s2.Err }
    $v2 = Test-OverlayHeaderBytes $s2.Buf $s2.Read $s2.Len $tempExe
    if ($v2.Ok) {
      Write-Log "Overlay exe first bytes were zero at source; using local copy: $tempExe"
      return $tempExe
    }
    if ($v2.Err) { throw $v2.Err }
    if ($v2.AllZeros) {
      $hz = $v2.Hex
      throw "Overlay exe still unreadable after copy to TEMP (first bytes: $hz). Source: $path size=$($s.Len). Replace the file with a known-good Windows build, try a local path (e.g. C:\Tools\qc_overlay_prepend.exe) via -QcOverlayExe, rule out antivirus blocking, or copy via USB (not a partial network copy)."
    }
    throw "Resolve-OverlayExePath: unexpected validation state for $tempExe"
  }
  throw "Resolve-OverlayExePath: unexpected state for $path"
}

function Prepend-PdfWithOverlay([string]$incomingPdf, [string]$historyPdf, [string]$outPdf, [string]$overlayExe) {
  if (-not (Test-Path $historyPdf)) {
    Copy-Item -Path $incomingPdf -Destination $outPdf -Force
    return
  }

  $exeToRun = Resolve-OverlayExePath $overlayExe

  # qc_overlay_prepend: page1 of history = Old/red, incoming = New/green + Current/black, prepended to history
  try {
    $overlayOut = & $exeToRun $incomingPdf $historyPdf -o $outPdf 2>&1
    $overlayExit = $LASTEXITCODE
    $overlayOut | ForEach-Object { Write-Log $_ }
  } catch {
    $m = $_.Exception.Message
    if ($m -match 'corrupted and unreadable') {
      throw "Windows could not load qc_overlay_prepend.exe ($exeToRun). Often: (1) exe on a slow/network/cloud path or blocked by AV - copy to a local folder (e.g. C:\Tools\) and set -QcOverlayExe; (2) incomplete or wrong-architecture copy - redeploy a Windows-built exe; (3) PyInstaller onedir - deploy the full folder including _internal. Original error: $m"
    }
    throw
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
Ensure-Dir $exportDir
Ensure-Dir $workDir
Ensure-Dir $tempWorkDir

# Validate qpdf presence (only needed when history exists and we need merge)
$haveQpdf = $false
try {
  Assert-Command $QpdfExe
  $haveQpdf = $true
} catch {
  # We'll only hard-fail later if we actually need to prepend.
  Write-Log $_.Exception.Message -Severity WARNING
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
Safe-RemoveFile $localIncoming   # best effort; PW may still have lock
$localIncoming = $localIncomingTmp
$fi = Get-Item -LiteralPath $localIncoming
if ($fi.Length -eq 0) {
  Safe-RemoveFile $localIncoming
  throw "Incoming export failed (zero file size - unmanaged copy?). Document may have no file in PW."
}
$header = [System.IO.File]::ReadAllBytes($localIncoming) | Select-Object -First 4
$isPdf = ($header.Count -ge 4) -and ([char]$header[0] -eq '%' -and [char]$header[1] -eq 'P' -and [char]$header[2] -eq 'D' -and [char]$header[3] -eq 'F')
if (-not $isPdf) {
  Safe-RemoveFile $localIncoming
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
    if (Test-Path $localIncoming) { Safe-RemoveFile $localIncoming }
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
Safe-RemoveFile $localHistoryExport   # best effort; PW may still have lock on export
$localHistory = $localHistoryTmp
$fh = Get-Item -LiteralPath $localHistory
if ($fh.Length -eq 0) {
  Safe-RemoveFile $localHistory
  throw "History export failed (zero file size - unmanaged copy?). Document may have no file in PW."
}
$header = [System.IO.File]::ReadAllBytes($localHistory) | Select-Object -First 4
$isPdf = ($header.Count -ge 4) -and ([char]$header[0] -eq '%' -and [char]$header[1] -eq 'P' -and [char]$header[2] -eq 'D' -and [char]$header[3] -eq 'F')
if (-not $isPdf) {
  Safe-RemoveFile $localHistory
  throw "History export failed (file not a valid PDF - Error 100 unmanaged copy?). Document may have no file in PW."
}
Write-Log "Local history file: $localHistory"

# Prepend merge locally (use overlay when available for layered Old/New/Current)
Write-Log "Merging (prepend) incoming -> history..."
if ($haveOverlay) {
  Write-Log "Using qc_overlay_prepend (layered output)..."
  Prepend-PdfWithOverlay -incomingPdf $localIncoming -historyPdf $localHistory -outPdf $localMerged -overlayExe $QcOverlayExe
} elseif ($haveQpdf) {
  Write-Log "Using qpdf (simple merge, no layers)..."
  Prepend-Pdf -newPdf $localIncoming -historyPdf $localHistory -outPdf $localMerged
} else {
  throw "Neither overlay exe nor qpdf found. Place dist\qc_overlay_prepend.exe next to this script or install qpdf."
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
  $args = @{ InputDocument = $historyDoc; $fileParamName = $localMerged }
  Update-PWDocumentFile @args | Out-Null

  Write-Log "Updated history document."
  # Clear working files after successful PW update (Safe-RemoveFile continues if antivirus blocks)
  @($localIncoming, $localHistory, $localMerged) | Where-Object { $_ } | ForEach-Object { Safe-RemoveFile $_ }
} else {
  Write-Log "WhatIf: would update history document file content."
}

Write-Log "Done."
Close-PWConnection -ErrorAction SilentlyContinue
