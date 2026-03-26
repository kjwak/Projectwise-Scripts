#requires -Version 5.1
# NOTE: no param block for watcher scripts

# Relaunch in 32-bit PowerShell if started in 64-bit
if ([Environment]::Is64BitProcess) {
    $x86 = "$env:WINDIR\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    & $x86 -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath
    exit
  }
  
  Import-Module "C:\Program Files (x86)\Bentley\ProjectWise\bin\PowerShell\pwps\pwps.psd1" -Force
  

# --- CONFIG ---
$Datasource   = "typsa-us-pw-03"   # your datasource name (as seen in Explorer login)
$ProjectPath  = "\AZTEC Engineering Group, Inc\AZDOT 2024\AZFWY1704-FD02-SR202 - I-10 to SR101\CADD\Working\TYPSA\Drainage\JFlint\Prepend Test"
$QcFolderPath = "$ProjectPath\QC"   # pick a real folder you want to use in PW
$HistoryName  = "QC_History.pdf"

$LocalWork    = "C:\PW_QC_LOCAL\work"
$QPDF         = "qpdf"              # must be on PATH

# input PDF you want to ingest this run
param([Parameter(Mandatory=$true)] [string] $NewPdfLocal)

$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force -Path $LocalWork | Out-Null

function Prepend($newPdf, $historyPdf, $outPdf) {
  if (-not (Test-Path $historyPdf)) {
    Copy-Item $newPdf $outPdf -Force
    return
  }
  & $QPDF $newPdf $historyPdf -- $outPdf
  if (-not (Test-Path $outPdf)) { throw "qpdf failed to create: $outPdf" }
}

# --- 1) Login to ProjectWise ---
# Typical pattern (varies by install):
# New-PWLogin -Datasource $Datasource -UseWindowsCredentials
# or: New-PWLogin -Datasource $Datasource -UserName ... -Password ...

# --- 2) Resolve folder in PW ---
# $folder = Get-PWFolder -Path $QcFolderPath

# --- 3) Upload the new QC pdf as its own document (optional but nice) ---
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$newDocName = "QC_$stamp.pdf"
# $newDoc = Add-PWDocument -Folder $folder -FilePath $NewPdfLocal -DocumentName $newDocName

# --- 4) Download existing history doc (if exists) ---
$localHistory = Join-Path $LocalWork $HistoryName
$localNewHist = Join-Path $LocalWork ("QC_History_NEW_$stamp.pdf")

# $historyDoc = Get-PWDocument -Folder $folder -DocumentName $HistoryName -ErrorAction SilentlyContinue
# if ($historyDoc) { Export-PWDocumentContent -Document $historyDoc -OutFile $localHistory }

# --- 5) Prepend merge locally ---
Prepend -newPdf $NewPdfLocal -historyPdf $localHistory -outPdf $localNewHist

# --- 6) Upload/replace history in PW ---
# if (-not $historyDoc) {
#   $historyDoc = Add-PWDocument -Folder $folder -FilePath $localNewHist -DocumentName $HistoryName
# } else {
#   Set-PWDocumentContent -Document $historyDoc -FilePath $localNewHist
# }

Write-Host "OK: updated history in PW (logic complete)."
