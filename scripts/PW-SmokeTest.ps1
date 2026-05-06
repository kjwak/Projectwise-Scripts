<#
.SYNOPSIS
End-to-end ProjectWise smoke test for this repo's automation assumptions.

.DESCRIPTION
Validates that this machine can:
  - Connect / disconnect (pwps)
  - Search a document (pwps_dab)
  - Export/download (pwps_dab)
  - Upload/update content (pwps_dab Update-PWDocumentFile)
  - Read and change Description (Update-PWDocumentProperties)

This script DOES perform writes (upload + description set/revert). Use a known safe test document.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$DatasourceName = 'typsa-us-pw.bentley.com:typsa-us-pw-03',

    [Parameter(Mandatory = $false)]
    [string]$CredentialPath = 'C:\PW_QC_LOCAL\pw_cred.txt',

    [Parameter(Mandatory = $false)]
    [string]$FolderPath = 'AZDOT 2024\AZFWY1704-FD02-SR202 - I-10 to SR101\CADD\Working\TYPSA\Drainage\JFlint\Prepend Test',

    [Parameter(Mandatory = $false)]
    [string]$DocumentName = 'input1.pdf'
)

$ErrorActionPreference = 'Stop'

function Get-PWCredFromKeyValueFile([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { throw "Credential file not found: $Path" }
    $lines = Get-Content -LiteralPath $Path -ErrorAction Stop
    $uLine = $lines | Where-Object { $_ -match '^\s*username\s*=' } | Select-Object -First 1
    $pLine = $lines | Where-Object { $_ -match '^\s*password\s*=' } | Select-Object -First 1
    if (-not $uLine -or -not $pLine) { throw "Bad credential file format (expected username=/password=): $Path" }
    $user = ($uLine -split '=', 2)[1].Trim()
    $pass = ($pLine -split '=', 2)[1].Trim()
    if ([string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($pass)) { throw "Credential file missing values: $Path" }
    $sec = ConvertTo-SecureString $pass -AsPlainText -Force
    return [pscredential]::new($user, $sec)
}

Import-Module pwps -Force
Import-Module pwps_dab -Force

$cred = Get-PWCredFromKeyValueFile -Path $CredentialPath

Write-Host ("Connecting to {0} as {1}" -f $DatasourceName, $cred.UserName) -ForegroundColor Cyan
Open-PWConnection -DatasourceName $DatasourceName -UserName $cred.UserName -Password $cred.Password | Out-Null

try {
    $doc = Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -DocumentName $DocumentName -PopulatePath -ErrorAction Stop
    if (-not $doc) { throw "Document not found: $FolderPath\\$DocumentName" }
    Write-Host ("Found DocumentID={0} FullPath={1}" -f $doc.DocumentID, $doc.FullPath) -ForegroundColor Green

    $tmp = Join-Path $env:TEMP ('pwps_smoke_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
    New-Item -ItemType Directory -Path $tmp -Force | Out-Null

    Write-Host ("Exporting to {0}" -f $tmp) -ForegroundColor Cyan
    Export-PWDocumentsSimple -InputDocuments $doc -TargetFolder $tmp | Out-Null
    Start-Sleep -Milliseconds 600

    $local = Join-Path $tmp $DocumentName
    if (-not (Test-Path -LiteralPath $local)) {
        $found = Get-ChildItem -LiteralPath $tmp -File | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($found) { $local = $found.FullName }
    }
    if (-not (Test-Path -LiteralPath $local)) { throw "Export did not produce a local file in $tmp" }
    Write-Host ("Export OK: {0}" -f $local) -ForegroundColor Green

    Write-Host "Uploading same file back (Update-PWDocumentFile)..." -ForegroundColor Cyan
    Update-PWDocumentFile -InputDocuments @($doc) -NewFilePathName $local | Out-Null
    Write-Host "Upload OK (Update-PWDocumentFile returned)." -ForegroundColor Green

    $doc2 = Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -DocumentName $DocumentName -PopulatePath -ErrorAction Stop
    $origDesc = [string]$doc2.Description
    $tag = 'SMOKETEST_' + (Get-Date -Format 'yyyyMMdd_HHmmss')
    $newDesc = (($origDesc + ' ' + $tag).Trim())

    Write-Host ("Original Description: [{0}]" -f $origDesc) -ForegroundColor Yellow
    Write-Host ("Setting Description:  [{0}]" -f $newDesc) -ForegroundColor Cyan
    $doc2.Description = $newDesc
    Update-PWDocumentProperties $doc2 -ErrorAction Stop | Out-Null

    $doc3 = Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -DocumentName $DocumentName -PopulatePath -ErrorAction Stop
    Write-Host ("Description after set: [{0}]" -f ([string]$doc3.Description)) -ForegroundColor Green

    Write-Host "Reverting Description..." -ForegroundColor Cyan
    $doc3.Description = $origDesc
    Update-PWDocumentProperties $doc3 -ErrorAction Stop | Out-Null

    $doc4 = Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -DocumentName $DocumentName -PopulatePath -ErrorAction Stop
    Write-Host ("Description after revert: [{0}]" -f ([string]$doc4.Description)) -ForegroundColor Green

    Write-Host ("DONE. TempDir={0}" -f $tmp) -ForegroundColor Green
} finally {
    try { Close-PWConnection | Out-Null } catch { }
}

