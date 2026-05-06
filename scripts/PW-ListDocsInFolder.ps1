<#
.SYNOPSIS
Lists documents in a ProjectWise folder (read-only).

.DESCRIPTION
Connects, runs Get-PWDocumentsBySearch -JustThisFolder -PopulatePath, and prints counts plus a small sample.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$DatasourceName = 'typsa-us-pw.bentley.com:typsa-us-pw-03',

    [Parameter(Mandatory = $false)]
    [string]$CredentialPath = 'C:\PW_QC_LOCAL\pw_cred.txt',

    [Parameter(Mandatory = $true)]
    [string]$FolderPath,

    [Parameter(Mandatory = $false)]
    [int]$Max = 25
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
Open-PWConnection -DatasourceName $DatasourceName -UserName $cred.UserName -Password $cred.Password | Out-Null

try {
    Write-Host ("FolderPath: {0}" -f $FolderPath) -ForegroundColor Cyan
    $docs = @(Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -PopulatePath -ErrorAction Stop)
    Write-Host ("TotalDocs: {0}" -f $docs.Count) -ForegroundColor Green

    $pdfs = @($docs | Where-Object { $_.Name -match '\.pdf$' })
    Write-Host ("PdfDocs:   {0}" -f $pdfs.Count) -ForegroundColor Green

    $pdfs | Select-Object -First $Max Name,DocumentID,Description,FullPath | Format-Table -AutoSize
} finally {
    try { Close-PWConnection | Out-Null } catch { }
}

