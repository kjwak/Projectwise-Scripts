<#
.SYNOPSIS
Connects to ProjectWise and prints child folders/docs for a given folder path.

.DESCRIPTION
Read-only helper to validate folder paths in the current datasource.
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
    [int]$Max = 40
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
    $folderViewCmd = Get-Command -Name Get-PWFolderView -ErrorAction Stop
    if ($folderViewCmd.Parameters.ContainsKey('InputFolder')) {
        $f = Get-PWFolders -FolderPath $FolderPath -JustOne -ErrorAction Stop
        if (-not $f) { throw "Folder not found: $FolderPath" }
        $v = $f | Get-PWFolderView -ErrorAction Stop
    } else {
        $v = Get-PWFolderView -FolderPath $FolderPath -ErrorAction Stop
    }

    $folders = @()
    if ($v.Folders) { $folders = @($v.Folders) }
    elseif ($v.Children) { $folders = @($v.Children | Where-Object { $_ -and $_.FolderPath -and -not $_.DocumentID }) }

    $docs = @()
    if ($v.Documents) { $docs = @($v.Documents) }
    elseif ($v.Children) { $docs = @($v.Children | Where-Object { $_ -and ($_.DocumentID -or $_.Name) }) }

    Write-Host ("Folders: {0}" -f $folders.Count) -ForegroundColor Green
    $folders | Select-Object -First $Max Name,FolderPath | Format-Table -AutoSize
    Write-Host ("Documents: {0}" -f $docs.Count) -ForegroundColor Green
    $docs | Select-Object -First $Max Name,Description,DocumentID,FullPath | Format-Table -AutoSize
} finally {
    try { Close-PWConnection | Out-Null } catch { }
}

