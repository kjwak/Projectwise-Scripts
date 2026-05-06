<#
.SYNOPSIS
Quick smoke test for ProjectWise connection module.

.DESCRIPTION
Loads appsettings.json and attempts to connect/disconnect to ProjectWise using pwps_dab.
No ProjectWise writes are performed.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$AppSettingsPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'appsettings.json')
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\PW.Connection.psm1') -Force

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config

$pw = @{}
if ($config.ContainsKey('projectWise') -and $config.projectWise) {
    $pwNorm = ConvertTo-HashtableDeep -Value $config.projectWise
    if ($pwNorm) { $pw = $pwNorm }
}

$ds = if ($pw.ContainsKey('datasourceName') -and $pw.datasourceName) { [string]$pw.datasourceName } else { 'typsa-us-pw.bentley.com:typsa-us-pw-03' }
$credPath = if ($pw.ContainsKey('credentialPath') -and $pw.credentialPath) { [string]$pw.credentialPath } else { 'C:\PW_QC_LOCAL\pw_cred.txt' }

Write-Host "Datasource: $ds" -ForegroundColor Cyan
Write-Host "CredentialPath: $credPath" -ForegroundColor Cyan

$credRes = Get-PWCredentialFromFile -CredentialPath $credPath
if (-not $credRes.IsSuccess) { throw ($credRes.Code + ': ' + $credRes.Message) }
$cred = [pscredential]$credRes.Data.credential

$conn = Connect-PW -DatasourceName $ds -Credential $cred
if (-not $conn.IsSuccess) { throw ($conn.Code + ': ' + $conn.Message) }
Write-Host "Connected as $($conn.Data.userName)" -ForegroundColor Green

$disc = Disconnect-PW
if (-not $disc.IsSuccess) { throw ($disc.Code + ': ' + $disc.Message) }
Write-Host "Disconnected OK" -ForegroundColor Green

