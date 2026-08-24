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
    [string]$AppSettingsPath = ''
)

# pwps_dab requires MTA; Cursor/VS Code terminals are usually STA.
$scriptPath = $PSCommandPath
if (-not $scriptPath) { $scriptPath = $MyInvocation.MyCommand.Path }
if (-not $scriptPath) { $scriptPath = Join-Path $PSScriptRoot 'Test-PWConnection.ps1' }
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -eq 'STA') {
    if (-not (Test-Path -LiteralPath $scriptPath)) {
        throw "MTA relaunch: could not resolve script path. Tried: $scriptPath"
    }
    $staMtaHelper = Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'legacy\StaMtaRelaunch.ps1'
    if (-not (Test-Path -LiteralPath $staMtaHelper)) {
        throw "MTA relaunch helper not found: $staMtaHelper"
    }
    . $staMtaHelper
    $exeArgs = Build-PowerShellExeFileArgs -ScriptPath $scriptPath -BoundParameters $PSBoundParameters
    & powershell.exe @exeArgs
    exit $LASTEXITCODE
}

$ErrorActionPreference = 'Stop'

function _Get-ThisScriptDir {
    try {
        if ($PSScriptRoot -and -not [string]::IsNullOrWhiteSpace($PSScriptRoot)) { return $PSScriptRoot }
    } catch { }
    try {
        $p = $MyInvocation.MyCommand.Path
        if ($p -and (Test-Path -LiteralPath $p)) { return (Split-Path -Parent $p) }
    } catch { }
    return (Get-Location).Path
}

$scriptDir = _Get-ThisScriptDir
$repoRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = (Join-Path $repoRoot 'appsettings.json')
}

. (Join-Path $PSScriptRoot '..\Restore-QCModuleExports.ps1') -RepoRoot $repoRoot
Import-QCModuleBootstrapSet -FeatureModules @(
    'ProjectWise\PW.Connection.psm1'
) -RequiredCommands @(
    'Read-QCAppSettings'
) -Context 'Test-PWConnection bootstrap'

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
Write-Host ("Apartment: {0}" -f [System.Threading.Thread]::CurrentThread.GetApartmentState()) -ForegroundColor Cyan

$credRes = Get-PWCredentialFromFile -CredentialPath $credPath
if (-not $credRes.IsSuccess) { throw ($credRes.Code + ': ' + $credRes.Message) }
$cred = [pscredential]$credRes.Data.credential

$conn = Connect-PW -DatasourceName $ds -Credential $cred
if (-not $conn.IsSuccess) {
    $detail = $null
    try { if ($conn.Data -and $conn.Data.errorMessage) { $detail = [string]$conn.Data.errorMessage } } catch { }
    if ([string]::IsNullOrWhiteSpace($detail)) { throw ($conn.Code + ': ' + $conn.Message) }
    throw ($conn.Code + ': ' + $conn.Message + ' ' + $detail)
}
Write-Host "Connected as $($conn.Data.userName)" -ForegroundColor Green

$disc = Disconnect-PW
if (-not $disc.IsSuccess) { throw ($disc.Code + ': ' + $disc.Message) }
Write-Host "Disconnected OK" -ForegroundColor Green

