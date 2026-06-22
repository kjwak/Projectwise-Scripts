#Requires -Version 5.1
<#
.SYNOPSIS
Backfill dbo.pw_users from ProjectWise for user numbers seen in audit_events (or explicit -UserNumber list).
#>
[CmdletBinding()]
param(
    [string]$AppSettingsPath = '',
    [int[]]$UserNumber = @(),
    [int]$MaxUsers = 200,
    [switch]$FromAuditEventsOnly
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

. (Join-Path $PSScriptRoot '..\Restore-QCModuleExports.ps1') -RepoRoot $repoRoot
Import-QCModuleBootstrapSet -FeatureModules @(
    'Database\Core.Database.psm1'
    'ProjectWise\PW.Connection.psm1'
    'ProjectWise\PW.Users.psm1'
) -RequiredCommands @(
    'Read-QCAppSettings'
    'Initialize-QCDatabaseSchema'
    'ConvertTo-HashtableDeep'
    'Get-PWCredentialFromFile'
    'Connect-PW'
    'Sync-PWUserDirectory'
    'Disconnect-PW'
) -Context 'Sync-PWUserDirectory bootstrap'

foreach ($moduleName in @('pwps', 'pwps_dab')) {
    if (-not (Get-Module -Name $moduleName -ErrorAction SilentlyContinue)) {
        $available = Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1
        if ($available) { Import-Module $available.Name -ErrorAction SilentlyContinue | Out-Null }
    }
}

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config

$init = Initialize-QCDatabaseSchema -Config $config
if (-not $init.IsSuccess) { throw "Schema init failed: $($init.Message)" }

$pw = @{}
if ($config.ContainsKey('projectWise') -and $config.projectWise) {
    $pwNorm = ConvertTo-HashtableDeep -Value $config.projectWise
    if ($pwNorm) { $pw = $pwNorm }
}
$ds = if ($pw.datasourceName) { [string]$pw.datasourceName } else { 'typsa-us-pw.bentley.com:typsa-us-pw-03' }
$credPath = if ($pw.credentialPath) { [string]$pw.credentialPath } else { 'C:\PW_QC_LOCAL\pw_cred.txt' }
$credRes = Get-PWCredentialFromFile -CredentialPath $credPath
if (-not $credRes.IsSuccess) { throw ($credRes.Code + ': ' + $credRes.Message) }
$cred = [pscredential]$credRes.Data.credential
$conn = Connect-PW -DatasourceName $ds -Credential $cred
if (-not $conn.IsSuccess) { throw ($conn.Code + ': ' + $conn.Message) }

try {
    $params = @{ Config = $config; MaxUsers = $MaxUsers }
    if ($UserNumber.Count -gt 0) { $params.UserNumbers = $UserNumber }
    if ($FromAuditEventsOnly -or $UserNumber.Count -eq 0) { $params.ResolveFromAuditEvents = $true }
    $sync = Sync-PWUserDirectory @params -Verbose:$VerbosePreference
    if (-not $sync.IsSuccess) { throw "User directory sync failed: $($sync.Message)" }
    Write-Host $sync.Message
    if ($sync.Data.requested) {
        Write-Host ("  requested={0} sqlResolved={1} synced={2} skipped={3}" -f `
            $sync.Data.requested, $sync.Data.sqlResolved, $sync.Data.synced, $sync.Data.skipped)
    }
    if ($sync.Code -eq 'PW_USER_SYNC_EMPTY' -and $sync.Data.sampleIds) {
        Write-Warning ("Sample unresolved pw_userno: {0}" -f ($sync.Data.sampleIds -join ', '))
    }
} finally {
    Disconnect-PW | Out-Null
}
