<#
.SYNOPSIS
Runs periodic QC_Pipeline audit_events retention.

.DESCRIPTION
Intended for scheduled Task Scheduler runs. Reads defaults from appsettings
database.retention (audit events only). Workflow telemetry is not aged automatically;
use Remove-QCWorkflowEvents.ps1 with -FolderPathLike for targeted cleanup.

.EXAMPLE
.\scripts\Invoke-QCDatabaseRetention.ps1
.\scripts\Invoke-QCDatabaseRetention.ps1 -ConfirmDeletes
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$AppSettingsPath = '',
    [switch]$ConfirmDeletes,
    [switch]$DryRun,
    [int]$AuditEventsDays = -1,
    [switch]$AuditIncludeUnprocessed,
    [switch]$Pretty
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

. (Join-Path $PSScriptRoot '..\Import-QCScriptModules.ps1') -RepoRoot $repoRoot

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config

function _QDR-GetRetentionInt {
    param([hashtable]$Map, [string]$Name, [int]$Fallback)
    if (-not $Map) { return $Fallback }
    if ($null -eq $Map[$Name]) { return $Fallback }
    try { return [int]$Map[$Name] } catch { return $Fallback }
}

function _QDR-GetRetentionBool {
    param([hashtable]$Map, [string]$Name, [bool]$Fallback)
    if (-not $Map) { return $Fallback }
    if ($null -eq $Map[$Name]) { return $Fallback }
    try { return [bool]$Map[$Name] } catch { return $Fallback }
}

$retention = $null
if ($config.database -and $config.database.retention) {
    $retention = $config.database.retention
    if ($retention -isnot [hashtable]) {
        $retention = @{}
        foreach ($p in $config.database.retention.PSObject.Properties) {
            $retention[$p.Name] = $p.Value
        }
    }
}

$auditDays = if ($AuditEventsDays -ge 0) { $AuditEventsDays } else { _QDR-GetRetentionInt -Map $retention -Name 'auditEventsDays' -Fallback 90 }
$auditProcessedOnly = -not $AuditIncludeUnprocessed.IsPresent
if (-not $AuditIncludeUnprocessed.IsPresent) {
    $auditProcessedOnly = _QDR-GetRetentionBool -Map $retention -Name 'auditEventsProcessedOnly' -Fallback $true
}

$auditArgs = @{
    AppSettingsPath = $AppSettingsPath
    OlderThanDays   = $auditDays
}
if ($ConfirmDeletes.IsPresent) { $auditArgs['ConfirmDeletes'] = $true }
if ($DryRun.IsPresent) { $auditArgs['DryRun'] = $true }
if ($auditProcessedOnly) { $auditArgs['ProcessedOnly'] = $true }
else { $auditArgs['IncludeUnprocessed'] = $true }
if ($Pretty) { $auditArgs['Pretty'] = $true }

Write-Host "=== audit_events (older than $auditDays days) ===" -ForegroundColor Cyan
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Remove-QCAuditEvents.ps1') @auditArgs
