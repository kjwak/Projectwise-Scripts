<#
.SYNOPSIS
Runs periodic QC_Pipeline telemetry retention (audit_events and qc_workflow_events).

.DESCRIPTION
Intended for scheduled Task Scheduler runs. Reads optional defaults from appsettings:

  database.retention.auditEventsDays
  database.retention.auditEventsProcessedOnly
  database.retention.workflowEventsDays
  database.retention.includeCommentTelemetry

CLI parameters override config. Default is preview; pass -ConfirmDeletes to apply.

.EXAMPLE
.\scripts\Invoke-QCDatabaseRetention.ps1
.\scripts\Invoke-QCDatabaseRetention.ps1 -ConfirmDeletes
.\scripts\Invoke-QCDatabaseRetention.ps1 -ConfirmDeletes -AuditEventsDays 60 -WorkflowEventsDays 120
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$AppSettingsPath = '',
    [switch]$ConfirmDeletes,
    [switch]$DryRun,
    [int]$AuditEventsDays = -1,
    [int]$WorkflowEventsDays = -1,
    [switch]$AuditIncludeUnprocessed,
    [switch]$IncludeCommentTelemetry,
    [switch]$SkipAuditEvents,
    [switch]$SkipWorkflowEvents,
    [switch]$Pretty
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force

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
$workflowDays = if ($WorkflowEventsDays -ge 0) { $WorkflowEventsDays } else { _QDR-GetRetentionInt -Map $retention -Name 'workflowEventsDays' -Fallback 180 }
$auditProcessedOnly = -not $AuditIncludeUnprocessed.IsPresent
if (-not $AuditIncludeUnprocessed.IsPresent) {
    $auditProcessedOnly = _QDR-GetRetentionBool -Map $retention -Name 'auditEventsProcessedOnly' -Fallback $true
}
$includeComments = $IncludeCommentTelemetry.IsPresent
if (-not $IncludeCommentTelemetry.IsPresent) {
    $includeComments = _QDR-GetRetentionBool -Map $retention -Name 'includeCommentTelemetry' -Fallback $false
}

function _QDR-InvokeCleanupScript {
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)][hashtable]$BoundParameters
    )
    $argList = [System.Collections.Generic.List[string]]::new()
    foreach ($key in $BoundParameters.Keys) {
        $val = $BoundParameters[$key]
        if ($val -is [switch]) {
            if ($val.IsPresent) { $argList.Add("-$key") }
            continue
        }
        if ($null -eq $val) { continue }
        $argList.Add("-$key")
        $argList.Add([string]$val)
    }
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $ScriptPath @argList
    if ($LASTEXITCODE -and $LASTEXITCODE -ne 0) { throw "Script failed ($ScriptPath) exit code $LASTEXITCODE" }
}

$common = @{
    AppSettingsPath = $AppSettingsPath
}
if ($ConfirmDeletes.IsPresent) { $common['ConfirmDeletes'] = $true }
if ($DryRun.IsPresent) { $common['DryRun'] = $true }

$results = [ordered]@{ audit = $null; workflow = $null }

if (-not $SkipAuditEvents.IsPresent) {
    Write-Host "=== audit_events (older than $auditDays days) ===" -ForegroundColor Cyan
    $auditArgs = @{} + $common
    $auditArgs['OlderThanDays'] = $auditDays
    if ($auditProcessedOnly) { $auditArgs['ProcessedOnly'] = $true }
    else { $auditArgs['IncludeUnprocessed'] = $true }
    if ($Pretty) { $auditArgs['Pretty'] = $true }
    _QDR-InvokeCleanupScript -ScriptPath (Join-Path $PSScriptRoot 'Remove-QCAuditEvents.ps1') -BoundParameters $auditArgs
}

if (-not $SkipWorkflowEvents.IsPresent) {
    Write-Host "=== qc_workflow_events (older than $workflowDays days) ===" -ForegroundColor Cyan
    $wfArgs = @{} + $common
    $wfArgs['OlderThanDays'] = $workflowDays
    if ($includeComments) { $wfArgs['IncludeCommentTelemetry'] = $true }
    if ($Pretty) { $wfArgs['Pretty'] = $true }
    _QDR-InvokeCleanupScript -ScriptPath (Join-Path $PSScriptRoot 'Remove-QCWorkflowEvents.ps1') -BoundParameters $wfArgs
}

if ($Pretty) { $results | ConvertTo-Json -Depth 6 }
