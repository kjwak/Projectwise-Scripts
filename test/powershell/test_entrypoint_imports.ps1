# Phase 4F/4G/4H: production entrypoints resolve folder implementation paths; flat shims removed.
$ErrorActionPreference = 'Stop'
$WarningPreference = 'SilentlyContinue'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$scriptsRoot = Join-Path $repoRoot 'scripts'
$modulesRoot = Join-Path $repoRoot 'modules'
$fail = 0

function Assert-Command($Name, $Msg) {
    if (-not (Get-Command $Name -ErrorAction SilentlyContinue)) {
        Write-Host "FAIL: $Msg ($Name)" -ForegroundColor Red
        $script:fail++
        return $false
    }
    Write-Host "OK:   $Msg" -ForegroundColor Green
    return $true
}

function Import-QCModuleImpl {
    param([Parameter(Mandatory)][string]$RelativePath)
    $path = Join-Path $modulesRoot ($RelativePath -replace '/', '\')
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Missing implementation module: $path"
    }
    Import-Module $path -Force -WarningAction SilentlyContinue | Out-Null
}

. (Join-Path $scriptsRoot 'Restore-QCModuleExports.ps1') -RepoRoot $repoRoot

Write-Host '=== Watcher bootstrap chain (folder paths) ===' -ForegroundColor Cyan
Import-QCModuleImpl 'Core\Core.Results.psm1'
Import-QCModuleImpl 'Core\Core.Runtime.psm1'
Import-QCModuleImpl 'Core\Core.Hashing.psm1'
Import-QCModuleImpl 'Core\Core.Paths.psm1'
Import-QCModuleImpl 'Queue\QC.Filters.psm1'
Import-QCModuleImpl 'Queue\QC.Triggers.psm1'
Import-QCModuleImpl 'Queue\QC.JobFactory.psm1'
Import-QCModuleImpl 'Queue\QC.Queue.Json.psm1'
Import-QCModuleImpl 'Database\Core.Database.psm1'
Import-QCModuleImpl 'ProjectWise\PW.Users.psm1'
Import-QCModuleImpl 'ProjectWise\PW.Discovery.psm1'
Import-QCModuleImpl 'ProjectWise\PW.AuditPoller.psm1'
Import-QCModuleImpl 'Processing\QC.StatusSet.psm1'
Import-QCModuleImpl 'Core\QC.WatcherOrchestration.psm1'
Assert-Command 'Test-QCDatabaseEnabled' 'watcher chain exposes Test-QCDatabaseEnabled'
Assert-Command 'Read-QCAppSettings' 'watcher chain exposes Read-QCAppSettings'
Assert-Command 'New-QCJobObject' 'watcher chain exposes New-QCJobObject'
Assert-Command 'Invoke-AuditTrailScan' 'watcher chain exposes Invoke-AuditTrailScan'

Write-Host '=== Processor bootstrap (shared restore helper) ===' -ForegroundColor Cyan
Import-QCModuleBootstrapSet -FeatureModules @(
    'Core\Core.Results.psm1'
    'Queue\QC.Queue.Json.psm1'
    'Queue\QC.HostThrottle.psm1'
    'Processing\QC.Processors.psm1'
    'Notifications\QC.Notifications.psm1'
    'Processing\QC.Rendition.psm1'
    'Queue\QC.Worker.psm1'
    'Core\Core.Telemetry.psm1'
) -RequiredCommands @(
    'Read-QCAppSettings'
    'Get-NextQCJob'
    'Lock-QCJob'
    'Move-QCJob'
    'Move-QCJobWithLockRetries'
    'Invoke-QCPrependProcessor'
    'Invoke-QCProcessorByType'
    'Write-QCJsonLog'
    'Test-QCDatabaseEnabled'
    'Get-QCHostThrottleClaimDecision'
) -Context 'processor entrypoint test'

Write-Host '=== Dashboard bootstrap (shared restore helper) ===' -ForegroundColor Cyan
Import-QCModuleBootstrapSet -FeatureModules @(
    'Core\Core.Results.psm1'
    'Core\Core.Config.psm1'
    'Queue\QC.Queue.Json.psm1'
    'Core\QC.WatcherOrchestration.psm1'
    'Notifications\QC.WatcherAlerts.psm1'
) -RequiredCommands @(
    'Get-QCAppSettingsConfig'
    'Get-QCTimestamp'
    'Recover-QCStaleJobs'
    'Clear-QCWatcherActive'
    'Read-AppConfig'
    'Send-QCWatcherSessionLostAlert'
    'Send-QCWatcherStallRecoveryAlert'
    'Write-QCWatcherPhaseHeartbeat'
) -Context 'dashboard entrypoint test'

Write-Host '=== run_prepend_qc -NoDashboard bootstrap (shared restore helper) ===' -ForegroundColor Cyan
Import-QCModuleBootstrapSet -FeatureModules @(
    'Core\Core.Results.psm1'
    'Queue\QC.Queue.Json.psm1'
) -RequiredCommands @(
    'Get-QCAppSettingsConfig'
    'Get-QCTimestamp'
    'Get-NextQCJob'
    'Get-QCQueueStats'
    'Invoke-QCQueueStartupCheck'
) -Context 'run_prepend_qc entrypoint test'

Write-Host '=== Processor StatusSet module path (Phase 4E repo root) ===' -ForegroundColor Cyan
$statusSetImpl = Join-Path $modulesRoot 'Processing\QC.StatusSet.psm1'
if (-not (Test-Path -LiteralPath $statusSetImpl)) {
    Write-Host "FAIL: missing StatusSet implementation $statusSetImpl" -ForegroundColor Red
    $fail++
} else {
    Write-Host 'OK:   StatusSet implementation path exists for STATUS_SET_GEN processor' -ForegroundColor Green
}

Write-Host '=== MCP diagnostic module (folder path) ===' -ForegroundColor Cyan
Import-QCModuleImpl 'Workflow\QC.ProcessType.psm1'
Import-QCModuleImpl 'Diagnostics\QC.DebugMcp.psm1'
Assert-Command 'Initialize-QCDebugMcpContext' 'QC.DebugMcp loads from Diagnostics folder'

Write-Host '=== Ops console bootstrap (shared restore helper) ===' -ForegroundColor Cyan
Import-QCModuleBootstrapSet -FeatureModules @(
    'Core\Core.Results.psm1'
    'Core\Core.Runtime.psm1'
    'Queue\QC.Queue.Json.psm1'
    'Ops\QC.OpsConsole.psm1'
) -RequiredCommands @(
    'Get-QCAppSettingsConfig'
    'Get-QCOpsPipelineStatus'
    'Start-QCOpsPipeline'
    'Stop-QCOpsPipeline'
    'Format-QCDisplayTime'
    'ConvertFrom-QCDisplayWallClock'
    'Get-QCWallClockNow'
) -Context 'ops console entrypoint test'

Write-Host '=== Flat compatibility shims removed (Phase 4H) ===' -ForegroundColor Cyan
$removedShims = @('Core.Results.psm1', 'QC.Queue.Json.psm1', 'QC.JobFactory.psm1')
foreach ($shim in $removedShims) {
    $shimPath = Join-Path $modulesRoot $shim
    if (Test-Path -LiteralPath $shimPath) {
        Write-Host "FAIL: flat shim should be removed: $shimPath" -ForegroundColor Red
        $fail++
    } else {
        Write-Host "OK:   flat shim removed: $shim" -ForegroundColor Green
    }
}

if ($fail -gt 0) {
    Write-Host "test_entrypoint_imports.ps1 failed: $fail" -ForegroundColor Red
    exit 1
}
Write-Host 'test_entrypoint_imports.ps1 passed' -ForegroundColor Green
