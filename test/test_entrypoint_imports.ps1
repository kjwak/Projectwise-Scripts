# Phase 4F: production entrypoints resolve folder implementation paths; flat shims still work.
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
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

Write-Host '=== Processor entry imports (folder paths) ===' -ForegroundColor Cyan
Import-QCModuleImpl 'Processing\QC.Processors.psm1'
Import-QCModuleImpl 'Queue\QC.Worker.psm1'
Assert-Command 'Invoke-QCPrependProcessor' 'processor map exposes Invoke-QCPrependProcessor'
Assert-Command 'Move-QCJobWithLockRetries' 'worker exposes Move-QCJobWithLockRetries'

Write-Host '=== Dashboard dependencies (folder paths) ===' -ForegroundColor Cyan
Import-QCModuleImpl 'Core\Core.Config.psm1'
Import-QCModuleImpl 'Notifications\QC.WatcherAlerts.psm1'
Assert-Command 'Read-AppConfig' 'dashboard config helper available'
Assert-Command 'Send-QCWatcherSessionLostAlert' 'dashboard watcher alerts available'

Write-Host '=== MCP diagnostic module (folder path) ===' -ForegroundColor Cyan
Import-QCModuleImpl 'Workflow\QC.ProcessType.psm1'
Import-QCModuleImpl 'Diagnostics\QC.DebugMcp.psm1'
Assert-Command 'Initialize-QCDebugMcpContext' 'QC.DebugMcp loads from Diagnostics folder'

Write-Host '=== Flat compatibility shims (legacy paths) ===' -ForegroundColor Cyan
$shimChecks = @(
    @{ Shim = 'Core.Results.psm1'; Cmd = 'New-QCResult' }
    @{ Shim = 'QC.Queue.Json.psm1'; Cmd = 'Get-NextQCJob' }
    @{ Shim = 'QC.JobFactory.psm1'; Cmd = 'New-QCJobObject' }
)
foreach ($check in $shimChecks) {
    $shimPath = Join-Path $modulesRoot $check.Shim
    if (-not (Test-Path -LiteralPath $shimPath)) {
        Write-Host "FAIL: missing shim $($check.Shim)" -ForegroundColor Red
        $fail++
        continue
    }
    Import-Module $shimPath -Force -WarningAction SilentlyContinue | Out-Null
    Assert-Command $check.Cmd "flat shim $($check.Shim) forwards $($check.Cmd)"
}

if ($fail -gt 0) {
    Write-Host "test_entrypoint_imports.ps1 failed: $fail" -ForegroundColor Red
    exit 1
}
Write-Host 'test_entrypoint_imports.ps1 passed' -ForegroundColor Green
