# Phase 4 module bootstrap restore: shared helper and clobber-prone import chains.
$ErrorActionPreference = 'Stop'
$WarningPreference = 'SilentlyContinue'

$repoRoot = Split-Path -Parent $PSScriptRoot
$scriptsRoot = Join-Path $repoRoot 'scripts'
$modulesRoot = Join-Path $repoRoot 'modules'
$fail = 0

function Assert-True($Cond, $Msg) {
    if (-not $Cond) {
        Write-Host "FAIL: $Msg" -ForegroundColor Red
        $script:fail++
        return $false
    }
    Write-Host "OK:   $Msg" -ForegroundColor Green
    return $true
}

. (Join-Path $scriptsRoot 'Restore-QCModuleExports.ps1') -RepoRoot $repoRoot

Write-Host '=== Resolve-QCModulePath uses folder implementations ===' -ForegroundColor Cyan
$coreRuntimePath = Resolve-QCModulePath -RelativePath 'Core.Runtime.psm1'
Assert-True ($coreRuntimePath -like '*\modules\Core\Core.Runtime.psm1') 'bare name maps to Core\Core.Runtime.psm1'
Assert-True (Test-Path -LiteralPath $coreRuntimePath) 'resolved Core.Runtime path exists'
$shimPath = Join-Path $modulesRoot 'Core.Runtime.psm1'
Assert-True ($coreRuntimePath -ne $shimPath) 'resolver does not target flat shim path'

Write-Host '=== Test-QCRequiredCommands error message ===' -ForegroundColor Cyan
$threw = $false
try {
    Test-QCRequiredCommands -Names @('Get-QCTimestamp', 'NotAReal-QCCommand-xyz') -Context 'unit test'
} catch {
    $threw = $true
    Assert-True ($_.Exception.Message -match 'unit test failed') 'throws with context prefix'
    Assert-True ($_.Exception.Message -match 'NotAReal-QCCommand-xyz') 'lists missing command'
}
Assert-True $threw 'Test-QCRequiredCommands throws on missing command'

Write-Host '=== Dashboard clobber chain post-restore guarantee ===' -ForegroundColor Cyan
Import-QCModuleGlobal -RelativePath 'Core\Core.Results.psm1'
Import-QCModuleGlobal -RelativePath 'Core\Core.Config.psm1'
Import-QCModuleGlobal -RelativePath 'Queue\QC.Queue.Json.psm1'
Import-QCModuleGlobal -RelativePath 'Core\QC.WatcherOrchestration.psm1'
Import-QCModuleGlobal -RelativePath 'Notifications\QC.WatcherAlerts.psm1'
Restore-QCFoundationModuleExports
foreach ($cmd in @('Get-QCAppSettingsConfig', 'Get-QCTimestamp', 'Recover-QCStaleJobs', 'Clear-QCWatcherActive')) {
    Assert-True (Get-Command -Name $cmd -ErrorAction SilentlyContinue) "after restore: $cmd"
}

Write-Host '=== Import-QCModuleBootstrapSet (processor chain) ===' -ForegroundColor Cyan
Import-QCModuleBootstrapSet -FeatureModules @(
    'Core\Core.Results.psm1'
    'Queue\QC.Queue.Json.psm1'
    'Processing\QC.Processors.psm1'
    'Queue\QC.Worker.psm1'
    'Core\Core.Telemetry.psm1'
) -RequiredCommands @(
    'Read-QCAppSettings'
    'Get-NextQCJob'
    'Invoke-QCPrependProcessor'
    'Write-QCJsonLog'
    'Test-QCDatabaseEnabled'
) -Context 'processor bootstrap test'

Write-Host '=== Import-QCModuleBootstrapSet (run_prepend_qc chain) ===' -ForegroundColor Cyan
Import-QCModuleBootstrapSet -FeatureModules @(
    'Core\Core.Results.psm1'
    'Queue\QC.Queue.Json.psm1'
) -RequiredCommands @(
    'Get-QCAppSettingsConfig'
    'Get-QCTimestamp'
    'Get-NextQCJob'
    'Invoke-QCQueueStartupCheck'
) -Context 'run_prepend_qc bootstrap test'

if ($fail -gt 0) {
    Write-Host "test_module_bootstrap_restore.ps1 failed: $fail" -ForegroundColor Red
    exit 1
}
Write-Host 'test_module_bootstrap_restore.ps1 passed' -ForegroundColor Green
