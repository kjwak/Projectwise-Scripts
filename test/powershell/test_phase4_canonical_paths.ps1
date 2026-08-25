# Phase 4H: canonical script/module paths exist; compatibility wrappers and flat shims are removed.
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$scriptsDir = Join-Path $repoRoot 'scripts'
$modulesDir = Join-Path $repoRoot 'modules'
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

$canonicalScripts = @(
    @{ Path = 'scripts\service\Start-QCPipelineDashboard.ps1' }
    @{ Path = 'scripts\service\Watch-QCTrigger.ps1' }
    @{ Path = 'scripts\service\Run-QCProcessor.ps1' }
    @{ Path = 'scripts\service\run_prepend_qc.ps1' }
    @{ Path = 'scripts\service\Stop-QCPipeline.ps1' }
    @{ Path = 'scripts\diagnostics\Show-QCStatus.ps1' }
    @{ Path = 'scripts\maintenance\Reset-QCFolderWorkflow.ps1' }
    @{ Path = 'scripts\processing\Combine-StatusSet.ps1' }
    @{ Path = 'scripts\deployment\Promote-DevToMain.ps1' }
    @{ Path = 'scripts\deployment\Register-QCPipelineDashboardTask.ps1' }
    @{ Path = 'scripts\deployment\Set-QCScheduledTaskOperatorAcl.ps1' }
    @{ Path = 'scripts\deployment\Register-QCPipelineDashboardLogonConsole.ps1' }
    @{ Path = 'scripts\service\Watch-QCPipelineDashboardConsole.ps1' }
    @{ Path = 'scripts\service\Start-QCOpsConsole.ps1' }
    @{ Path = 'scripts\service\Invoke-QCOpsPwCompare.ps1' }
    @{ Path = 'scripts\deployment\Register-QCRemoteWorkerHostTask.ps1' }
)
Write-Host '=== Canonical script paths ===' -ForegroundColor Cyan
foreach ($item in $canonicalScripts) {
    Assert-True (Test-Path -LiteralPath (Join-Path $repoRoot $item.Path)) "canonical exists: $($item.Path)"
}

$removedScriptWrappers = @(
    'Start-QCPipelineDashboard.ps1', 'Watch-QCTrigger.ps1', 'Run-QCProcessor.ps1', 'run_prepend_qc.ps1', 'Stop-QCPipeline.ps1',
    'Show-QCStatus.ps1', 'Reset-QCFolderWorkflow.ps1', 'Combine-StatusSet.ps1', 'Promote-DevToMain.ps1'
)
Write-Host '=== Removed scripts/ wrappers ===' -ForegroundColor Cyan
foreach ($name in $removedScriptWrappers) {
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $scriptsDir $name))) "wrapper removed: scripts\$name"
}

$removedRootShims = @('Start-QCPipelineDashboard.ps1', 'Watch-QCTrigger.ps1', 'Run-QCProcessor.ps1', 'run_prepend_qc.ps1')
Write-Host '=== Removed repo-root shims ===' -ForegroundColor Cyan
foreach ($name in $removedRootShims) {
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $repoRoot $name))) "root shim removed: $name"
}

$removedModuleShims = @(
    'Core.Results.psm1', 'Core.Runtime.psm1', 'QC.Queue.Json.psm1', 'PW.Discovery.psm1', 'QC.Processors.psm1'
)
Write-Host '=== Removed flat module shims ===' -ForegroundColor Cyan
foreach ($name in $removedModuleShims) {
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $modulesDir $name))) "flat shim removed: modules\$name"
}

Write-Host '=== Publish copy plan (canonical only) ===' -ForegroundColor Cyan
$publishScript = Join-Path $scriptsDir 'Publish-QCPipelineCode.ps1'
$publishText = Get-Content -LiteralPath $publishScript -Raw
$publishSources = @(
    'scripts\service'
    'scripts\Restore-QCModuleExports.ps1'
    'scripts\Import-QCScriptModules.ps1'
    'scripts\maintenance\Reset-QCFolderWorkflow.ps1'
)
foreach ($rel in $publishSources) {
    Assert-True (Test-Path -LiteralPath (Join-Path $repoRoot $rel)) "publish source exists: $rel"
    Assert-True ($publishText -match [regex]::Escape($rel)) "publish references: $rel"
}
Assert-True ($publishText -notmatch 'scripts\\Watch-QCTrigger\.ps1') 'publish does not copy Watch wrapper'
Assert-True ($publishText -match 'scripts\\service\\Stop-QCPipeline\.ps1') 'publish restart uses service Stop-QCPipeline'

Write-Host '=== Service spawn paths ===' -ForegroundColor Cyan
$dashText = Get-Content -LiteralPath (Join-Path $repoRoot 'scripts\service\Start-QCPipelineDashboard.ps1') -Raw
Assert-True ($dashText -match 'scripts\\service\\Watch-QCTrigger\.ps1') 'dashboard spawns service watcher'
Assert-True ($dashText -match 'scripts\\service\\Run-QCProcessor\.ps1') 'dashboard spawns service processor'

if ($fail -gt 0) {
    Write-Host "test_phase4_canonical_paths.ps1 failed: $fail" -ForegroundColor Red
    exit 1
}
Write-Host 'test_phase4_canonical_paths.ps1 passed' -ForegroundColor Green
exit 0
