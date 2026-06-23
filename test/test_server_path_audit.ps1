# Static in-repo inventory for Phase 4 server path audit.
# Does not contact production servers — validates documented path surfaces exist in the repo.
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
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

Write-Host '=== Root compatibility shims ===' -ForegroundColor Cyan
$rootShims = @(
    @{ Shim = 'Start-QCPipelineDashboard.ps1'; Target = 'scripts\Start-QCPipelineDashboard.ps1' }
    @{ Shim = 'Watch-QCTrigger.ps1'; Target = 'scripts\Watch-QCTrigger.ps1' }
    @{ Shim = 'Run-QCProcessor.ps1'; Target = 'scripts\Run-QCProcessor.ps1' }
    @{ Shim = 'run_prepend_qc.ps1'; Target = 'scripts\run_prepend_qc.ps1' }
)
foreach ($item in $rootShims) {
    $shimPath = Join-Path $repoRoot $item.Shim
    $targetPath = Join-Path $repoRoot $item.Target
    Assert-True (Test-Path -LiteralPath $shimPath) "root shim exists: $($item.Shim)"
    Assert-True (Test-Path -LiteralPath $targetPath) "shim target exists: $($item.Target)"
    $text = Get-Content -LiteralPath $shimPath -Raw
    Assert-True ($text -match [regex]::Escape($item.Target)) "root shim forwards to $($item.Target)"
}

Write-Host '=== Publish-QCPipelineCode copy plan ===' -ForegroundColor Cyan
$publishScript = Join-Path $repoRoot 'scripts\Publish-QCPipelineCode.ps1'
Assert-True (Test-Path -LiteralPath $publishScript) 'Publish-QCPipelineCode.ps1 exists'
$publishSources = @(
    'modules'
    'email'
    'scripts\service'
    'scripts\Restore-QCModuleExports.ps1'
    'scripts\Watch-QCTrigger.ps1'
    'scripts\Run-QCProcessor.ps1'
    'scripts\Import-QCScriptModules.ps1'
    'scripts\Start-QCPipelineDashboard.ps1'
    'scripts\run_prepend_qc.ps1'
    'scripts\Stop-QCPipeline.ps1'
    'scripts\Reset-QCFolderWorkflow.ps1'
)
foreach ($rel in $publishSources) {
    Assert-True (Test-Path -LiteralPath (Join-Path $repoRoot $rel)) "publish source exists: $rel"
}
$publishText = Get-Content -LiteralPath $publishScript -Raw
foreach ($rel in $publishSources) {
    Assert-True ($publishText -match [regex]::Escape($rel)) "publish script references: $rel"
}

Write-Host '=== Publish email assets and config exclusions ===' -ForegroundColor Cyan
Assert-True (Test-Path -LiteralPath (Join-Path $repoRoot 'email\templates\qc_notification.html')) 'email template exists in repo'
if ($publishText -match '(?s)\$copyPlan = @\((.*?)\)\s*\r?\n\r?\nWrite-Host') {
    $copyPlanBlock = $Matches[1]
} else {
    $copyPlanBlock = ''
}
Assert-True ($copyPlanBlock -match "repoRoot 'email'") 'publish copy plan includes email/ directory'
Assert-True ($copyPlanBlock -match 'Restore-QCModuleExports') 'publish copy plan includes Restore-QCModuleExports.ps1'
Assert-True ($copyPlanBlock -notmatch 'appsettings') 'publish copy plan does not include appsettings files'

Write-Host '=== Service spawn path references (implementations) ===' -ForegroundColor Cyan
$dashText = Get-Content -LiteralPath (Join-Path $repoRoot 'scripts\service\Start-QCPipelineDashboard.ps1') -Raw
Assert-True ($dashText -match 'scripts\\service\\Watch-QCTrigger\.ps1') 'dashboard spawns service Watch-QCTrigger'
Assert-True ($dashText -match 'scripts\\service\\Run-QCProcessor\.ps1') 'dashboard spawns service Run-QCProcessor'

$prependText = Get-Content -LiteralPath (Join-Path $repoRoot 'scripts\service\run_prepend_qc.ps1') -Raw
Assert-True ($prependText -match 'scripts\\service\\Start-QCPipelineDashboard\.ps1') 'run_prepend_qc references service dashboard'
Assert-True ($prependText -match 'prepend_qc_on_trigger\.ps1') 'run_prepend_qc -Legacy path documented'

Write-Host '=== Processor default paths ===' -ForegroundColor Cyan
$procText = Get-Content -LiteralPath (Join-Path $repoRoot 'modules\Processing\QC.Processors.psm1') -Raw
Assert-True ($procText -match 'scripts\\processing\\Invoke-QCPrependPw\.ps1') 'default projectWise prepend path'
Assert-True ($procText -match 'legacy\\combine_status_set\.ps1') 'legacy status set fallback path'
Assert-True ($procText -match 'dist\\qc_overlay_prepend\\qc_overlay_prepend\.exe') 'default overlay exe path'

Write-Host '=== Legacy prepend shim ===' -ForegroundColor Cyan
$legacyShim = Join-Path $repoRoot 'legacy\prepend_qc.ps1'
Assert-True (Test-Path -LiteralPath $legacyShim) 'legacy/prepend_qc.ps1 wrapper exists'
$legacyText = Get-Content -LiteralPath $legacyShim -Raw
Assert-True ($legacyText -match 'scripts\\processing\\Invoke-QCPrependPw\.ps1') 'legacy prepend forwards to Invoke-QCPrependPw'

Write-Host '=== No in-repo scheduled task registration ===' -ForegroundColor Cyan
$repoPs1 = Get-ChildItem -Path $repoRoot -Recurse -Filter '*.ps1' -File |
    Where-Object {
        $_.FullName -notmatch '\\\.git\\' -and
        $_.FullName -notmatch '\\test\\' -and
        $_.FullName -notmatch '\\archive\\'
    }
$schedHits = @($repoPs1 | Where-Object {
    $c = Get-Content -LiteralPath $_.FullName -Raw -ErrorAction SilentlyContinue
    $c -and ($c -match 'Register-ScheduledTask|\bschtasks\b')
})
Assert-True ($schedHits.Count -eq 0) 'no Register-ScheduledTask or schtasks in production ps1 files'

if ($fail -gt 0) {
    Write-Host "test_server_path_audit.ps1 failed: $fail" -ForegroundColor Red
    exit 1
}
Write-Host 'test_server_path_audit.ps1 passed' -ForegroundColor Green
