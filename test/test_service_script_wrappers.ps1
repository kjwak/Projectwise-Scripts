# Assert Phase 4 service script wrappers forward to scripts/service/ implementations.
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$scriptsDir = Join-Path $repoRoot 'scripts'
$serviceDir = Join-Path $scriptsDir 'service'

$expected = @(
    'Start-QCPipelineDashboard.ps1'
    'Watch-QCTrigger.ps1'
    'Run-QCProcessor.ps1'
    'run_prepend_qc.ps1'
    'Stop-QCPipeline.ps1'
)

$fail = 0
foreach ($name in $expected) {
    $wrapperPath = Join-Path $scriptsDir $name
    $targetPath = Join-Path $serviceDir $name
    $relTarget = "service\$name"

    if (-not (Test-Path -LiteralPath $wrapperPath)) {
        Write-Host "FAIL: missing wrapper $wrapperPath" -ForegroundColor Red
        $fail++
        continue
    }
    if (-not (Test-Path -LiteralPath $targetPath)) {
        Write-Host "FAIL: missing implementation $targetPath" -ForegroundColor Red
        $fail++
        continue
    }

    $wrapperText = Get-Content -LiteralPath $wrapperPath -Raw
    if ($wrapperText -notmatch [regex]::Escape($relTarget)) {
        Write-Host "FAIL: wrapper does not reference $relTarget : $wrapperPath" -ForegroundColor Red
        $fail++
        continue
    }

    $tokens = $null
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($wrapperPath, [ref]$tokens, [ref]$errors)
    if ($errors -and $errors.Count -gt 0) {
        Write-Host "FAIL: wrapper parse errors: $wrapperPath" -ForegroundColor Red
        $fail++
        continue
    }

    Write-Host "OK:   scripts\$name -> $relTarget" -ForegroundColor Green
}

Write-Host '=== Root shims forward to scripts/ wrappers ===' -ForegroundColor Cyan
$rootShims = @(
    @{ Shim = 'Start-QCPipelineDashboard.ps1'; Target = 'scripts\Start-QCPipelineDashboard.ps1' }
    @{ Shim = 'Watch-QCTrigger.ps1'; Target = 'scripts\Watch-QCTrigger.ps1' }
    @{ Shim = 'Run-QCProcessor.ps1'; Target = 'scripts\Run-QCProcessor.ps1' }
    @{ Shim = 'run_prepend_qc.ps1'; Target = 'scripts\run_prepend_qc.ps1' }
)
foreach ($item in $rootShims) {
    $shimPath = Join-Path $repoRoot $item.Shim
    $text = Get-Content -LiteralPath $shimPath -Raw
    if ($text -notmatch [regex]::Escape($item.Target)) {
        Write-Host "FAIL: root shim does not forward to $($item.Target): $shimPath" -ForegroundColor Red
        $fail++
    } else {
        Write-Host "OK:   root\$($item.Shim) -> $($item.Target)" -ForegroundColor Green
    }
}

Write-Host '=== Internal service spawn paths ===' -ForegroundColor Cyan
$dashText = Get-Content -LiteralPath (Join-Path $serviceDir 'Start-QCPipelineDashboard.ps1') -Raw
if ($dashText -notmatch 'scripts\\service\\Watch-QCTrigger\.ps1') {
    Write-Host 'FAIL: dashboard must spawn scripts\service\Watch-QCTrigger.ps1' -ForegroundColor Red
    $fail++
} else { Write-Host 'OK:   dashboard spawns service Watch-QCTrigger' -ForegroundColor Green }
if ($dashText -notmatch 'scripts\\service\\Run-QCProcessor\.ps1') {
    Write-Host 'FAIL: dashboard must spawn scripts\service\Run-QCProcessor.ps1' -ForegroundColor Red
    $fail++
} else { Write-Host 'OK:   dashboard spawns service Run-QCProcessor' -ForegroundColor Green }

$prependText = Get-Content -LiteralPath (Join-Path $serviceDir 'run_prepend_qc.ps1') -Raw
$prependOk = $true
foreach ($needle in @('scripts\\service\\Watch-QCTrigger\.ps1', 'scripts\\service\\Run-QCProcessor\.ps1', 'scripts\\service\\Start-QCPipelineDashboard\.ps1')) {
    if ($prependText -notmatch $needle) {
        Write-Host "FAIL: run_prepend_qc missing spawn path $needle" -ForegroundColor Red
        $fail++
        $prependOk = $false
    }
}
if ($prependOk) { Write-Host 'OK:   run_prepend_qc spawns service entrypoints' -ForegroundColor Green }

Write-Host '=== Stop-QCPipeline process matching ===' -ForegroundColor Cyan
$stopText = Get-Content -LiteralPath (Join-Path $serviceDir 'Stop-QCPipeline.ps1') -Raw
$stopOk = $true
foreach ($pat in @('Start-QCPipelineDashboard', 'Watch-QCTrigger', 'Run-QCProcessor', 'run_prepend_qc')) {
    if ($stopText -notmatch $pat) {
        Write-Host "FAIL: Stop-QCPipeline missing pattern $pat" -ForegroundColor Red
        $fail++
        $stopOk = $false
    }
}
if ($stopOk) { Write-Host 'OK:   Stop-QCPipeline matches service script names' -ForegroundColor Green }

Write-Host '=== Publish copy plan (service + wrappers) ===' -ForegroundColor Cyan
$publishText = Get-Content -LiteralPath (Join-Path $scriptsDir 'Publish-QCPipelineCode.ps1') -Raw
if ($publishText -notmatch 'scripts\\service') {
    Write-Host 'FAIL: Publish must copy scripts\service' -ForegroundColor Red
    $fail++
} else { Write-Host 'OK:   Publish copies scripts\service' -ForegroundColor Green }
if ($publishText -notmatch 'Restore-QCModuleExports') {
    Write-Host 'FAIL: Publish must copy Restore-QCModuleExports.ps1' -ForegroundColor Red
    $fail++
} else { Write-Host 'OK:   Publish copies Restore-QCModuleExports.ps1' -ForegroundColor Green }

if ($fail -gt 0) {
    Write-Host "test_service_script_wrappers.ps1 failed: $fail" -ForegroundColor Red
    exit 1
}
Write-Host "All service script wrapper checks passed ($($expected.Count) scripts)." -ForegroundColor Green
exit 0
