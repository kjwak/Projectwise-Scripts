# Assert Phase 4C diagnostic script wrappers forward to scripts/diagnostics/ targets.
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$scriptsDir = Join-Path $repoRoot 'scripts'
$diagDir = Join-Path $scriptsDir 'diagnostics'

$expected = @(
    'Get-PWFolderStateCounts.ps1',
    'Scan-PWProjectMetrics.ps1',
    'Show-QCStatus.ps1',
    'Show-QCQueueDiag.ps1',
    'PW-BrowseFolder.ps1',
    'PW-ListDocsInFolder.ps1',
    'PW-TestWatchRoots.ps1',
    'PW-SmokeTest.ps1',
    'Test-PWConnection.ps1',
    'Test-PWDocumentStateChange.ps1',
    'Test-PWEmailAttributes.ps1',
    'Test-PWEmailAttributes-AttributesBag.ps1',
    'Test-PWEmailAttributes-Caltrans.ps1',
    'Test-PWEmailAttributes-DeepProbe.ps1',
    'Test-PWEmailAttributes-DumpBag.ps1',
    'Test-PWEmailAttributes-EnvCount.ps1',
    'Test-PWEmailAttributes-Extract.ps1',
    'Test-PWEmailAttributes-FolderEnv.ps1',
    'Test-PWEmailAttributes-InspectOne.ps1',
    'Test-PWEmailAttributes-ScanPdfs.ps1',
    'Test-QCEmailTemplate.ps1',
    'Test-QCNotificationGraph.ps1',
    'Test-QCWatcherSessionAlert.ps1'
)

$fail = 0
foreach ($name in $expected) {
    $wrapperPath = Join-Path $scriptsDir $name
    $targetPath = Join-Path $diagDir $name
    $relTarget = "diagnostics\$name"

    if (-not (Test-Path -LiteralPath $wrapperPath)) {
        Write-Host "FAIL: missing wrapper $wrapperPath" -ForegroundColor Red
        $fail++
        continue
    }
    if (-not (Test-Path -LiteralPath $targetPath)) {
        Write-Host "FAIL: missing target $targetPath" -ForegroundColor Red
        $fail++
        continue
    }

    $wrapperText = Get-Content -LiteralPath $wrapperPath -Raw
    if ($wrapperText -notmatch [regex]::Escape($relTarget)) {
        Write-Host "FAIL: wrapper does not reference $relTarget : $wrapperPath" -ForegroundColor Red
        $fail++
        continue
    }

    $resolved = Join-Path $scriptsDir $relTarget
    if (-not (Test-Path -LiteralPath $resolved)) {
        Write-Host "FAIL: wrapper target not resolvable: $resolved" -ForegroundColor Red
        $fail++
        continue
    }

    Write-Host "OK:   $name" -ForegroundColor Green
}

if ($fail -gt 0) {
    Write-Host "Diagnostic wrapper tests failed: $fail" -ForegroundColor Red
    exit 1
}

Write-Host "All diagnostic script wrapper checks passed ($($expected.Count) scripts)." -ForegroundColor Green
exit 0
