# Assert Phase 4D maintenance script wrappers forward to scripts/maintenance/ targets.
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$scriptsDir = Join-Path $repoRoot 'scripts'
$maintDir = Join-Path $scriptsDir 'maintenance'

$expected = @(
    'Reset-QCFolderWorkflow.ps1',
    'Purge-QCPendingByFilters.ps1',
    'Requeue-QCJobs.ps1',
    'Repair-QCQueueDuplicates.ps1',
    'Repair-QCDocumentsFolderPaths.ps1',
    'Invoke-QCDatabaseRetention.ps1',
    'Remove-QCAuditEvents.ps1',
    'Remove-QCWorkflowEvents.ps1',
    'Remove-LegacyQcPdfDatabaseRows.ps1',
    'Remove-InvalidSheetIndexRows.ps1',
    'Import-QCJsonlLogsToAutomationEvents.ps1',
    'Sync-QCFolderSheetIndex.ps1',
    'Refresh-SheetIndexStates.ps1',
    'Reconcile-QCSheetOwnership.ps1',
    'Reconcile-QCStatusSets.ps1',
    'Sync-PWUserDirectory.ps1'
)

$fail = 0
foreach ($name in $expected) {
    $wrapperPath = Join-Path $scriptsDir $name
    $targetPath = Join-Path $maintDir $name
    $relTarget = "maintenance\$name"

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

    $tokens = $null
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($wrapperPath, [ref]$tokens, [ref]$errors)
    if ($errors -and $errors.Count -gt 0) {
        Write-Host "FAIL: wrapper parse errors: $wrapperPath" -ForegroundColor Red
        $fail++
        continue
    }

    Write-Host "OK:   $name" -ForegroundColor Green
}

if ($fail -gt 0) {
    Write-Host "Maintenance wrapper tests failed: $fail" -ForegroundColor Red
    exit 1
}

Write-Host "All maintenance script wrapper checks passed ($($expected.Count) scripts)." -ForegroundColor Green
exit 0
