# Phase 4 diagnostic/maintenance bootstrap: shared restore helper usage and clobber-prone chains.
$ErrorActionPreference = 'Stop'
$WarningPreference = 'SilentlyContinue'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$scriptsRoot = Join-Path $repoRoot 'scripts'
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

# Scripts that use QC repo modules and must dot-source Restore-QCModuleExports + Import-QCModuleBootstrapSet.
$qcBootstrapScripts = @(
    @{ Path = 'diagnostics\Show-QCStatus.ps1'; Context = 'Show-QCStatus bootstrap' }
    @{ Path = 'diagnostics\Show-QCQueueDiag.ps1'; Context = 'Show-QCQueueDiag bootstrap' }
    @{ Path = 'diagnostics\Get-PWFolderStateCounts.ps1'; Context = 'Get-PWFolderStateCounts bootstrap' }
    @{ Path = 'diagnostics\Scan-PWProjectMetrics.ps1'; Context = 'Scan-PWProjectMetrics bootstrap' }
    @{ Path = 'diagnostics\PW-BrowseFolder.ps1'; Context = 'PW-BrowseFolder bootstrap' }
    @{ Path = 'diagnostics\PW-ListDocsInFolder.ps1'; Context = 'PW-ListDocsInFolder bootstrap' }
    @{ Path = 'diagnostics\PW-TestWatchRoots.ps1'; Context = 'PW-TestWatchRoots bootstrap' }
    @{ Path = 'diagnostics\Test-PWConnection.ps1'; Context = 'Test-PWConnection bootstrap' }
    @{ Path = 'diagnostics\Test-QCEmailTemplate.ps1'; Context = 'Test-QCEmailTemplate bootstrap' }
    @{ Path = 'diagnostics\Test-QCNotificationGraph.ps1'; Context = 'Test-QCNotificationGraph bootstrap' }
    @{ Path = 'diagnostics\Test-QCWatcherSessionAlert.ps1'; Context = 'Test-QCWatcherSessionAlert bootstrap' }
    @{ Path = 'maintenance\Remove-QCAuditEvents.ps1'; Context = 'Remove-QCAuditEvents bootstrap' }
    @{ Path = 'maintenance\Remove-QCWorkflowEvents.ps1'; Context = 'Remove-QCWorkflowEvents bootstrap' }
    @{ Path = 'maintenance\Remove-LegacyQcPdfDatabaseRows.ps1'; Context = 'Remove-LegacyQcPdfDatabaseRows bootstrap' }
    @{ Path = 'maintenance\Invoke-QCDatabaseRetention.ps1'; Context = 'Invoke-QCDatabaseRetention bootstrap' }
    @{ Path = 'maintenance\Reset-QCFolderWorkflow.ps1'; Context = 'Reset-QCFolderWorkflow bootstrap' }
    @{ Path = 'maintenance\Sync-QCFolderSheetIndex.ps1'; Context = 'Sync-QCFolderSheetIndex bootstrap' }
    @{ Path = 'maintenance\Refresh-SheetIndexStates.ps1'; Context = 'Refresh-SheetIndexStates bootstrap' }
    @{ Path = 'maintenance\Repair-QCDocumentsFolderPaths.ps1'; Context = 'Repair-QCDocumentsFolderPaths bootstrap' }
    @{ Path = 'maintenance\Reconcile-QCSheetOwnership.ps1'; Context = 'Reconcile-QCSheetOwnership bootstrap' }
    @{ Path = 'maintenance\Import-QCJsonlLogsToAutomationEvents.ps1'; Context = 'Import-QCJsonlLogsToAutomationEvents bootstrap' }
    @{ Path = 'maintenance\Purge-QCPendingByFilters.ps1'; Context = 'Purge-QCPendingByFilters bootstrap' }
    @{ Path = 'maintenance\Sync-PWUserDirectory.ps1'; Context = 'Sync-PWUserDirectory bootstrap' }
    @{ Path = 'maintenance\Requeue-QCJobs.ps1'; Context = 'Requeue-QCJobs bootstrap' }
    @{ Path = 'maintenance\Reconcile-QCStatusSets.ps1'; Context = 'Reconcile-QCStatusSets bootstrap' }
    @{ Path = 'maintenance\Remove-InvalidSheetIndexRows.ps1'; Context = 'Remove-InvalidSheetIndexRows bootstrap' }
    @{ Path = 'maintenance\Repair-QCQueueDuplicates.ps1'; Context = 'Repair-QCQueueDuplicates bootstrap' }
    @{ Path = 'processing\Combine-StatusSet.ps1'; Context = 'Combine-StatusSet bootstrap' }
)

# PW-only probes — no QC module bootstrap required.
$pwOnlyScripts = @(
    'diagnostics\Test-PWEmailAttributes.ps1'
    'diagnostics\Test-PWEmailAttributes-AttributesBag.ps1'
    'diagnostics\Test-PWEmailAttributes-Caltrans.ps1'
    'diagnostics\Test-PWEmailAttributes-DeepProbe.ps1'
    'diagnostics\Test-PWEmailAttributes-DumpBag.ps1'
    'diagnostics\Test-PWEmailAttributes-EnvCount.ps1'
    'diagnostics\Test-PWEmailAttributes-Extract.ps1'
    'diagnostics\Test-PWEmailAttributes-FolderEnv.ps1'
    'diagnostics\Test-PWEmailAttributes-InspectOne.ps1'
    'diagnostics\Test-PWEmailAttributes-ScanPdfs.ps1'
    'diagnostics\PW-SmokeTest.ps1'
    'diagnostics\Test-PWDocumentStateChange.ps1'
)

Write-Host '=== QC scripts reference shared bootstrap helper ===' -ForegroundColor Cyan
foreach ($item in $qcBootstrapScripts) {
    $full = Join-Path $scriptsRoot $item.Path
    if (-not (Test-Path -LiteralPath $full)) {
        Assert-True $false "missing script $($item.Path)"
        continue
    }
    $text = Get-Content -LiteralPath $full -Raw
    Assert-True ($text -match 'Restore-QCModuleExports\.ps1') "$($item.Path) dot-sources Restore-QCModuleExports.ps1"
    Assert-True ($text -match 'Import-QCModuleBootstrapSet') "$($item.Path) calls Import-QCModuleBootstrapSet"
    Assert-True ($text -match [regex]::Escape($item.Context)) "$($item.Path) has context $($item.Context)"
    Assert-True ($text -notmatch 'Import-QCScriptModules\.ps1') "$($item.Path) no longer uses Import-QCScriptModules.ps1"
}

Write-Host '=== PW-only scripts exempt from QC bootstrap ===' -ForegroundColor Cyan
foreach ($rel in $pwOnlyScripts) {
    $full = Join-Path $scriptsRoot $rel
    if (-not (Test-Path -LiteralPath $full)) {
        Assert-True $false "missing pw-only script $rel"
        continue
    }
    $text = Get-Content -LiteralPath $full -Raw
    Assert-True ($text -notmatch 'Restore-QCModuleExports\.ps1') "$rel remains PW-only (no restore helper)"
}

. (Join-Path $scriptsRoot 'Restore-QCModuleExports.ps1') -RepoRoot $repoRoot

Write-Host '=== Watcher session alert chain (pre-restore failure) ===' -ForegroundColor Cyan
Import-QCModuleBootstrapSet -FeatureModules @(
    'Notifications\QC.NotificationGraph.psm1'
    'Notifications\QC.WatcherAlerts.psm1'
) -RequiredCommands @(
    'Write-QCJsonLog'
    'Get-QCWatcherSessionAlertSettings'
    'Send-QCWatcherSessionLostAlert'
) -Context 'Test-QCWatcherSessionAlert bootstrap test'

Write-Host '=== DB maintenance chain ===' -ForegroundColor Cyan
Import-QCModuleBootstrapSet -FeatureModules @(
    'Database\Core.Database.psm1'
) -RequiredCommands @(
    'Read-QCAppSettings'
    'Test-QCDatabaseEnabled'
    'Invoke-QCDatabaseScalar'
    'Invoke-QCDatabaseNonQuery'
    'Invoke-QCDatabaseQuery'
) -Context 'DB maintenance bootstrap test'

Write-Host '=== Reset-QCFolderWorkflow module chain ===' -ForegroundColor Cyan
Import-QCModuleBootstrapSet -FeatureModules @(
    'Database\Core.Database.psm1'
    'ProjectWise\PW.Connection.psm1'
    'ProjectWise\PW.Discovery.psm1'
) -RequiredCommands @(
    'Read-QCAppSettings'
    'Test-QCDatabaseEnabled'
    'New-QCDatabaseSession'
) -Context 'Reset-QCFolderWorkflow bootstrap test'
Import-QCModuleGlobal -RelativePath 'Core\Core.Paths.psm1'
Assert-True (Get-Command Normalize-QCDocumentsFolderPath -ErrorAction SilentlyContinue) 'Reset chain: Normalize-QCDocumentsFolderPath after paths post-restore'

Write-Host '=== Combine-StatusSet chain ===' -ForegroundColor Cyan
Import-QCModuleBootstrapSet -FeatureModules @(
    'Processing\QC.StatusSet.psm1'
) -RequiredCommands @(
    'Invoke-StatusSetNativeJob'
) -Context 'Combine-StatusSet bootstrap test'

if ($fail -gt 0) {
    Write-Host "test_diagnostic_maintenance_bootstrap.ps1 failed: $fail" -ForegroundColor Red
    exit 1
}
Write-Host 'test_diagnostic_maintenance_bootstrap.ps1 passed' -ForegroundColor Green
