$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$modulesRoot = Join-Path $repoRoot 'modules'

$watcherText = Get-Content -LiteralPath (Join-Path $repoRoot 'scripts\service\Watch-QCTrigger.ps1') -Raw
if ($watcherText -notmatch '(?s)\$script:WatchModuleRestoreOrder = @\((.*?)\)') {
    throw 'WatchModuleRestoreOrder not found in Watch-QCTrigger.ps1'
}
$watcherRestore = @([regex]::Matches($Matches[1], "'([^']+)'") | ForEach-Object { $_.Groups[1].Value })
$publicationIndex = [array]::IndexOf($watcherRestore, 'Processing\QC.StatusSetPublication.psm1')
$resultsIndex = [array]::IndexOf($watcherRestore, 'Core\Core.Results.psm1')
if ($publicationIndex -lt 0 -or $resultsIndex -lt 0 -or $publicationIndex -ge $resultsIndex) {
    throw 'QC.StatusSetPublication.psm1 must load before the foundation modules in WatchModuleRestoreOrder'
}

$restoreOrder = @(
    'Processing\QC.StatusSetPublication.psm1'
    'Core\Core.Results.psm1'
    'Core\Core.Paths.psm1'
    'Core\Core.Runtime.psm1'
    'Core\Core.Hashing.psm1'
    'Database\Core.Database.psm1'
    'Notifications\QC.Notifications.psm1'
    'Processing\QC.StatusSet.psm1'
    'ProjectWise\PW.Connection.psm1'
    'ProjectWise\PW.AuditPoller.psm1'
    'ProjectWise\PW.Discovery.psm1'
)
$loadOrder = @(
    'Core\Core.Results.psm1'
    'Core\Core.Paths.psm1'
    'Core\Core.Runtime.psm1'
    'Core\Core.Hashing.psm1'
    'Database\Core.Database.psm1'
    'Queue\QC.Filters.psm1'
    'Queue\QC.Triggers.psm1'
    'Queue\QC.JobFactory.psm1'
    'Queue\QC.Queue.Json.psm1'
    'Notifications\QC.Notifications.psm1'
    'Workflow\QC.Workflow.psm1'
    'Processing\QC.Rendition.psm1'
    'Processing\QC.Processors.psm1'
    'Core\QC.WatcherOrchestration.psm1'
    'Processing\QC.StatusSet.psm1'
    'Processing\QC.StatusSetPublication.psm1'
    'ProjectWise\PW.Connection.psm1'
    'ProjectWise\PW.Users.psm1'
    'ProjectWise\PW.Discovery.psm1'
    'ProjectWise\PW.AuditPoller.psm1'
)
$need = @(
    'Get-Sha256TextHex'
    'Test-QCDatabaseEnabled'
    'Invoke-StatusSetScheduledPublication'
    'Write-QCPollRunTelemetry'
    'Mark-QCAuditEventsProcessed'
    'Write-QCSheetIndex'
    'Write-QCSheetIndexBatch'
    'Invoke-AuditTrailScan'
    'Sync-PWAssociatedSheetWorkflowState'
    'Invoke-QCSheetGroupWorkflowTransition'
    'Test-QCNotificationsEnqueueAsJob'
    'Invoke-QCNotificationForStateChange'
    'Get-QCNotificationDedupeKey'
    'Get-PWObjectPropertyValue'
    'Ensure-PWDiscoveryModuleLoaded'
    'Get-PWCredentialFromFile'
    'Connect-PW'
    'Get-PWImmediateChildFolders'
)

function Restore-Foundation {
    foreach ($m in $restoreOrder) {
        Import-Module (Join-Path $modulesRoot $m) -Force -WarningAction SilentlyContinue | Out-Null
    }
    if (Get-Command -Name 'Ensure-PWDiscoveryModuleLoaded' -ErrorAction SilentlyContinue) {
        [void](Ensure-PWDiscoveryModuleLoaded)
    }
}
function Reload-All {
    foreach ($m in $loadOrder) {
        Import-Module (Join-Path $modulesRoot $m) -Force -WarningAction SilentlyContinue | Out-Null
    }
    if (Get-Command -Name 'Ensure-PWDiscoveryModuleLoaded' -ErrorAction SilentlyContinue) {
        [void](Ensure-PWDiscoveryModuleLoaded)
    }
    Restore-Foundation
}

Import-Module (Join-Path $modulesRoot 'Core\Core.Results.psm1') -Force -WarningAction SilentlyContinue | Out-Null
Import-Module (Join-Path $modulesRoot 'Core\Core.Runtime.psm1') -Force -WarningAction SilentlyContinue | Out-Null
Import-Module (Join-Path $modulesRoot 'Core\Core.Hashing.psm1') -Force -WarningAction SilentlyContinue | Out-Null
Import-Module (Join-Path $modulesRoot 'Database\Core.Database.psm1') -Force -WarningAction SilentlyContinue | Out-Null
Import-Module (Join-Path $modulesRoot 'ProjectWise\PW.Discovery.psm1') -Force -WarningAction SilentlyContinue | Out-Null
Import-Module (Join-Path $modulesRoot 'ProjectWise\PW.AuditPoller.psm1') -Force -WarningAction SilentlyContinue | Out-Null
Import-Module (Join-Path $modulesRoot 'Processing\QC.StatusSet.psm1') -Force -WarningAction SilentlyContinue | Out-Null
Import-Module (Join-Path $modulesRoot 'Core\QC.WatcherOrchestration.psm1') -Force -WarningAction SilentlyContinue | Out-Null
Restore-Foundation
$missing = @($need | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) })
if ($missing.Count -gt 0) { throw ('Still missing after restore: ' + ($missing -join ', ')) }

Reload-All
$missing2 = @($need | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) })
if ($missing2.Count -gt 0) { throw ('Still missing after reload+restore: ' + ($missing2 -join ', ')) }

Write-Host 'test_watch_foundation_restore.ps1 passed'
