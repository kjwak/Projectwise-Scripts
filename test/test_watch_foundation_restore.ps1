$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$modulesRoot = Join-Path $repoRoot 'modules'

$restoreOrder = @(
    'Core.Results.psm1'
    'Core.Paths.psm1'
    'Core.Runtime.psm1'
    'Core.Hashing.psm1'
    'Core.Database.psm1'
    'QC.Notifications.psm1'
    'QC.StatusSet.psm1'
    'PW.Connection.psm1'
    'PW.AuditPoller.psm1'
    'PW.Discovery.psm1'
)
$loadOrder = @(
    'Core.Results.psm1'
    'Core.Paths.psm1'
    'Core.Runtime.psm1'
    'Core.Hashing.psm1'
    'Core.Database.psm1'
    'QC.Filters.psm1'
    'QC.Triggers.psm1'
    'QC.JobFactory.psm1'
    'QC.Queue.Json.psm1'
    'QC.Notifications.psm1'
    'QC.Workflow.psm1'
    'QC.Rendition.psm1'
    'QC.Processors.psm1'
    'QC.WatcherOrchestration.psm1'
    'QC.StatusSet.psm1'
    'PW.Connection.psm1'
    'PW.Users.psm1'
    'PW.Discovery.psm1'
    'PW.AuditPoller.psm1'
)
$need = @(
    'Get-Sha256TextHex'
    'Test-QCDatabaseEnabled'
    'Write-QCPollRunTelemetry'
    'Mark-QCAuditEventsProcessed'
    'Write-QCSheetIndex'
    'Write-QCSheetIndexBatch'
    'Invoke-AuditTrailScan'
    'Sync-PWAssociatedSheetWorkflowState'
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

Import-Module (Join-Path $modulesRoot 'Core.Results.psm1') -Force -WarningAction SilentlyContinue | Out-Null
Import-Module (Join-Path $modulesRoot 'Core.Runtime.psm1') -Force -WarningAction SilentlyContinue | Out-Null
Import-Module (Join-Path $modulesRoot 'Core.Hashing.psm1') -Force -WarningAction SilentlyContinue | Out-Null
Import-Module (Join-Path $modulesRoot 'Core.Database.psm1') -Force -WarningAction SilentlyContinue | Out-Null
Import-Module (Join-Path $modulesRoot 'PW.Discovery.psm1') -Force -WarningAction SilentlyContinue | Out-Null
Import-Module (Join-Path $modulesRoot 'PW.AuditPoller.psm1') -Force -WarningAction SilentlyContinue | Out-Null
Import-Module (Join-Path $modulesRoot 'QC.StatusSet.psm1') -Force -WarningAction SilentlyContinue | Out-Null
Import-Module (Join-Path $modulesRoot 'QC.WatcherOrchestration.psm1') -Force -WarningAction SilentlyContinue | Out-Null
Restore-Foundation
$missing = @($need | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) })
if ($missing.Count -gt 0) { throw ('Still missing after restore: ' + ($missing -join ', ')) }

Reload-All
$missing2 = @($need | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) })
if ($missing2.Count -gt 0) { throw ('Still missing after reload+restore: ' + ($missing2 -join ', ')) }

Write-Host 'test_watch_foundation_restore.ps1 passed'
