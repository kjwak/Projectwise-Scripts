# Assert Phase 4E folder module implementations exist; flat compatibility shims removed (Phase 4H).
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$modulesDir = Join-Path $repoRoot 'modules'

$folderMap = @{
    'Core.Results.psm1' = 'Core'
    'Core.Runtime.psm1' = 'Core'
    'Core.Paths.psm1' = 'Core'
    'Core.Config.psm1' = 'Core'
    'Core.Logging.psm1' = 'Core'
    'Core.Hashing.psm1' = 'Core'
    'Core.Telemetry.psm1' = 'Core'
    'QC.WatcherOrchestration.psm1' = 'Core'
    'Core.Database.psm1' = 'Database'
    'PW.Connection.psm1' = 'ProjectWise'
    'PW.Discovery.psm1' = 'ProjectWise'
    'PW.AuditPoller.psm1' = 'ProjectWise'
    'PW.Users.psm1' = 'ProjectWise'
    'QC.Workflow.psm1' = 'Workflow'
    'QC.AuditTriggers.psm1' = 'Workflow'
    'QC.ProcessType.psm1' = 'Workflow'
    'QC.Queue.Json.psm1' = 'Queue'
    'QC.JobFactory.psm1' = 'Queue'
    'QC.Worker.psm1' = 'Queue'
    'QC.Filters.psm1' = 'Queue'
    'QC.Triggers.psm1' = 'Queue'
    'QC.Processors.psm1' = 'Processing'
    'QC.StatusSet.psm1' = 'Processing'
    'QC.Rendition.psm1' = 'Processing'
    'QC.ReviewStamp.psm1' = 'Processing'
    'QC.PdfExport.psm1' = 'Processing'
    'QC.CommentExtract.psm1' = 'Processing'
    'QC.CommentStatusDecision.psm1' = 'Processing'
    'QC.CommentStatusProcessor.psm1' = 'Processing'
    'QC.CommentSync.Job.psm1' = 'Processing'
    'QC.CommentSync.State.psm1' = 'Processing'
    'QC.CommentSync.Database.psm1' = 'Processing'
    'QC.CommentSync.Notifications.psm1' = 'Processing'
    'QC.Notifications.psm1' = 'Notifications'
    'QC.NotificationGraph.psm1' = 'Notifications'
    'QC.NotificationMock.psm1' = 'Notifications'
    'QC.NotificationTemplates.psm1' = 'Notifications'
    'QC.NotificationThreads.psm1' = 'Notifications'
    'QC.WatcherAlerts.psm1' = 'Notifications'
    'QC.Reporting.psm1' = 'Reporting'
    'QC.DebugMcp.psm1' = 'Diagnostics'
}

$fail = 0
foreach ($name in ($folderMap.Keys | Sort-Object)) {
    $folder = $folderMap[$name]
    $shimPath = Join-Path $modulesDir $name
    $implPath = Join-Path $modulesDir (Join-Path $folder $name)

    if (Test-Path -LiteralPath $shimPath) {
        Write-Host "FAIL: flat shim should be removed: $shimPath" -ForegroundColor Red
        $fail++
        continue
    }
    if (-not (Test-Path -LiteralPath $implPath)) {
        Write-Host "FAIL: missing implementation $implPath" -ForegroundColor Red
        $fail++
        continue
    }

    Write-Host "OK:   $folder\$name (no flat shim)" -ForegroundColor Green
}

if ($fail -gt 0) {
    Write-Host "Module folder layout failed: $fail" -ForegroundColor Red
    exit 1
}

# Import smoke via folder paths (offline; no PW/SQL mutations).
$importChecks = @(
    @{ Module = 'Core\Core.Results.psm1'; Command = 'New-QCResult' }
    @{ Module = 'Core\Core.Runtime.psm1'; Command = 'Get-QCTimestamp' }
    @{ Module = 'Database\Core.Database.psm1'; Command = 'Test-QCDatabaseEnabled' }
    @{ Module = 'Queue\QC.Queue.Json.psm1'; Command = 'Get-NextQCJob' }
    @{ Module = 'Queue\QC.JobFactory.psm1'; Command = 'New-QCJobObject' }
    @{ Module = 'Workflow\QC.Workflow.psm1'; Command = 'Get-QCWorkflowSettings' }
    @{ Module = 'Notifications\QC.Notifications.psm1'; Command = 'Get-QCNotificationSettings' }
    @{ Module = 'Processing\QC.Processors.psm1'; Command = 'Invoke-QCPrependProcessor' }
    @{ Module = 'ProjectWise\PW.Connection.psm1'; Command = 'Ensure-PWDiscoveryCmdlets' }
    @{ Module = 'ProjectWise\PW.Discovery.psm1'; Command = 'Get-PWDocName' }
    @{ Module = 'Diagnostics\QC.DebugMcp.psm1'; Command = 'Initialize-QCDebugMcpContext' }
)

foreach ($check in $importChecks) {
    $modPath = Join-Path $modulesDir $check.Module
    try {
        Import-Module $modPath -Force -ErrorAction Stop
    } catch {
        Write-Host "FAIL: import $($check.Module): $($_.Exception.Message)" -ForegroundColor Red
        $fail++
        continue
    }
    if (-not (Get-Command $check.Command -ErrorAction SilentlyContinue)) {
        Write-Host "FAIL: command missing after import $($check.Module): $($check.Command)" -ForegroundColor Red
        $fail++
        continue
    }
    Write-Host "OK:   import $($check.Module) -> $($check.Command)" -ForegroundColor Green
}

if ($fail -gt 0) {
    Write-Host "Module folder import checks failed: $fail" -ForegroundColor Red
    exit 1
}

Write-Host "All module folder layout checks passed ($($folderMap.Count) modules, $($importChecks.Count) import probes)." -ForegroundColor Green
exit 0
