# Assert Phase 4E flat module shims forward to folder implementations and key exports load.
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
    $wrapperPath = Join-Path $modulesDir $name
    $implPath = Join-Path $modulesDir (Join-Path $folder $name)
    $relTarget = "$folder\$name"

    if (-not (Test-Path -LiteralPath $wrapperPath)) {
        Write-Host "FAIL: missing shim $wrapperPath" -ForegroundColor Red
        $fail++
        continue
    }
    if (-not (Test-Path -LiteralPath $implPath)) {
        Write-Host "FAIL: missing implementation $implPath" -ForegroundColor Red
        $fail++
        continue
    }

    $wrapperText = Get-Content -LiteralPath $wrapperPath -Raw
    if ($wrapperText -notmatch [regex]::Escape($relTarget)) {
        Write-Host "FAIL: shim does not reference $relTarget : $wrapperPath" -ForegroundColor Red
        $fail++
        continue
    }

    Write-Host "OK:   $name -> $relTarget" -ForegroundColor Green
}

if ($fail -gt 0) {
    Write-Host "Module folder shim layout failed: $fail" -ForegroundColor Red
    exit 1
}

# Import smoke via flat shims (offline; no PW/SQL mutations).
$importChecks = @(
    @{ Module = 'Core.Results.psm1'; Command = 'New-QCResult' }
    @{ Module = 'Core.Runtime.psm1'; Command = 'Get-QCTimestamp' }
    @{ Module = 'Core.Database.psm1'; Command = 'Test-QCDatabaseEnabled' }
    @{ Module = 'QC.Queue.Json.psm1'; Command = 'Get-NextQCJob' }
    @{ Module = 'QC.JobFactory.psm1'; Command = 'New-QCJobObject' }
    @{ Module = 'QC.Workflow.psm1'; Command = 'Get-QCWorkflowSettings' }
    @{ Module = 'QC.Notifications.psm1'; Command = 'Get-QCNotificationSettings' }
    @{ Module = 'QC.Processors.psm1'; Command = 'Invoke-QCPrependProcessor' }
    @{ Module = 'PW.Connection.psm1'; Command = 'Ensure-PWDiscoveryCmdlets' }
    @{ Module = 'PW.Discovery.psm1'; Command = 'Get-PWDocName' }
    @{ Module = 'QC.DebugMcp.psm1'; Command = 'Initialize-QCDebugMcpContext' }
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
    Write-Host "Module folder shim import checks failed: $fail" -ForegroundColor Red
    exit 1
}

Write-Host "All module folder shim checks passed ($($folderMap.Count) modules, $($importChecks.Count) import probes)." -ForegroundColor Green
exit 0
