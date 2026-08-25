# Static tests for the logon ops console: status, ops-request, requeue guards, stop exclusions.
$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ("$Actual" -ne "$Expected") { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$modulesRoot = Join-Path $repoRoot 'modules'
Import-Module (Join-Path $modulesRoot 'Core\Core.Results.psm1') -Force
Import-Module (Join-Path $modulesRoot 'Core\Core.Runtime.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $modulesRoot 'Queue\QC.Queue.Json.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $modulesRoot 'Core\QC.WatcherOrchestration.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $modulesRoot 'Ops\QC.OpsConsole.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $modulesRoot 'Core\QC.StatusSetBatching.psm1') -Force -WarningAction SilentlyContinue

Assert-True (Get-Command Get-QCOpsPipelineStatus -ErrorAction SilentlyContinue) 'Get-QCOpsPipelineStatus exported'
Assert-Eq (Get-QCOpsExpectedHost) 'PXBENTLEY01' 'ops console host is PXBENTLEY01'
Assert-True (Test-QCOpsHostAllowed -ComputerName 'PXBENTLEY01') 'PXBENTLEY01 allowed'
Assert-True (-not (Test-QCOpsHostAllowed -ComputerName 'AZTEC002799')) 'modelling host not allowed without -Force'
Assert-True (Test-QCOpsHostAllowed -ComputerName 'AZTEC002799' -Force) '-Force overrides host guard'
Assert-Eq (Get-QCOpsDryRunHeaderText -Policy @{ globalDryRun = $false; sources = @{ qcWorkflowDryRunWriteback = $false; notificationsEnabled = $true; notificationsDryRun = $false } }) 'live' 'global off is live'
Assert-Eq (Get-QCOpsDryRunHeaderText -Policy @{ globalDryRun = $true }) 'ON (global)' 'global dry-run label'

$tempRoot = Join-Path $env:TEMP ('qc-ops-console-' + [guid]::NewGuid().ToString('N'))
$queueRoot = Join-Path $tempRoot 'queue'
New-Item -ItemType Directory -Path (Join-Path $queueRoot 'running') -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $queueRoot 'failed') -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $queueRoot 'pending') -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $queueRoot '_watcher') -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $queueRoot '_logs') -Force | Out-Null

$cfg = @{
    queue = @{ rootDir = $queueRoot }
    database = @{ enabled = $false }
    dryRun = $false
}

try {
    $lockPath = Join-Path $queueRoot '_dashboard.lock'
    @{ pid = 1; host = 'PXBENTLEY01' } | ConvertTo-Json | Set-Content -LiteralPath $lockPath -Encoding utf8
    $lock = Get-QCOpsDashboardLockStatus -QueueRoot $queueRoot
    Assert-True $lock.exists 'fake lock exists'
    Assert-True (-not $lock.alive) 'pid 1 is not a live dashboard'

    $st = Get-QCOpsPipelineStatus -Config $cfg -TaskName 'QC-OpsConsole-Test-DoesNotExist'
    Assert-True $st.IsSuccess 'status result succeeds when task is missing'
    Assert-Eq $st.Data.task.registered $false 'missing task is not registered'
    Assert-True ($st.Data.stateLabel -in @('Not registered', 'Stale lock', 'Stopped', 'Disabled')) ('state label set, got ' + $st.Data.stateLabel)
    Assert-Eq $st.Data.pipelineOn $false 'pipeline is off when task is missing'
    Assert-Eq $st.Data.queueRoot $queueRoot 'status uses fake queue root'

    $missing = Start-QCOpsPipeline -TaskName 'QC-OpsConsole-Test-DoesNotExist'
    Assert-True (-not $missing.IsSuccess) 'start fails when task is not registered'
    Assert-Eq $missing.Code 'OPS_TASK_MISSING' 'start missing-task code'

    $jobPath = Join-Path (Join-Path $queueRoot 'running') 'job-prepend.json'
    @{
        id = 'job-prepend'
        type = 'QC_PREPEND'
        sourceName = 'ea003-rev.pdf'
        checkpoint = 'prepend_complete'
        status = 'running'
        machineName = 'PXBENTLEY01'
    } | ConvertTo-Json | Set-Content -LiteralPath $jobPath -Encoding utf8

    $inFlight = @(Get-QCOpsInFlightPrependJobs -QueueRoot $queueRoot)
    Assert-Eq $inFlight.Count 1 'prepend_complete running job is in-flight'
    $stopBlocked = Stop-QCOpsPipeline -Config $cfg -TaskName 'QC-OpsConsole-Test-DoesNotExist' -SkipProcessKill
    Assert-True (-not $stopBlocked.IsSuccess) 'stop blocked by in-flight prepend'
    Assert-Eq $stopBlocked.Code 'OPS_STOP_IN_FLIGHT_PREPEND' 'in-flight prepend code'
    $stopForced = Stop-QCOpsPipeline -Config $cfg -TaskName 'QC-OpsConsole-Test-DoesNotExist' -Force -SkipProcessKill
    Assert-True $stopForced.IsSuccess 'force stop proceeds without killing GUI'

    $failedPath = Join-Path (Join-Path $queueRoot 'failed') 'job-prepend.json'
    Move-Item -LiteralPath $jobPath -Destination $failedPath -Force
    $requeueSkip = Invoke-QCOpsRequeueJobs -Config $cfg -JobPaths @($failedPath)
    Assert-Eq $requeueSkip.Data.moved 0 'requeue skips prepend_complete'
    Assert-Eq $requeueSkip.Data.skipped 1 'prepend_complete counted as skipped'
    Assert-True (Test-Path -LiteralPath $failedPath) 'skipped job stays in failed'
    $requeueWb = Invoke-QCOpsRequeueJobs -Config $cfg -JobPaths @($failedPath) -WritebackOnly
    Assert-Eq $requeueWb.Data.moved 1 'writeback-only requeue moves prepend_complete'
    Assert-True (Test-Path -LiteralPath (Join-Path (Join-Path $queueRoot 'pending') 'job-prepend.json')) 'job landed in pending'

    $blind = Invoke-QCOpsEnqueueJobType -Config $cfg -JobType 'QC_PREPEND'
    Assert-True (-not $blind.IsSuccess) 'ops console refuses blind QC_PREPEND'
    Assert-Eq $blind.Code 'OPS_NO_BLIND_PREPEND' 'no-blind-prepend code'

    $wmDenied = Set-QCOpsAuditWatermark -Config $cfg -WatermarkUtc ([datetime]::UtcNow) -ConfirmText 'nope'
    Assert-Eq $wmDenied.Code 'OPS_WATERMARK_CONFIRM' 'watermark rewind requires confirm text'

    $req = Set-QCWatcherOpsRequest -Config $cfg -QueueRoot $queueRoot -Action 'fullScan' -RequestedBy 'test'
    Assert-True $req.IsSuccess 'ops-request write succeeds'
    $reqPath = Get-QCWatcherOpsRequestPath -Config $cfg -QueueRoot $queueRoot
    Assert-True (Test-Path -LiteralPath $reqPath) 'ops-request.json exists'
    $plan = Get-QCFullFolderScanReconciliationPlan -Config $cfg -QueueRoot $queueRoot
    Assert-True $plan.due 'full-scan plan is due for ops request'
    Assert-Eq $plan.mode 'ops_request' 'full-scan mode is ops_request'
    Assert-True ([string]$plan.slotKey -like 'ops_request|*') 'ops_request slot key'
    Assert-True (Clear-QCWatcherOpsRequest -Config $cfg -QueueRoot $queueRoot) 'ops-request ack deletes file'
    Assert-True (-not (Test-Path -LiteralPath $reqPath)) 'ops-request.json removed after ack'
    $planAfter = Get-QCFullFolderScanReconciliationPlan -Config $cfg -QueueRoot $queueRoot
    Assert-True (-not $planAfter.due -or $planAfter.mode -ne 'ops_request') 'cleared request is not due as ops_request'

    $stopText = Get-Content -LiteralPath (Join-Path $repoRoot 'scripts\service\Stop-QCPipeline.ps1') -Raw
    Assert-True ($stopText -match 'Start-QCOpsConsole') 'Stop-QCPipeline excludes Start-QCOpsConsole'
    Assert-True ($stopText -match 'Watch-QCPipelineDashboardConsole') 'Stop-QCPipeline excludes logon text console'
    $includeRx = 'Start-QCPipelineDashboard|Start-QCRemoteWorkerHost|Watch-QCTrigger|Run-QCProcessor|run_prepend_qc'
    $excludeRx = 'Start-QCOpsConsole|Watch-QCPipelineDashboardConsole|Invoke-QCOpsPwCompare'
    $guiCmd = 'powershell.exe -STA -File C:\repo\scripts\service\Start-QCOpsConsole.ps1'
    $watchCmd = 'powershell.exe -File C:\repo\scripts\service\Watch-QCTrigger.ps1'
    $dashCmd = 'powershell.exe -File C:\repo\scripts\service\Start-QCPipelineDashboard.ps1'
    Assert-True ($guiCmd -notmatch $includeRx -or $guiCmd -match $excludeRx) 'GUI command is excluded from Stop-QCPipeline'
    Assert-True (($watchCmd -match $includeRx) -and ($watchCmd -notmatch $excludeRx)) 'watcher remains stoppable'
    Assert-True (($dashCmd -match $includeRx) -and ($dashCmd -notmatch $excludeRx)) 'dashboard remains stoppable'

    Assert-Eq (Get-QCOpsPendingJobPriority -JobType 'QC_PREPEND' -PreferJobTypes @('QC_PREPEND', 'STATUS_SET_GEN')) 1 'prepend is priority 1'
    Assert-Eq (Get-QCOpsPendingJobPriority -JobType 'STATUS_SET_GEN' -PreferJobTypes @('QC_PREPEND', 'STATUS_SET_GEN')) 2 'status set is priority 2'
    Assert-Eq (Get-QCOpsPendingJobPriority -JobType 'QC_REPORTING_SCAN' -PreferJobTypes @('QC_PREPEND', 'STATUS_SET_GEN')) 3 'unlisted types sort after prefer list'

    $batchCfg = @{
        queue = @{ rootDir = $queueRoot }
        statusSetBatching = @{
            enabled = $true
            intervalMinutes = 15
            dirtyFolderStorePath = (Join-Path (Join-Path $queueRoot '_watcher') 'statusset-dirty-folders.json')
        }
    }
    $now = [datetime]::SpecifyKind(([datetime]::UtcNow), 'Utc')
    $last = $now.AddMinutes(-5)
    $watcherDir = Join-Path $queueRoot '_watcher'
    New-Item -ItemType Directory -Path $watcherDir -Force | Out-Null
    @{
        version = 1
        lastBatchRunUtc = $last.ToString('o')
        folders = @{
            'documents\proj\cadd\sheets' = @{ folderPath = 'Documents\Proj\CADD\Sheets'; lastSeenUtc = $now.ToString('o') }
        }
    } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $batchCfg.statusSetBatching.dirtyFolderStorePath -Encoding utf8
    $sched = Get-QCStatusSetBatchSchedule -Config $batchCfg -NowUtc $now
    Assert-True $sched.IsSuccess 'batch schedule succeeds'
    Assert-Eq $sched.Data.dirtyFolderCount 1 'one dirty folder'
    Assert-True (-not [bool]$sched.Data.due) 'batch is not due before interval'
    Assert-True ($sched.Data.minutesUntil -ge 9 -and $sched.Data.minutesUntil -le 11) 'about 10 minutes until next batch'
    $batchTxt = Get-QCOpsStatusSetBatchHeaderText -Schedule $sched
    Assert-True ($batchTxt -like 'Status-set batch:*dirty*') 'header mentions dirty folders'
    Assert-True ($batchTxt -like '*in *m*') 'header shows minutes until next run'

    $pendingPath = Join-Path (Join-Path $queueRoot 'pending') 'job-status.json'
    @{
        id = 'job-status'
        type = 'STATUS_SET_GEN'
        sourceName = 'Sheets'
        checkpoint = ''
    } | ConvertTo-Json | Set-Content -LiteralPath $pendingPath -Encoding utf8
    $pendingRows = @(Get-QCOpsRunningJobRows -QueueRoot $queueRoot -State 'pending' -PreferJobTypes @('QC_PREPEND', 'STATUS_SET_GEN'))
    Assert-True ($pendingRows.Count -ge 1) 'pending row exists'
    $ssRow = @($pendingRows | Where-Object { $_.jobId -eq 'job-status' })[0]
    Assert-Eq $ssRow.priority 2 'pending STATUS_SET_GEN row has prefer rank 2'

    $stBatch = Get-QCOpsPipelineStatus -Config $batchCfg -TaskName 'QC-OpsConsole-Test-DoesNotExist'
    Assert-True ($stBatch.Data.statusSetBatchText -like 'Status-set batch:*') 'pipeline status includes status-set batch text'

    $guiPath = Join-Path $repoRoot 'scripts\service\Start-QCOpsConsole.ps1'
    $guiText = Get-Content -LiteralPath $guiPath -Raw
    Assert-True ($guiText -notmatch "(?m)^\s*(Start-Process|&).{0,120}Start-QCPipelineDashboard") 'GUI does not start the dashboard'
    Assert-True ($guiText -match '-STA') 'GUI requests STA'
    Assert-True ($guiText -match 'NoGui') 'GUI keeps text console fallback'
    Assert-True ($guiText -notmatch "_New-Tab 'Hosts'") 'Hosts tab removed'
    Assert-True ($guiText -notmatch 'QC_COMMENT_STATUS_SYNC') 'Runs tab does not enqueue comment sync'
    Assert-True ($guiText -match 'IncludePriority') 'pending queue grid includes priority'
    Assert-True ($guiText -match 'statusSetBatchText') 'header binds status-set batch schedule'
    Assert-True ($guiText -match 'QueueScoreLabels') 'header queue counts use stretching scoreboard'
    Assert-True ($guiText -notmatch 'gridQueueCounts') 'header no longer uses a fixed DataGridView for counts'
    Assert-True ($guiText -match "Source folder") 'SQL filter is source folder'
    Assert-True ($guiText -match 'cmbSqlSourceFolder') 'SQL folder combo exists'
    Assert-True ($guiText -notmatch "Source path") 'SQL filter is not source path'

    $sqlCmd = Get-QCOpsSqlPreviewCommand -TableOrView 'processing_jobs' -SourceFolder 'Documents\Proj\CADD\Sheets' -JobType 'STATUS_SET_GEN'
    Assert-True $sqlCmd.IsSuccess 'preview SQL builds'
    Assert-True ($sqlCmd.Data.sql -match 'source_folder LIKE') 'preview filters source_folder'
    Assert-True ($sqlCmd.Data.sql -notmatch 'source_path LIKE') 'preview does not filter source_path'
    Assert-True ($sqlCmd.Data.parameters.ContainsKey('sourceFolder')) 'sourceFolder parameter set'

    $regPath = Join-Path $repoRoot 'scripts\deployment\Register-QCPipelineDashboardLogonConsole.ps1'
    $regText = Get-Content -LiteralPath $regPath -Raw
    Assert-True ($regText -match 'Start-QCOpsConsole\.ps1') 'logon registrar points at ops GUI'
    Assert-True ($regText -match "-STA") 'logon shortcut launches -STA'
    Assert-True ($regText -match 'Watch-QCPipelineDashboardConsole') 'text console remains fallback'

    $parseFiles = @(
        'scripts\service\Start-QCOpsConsole.ps1'
        'scripts\service\Invoke-QCOpsPwCompare.ps1'
        'modules\Ops\QC.OpsConsole.psm1'
        'scripts\deployment\Register-QCPipelineDashboardLogonConsole.ps1'
    )
    foreach ($rel in $parseFiles) {
        $full = Join-Path $repoRoot $rel
        $tokens = $null
        $errors = $null
        [void][System.Management.Automation.Language.Parser]::ParseFile($full, [ref]$tokens, [ref]$errors)
        if ($errors -and $errors.Count -gt 0) {
            $errors | Select-Object -First 5 | ForEach-Object {
                Write-Host ("PARSE {0} L{1}: {2}" -f $rel, $_.Extent.StartLineNumber, $_.Message) -ForegroundColor Yellow
            }
        }
        Assert-True (-not $errors -or $errors.Count -eq 0) ('parse-check ' + $rel)
    }
}
finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host 'test_ops_console: passed' -ForegroundColor Green
exit 0
