# QC.OpsConsole.psm1
# Logon-session control plane for the Session 0 dashboard. Does not start Start-QCPipelineDashboard.ps1.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$modulesRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $modulesRoot 'Core\Core.Results.psm1') -Force
Import-Module (Join-Path $modulesRoot 'Core\Core.Runtime.psm1') -Force
Import-Module (Join-Path $modulesRoot 'Queue\QC.Queue.Json.psm1') -Force
Import-Module (Join-Path $modulesRoot 'Core\QC.WatcherOrchestration.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $modulesRoot 'Database\Core.Database.psm1') -Force -WarningAction SilentlyContinue

$script:QCOpsDefaultTaskName = 'QC-PipelineDashboard'
$script:QCOpsExpectedHost = 'PXBENTLEY01'
$script:QCOpsModellingHost = 'AZTEC002799'

function Get-QCOpsExpectedHost { return $script:QCOpsExpectedHost }
function Get-QCOpsDefaultTaskName { return $script:QCOpsDefaultTaskName }

function Test-QCOpsHostAllowed {
    [CmdletBinding()]
    param(
        [string]$ComputerName = '',
        [switch]$Force
    )
    if ($Force.IsPresent) { return $true }
    $name = $ComputerName
    if ([string]::IsNullOrWhiteSpace($name)) { $name = [string]$env:COMPUTERNAME }
    return ($name -ieq $script:QCOpsExpectedHost)
}

function Get-QCOpsQueueRoot {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)
    $root = $null
    try {
        if ($Config.queue -and $Config.queue.rootDir) { $root = [string]$Config.queue.rootDir }
        elseif ($Config.queue -and $Config.queue.root) { $root = [string]$Config.queue.root }
    } catch { }
    if ([string]::IsNullOrWhiteSpace($root)) {
        $repoRoot = Split-Path -Parent $modulesRoot
        $root = Join-Path $repoRoot 'queue'
    }
    return $root
}

function Get-QCOpsRepoRoot {
    return (Split-Path -Parent $modulesRoot)
}

function Get-QCOpsScheduledTaskState {
    [CmdletBinding()]
    param([string]$TaskName = '')
    if ([string]::IsNullOrWhiteSpace($TaskName)) { $TaskName = $script:QCOpsDefaultTaskName }
    try {
        $t = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
        $info = $null
        try { $info = Get-ScheduledTaskInfo -TaskName $TaskName -ErrorAction SilentlyContinue } catch { }
        $enabled = $true
        try { $enabled = ([string]$t.Settings.Enabled -ne 'False') -and ([bool]$t.Settings.Enabled) } catch { }
        return @{
            name = $TaskName
            state = [string]$t.State
            enabled = [bool]$enabled
            lastRunTime = $(if ($info) { $info.LastRunTime } else { $null })
            lastTaskResult = $(if ($info) { $info.LastTaskResult } else { $null })
            registered = $true
        }
    } catch {
        return @{
            name = $TaskName
            state = 'not registered'
            enabled = $false
            lastRunTime = $null
            lastTaskResult = $null
            registered = $false
        }
    }
}

function Get-QCOpsDashboardLockStatus {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$QueueRoot)
    $lockPath = Join-Path $QueueRoot '_dashboard.lock'
    if (-not (Test-Path -LiteralPath $lockPath)) {
        return @{ exists = $false; alive = $false; pid = 0; text = 'no lock'; host = '' }
    }
    $pl = $null
    try { $pl = Get-Content -LiteralPath $lockPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop } catch { $pl = $null }
    $pidVal = 0
    try { if ($pl -and $pl.pid) { $pidVal = [int]$pl.pid } } catch { $pidVal = 0 }
    $hostName = ''
    try { if ($pl -and $pl.host) { $hostName = [string]$pl.host } } catch { }
    if ($pidVal -le 0) {
        return @{ exists = $true; alive = $false; pid = 0; text = 'lock unreadable'; host = $hostName; path = $lockPath }
    }
    $proc = Get-Process -Id $pidVal -ErrorAction SilentlyContinue
    if (-not $proc) {
        return @{ exists = $true; alive = $false; pid = $pidVal; text = ('stale pid=' + $pidVal); host = $hostName; path = $lockPath }
    }
    $isPw = ($proc.ProcessName -in @('powershell', 'pwsh'))
    if (-not $isPw) {
        return @{ exists = $true; alive = $false; pid = $pidVal; text = ('pid=' + $pidVal + ' not powershell'); host = $hostName; path = $lockPath }
    }
    $cmdOk = $false
    try {
        $p2 = Get-CimInstance Win32_Process -Filter ("ProcessId={0}" -f $pidVal) -ErrorAction SilentlyContinue
        if ($p2 -and $p2.CommandLine) {
            if ([string]$p2.CommandLine -match '(?i)Start-QCPipelineDashboard\.ps1') { $cmdOk = $true }
        }
    } catch { }
    $alive = $isPw -and ($cmdOk -or $true)
    $text = if ($cmdOk) { ('alive pid=' + $pidVal) } else { ('powershell pid=' + $pidVal) }
    return @{ exists = $true; alive = [bool]$alive; pid = $pidVal; text = $text; host = $hostName; path = $lockPath; confirmedDashboard = $cmdOk }
}

function Get-QCOpsLogDir {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$QueueRoot)
    return (Join-Path $QueueRoot '_logs')
}

function _QCOps-ReadJsonlTail {
    param(
        [string]$Path,
        [int]$MaxLines = 80,
        [scriptblock]$Filter
    )
    $out = New-Object System.Collections.Generic.List[object]
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return @() }
    $lines = @(Get-Content -LiteralPath $Path -Tail ($MaxLines * 4) -ErrorAction SilentlyContinue)
    foreach ($line in $lines) {
        $t = ([string]$line).Trim()
        if (-not $t -or -not $t.StartsWith('{')) { continue }
        try {
            $o = $t | ConvertFrom-Json -ErrorAction Stop
            if ($Filter -and -not (& $Filter $o)) { continue }
            [void]$out.Add($o)
        } catch { }
    }
    if ($out.Count -le $MaxLines) { return @($out.ToArray()) }
    return @($out.ToArray() | Select-Object -Last $MaxLines)
}

function Get-QCOpsRecentLogEvents {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$QueueRoot,
        [int]$MaxLines = 40
    )
    $logDir = Get-QCOpsLogDir -QueueRoot $QueueRoot
    $hour = Get-QCLogHourStamp
    $watchPath = Join-Path $logDir ("Watch-QCTrigger_${hour}.jsonl")
    $procPath = Join-Path $logDir ("Run-QCProcessor_${hour}.jsonl")
    $watchFilter = {
        param($Obj)
        $level = ''; $code = ''
        try { $level = [string]$Obj.level } catch { }
        try { $code = [string]$Obj.code } catch { }
        if ($level -match '(?i)error|warning') { return $true }
        if ($code -match '(FAILED|STALE|ALERT|STALL|ACCEPTED|CONNECT_OK|WATCH_START$|ENQUEUE|DASH_WATCHER)') { return $true }
        return $false
    }
    $watch = @(_QCOps-ReadJsonlTail -Path $watchPath -MaxLines $MaxLines -Filter $watchFilter)
    $proc = @(_QCOps-ReadJsonlTail -Path $procPath -MaxLines $MaxLines)
    return @{ watcher = $watch; processor = $proc; watchPath = $watchPath; processorPath = $procPath }
}

function Get-QCOpsJsonlLastCodeTime {
    param([string]$LogDir, [string]$Tag, [string]$CodePattern)
    $hour = Get-QCLogHourStamp
    $path = Join-Path $LogDir ("${Tag}_${hour}.jsonl")
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    $lines = @(Get-Content -LiteralPath $path -Tail 400 -ErrorAction SilentlyContinue)
    [array]::Reverse($lines)
    foreach ($line in $lines) {
        $t = ([string]$line).Trim()
        if (-not $t.StartsWith('{')) { continue }
        try {
            $o = $t | ConvertFrom-Json -ErrorAction Stop
            $code = [string]$o.code
            if ($code -match $CodePattern) {
                $ts = $null
                try { if ($o.ts) { $ts = [string]$o.ts } } catch { }
                return @{ code = $code; ts = $ts; message = [string]$o.message }
            }
        } catch { }
    }
    return $null
}

function Get-QCOpsRunningJobRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$QueueRoot,
        [string]$State = 'running'
    )
    $dir = Join-Path $QueueRoot $State
    $out = New-Object System.Collections.Generic.List[object]
    if (-not (Test-Path -LiteralPath $dir)) { return @() }
    foreach ($f in @(Get-ChildItem -LiteralPath $dir -Filter '*.json' -File -ErrorAction SilentlyContinue)) {
        $j = $null
        try { $j = Get-Content -LiteralPath $f.FullName -Raw | ConvertFrom-Json } catch { $j = $null }
        $id = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
        $type = ''; $src = ''; $machine = ''; $ckpt = ''; $folder = ''
        if ($j) {
            try { if ($j.id) { $id = [string]$j.id } } catch { }
            try { if ($j.type) { $type = [string]$j.type } } catch { }
            try {
                if ($j.sourceName) { $src = [string]$j.sourceName }
                elseif ($j.sourcePath) { $src = [System.IO.Path]::GetFileName([string]$j.sourcePath) }
            } catch { }
            try { if ($j.machineName) { $machine = [string]$j.machineName } } catch { }
            try { if ($j.sourceFolder) { $folder = [string]$j.sourceFolder } } catch { }
            try {
                if ($j.checkpoint) { $ckpt = [string]$j.checkpoint }
                elseif ($j.PSObject.Properties['checkpoint']) { $ckpt = [string]$j.checkpoint }
            } catch { }
        }
        [void]$out.Add([pscustomobject]@{
            state = $State
            jobId = $id
            type = $type
            sourceName = $src
            sourceFolder = $folder
            machineName = $machine
            checkpoint = $ckpt
            path = $f.FullName
            lastWriteTimeUtc = $f.LastWriteTimeUtc
        })
    }
    return @($out.ToArray())
}

function Get-QCOpsInFlightPrependJobs {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$QueueRoot)
    $rows = @(Get-QCOpsRunningJobRows -QueueRoot $QueueRoot -State 'running')
    return @($rows | Where-Object {
        ([string]$_.type -eq 'QC_PREPEND') -and ([string]$_.checkpoint -in @('prepend_complete', 'writeback_running'))
    })
}

function Test-QCOpsCanRequeueJob {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$JobRow,
        [switch]$WritebackOnly
    )
    $type = [string]$JobRow.type
    $ckpt = [string]$JobRow.checkpoint
    if ($type -eq 'QC_PREPEND' -and $ckpt -in @('prepend_complete', 'writeback_running') -and -not $WritebackOnly.IsPresent) {
        return $false
    }
    return $true
}

function Get-QCOpsPipelineStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$TaskName = ''
    )
    if ([string]::IsNullOrWhiteSpace($TaskName)) { $TaskName = $script:QCOpsDefaultTaskName }
    $queueRoot = Get-QCOpsQueueRoot -Config $Config
    $task = Get-QCOpsScheduledTaskState -TaskName $TaskName
    $lock = Get-QCOpsDashboardLockStatus -QueueRoot $queueRoot
    $queue = @{ pending = 0; running = 0; succeeded = 0; failed = 0; locks = 0 }
    try {
        $statsRes = Get-QCQueueStats -Config $Config
        if ($statsRes.IsSuccess -and $statsRes.Data) {
            $st = $statsRes.Data.states
            $queue.pending = [int]$st.pending
            $queue.running = [int]$st.running
            $queue.succeeded = [int]$st.succeeded
            $queue.failed = [int]$st.failed
            if ($statsRes.Data.locks) { $queue.locks = [int]$statsRes.Data.locks.count }
        }
    } catch { }
    $dry = $null
    try { $dry = Get-QCEffectiveDryRunPolicy -Config $Config -Role 'worker' } catch { }
    $sqlEnabled = $false
    try { $sqlEnabled = [bool](Test-QCDatabaseEnabled -Config $Config) } catch { }
    $wm = $null
    $wmAge = $null
    try {
        if (Get-Command Get-QCAuditWatermarkUtc -ErrorAction SilentlyContinue) {
            $wm = Get-QCAuditWatermarkUtc -Config $Config
            if ($wm) { $wmAge = [int](([DateTime]::UtcNow - $wm).TotalSeconds) }
        }
    } catch { }
    $logDir = Get-QCOpsLogDir -QueueRoot $queueRoot
    $pwEvt = Get-QCOpsJsonlLastCodeTime -LogDir $logDir -Tag 'Watch-QCTrigger' -CodePattern 'WATCH_PW_CONNECT_OK|WATCH_PW_SESSION_STALE|WATCH_PW_CONNECT_START'
    $tickEvt = Get-QCOpsJsonlLastCodeTime -LogDir $logDir -Tag 'Watch-QCTrigger' -CodePattern 'WATCH_TICK_START|WATCH_AUDIT_SCAN_DONE'
    $stallEvt = Get-QCOpsJsonlLastCodeTime -LogDir $logDir -Tag 'Watch-QCTrigger' -CodePattern 'DASH_WATCHER_STALL|WATCH_STALL'
    $notifFailEvt = Get-QCOpsJsonlLastCodeTime -LogDir $logDir -Tag 'Run-QCProcessor' -CodePattern 'NOTIF.*FAIL|GRAPH_.*FAIL|NOTIFICATION_FAILED|WATCHER_ALERT'
    $runningRows = @(Get-QCOpsRunningJobRows -QueueRoot $queueRoot -State 'running')
    $localHost = [string]$env:COMPUTERNAME
    $localN = @($runningRows | Where-Object { $_.machineName -and ([string]$_.machineName -ieq $localHost) }).Count
    $remoteN = @($runningRows | Where-Object { $_.machineName -and ([string]$_.machineName -ine $localHost) }).Count
    $pwText = 'unknown'
    if ($pwEvt) {
        if ($pwEvt.code -match 'CONNECT_OK') { $pwText = 'SESSION ACTIVE' }
        elseif ($pwEvt.code -match 'STALE') { $pwText = 'SESSION STALE' }
        elseif ($pwEvt.code -match 'CONNECT_START') { $pwText = 'CONNECTING' }
    } elseif (-not $lock.alive) {
        $pwText = 'NOT RUNNING'
    }
    $pipelineOn = [bool]($task.enabled -and ($task.state -eq 'Running' -or $lock.alive))
    $stateLabel = 'Stopped'
    if ($task.state -eq 'Disabled' -or -not $task.enabled) { $stateLabel = 'Disabled' }
    elseif ($task.state -eq 'Running' -or $lock.alive) { $stateLabel = 'Running' }
    elseif ($lock.exists -and -not $lock.alive) { $stateLabel = 'Stale lock' }
    elseif ($task.registered) { $stateLabel = 'Ready' }
    else { $stateLabel = 'Not registered' }

    return New-QCSuccessResult -Code 'OPS_STATUS' -Message $stateLabel -Data @{
        hostName = $localHost
        hostAllowed = (Test-QCOpsHostAllowed)
        modellingHost = $script:QCOpsModellingHost
        task = $task
        lock = $lock
        pipelineOn = $pipelineOn
        stateLabel = $stateLabel
        queueRoot = $queueRoot
        queue = $queue
        dryRun = $dry
        sqlEnabled = $sqlEnabled
        watermarkUtc = $wm
        watermarkAgeSeconds = $wmAge
        pwText = $pwText
        pwEvent = $pwEvt
        lastTick = $tickEvt
        lastStall = $stallEvt
        lastNotificationFail = $notifFailEvt
        localRunning = $localN
        remoteRunning = $remoteN
        inFlightPrepend = @(Get-QCOpsInFlightPrependJobs -QueueRoot $queueRoot)
        opsRequest = (Get-QCWatcherOpsRequest -Config $Config -QueueRoot $queueRoot)
    }
}

function Start-QCOpsPipeline {
    [CmdletBinding()]
    param(
        [string]$TaskName = '',
        [hashtable]$Config
    )
    if ([string]::IsNullOrWhiteSpace($TaskName)) { $TaskName = $script:QCOpsDefaultTaskName }
    $task = Get-QCOpsScheduledTaskState -TaskName $TaskName
    if (-not $task.registered) {
        return New-QCFailureResult -Code 'OPS_TASK_MISSING' -Message ("Scheduled task '{0}' is not registered." -f $TaskName)
    }
    try {
        Enable-ScheduledTask -TaskName $TaskName -ErrorAction Stop | Out-Null
    } catch {
        return New-QCFailureResult -Code 'OPS_TASK_ACCESS_DENIED' -Message ("Enable-ScheduledTask failed (re-run elevated): {0}" -f $_.Exception.Message)
    }
    try {
        Start-ScheduledTask -TaskName $TaskName -ErrorAction Stop
    } catch {
        return New-QCFailureResult -Code 'OPS_TASK_START_FAILED' -Message ("Start-ScheduledTask failed: {0}" -f $_.Exception.Message)
    }
    return New-QCSuccessResult -Code 'OPS_PIPELINE_STARTED' -Message 'Pipeline task enabled and started.'
}

function Stop-QCOpsPipeline {
    [CmdletBinding()]
    param(
        [hashtable]$Config,
        [string]$TaskName = '',
        [switch]$Force,
        [switch]$SkipProcessKill
    )
    if ([string]::IsNullOrWhiteSpace($TaskName)) { $TaskName = $script:QCOpsDefaultTaskName }
    $queueRoot = $null
    if ($Config) { $queueRoot = Get-QCOpsQueueRoot -Config $Config }
    $inFlight = @()
    if ($queueRoot) { $inFlight = @(Get-QCOpsInFlightPrependJobs -QueueRoot $queueRoot) }
    if (@($inFlight).Count -gt 0 -and -not $Force.IsPresent) {
        $names = @($inFlight | ForEach-Object { $_.sourceName } | Select-Object -First 5) -join ', '
        return New-QCFailureResult -Code 'OPS_STOP_IN_FLIGHT_PREPEND' -Message 'In-flight prepend/writeback jobs are running. Wait or pass -Force.' -Data @{
            jobs = $inFlight
            summary = $names
        }
    }
    $task = Get-QCOpsScheduledTaskState -TaskName $TaskName
    if ($task.registered) {
        try {
            Disable-ScheduledTask -TaskName $TaskName -ErrorAction Stop | Out-Null
        } catch {
            return New-QCFailureResult -Code 'OPS_TASK_ACCESS_DENIED' -Message ("Disable-ScheduledTask failed (re-run elevated): {0}" -f $_.Exception.Message)
        }
        try { Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue } catch { }
    }
    if (-not $SkipProcessKill.IsPresent) {
        $stopScript = Join-Path (Get-QCOpsRepoRoot) 'scripts\service\Stop-QCPipeline.ps1'
        if (Test-Path -LiteralPath $stopScript) {
            $oldConfirm = $ConfirmPreference
            $ConfirmPreference = 'None'
            try {
                & $stopScript -Confirm:$false
            } catch {
                return New-QCFailureResult -Code 'OPS_STOP_KILL_FAILED' -Message $_.Exception.Message
            } finally {
                $ConfirmPreference = $oldConfirm
            }
        }
    }
    return New-QCSuccessResult -Code 'OPS_PIPELINE_STOPPED' -Message 'Pipeline task disabled and processes stopped.' -Data @{ forced = [bool]$Force.IsPresent; inFlightSkipped = @($inFlight).Count }
}

function Request-QCOpsFullScan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$RequestedBy = ''
    )
    return (Set-QCWatcherOpsRequest -Config $Config -Action 'fullScan' -RequestedBy $RequestedBy)
}

function Invoke-QCOpsRequeueJobs {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string[]]$JobPaths,
        [switch]$WritebackOnly
    )
    $moved = 0
    $skipped = 0
    $errors = New-Object System.Collections.Generic.List[string]
    $queueRoot = Get-QCOpsQueueRoot -Config $Config
    $pendingDir = Join-Path $queueRoot 'pending'
    if (-not (Test-Path -LiteralPath $pendingDir)) { New-Item -ItemType Directory -Path $pendingDir -Force | Out-Null }
    foreach ($path in @($JobPaths)) {
        if (-not (Test-Path -LiteralPath $path)) { $skipped++; continue }
        $job = $null
        try { $job = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json } catch { $job = $null }
        if (-not $job) { $skipped++; continue }
        $row = [pscustomobject]@{
            type = [string]$job.type
            checkpoint = $(try { [string]$job.checkpoint } catch { '' })
        }
        if (-not (Test-QCOpsCanRequeueJob -JobRow $row -WritebackOnly:$WritebackOnly)) {
            $skipped++
            continue
        }
        $jobId = [string]$job.id
        if ([string]::IsNullOrWhiteSpace($jobId)) { $jobId = [System.IO.Path]::GetFileNameWithoutExtension($path) }
        $dst = Join-Path $pendingDir ($jobId + '.json')
        if (-not $PSCmdlet.ShouldProcess($path, 'Requeue to pending')) { continue }
        try {
            $job | Add-Member -MemberType NoteProperty -Name 'status' -Value 'pending' -Force
            $ht = @{}
            foreach ($p in $job.PSObject.Properties) { $ht[$p.Name] = $p.Value }
            $ht['status'] = 'pending'
            $ht['updatedAtUtc'] = (Get-QCTimestamp)
            ($ht | ConvertTo-Json -Depth 30) | Set-Content -LiteralPath $path -Encoding utf8
            Move-Item -LiteralPath $path -Destination $dst -Force
            $moved++
        } catch {
            [void]$errors.Add($_.Exception.Message)
        }
    }
    return New-QCSuccessResult -Code 'OPS_REQUEUE' -Message ('Moved {0}, skipped {1}.' -f $moved, $skipped) -Data @{
        moved = $moved
        skipped = $skipped
        errors = @($errors.ToArray())
    }
}

function Invoke-QCOpsRecoverStaleJobs {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)
    return (Recover-QCStaleJobs -Config $Config)
}

function _QCOps-TableToObjects {
    param($Table)
    $list = New-Object System.Collections.Generic.List[object]
    if (-not $Table -or $Table.Rows.Count -eq 0) { return @() }
    foreach ($dataRow in @($Table.Rows)) {
        $h = @{}
        foreach ($col in $Table.Columns) {
            $val = $dataRow[$col.ColumnName]
            if ($val -is [DBNull]) { $h[$col.ColumnName] = $null } else { $h[$col.ColumnName] = $val }
        }
        [void]$list.Add([pscustomobject]$h)
    }
    return @($list.ToArray())
}

function Get-QCOpsLastRuns {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)
    $rows = New-Object System.Collections.Generic.List[object]
    $add = {
        param($Kind, $LastUtc, $Detail)
        [void]$rows.Add([pscustomobject]@{ kind = $Kind; lastUtc = $LastUtc; detail = $Detail })
    }
    if (Test-QCDatabaseEnabled -Config $Config) {
        try {
            $poll = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP (1) started_at, is_reconciliation, jobs_enqueued, error_message
FROM poll_runs
WHERE ISNULL(is_reconciliation, 0) = 0
ORDER BY started_at DESC
"@
            $pollRows = @(_QCOps-TableToObjects -Table $(if ($poll.IsSuccess) { $poll.Data.table } else { $null }))
            if ($pollRows.Count -gt 0) {
                $r = $pollRows[0]
                & $add 'Audit tick' $r.started_at ("enqueued=" + $r.jobs_enqueued)
            }
        } catch { }
        try {
            $full = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP (1) started_at, duration_ms, error_message
FROM poll_runs
WHERE is_reconciliation = 1
ORDER BY started_at DESC
"@
            $fullRows = @(_QCOps-TableToObjects -Table $(if ($full.IsSuccess) { $full.Data.table } else { $null }))
            if ($fullRows.Count -gt 0) {
                $r = $fullRows[0]
                & $add 'Full reconcile' $r.started_at ("duration_ms=" + $r.duration_ms)
            }
        } catch { }
        foreach ($pair in @(
                @{ kind = 'QC_PREPEND'; type = 'QC_PREPEND' }
                @{ kind = 'STATUS_SET_GEN'; type = 'STATUS_SET_GEN' }
                @{ kind = 'QC_NOTIFICATION'; type = 'QC_NOTIFICATION' }
                @{ kind = 'QC_COMMENT_STATUS_SYNC'; type = 'QC_STATE' }
                @{ kind = 'QC_REPORTING_SCAN'; type = 'QC_REPORTING_SCAN' }
            )) {
            try {
                $q = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP (1) job_type, status, completed_at, started_at, created_at, error_code
FROM processing_jobs
WHERE job_type = @t
ORDER BY COALESCE(completed_at, started_at, created_at) DESC
"@ -Parameters @{ t = $pair.type }
                $jobRows = @(_QCOps-TableToObjects -Table $(if ($q.IsSuccess) { $q.Data.table } else { $null }))
                if ($jobRows.Count -gt 0) {
                    $r = $jobRows[0]
                    $when = $r.completed_at
                    if (-not $when) { $when = $r.started_at }
                    if (-not $when) { $when = $r.created_at }
                    & $add $pair.kind $when ("status=" + $r.status + ' ' + $r.error_code)
                }
            } catch { }
        }
        try {
            $n = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP (1) sent_at, event_type, success, document_name, error_message
FROM notification_log
ORDER BY sent_at DESC
"@
            $nRows = @(_QCOps-TableToObjects -Table $(if ($n.IsSuccess) { $n.Data.table } else { $null }))
            if ($nRows.Count -gt 0) {
                $r = $nRows[0]
                & $add 'Notification log' $r.sent_at ("{0} success={1} {2}" -f $r.event_type, $r.success, $r.document_name)
            }
        } catch { }
    }
    $req = Get-QCWatcherOpsRequest -Config $Config
    if ($req) {
        & $add 'Ops full-scan request' $req.requestedAtUtc ('pending by ' + $req.requestedBy)
    }
    return New-QCSuccessResult -Code 'OPS_LAST_RUNS' -Message 'Last-run snapshot.' -Data @{ rows = @($rows.ToArray()) }
}

function Invoke-QCOpsEnqueueJobType {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$JobType,
        [string]$SourcePath = 'ops-console',
        [string]$SourceName = ''
    )
    if ($JobType -eq 'QC_PREPEND') {
        return New-QCFailureResult -Code 'OPS_NO_BLIND_PREPEND' -Message 'Do not enqueue QC_PREPEND from the ops console. Requeue failed jobs or resume writeback via stale recovery.'
    }
    if ([string]::IsNullOrWhiteSpace($SourceName)) { $SourceName = ('manual-' + $JobType) }
    $id = 'ops-' + $JobType.ToLowerInvariant() + '-' + [guid]::NewGuid().ToString('N').Substring(0, 12)
    $job = @{
        id = $id
        type = $JobType
        sourcePath = $SourcePath
        sourceName = $SourceName
        sourceFolder = $SourcePath
        triggerRule = @{ id = 'ops-console'; jobType = $JobType; triggerType = 'manual' }
        dedupeKey = $id
        status = 'pending'
        createdAt = (Get-QCTimestamp)
        attempts = 0
        metadata = @{ requestedBy = [string]$env:USERNAME; source = 'ops-console' }
    }
    return (Add-QCQueueJob -Job $job -Config $Config)
}

function Get-QCOpsSqlTablePreview {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$TableOrView,
        [int]$Top = 50
    )
    if (-not (Test-QCDatabaseEnabled -Config $Config)) {
        return New-QCFailureResult -Code 'OPS_SQL_DISABLED' -Message 'database.enabled is false.'
    }
    $safe = $TableOrView -replace '[^A-Za-z0-9_]', ''
    if ([string]::IsNullOrWhiteSpace($safe)) {
        return New-QCFailureResult -Code 'OPS_SQL_BAD_NAME' -Message 'Invalid table or view name.'
    }
    $sql = "SELECT TOP ($Top) * FROM [$safe]"
    return (Invoke-QCDatabaseQuery -Config $Config -Sql $sql)
}

function Search-QCOpsSheets {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$Query = ''
    )
    if ([string]::IsNullOrWhiteSpace($Query)) {
        return New-QCFailureResult -Code 'OPS_SEARCH_EMPTY' -Message 'Enter a sheet stem, GUID, or folder fragment.'
    }
    if (-not (Test-QCDatabaseEnabled -Config $Config)) {
        return New-QCFailureResult -Code 'OPS_SQL_DISABLED' -Message 'database.enabled is false.'
    }
    $mcp = Join-Path $modulesRoot 'Diagnostics\QC.DebugMcp.psm1'
    Import-Module $mcp -Force
    $appSettings = $null
    try {
        if ($Config.ContainsKey('_appSettingsPath')) { $appSettings = [string]$Config['_appSettingsPath'] }
    } catch { }
    if ([string]::IsNullOrWhiteSpace($appSettings)) {
        $appSettings = Join-Path (Get-QCOpsRepoRoot) 'appsettings.json'
    }
    Initialize-QCDebugMcpContext -AppSettingsPath $appSettings | Out-Null
    $raw = $null
    $guid = $null
    try { $guid = [guid]$Query } catch { $guid = $null }
    try {
        if ($guid) {
            $raw = Search-QCDebugSheet -DocumentGuid $Query
        } else {
            $raw = Search-QCDebugSheet -SheetNumber $Query
        }
    } catch {
        return New-QCFailureResult -Code 'OPS_SEARCH_FAILED' -Message $_.Exception.Message
    }
    return New-QCSuccessResult -Code 'OPS_SEARCH' -Message 'Search finished.' -Data $raw
}

function Get-QCOpsNotificationDedupePath {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)
    $repo = Get-QCOpsRepoRoot
    $p = Join-Path $repo 'notifications\dedupe\sent-keys.jsonl'
    try {
        if ($Config.notifications -and $Config.notifications.dedupePath) {
            $p = [string]$Config.notifications.dedupePath
        }
    } catch { }
    return $p
}

function Get-QCOpsNotificationSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [int]$Tail = 30
    )
    $path = Get-QCOpsNotificationDedupePath -Config $Config
    $keys = @()
    if (Test-Path -LiteralPath $path) {
        $keys = @(Get-Content -LiteralPath $path -Tail $Tail -ErrorAction SilentlyContinue)
    }
    $sqlRows = @()
    if (Test-QCDatabaseEnabled -Config $Config) {
        try {
            $q = Invoke-QCDatabaseQuery -Config $Config -Sql "SELECT TOP ($Tail) sent_at, event_type, success, document_name, error_message FROM notification_log ORDER BY sent_at DESC"
            if ($q.IsSuccess) { $sqlRows = @(_QCOps-TableToObjects -Table $q.Data.table) }
        } catch { }
    }
    return New-QCSuccessResult -Code 'OPS_NOTIF' -Message 'Notification snapshot.' -Data @{
        dedupePath = $path
        recentKeys = $keys
        sqlRows = $sqlRows
    }
}

function Get-QCOpsHostSummary {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)
    $queueRoot = Get-QCOpsQueueRoot -Config $Config
    $running = @(Get-QCOpsRunningJobRows -QueueRoot $queueRoot -State 'running')
    $byHost = @{}
    foreach ($j in $running) {
        $h = [string]$j.machineName
        if ([string]::IsNullOrWhiteSpace($h)) { $h = '(unknown)' }
        if (-not $byHost.ContainsKey($h)) { $byHost[$h] = New-Object System.Collections.Generic.List[object] }
        [void]$byHost[$h].Add($j)
    }
    $throttle = @()
    foreach ($f in @(Get-ChildItem -LiteralPath $queueRoot -Filter '_remote_worker.*.throttle.json' -File -ErrorAction SilentlyContinue)) {
        try {
            $doc = Get-Content -LiteralPath $f.FullName -Raw | ConvertFrom-Json
            $throttle += [pscustomobject]@{ file = $f.Name; doc = $doc }
        } catch { }
    }
    return New-QCSuccessResult -Code 'OPS_HOSTS' -Message 'Host snapshot.' -Data @{
        byHost = $byHost
        throttleFiles = $throttle
        localHost = [string]$env:COMPUTERNAME
        modellingHost = $script:QCOpsModellingHost
    }
}

function Get-QCOpsMaintenanceCatalog {
    [CmdletBinding()]
    param()
    $repo = Get-QCOpsRepoRoot
    $m = Join-Path $repo 'scripts\maintenance'
    return @(
        @{ id = 'reconcile-status-sets'; label = 'Reconcile status sets'; script = (Join-Path $m 'Reconcile-QCStatusSets.ps1'); danger = 'Low'; args = @() }
        @{ id = 'sync-sheet-index'; label = 'Sync sheet index for folder'; script = (Join-Path $m 'Sync-QCFolderSheetIndex.ps1'); danger = 'Low'; args = @(); needsFolder = $true }
        @{ id = 'refresh-sheet-states'; label = 'Refresh sheet_index states from PW'; script = (Join-Path $m 'Refresh-SheetIndexStates.ps1'); danger = 'Medium'; args = @() }
        @{ id = 'repair-folder-paths'; label = 'Repair folder path casing'; script = (Join-Path $m 'Repair-QCDocumentsFolderPaths.ps1'); danger = 'Low'; args = @('-Table', 'all') }
        @{ id = 'db-retention'; label = 'Database retention (audit_events)'; script = (Join-Path $m 'Invoke-QCDatabaseRetention.ps1'); danger = 'Medium'; args = @() }
        @{ id = 'reset-folder-workflow'; label = 'Reset folder workflow + telemetry'; script = (Join-Path $m 'Reset-QCFolderWorkflow.ps1'); danger = 'High'; args = @('-ConfirmReset'); needsFolder = $true }
        @{ id = 'rewind-watermark'; label = 'Rewind audit watermark'; script = '(module) Set-QCOpsAuditWatermark'; danger = 'High'; args = @(); confirmText = 'REWIND WATERMARK' }
    )
}

function Invoke-QCOpsMaintenanceScript {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [string[]]$ArgumentList = @(),
        [string]$AppSettingsPath = '',
        [int]$TimeoutMs = 300000
    )
    if (-not (Test-Path -LiteralPath $ScriptPath)) {
        return New-QCFailureResult -Code 'OPS_MAINT_MISSING' -Message ("Script not found: {0}" -f $ScriptPath)
    }
    $arg = New-Object System.Collections.Generic.List[string]
    [void]$arg.Add('-NoProfile'); [void]$arg.Add('-ExecutionPolicy'); [void]$arg.Add('Bypass')
    [void]$arg.Add('-MTA'); [void]$arg.Add('-File'); [void]$arg.Add($ScriptPath)
    foreach ($a in @($ArgumentList)) { [void]$arg.Add($a) }
    if ($AppSettingsPath -and ($ArgumentList -notcontains '-AppSettingsPath')) {
        [void]$arg.Add('-AppSettingsPath'); [void]$arg.Add($AppSettingsPath)
    }
    $outFile = Join-Path $env:TEMP ('qc-ops-maint-' + [guid]::NewGuid().ToString('N') + '.txt')
    try {
        $p = Start-Process -FilePath 'powershell.exe' -ArgumentList @($arg.ToArray()) -Wait -PassThru -NoNewWindow -RedirectStandardOutput $outFile -RedirectStandardError $outFile
        $text = ''
        if (Test-Path -LiteralPath $outFile) { $text = Get-Content -LiteralPath $outFile -Raw -ErrorAction SilentlyContinue }
        $ok = ($p.ExitCode -eq 0)
        $code = if ($ok) { 'OPS_MAINT_OK' } else { 'OPS_MAINT_FAILED' }
        if ($ok) {
            return New-QCSuccessResult -Code $code -Message 'Maintenance script finished.' -Data @{ exitCode = $p.ExitCode; output = $text; script = $ScriptPath }
        }
        return New-QCFailureResult -Code $code -Message ('Maintenance script exit ' + $p.ExitCode) -Data @{ exitCode = $p.ExitCode; output = $text; script = $ScriptPath }
    } finally {
        Remove-Item -LiteralPath $outFile -Force -ErrorAction SilentlyContinue
    }
}

function Invoke-QCOpsPwCompareChild {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AppSettingsPath,
        [string]$SheetNumber = '',
        [string]$DocumentGuid = ''
    )
    $helper = Join-Path (Get-QCOpsRepoRoot) 'scripts\service\Invoke-QCOpsPwCompare.ps1'
    if (-not (Test-Path -LiteralPath $helper)) {
        return New-QCFailureResult -Code 'OPS_PW_COMPARE_MISSING' -Message 'Invoke-QCOpsPwCompare.ps1 is missing.'
    }
    $outFile = Join-Path $env:TEMP ('qc-ops-pwcompare-' + [guid]::NewGuid().ToString('N') + '.json')
    $arg = @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-MTA', '-File', $helper,
        '-AppSettingsPath', $AppSettingsPath,
        '-OutputPath', $outFile
    )
    if ($SheetNumber) { $arg += @('-SheetNumber', $SheetNumber) }
    if ($DocumentGuid) { $arg += @('-DocumentGuid', $DocumentGuid) }
    try {
        $p = Start-Process -FilePath 'powershell.exe' -ArgumentList $arg -Wait -PassThru -WindowStyle Hidden
        $json = $null
        if (Test-Path -LiteralPath $outFile) {
            $json = Get-Content -LiteralPath $outFile -Raw | ConvertFrom-Json
        }
        if ($p.ExitCode -ne 0 -and -not $json) {
            return New-QCFailureResult -Code 'OPS_PW_COMPARE_FAILED' -Message ('PW compare child exit ' + $p.ExitCode)
        }
        return New-QCSuccessResult -Code 'OPS_PW_COMPARE' -Message 'PW compare finished.' -Data $json
    } finally {
        Remove-Item -LiteralPath $outFile -Force -ErrorAction SilentlyContinue
    }
}

function Set-QCOpsAuditWatermark {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][datetime]$WatermarkUtc,
        [string]$ConfirmText = ''
    )
    if ($ConfirmText -ne 'REWIND WATERMARK') {
        return New-QCFailureResult -Code 'OPS_WATERMARK_CONFIRM' -Message 'Type REWIND WATERMARK to confirm.'
    }
    $ok = Set-QCAuditWatermarkUtc -Config $Config -WatermarkUtc $WatermarkUtc
    if ($ok) {
        return New-QCSuccessResult -Code 'OPS_WATERMARK_SET' -Message 'Watermark updated.' -Data @{ watermarkUtc = $WatermarkUtc }
    }
    return New-QCFailureResult -Code 'OPS_WATERMARK_FAILED' -Message 'Set-QCAuditWatermarkUtc returned false.'
}

Export-ModuleMember -Function *
