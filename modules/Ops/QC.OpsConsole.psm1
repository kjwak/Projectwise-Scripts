# QC.OpsConsole.psm1
# Logon-session control plane for the Session 0 dashboard. Does not start Start-QCPipelineDashboard.ps1.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$modulesRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $modulesRoot 'Core\Core.Results.psm1') -Force
Import-Module (Join-Path $modulesRoot 'Core\Core.Runtime.psm1') -Force
Import-Module (Join-Path $modulesRoot 'Queue\QC.Queue.Json.psm1') -Force
Import-Module (Join-Path $modulesRoot 'Queue\QC.HostThrottle.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $modulesRoot 'Core\QC.WatcherOrchestration.psm1') -Force -WarningAction SilentlyContinue
Import-Module (Join-Path $modulesRoot 'Core\QC.StatusSetBatching.psm1') -Force -WarningAction SilentlyContinue
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
    return @{ exists = $true; alive = $true; pid = $pidVal; text = ('powershell pid=' + $pidVal); host = $hostName; path = $lockPath; confirmedDashboard = $false }
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
    $tailN = $MaxLines
    if ($Filter) { $tailN = [Math]::Min(800, $MaxLines * 3) }
    $lines = @(Get-Content -LiteralPath $Path -Tail $tailN -ErrorAction SilentlyContinue)
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

function Get-QCOpsQueueDirCounts {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$QueueRoot)
    $q = @{ pending = 0; running = 0; succeeded = 0; failed = 0; locks = 0 }
    foreach ($s in @('pending', 'running', 'succeeded', 'failed')) {
        $dir = Join-Path $QueueRoot $s
        if (-not (Test-Path -LiteralPath $dir)) { continue }
        $q[$s] = @(Get-ChildItem -LiteralPath $dir -Filter '*.json' -File -ErrorAction SilentlyContinue).Count
    }
    $locksDir = Join-Path $QueueRoot 'locks'
    if (Test-Path -LiteralPath $locksDir) {
        $q.locks = @(Get-ChildItem -LiteralPath $locksDir -Filter '*.lock' -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -ne '_queue_write.lock' }).Count
    }
    return $q
}

function Get-QCOpsJsonlHeadlineScan {
    [CmdletBinding()]
    param(
        [string]$Path,
        [hashtable]$Patterns,
        [int]$Tail = 80
    )
    $found = @{}
    foreach ($k in @($Patterns.Keys)) { $found[$k] = $null }
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return $found }
    $remaining = $Patterns.Count
    $lines = @(Get-Content -LiteralPath $Path -Tail $Tail -ErrorAction SilentlyContinue)
    [array]::Reverse($lines)
    foreach ($line in $lines) {
        $t = ([string]$line).Trim()
        if (-not $t.StartsWith('{')) { continue }
        try {
            $o = $t | ConvertFrom-Json -ErrorAction Stop
            $code = [string]$o.code
            foreach ($k in @($Patterns.Keys)) {
                if ($found[$k]) { continue }
                if ($code -match $Patterns[$k]) {
                    $ts = ''
                    try { if ($o.ts) { $ts = [string]$o.ts } } catch { }
                    $found[$k] = @{ code = $code; ts = $ts; message = [string]$o.message }
                    $remaining--
                }
            }
            if ($remaining -le 0) { break }
        } catch { }
    }
    return $found
}

function Get-QCOpsDryRunHeaderText {
    [CmdletBinding()]
    param($Policy)
    if (-not $Policy) { return 'n/a' }
    $global = $false
    try { $global = [bool]$Policy.globalDryRun } catch { }
    if ($global) { return 'ON (global)' }
    $layers = New-Object System.Collections.Generic.List[string]
    $src = $null
    try { $src = $Policy.sources } catch { }
    if ($src) {
        try {
            if ($src.qcWorkflowDryRunWriteback) { [void]$layers.Add('workflow writeback') }
        } catch { }
        try {
            if ($src.notificationsEnabled -and $src.notificationsDryRun) { [void]$layers.Add('notifications') }
        } catch { }
    }
    if ($layers.Count -gt 0) {
        return ('live + {0} dry-run' -f ($layers -join ', '))
    }
    return 'live'
}

function Get-QCOpsJobDisplayName {
    [CmdletBinding()]
    param(
        [string]$Type = '',
        [string]$SourceName = '',
        [string]$SourceFolder = ''
    )
    $src = [string]$SourceName
    $folder = [string]$SourceFolder
    $isFolderJob = ([string]$Type -eq 'STATUS_SET_GEN') -or ($src -eq '_folder_')
    if ($isFolderJob -and -not [string]::IsNullOrWhiteSpace($folder)) { return $folder }
    if (([string]::IsNullOrWhiteSpace($src) -or $src -eq '_folder_') -and -not [string]::IsNullOrWhiteSpace($folder)) {
        return $folder
    }
    return $src
}

function Get-QCOpsHostHealthColorArgb {
    [CmdletBinding()]
    param([string]$Health = '')
    switch ([string]$Health) {
        'healthy' { return @{ R = 25; G = 135; B = 84 } }
        'throttled' { return @{ R = 201; G = 154; B = 22 } }
        'stale' { return @{ R = 220; G = 53; B = 69 } }
        default { return @{ R = 108; G = 117; B = 125 } }
    }
}

function Get-QCOpsLocalHostHealth {
    [CmdletBinding()]
    param(
        $Task,
        $Lock
    )
    $alive = $false
    $exists = $false
    try { if ($Lock) { $alive = [bool]$Lock.alive; $exists = [bool]$Lock.exists } } catch { }
    if ($exists -and -not $alive) { return 'stale' }
    $state = ''
    $enabled = $true
    try { if ($Task) { $state = [string]$Task.state; $enabled = [bool]$Task.enabled } } catch { }
    if ($alive -or $state -eq 'Running') { return 'healthy' }
    if ((-not $enabled) -or $state -eq 'Disabled' -or $state -eq 'not registered') { return 'disabled' }
    return 'disabled'
}

function Get-QCOpsThrottleModeLabel {
    [CmdletBinding()]
    param(
        [string]$Health = '',
        [hashtable]$Status
    )
    switch ([string]$Health) {
        'disabled' { return 'OFF' }
        'stale' {
            if ($Status) { return 'STALE' }
            return 'NONE'
        }
        'throttled' {
            $reason = ''
            try { if ($Status) { $reason = [string]$Status.reason } } catch { }
            if ($reason -eq 'sample_error') { return 'ERROR' }
            $pause = $false
            try { if ($Status) { $pause = [bool]$Status.pauseNewClaims } } catch { }
            if ($pause) { return 'PAUSE' }
            return 'REDUCE'
        }
        default { return 'RUN' }
    }
}

function Get-QCOpsHostStatusLine {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$HostName,
        [int]$Running = 0,
        [string]$Health = 'disabled',
        [hashtable]$Status,
        [switch]$IncludeThrottle
    )
    $bits = New-Object System.Collections.Generic.List[string]
    [void]$bits.Add(('{0}: {1} running' -f $HostName, [int]$Running))
    if ($IncludeThrottle.IsPresent) {
        [void]$bits.Add(('throttle={0}' -f (Get-QCOpsThrottleModeLabel -Health $Health -Status $Status)))
        $slots = $null
        $max = $null
        if ($Status) {
            if ($Status.ContainsKey('recommendedSlots') -and $null -ne $Status.recommendedSlots) {
                try { $slots = [int]$Status.recommendedSlots } catch { $slots = $null }
            }
            if ($Status.ContainsKey('maxParallel') -and $null -ne $Status.maxParallel) {
                try { $max = [int]$Status.maxParallel } catch { $max = $null }
            }
        }
        if ($null -ne $slots) {
            if ($null -ne $max -and $max -gt 0) {
                [void]$bits.Add(('slots={0}/{1}' -f $slots, $max))
            } else {
                [void]$bits.Add(('slots={0}' -f $slots))
            }
        }
        $cpu = $null
        $mem = $null
        if ($Status) {
            if ($Status.ContainsKey('cpuPercent') -and $null -ne $Status.cpuPercent) {
                try { $cpu = [int][math]::Round([double]$Status.cpuPercent) } catch { $cpu = $null }
            }
            if ($Status.ContainsKey('memoryPercent') -and $null -ne $Status.memoryPercent) {
                try { $mem = [int][math]::Round([double]$Status.memoryPercent) } catch { $mem = $null }
            }
        }
        if ($null -ne $cpu) { [void]$bits.Add(('CPU={0}%' -f $cpu)) }
        if ($null -ne $mem) { [void]$bits.Add(('MEM={0}%' -f $mem)) }
    }
    return ($bits -join ' | ')
}

function Get-QCOpsHostStatusRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$QueueRoot,
        [string]$LocalHost = '',
        [string]$ModellingHost = '',
        [int]$LocalRunning = 0,
        [int]$RemoteRunning = 0,
        $Task,
        $Lock,
        [datetime]$NowUtc = [datetime]::MinValue
    )
    if ([string]::IsNullOrWhiteSpace($LocalHost)) { $LocalHost = [string]$env:COMPUTERNAME }
    if ([string]::IsNullOrWhiteSpace($ModellingHost)) { $ModellingHost = $script:QCOpsModellingHost }
    $localHealth = Get-QCOpsLocalHostHealth -Task $Task -Lock $Lock
    $localText = Get-QCOpsHostStatusLine -HostName $LocalHost -Running $LocalRunning -Health $localHealth
    $rows = New-Object System.Collections.Generic.List[object]
    [void]$rows.Add([pscustomobject]@{
        hostName = $LocalHost
        role = 'local'
        running = [int]$LocalRunning
        health = $localHealth
        text = $localText
    })

    $status = $null
    try {
        if (Get-Command Read-QCHostThrottleStatus -ErrorAction SilentlyContinue) {
            $thPath = Get-QCHostThrottleStatusPath -QueueRoot $QueueRoot -HostName $ModellingHost
            $status = Read-QCHostThrottleStatus -Path $thPath
        }
    } catch { $status = $null }
    $sampleSeconds = 10
    if ($status -and $status.ContainsKey('sampleSeconds') -and $null -ne $status.sampleSeconds) {
        try { $sampleSeconds = [int]$status.sampleSeconds } catch { $sampleSeconds = 10 }
    }
    $remoteHealth = 'stale'
    try {
        if (Get-Command Get-QCHostThrottleHealth -ErrorAction SilentlyContinue) {
            $remoteHealth = Get-QCHostThrottleHealth -Status $status -SampleSeconds $sampleSeconds -NowUtc $NowUtc
        } elseif (-not $status) {
            $remoteHealth = 'stale'
        }
    } catch { $remoteHealth = 'stale' }
    $remoteText = Get-QCOpsHostStatusLine -HostName $ModellingHost -Running $RemoteRunning -Health $remoteHealth -Status $status -IncludeThrottle
    [void]$rows.Add([pscustomobject]@{
        hostName = $ModellingHost
        role = 'remote'
        running = [int]$RemoteRunning
        health = $remoteHealth
        text = $remoteText
    })
    return @($rows.ToArray())
}

function Get-QCOpsRecentLogEvents {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$QueueRoot,
        [int]$MaxLines = 80,
        [int]$MaxFiles = 6,
        [switch]$Unfiltered
    )
    $logDir = Get-QCOpsLogDir -QueueRoot $QueueRoot
    $watchFilter = {
        param($Obj)
        $level = ''; $code = ''
        try { $level = [string]$Obj.level } catch { }
        try { $code = [string]$Obj.code } catch { }
        if ($level -match '(?i)error|warning') { return $true }
        if ($code -match '(FAILED|STALE|ALERT|STALL|ACCEPTED|CONNECT_OK|WATCH_START$|ENQUEUE|DASH_WATCHER)') { return $true }
        return $false
    }
    $events = New-Object System.Collections.Generic.List[object]
    $watch = New-Object System.Collections.Generic.List[object]
    $proc = New-Object System.Collections.Generic.List[object]
    foreach ($pair in @(
            @{ tag = 'Watch-QCTrigger'; source = 'watcher' }
            @{ tag = 'Run-QCProcessor'; source = 'processor' }
        )) {
        $files = @()
        if (Test-Path -LiteralPath $logDir) {
            $files = @(Get-ChildItem -LiteralPath $logDir -Filter ($pair.tag + '_*.jsonl') -File -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTimeUtc -Descending |
                Select-Object -First $MaxFiles)
        }
        $filter = $null
        if (-not $Unfiltered.IsPresent -and $pair.source -eq 'watcher') { $filter = $watchFilter }
        foreach ($f in $files) {
            foreach ($o in @(_QCOps-ReadJsonlTail -Path $f.FullName -MaxLines $MaxLines -Filter $filter)) {
                try { $o | Add-Member -NotePropertyName 'logSource' -NotePropertyValue $pair.source -Force } catch { }
                try { $o | Add-Member -NotePropertyName 'logFile' -NotePropertyValue $f.Name -Force } catch { }
                [void]$events.Add($o)
                if ($pair.source -eq 'watcher') { [void]$watch.Add($o) } else { [void]$proc.Add($o) }
            }
        }
    }
    return @{
        events = @($events.ToArray())
        watcher = @($watch.ToArray())
        processor = @($proc.ToArray())
        logDir = $logDir
    }
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

function Get-QCOpsPreferJobTypes {
    [CmdletBinding()]
    param([hashtable]$Config)
    $prefer = New-Object System.Collections.Generic.List[string]
    try {
        $sel = $null
        if ($Config -and $Config.queue) {
            $q = $Config.queue
            if ($q -is [hashtable] -and $q.ContainsKey('selection')) { $sel = $q.selection }
            elseif ($q.PSObject.Properties['selection']) { $sel = $q.selection }
        }
        $raw = $null
        if ($sel -is [hashtable] -and $sel.ContainsKey('preferJobTypes')) { $raw = $sel.preferJobTypes }
        elseif ($sel -and $sel.PSObject.Properties['preferJobTypes']) { $raw = $sel.preferJobTypes }
        foreach ($t in @($raw)) {
            $s = ([string]$t).Trim()
            if ($s) { [void]$prefer.Add($s) }
        }
    } catch { }
    return @($prefer.ToArray())
}

function Get-QCOpsPendingJobPriority {
    [CmdletBinding()]
    param(
        [string]$JobType,
        [string[]]$PreferJobTypes = @()
    )
    $list = @($PreferJobTypes | Where-Object { $_ })
    if ($list.Count -eq 0) { return 1 }
    $jt = ([string]$JobType).Trim()
    for ($i = 0; $i -lt $list.Count; $i++) {
        if ([string]$list[$i] -eq $jt) { return ($i + 1) }
    }
    return ($list.Count + 1)
}

function Get-QCOpsStatusSetBatchHeaderText {
    [CmdletBinding()]
    param($Schedule)
    $d = $null
    try {
        if ($Schedule -and $Schedule.PSObject.Properties['IsSuccess'] -and $Schedule.IsSuccess) { $d = $Schedule.Data }
        elseif ($Schedule -and $Schedule.PSObject.Properties['enabled']) { $d = $Schedule }
        elseif ($Schedule -is [hashtable]) { $d = $Schedule }
    } catch { }
    if (-not $d) { return 'Status-set batch: n/a' }
    $enabled = $true
    try { $enabled = [bool]$d.enabled } catch { }
    $interval = 15
    try { $interval = [int]$d.intervalMinutes } catch { }
    $dirty = 0
    try { $dirty = [int]$d.dirtyFolderCount } catch { }
    if (-not $enabled) { return 'Status-set batch: off' }
    $nextLocal = ''
    try {
        if ($d.nextBatchUtc) {
            $nextLocal = Format-QCDisplayTime -Value $d.nextBatchUtc -Format 'yyyy-MM-dd HH:mm'
        }
    } catch { }
    $due = $false
    try { $due = [bool]$d.due } catch { }
    if ($dirty -le 0 -and -not $due) {
        if ($nextLocal) {
            return ('Status-set batch: idle (0 dirty); next {0} AZ; every {1}m' -f $nextLocal, $interval)
        }
        return ('Status-set batch: idle (0 dirty); every {0}m' -f $interval)
    }
    if ($due) {
        return ('Status-set batch: {0} dirty folder(s); due now (every {1}m)' -f $dirty, $interval)
    }
    $inMin = $null
    try { if ($null -ne $d.minutesUntil) { $inMin = [int]$d.minutesUntil } } catch { }
    if ($nextLocal -and $null -ne $inMin) {
        return ('Status-set batch: {0} dirty folder(s); next {1} AZ (in {2}m; every {3}m)' -f $dirty, $nextLocal, $inMin, $interval)
    }
    if ($nextLocal) {
        return ('Status-set batch: {0} dirty folder(s); next {1} AZ; every {2}m' -f $dirty, $nextLocal, $interval)
    }
    return ('Status-set batch: {0} dirty folder(s); every {1}m' -f $dirty, $interval)
}

function Get-QCOpsFolderKey {
    [CmdletBinding()]
    param([string]$FolderPath = '')
    if ([string]::IsNullOrWhiteSpace($FolderPath)) { return '' }
    if (Get-Command Normalize-QCDocumentsFolderPath -ErrorAction SilentlyContinue) {
        try {
            $r = Normalize-QCDocumentsFolderPath -Path $FolderPath
            if ($r -and $r.IsSuccess -and $r.Data -and $r.Data.path) { return [string]$r.Data.path }
        } catch { }
    }
    return ([string]$FolderPath).Trim().TrimEnd('\', '/').Replace('/', '\').ToLowerInvariant()
}

function Get-QCOpsDeferredStatusSetCheckpointText {
    [CmdletBinding()]
    param($Schedule)
    $d = $null
    try {
        if ($Schedule -and $Schedule.PSObject.Properties['IsSuccess'] -and $Schedule.IsSuccess) { $d = $Schedule.Data }
        elseif ($Schedule -and $Schedule.PSObject.Properties['enabled']) { $d = $Schedule }
        elseif ($Schedule -and $Schedule.Data) { $d = $Schedule.Data }
    } catch { $d = $null }
    if (-not $d) { return 'wait (batch)' }
    $due = $false
    try { $due = [bool]$d.due } catch { }
    if ($due) { return 'wait (due)' }
    $inMin = $null
    try { if ($null -ne $d.minutesUntil) { $inMin = [int]$d.minutesUntil } } catch { }
    if ($null -ne $inMin) { return ('wait {0}m' -f $inMin) }
    return 'wait (batch)'
}

function Get-QCOpsDeferredStatusSetRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$QueueRoot = '',
        [string[]]$PreferJobTypes = @(),
        [string[]]$ExcludeFolderKeys = @()
    )
    $out = New-Object System.Collections.Generic.List[object]
    if (-not (Get-Command Get-QCStatusSetDirtyFolders -ErrorAction SilentlyContinue)) { return @() }
    $dirtyRes = $null
    try { $dirtyRes = Get-QCStatusSetDirtyFolders -Config $Config } catch { return @() }
    if (-not $dirtyRes -or -not $dirtyRes.IsSuccess) { return @() }
    if (-not [bool]$dirtyRes.Data.enabled) { return @() }
    $sched = $null
    try { $sched = Get-QCStatusSetBatchSchedule -Config $Config } catch { $sched = $null }
    $ckpt = Get-QCOpsDeferredStatusSetCheckpointText -Schedule $sched
    $pri = Get-QCOpsPendingJobPriority -JobType 'STATUS_SET_GEN' -PreferJobTypes $PreferJobTypes
    $exclude = @{}
    foreach ($k in @($ExcludeFolderKeys)) {
        $nk = Get-QCOpsFolderKey -FolderPath $k
        if ($nk) { $exclude[$nk] = $true }
    }
    foreach ($f in @($dirtyRes.Data.folders)) {
        $folder = [string]$f.folderPath
        $key = [string]$f.folderKey
        if ([string]::IsNullOrWhiteSpace($key)) { $key = Get-QCOpsFolderKey -FolderPath $folder }
        if ($key -and $exclude.ContainsKey($key)) { continue }
        $when = $null
        foreach ($raw in @($f.lastSeenUtc, $f.firstSeenUtc)) {
            if ([string]::IsNullOrWhiteSpace([string]$raw)) { continue }
            try {
                $when = [datetimeoffset]::Parse([string]$raw, [System.Globalization.CultureInfo]::InvariantCulture).UtcDateTime
                break
            } catch {
                try { $when = [datetime]::Parse([string]$raw, $null, [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal).ToUniversalTime(); break } catch { }
            }
        }
        if (-not $when) { $when = [datetime]::UtcNow }
        $leaf = [System.IO.Path]::GetFileName($folder.Trim().TrimEnd('\', '/'))
        if ([string]::IsNullOrWhiteSpace($leaf)) { $leaf = 'status-set' }
        [void]$out.Add([pscustomobject]@{
            state = 'pending'
            jobId = $leaf
            type = 'STATUS_SET_GEN'
            priority = $pri
            sourceName = $folder
            sourceFolder = $folder
            machineName = ''
            checkpoint = $ckpt
            locked = $false
            deferred = $true
            path = ''
            lastWriteTimeUtc = $when
            lastWriteTime = (Format-QCDisplayTime -Value $when)
        })
    }
    return @($out.ToArray())
}

function Get-QCOpsRunningJobRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$QueueRoot,
        [string]$State = 'running',
        [int]$Limit = 0,
        [string[]]$PreferJobTypes = @(),
        [hashtable]$Config,
        [switch]$IncludeDeferredStatusSet
    )
    $dir = Join-Path $QueueRoot $State
    $out = New-Object System.Collections.Generic.List[object]
    if (Test-Path -LiteralPath $dir) {
        $files = @(Get-ChildItem -LiteralPath $dir -Filter '*.json' -File -ErrorAction SilentlyContinue)
        if ($Limit -gt 0 -and $State -ne 'pending' -and $files.Count -gt $Limit) {
            $files = @($files | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First $Limit)
        }
        foreach ($f in $files) {
            $j = $null
            try { $j = Get-Content -LiteralPath $f.FullName -Raw | ConvertFrom-Json } catch { $j = $null }
            $id = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
            $type = ''; $src = ''; $machine = ''; $ckpt = ''; $folder = ''
            if ($j) {
                try { if ($j.id) { $id = [string]$j.id } } catch { }
                try { if ($j.type) { $type = [string]$j.type } } catch { }
                try { if ($j.sourceFolder) { $folder = [string]$j.sourceFolder } } catch { }
                try {
                    $rawName = ''
                    if ($j.sourceName) { $rawName = [string]$j.sourceName }
                    elseif ($j.sourcePath) { $rawName = [System.IO.Path]::GetFileName([string]$j.sourcePath) }
                    $src = Get-QCOpsJobDisplayName -Type $type -SourceName $rawName -SourceFolder $folder
                } catch { }
                try { if ($j.machineName) { $machine = [string]$j.machineName } } catch { }
                try {
                    if ($j.checkpoint) { $ckpt = [string]$j.checkpoint }
                    elseif ($j.PSObject.Properties['checkpoint']) { $ckpt = [string]$j.checkpoint }
                } catch { }
            }
            $pri = 0
            if ($State -eq 'pending') {
                $pri = Get-QCOpsPendingJobPriority -JobType $type -PreferJobTypes $PreferJobTypes
            }
            $locked = $false
            try {
                $lockPath = Join-Path (Join-Path $QueueRoot 'locks') ($id + '.lock')
                $locked = Test-Path -LiteralPath $lockPath
            } catch { }
            [void]$out.Add([pscustomobject]@{
                state = $State
                jobId = $id
                type = $type
                priority = $pri
                sourceName = $src
                sourceFolder = $folder
                machineName = $machine
                checkpoint = $ckpt
                locked = [bool]$locked
                deferred = $false
                path = $f.FullName
                lastWriteTimeUtc = $f.LastWriteTimeUtc
                lastWriteTime = (Format-QCDisplayTime -Value $f.LastWriteTimeUtc)
            })
        }
    }
    if ($State -eq 'pending' -and $IncludeDeferredStatusSet.IsPresent -and $Config) {
        $exclude = New-Object System.Collections.Generic.List[string]
        foreach ($row in $out) {
            if ([string]$row.type -eq 'STATUS_SET_GEN' -and $row.sourceFolder) {
                [void]$exclude.Add([string]$row.sourceFolder)
            }
        }
        $runDir = Join-Path $QueueRoot 'running'
        if (Test-Path -LiteralPath $runDir) {
            foreach ($rf in @(Get-ChildItem -LiteralPath $runDir -Filter '*.json' -File -ErrorAction SilentlyContinue)) {
                $rj = $null
                try { $rj = Get-Content -LiteralPath $rf.FullName -Raw | ConvertFrom-Json } catch { $rj = $null }
                if (-not $rj) { continue }
                $rt = ''
                $rfolder = ''
                try { if ($rj.type) { $rt = [string]$rj.type } } catch { }
                try { if ($rj.sourceFolder) { $rfolder = [string]$rj.sourceFolder } } catch { }
                if ($rt -eq 'STATUS_SET_GEN' -and $rfolder) { [void]$exclude.Add($rfolder) }
            }
        }
        foreach ($drow in @(Get-QCOpsDeferredStatusSetRows -Config $Config -QueueRoot $QueueRoot -PreferJobTypes $PreferJobTypes -ExcludeFolderKeys @($exclude.ToArray()))) {
            [void]$out.Add($drow)
        }
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
    $deferred = $false
    try { $deferred = [bool]$JobRow.deferred } catch { }
    if ($deferred) { return $false }
    if ($type -eq 'QC_PREPEND' -and $ckpt -in @('prepend_complete', 'writeback_running') -and -not $WritebackOnly.IsPresent) {
        return $false
    }
    return $true
}

function Get-QCOpsPipelineStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$TaskName = '',
        [switch]$Light
    )
    if ([string]::IsNullOrWhiteSpace($TaskName)) { $TaskName = $script:QCOpsDefaultTaskName }
    $queueRoot = Get-QCOpsQueueRoot -Config $Config
    $task = Get-QCOpsScheduledTaskState -TaskName $TaskName
    $lock = Get-QCOpsDashboardLockStatus -QueueRoot $queueRoot
    $queue = @{ pending = 0; running = 0; succeeded = 0; failed = 0; locks = 0 }
    if (-not $Light.IsPresent) {
        try { $queue = Get-QCOpsQueueDirCounts -QueueRoot $queueRoot } catch { }
    }
    $dry = $null
    try { $dry = Get-QCEffectiveDryRunPolicy -Config $Config -Role 'worker' } catch { }
    $sqlEnabled = $false
    try { $sqlEnabled = [bool](Test-QCDatabaseEnabled -Config $Config) } catch { }
    $wm = $null
    $wmAge = $null
    $pwEvt = $null
    $tickEvt = $null
    $stallEvt = $null
    $notifFailEvt = $null
    if (-not $Light.IsPresent) {
        try {
            if (Get-Command Get-QCAuditWatermarkUtc -ErrorAction SilentlyContinue) {
                $wm = Get-QCAuditWatermarkUtc -Config $Config
                if ($wm) { $wmAge = [int](([DateTime]::UtcNow - $wm).TotalSeconds) }
            }
        } catch { }
        $logDir = Get-QCOpsLogDir -QueueRoot $queueRoot
        $hour = Get-QCLogHourStamp
        $watchPath = Join-Path $logDir ("Watch-QCTrigger_${hour}.jsonl")
        $procPath = Join-Path $logDir ("Run-QCProcessor_${hour}.jsonl")
        $watchHit = Get-QCOpsJsonlHeadlineScan -Path $watchPath -Tail 80 -Patterns @{
            pw = 'WATCH_PW_CONNECT_OK|WATCH_PW_SESSION_STALE|WATCH_PW_CONNECT_START'
            tick = 'WATCH_TICK_START|WATCH_AUDIT_SCAN_DONE'
            stall = 'DASH_WATCHER_STALL|WATCH_STALL'
        }
        $procHit = Get-QCOpsJsonlHeadlineScan -Path $procPath -Tail 80 -Patterns @{
            notifFail = 'NOTIF.*FAIL|GRAPH_.*FAIL|NOTIFICATION_FAILED|WATCHER_ALERT'
        }
        $pwEvt = $watchHit.pw
        $tickEvt = $watchHit.tick
        $stallEvt = $watchHit.stall
        $notifFailEvt = $procHit.notifFail
    }
    $runningRows = @(Get-QCOpsRunningJobRows -QueueRoot $queueRoot -State 'running' -Limit 50)
    $localHost = [string]$env:COMPUTERNAME
    $modellingHost = $script:QCOpsModellingHost
    $localRows = @($runningRows | Where-Object { $_.machineName -and ([string]$_.machineName -ieq $localHost) })
    $remoteRows = @($runningRows | Where-Object { $_.machineName -and ([string]$_.machineName -ieq $modellingHost) })
    $localN = $localRows.Count
    $remoteN = $remoteRows.Count
    $remoteJobNames = @($remoteRows | ForEach-Object { $_.sourceName } | Where-Object { $_ } | Select-Object -First 3)
    $hostRows = @(Get-QCOpsHostStatusRows -QueueRoot $queueRoot -LocalHost $localHost -ModellingHost $modellingHost `
        -LocalRunning $localN -RemoteRunning $remoteN -Task $task -Lock $lock)
    $throttleSummary = @($hostRows | ForEach-Object { $_.text }) -join ' | '
    $pwText = 'unknown'
    if ($Light.IsPresent) {
        $pwText = $(if ($lock.alive) { 'SESSION ?' } else { 'NOT RUNNING' })
    } elseif ($pwEvt) {
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
    $inFlight = @($runningRows | Where-Object {
            ([string]$_.type -eq 'QC_PREPEND') -and ([string]$_.checkpoint -in @('prepend_complete', 'writeback_running'))
        })
    $opsReq = $null
    if (-not $Light.IsPresent) {
        $opsReq = (Get-QCWatcherOpsRequest -Config $Config -QueueRoot $queueRoot)
    }
    $statusSetBatch = $null
    try {
        if (Get-Command Get-QCStatusSetBatchSchedule -ErrorAction SilentlyContinue) {
            $statusSetBatch = Get-QCStatusSetBatchSchedule -Config $Config
        }
    } catch { }

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
        dryRunText = (Get-QCOpsDryRunHeaderText -Policy $dry)
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
        remoteJobNames = $remoteJobNames
        hostRows = $hostRows
        throttleSummary = $throttleSummary
        inFlightPrepend = $inFlight
        opsRequest = $opsReq
        statusSetBatch = $statusSetBatch
        statusSetBatchText = (Get-QCOpsStatusSetBatchHeaderText -Schedule $statusSetBatch)
        light = [bool]$Light.IsPresent
    }
}

function Test-QCOpsAccessDeniedMessage {
    param([string]$Message)
    return [bool]($Message -match 'Access is denied|UnauthorizedAccess|0x80070005|E_ACCESSDENIED')
}

function Invoke-QCOpsScheduledTaskCom {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TaskName,
        [Parameter(Mandatory)][ValidateSet('Enable', 'Disable', 'Start', 'Stop')][string]$Action
    )
    # CIM Enable/Disable/Start-ScheduledTask ignore the task DACL and demand admin.
    # IRegisteredTask honors the Full Control ACE granted by Set-QCScheduledTaskOperatorAcl.
    $svc = New-Object -ComObject 'Schedule.Service'
    $svc.Connect()
    $folder = $svc.GetFolder('\')
    $task = $null
    try { $task = $folder.GetTask($TaskName) } catch { $task = $null }
    if (-not $task) {
        throw ("Scheduled task '{0}' was not found." -f $TaskName)
    }
    switch ($Action) {
        'Enable' { $task.Enabled = $true }
        'Disable' { $task.Enabled = $false }
        'Start' {
            try {
                [void]$task.Run($null)
            } catch {
                $err = [string]$_.Exception.Message
                if ($err -notmatch 'already running|0x8004131F|0x41301') { throw }
            }
        }
        'Stop' { $task.Stop(0) }
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
        Invoke-QCOpsScheduledTaskCom -TaskName $TaskName -Action Enable
    } catch {
        $err = [string]$_.Exception.Message
        if (Test-QCOpsAccessDeniedMessage -Message $err) {
            return New-QCFailureResult -Code 'OPS_TASK_ACCESS_DENIED' -Message (
                "Enable task access denied. Once, elevated: .\scripts\deployment\Register-QCPipelineDashboardTask.ps1 -GrantOperatorAccess"
            )
        }
        return New-QCFailureResult -Code 'OPS_TASK_ACCESS_DENIED' -Message ("Enable task failed: {0}" -f $err)
    }
    try {
        Invoke-QCOpsScheduledTaskCom -TaskName $TaskName -Action Start
    } catch {
        $err = [string]$_.Exception.Message
        if (Test-QCOpsAccessDeniedMessage -Message $err) {
            return New-QCFailureResult -Code 'OPS_TASK_ACCESS_DENIED' -Message (
                "Start task access denied. Once, elevated: .\scripts\deployment\Register-QCPipelineDashboardTask.ps1 -GrantOperatorAccess"
            )
        }
        return New-QCFailureResult -Code 'OPS_TASK_START_FAILED' -Message ("Start task failed: {0}" -f $err)
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
            Invoke-QCOpsScheduledTaskCom -TaskName $TaskName -Action Disable
        } catch {
            $err = [string]$_.Exception.Message
            if (Test-QCOpsAccessDeniedMessage -Message $err) {
                return New-QCFailureResult -Code 'OPS_TASK_ACCESS_DENIED' -Message (
                    "Disable task access denied. Once, elevated: .\scripts\deployment\Register-QCPipelineDashboardTask.ps1 -GrantOperatorAccess"
                )
            }
            return New-QCFailureResult -Code 'OPS_TASK_ACCESS_DENIED' -Message ("Disable task failed: {0}" -f $err)
        }
        try { Invoke-QCOpsScheduledTaskCom -TaskName $TaskName -Action Stop } catch { }
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
        [void]$rows.Add([pscustomobject]@{ kind = $Kind; lastTime = (Format-QCDisplayTime -Value $LastUtc); detail = $Detail })
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

function Get-QCOpsSqlTimeRangeChoices {
    return @(
        'Last 15 min'
        'Last 1 hour'
        'Last 6 hours'
        'Last 24 hours'
        'Last 7 days'
        'Last 30 days'
        'All'
        'Custom'
    )
}

function Get-QCOpsSqlJobTypeChoices {
    return @(
        'All'
        'QC_PREPEND'
        'STATUS_SET_GEN'
        'QC_RENDITION'
        'QC_COMMENT_STATUS_SYNC'
        'QC_REPORTING_SCAN'
        'QC_NOTIFICATION'
        'QC_STATE'
    )
}

function Resolve-QCOpsSqlTimeRange {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Range,
        [datetime]$CustomFrom,
        [datetime]$CustomTo,
        [datetime]$NowUtc = [datetime]::UtcNow
    )
    $now = $NowUtc
    if ($now.Kind -ne [DateTimeKind]::Utc) { $now = $now.ToUniversalTime() }
    $toUtc = $now
    $fromUtc = $null
    $unbounded = $false
    switch -Regex ($Range.Trim()) {
        '^Last 15 min$' { $fromUtc = $now.AddMinutes(-15) }
        '^Last 1 hour$' { $fromUtc = $now.AddHours(-1) }
        '^Last 6 hours$' { $fromUtc = $now.AddHours(-6) }
        '^Last 24 hours$' { $fromUtc = $now.AddHours(-24) }
        '^Last 7 days$' { $fromUtc = $now.AddDays(-7) }
        '^Last 30 days$' { $fromUtc = $now.AddDays(-30) }
        '^All$' { $unbounded = $true }
        '^Custom$' {
            if ($CustomFrom -eq [datetime]::MinValue -or $CustomTo -eq [datetime]::MinValue) {
                return New-QCFailureResult -Code 'OPS_SQL_CUSTOM_RANGE' -Message 'Pick a custom from and to datetime.'
            }
            $fromUtc = ConvertFrom-QCDisplayWallClock -WallClock $CustomFrom
            $toUtc = ConvertFrom-QCDisplayWallClock -WallClock $CustomTo
            if ($fromUtc -gt $toUtc) {
                return New-QCFailureResult -Code 'OPS_SQL_RANGE_ORDER' -Message 'Custom range From must be earlier than To.'
            }
        }
        default {
            return New-QCFailureResult -Code 'OPS_SQL_BAD_RANGE' -Message ('Unknown time range: ' + $Range)
        }
    }
    return New-QCSuccessResult -Code 'OPS_SQL_RANGE' -Message 'Time range resolved.' -Data @{
        range = $Range
        fromUtc = $fromUtc
        toUtc = $(if ($unbounded) { $null } else { $toUtc })
        unbounded = $unbounded
    }
}

function Get-QCOpsSqlPreviewCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TableOrView,
        [int]$Top = 80,
        [AllowNull()]$FromUtc = $null,
        [AllowNull()]$ToUtc = $null,
        [string]$JobType = '',
        [string]$SourceFolder = ''
    )
    $safe = $TableOrView -replace '[^A-Za-z0-9_]', ''
    $catalog = @{
        v_poller_health = @{ kind = 'poller'; timeCol = 'started_at' }
        v_job_summary   = @{ kind = 'job_summary'; timeCol = 'created_at' }
        poll_runs       = @{ kind = 'table'; timeCol = 'started_at'; from = 'poll_runs' }
        processing_jobs = @{ kind = 'table'; timeCol = 'created_at'; from = 'processing_jobs' }
        audit_events    = @{ kind = 'table'; timeCol = 'captured_at'; from = 'audit_events' }
        notification_log = @{ kind = 'table'; timeCol = 'sent_at'; from = 'notification_log' }
    }
    if (-not $catalog.ContainsKey($safe)) {
        return New-QCFailureResult -Code 'OPS_SQL_BAD_NAME' -Message 'Invalid table or view name.'
    }
    $meta = $catalog[$safe]
    $topN = [Math]::Max(1, [Math]::Min(500, [int]$Top))
    $params = @{}
    $where = New-Object System.Collections.Generic.List[string]
    if ($null -ne $FromUtc -and $null -ne $ToUtc) {
        $fromDt = [datetime]$FromUtc
        $toDt = [datetime]$ToUtc
        if ($fromDt.Kind -ne [DateTimeKind]::Utc) { $fromDt = $fromDt.ToUniversalTime() }
        if ($toDt.Kind -ne [DateTimeKind]::Utc) { $toDt = $toDt.ToUniversalTime() }
        [void]$where.Add(('[{0}] >= @fromUtc' -f $meta.timeCol))
        [void]$where.Add(('[{0}] <= @toUtc' -f $meta.timeCol))
        $params.fromUtc = [datetimeoffset]([datetime]::SpecifyKind($fromDt, [DateTimeKind]::Utc))
        $params.toUtc = [datetimeoffset]([datetime]::SpecifyKind($toDt, [DateTimeKind]::Utc))
    }
    $jt = ([string]$JobType).Trim()
    if ($jt -and $jt -ne 'All') {
        if ($jt -notmatch '^[A-Za-z0-9_]+$') {
            return New-QCFailureResult -Code 'OPS_SQL_BAD_JOB_TYPE' -Message 'Job type contains invalid characters.'
        }
        if ($safe -in @('processing_jobs', 'v_job_summary')) {
            [void]$where.Add('job_type = @jobType')
            $params.jobType = $jt
        }
    }
    $sf = ([string]$SourceFolder).Trim()
    if ($sf) {
        $like = '%' + (($sf -replace '\[', '[[]') -replace '%', '[%]' -replace '_', '[_]') + '%'
        if ($safe -in @('processing_jobs', 'v_job_summary')) {
            [void]$where.Add('source_folder LIKE @sourceFolder')
            $params.sourceFolder = $like
        } elseif ($safe -eq 'audit_events') {
            [void]$where.Add('resolved_folder LIKE @sourceFolder')
            $params.sourceFolder = $like
        } elseif ($safe -eq 'notification_log') {
            [void]$where.Add('folder_path LIKE @sourceFolder')
            $params.sourceFolder = $like
        }
    }
    $whereSql = ''
    if ($where.Count -gt 0) { $whereSql = ' WHERE ' + ($where -join ' AND ') }
    $sql = ''
    switch ($meta.kind) {
        'poller' {
            $sql = @"
SELECT TOP ($topN)
    id, started_at, duration_ms, events_fetched, events_relevant, jobs_enqueued,
    CASE WHEN error_message IS NOT NULL THEN 'ERROR' ELSE 'OK' END AS run_status,
    watermark_after
FROM poll_runs
$whereSql
ORDER BY started_at DESC
"@
        }
        'job_summary' {
            $sql = @"
SELECT job_type, status, COUNT(*) AS job_count, AVG(duration_ms) AS avg_duration_ms, MAX(completed_at) AS last_completed
FROM processing_jobs
$whereSql
GROUP BY job_type, status
ORDER BY job_type, status
"@
        }
        default {
            $sql = "SELECT TOP ($topN) * FROM [$($meta.from)]$whereSql ORDER BY [$($meta.timeCol)] DESC"
        }
    }
    return New-QCSuccessResult -Code 'OPS_SQL_PREVIEW_SQL' -Message 'Preview SQL built.' -Data @{
        sql = $sql
        parameters = $params
        hideColumns = @('source_folder', 'sourceFolder')
        table = $safe
    }
}

function Get-QCOpsSqlTablePreview {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$TableOrView,
        [int]$Top = 50,
        [AllowNull()]$FromUtc = $null,
        [AllowNull()]$ToUtc = $null,
        [string]$JobType = '',
        [string]$SourceFolder = ''
    )
    if (-not (Test-QCDatabaseEnabled -Config $Config)) {
        return New-QCFailureResult -Code 'OPS_SQL_DISABLED' -Message 'database.enabled is false.'
    }
    $cmd = Get-QCOpsSqlPreviewCommand -TableOrView $TableOrView -Top $Top -FromUtc $FromUtc -ToUtc $ToUtc -JobType $JobType -SourceFolder $SourceFolder
    if (-not $cmd.IsSuccess) { return $cmd }
    $r = Invoke-QCDatabaseQuery -Config $Config -Sql $cmd.Data.sql -Parameters $cmd.Data.parameters
    if (-not $r.IsSuccess) { return $r }
    $r.Data.hideColumns = @($cmd.Data.hideColumns)
    return $r
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

function Get-QCOpsWatchFolderChoices {
    [CmdletBinding()]
    param([hashtable]$Config)
    $out = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    $add = {
        param([string]$Path)
        $p = ([string]$Path).Trim()
        if ([string]::IsNullOrWhiteSpace($p)) { return }
        if ($seen.ContainsKey($p.ToLowerInvariant())) { return }
        $seen[$p.ToLowerInvariant()] = $true
        [void]$out.Add($p)
    }
    try {
        $roots = $null
        if ($Config -and $Config.projectWise -and $Config.projectWise.watchList) {
            $roots = $Config.projectWise.watchList.roots
        }
        foreach ($r in @($roots)) {
            $path = ''
            try { $path = [string]$r.path } catch { }
            & $add $path
            $sheets = ''
            try { $sheets = [string]$r.sheetsPathFromProject } catch { }
            if ($path -and $sheets) { & $add ($path.TrimEnd('\') + '\' + $sheets.TrimStart('\')) }
        }
    } catch { }
    return @($out.ToArray())
}

function Get-QCOpsMaintenanceCatalog {
    [CmdletBinding()]
    param()
    $repo = Get-QCOpsRepoRoot
    $m = Join-Path $repo 'scripts\maintenance'
    return @(
        @{
            id = 'reconcile-status-sets'
            label = 'Reconcile status sets'
            script = (Join-Path $m 'Reconcile-QCStatusSets.ps1')
            danger = 'Low'
            flags = @(
                @{ name = 'DryRun'; kind = 'bool'; param = '-DryRun'; default = 'Off' }
            )
        }
        @{
            id = 'sync-sheet-index'
            label = 'Sync sheet index for folder'
            script = (Join-Path $m 'Sync-QCFolderSheetIndex.ps1')
            danger = 'Low'
            flags = @(
                @{ name = 'FolderPath'; kind = 'folder'; param = '-FolderPath'; required = $true }
                @{ name = 'ConfirmWrites'; kind = 'bool'; param = '-ConfirmWrites'; default = 'Off' }
                @{ name = 'DryRun'; kind = 'bool'; param = '-DryRun'; default = 'Off' }
                @{ name = 'OneLevelDeep'; kind = 'bool'; param = '-OneLevelDeep'; default = 'Off' }
            )
        }
        @{
            id = 'refresh-sheet-states'
            label = 'Refresh sheet_index states from PW'
            script = (Join-Path $m 'Refresh-SheetIndexStates.ps1')
            danger = 'Medium'
            flags = @(
                @{ name = 'FolderPathFilter'; kind = 'folder'; param = '-FolderPathFilter' }
                @{ name = 'ConfirmWrites'; kind = 'bool'; param = '-ConfirmWrites'; default = 'Off' }
                @{ name = 'DryRun'; kind = 'bool'; param = '-DryRun'; default = 'Off' }
                @{ name = 'IncludeQcPdfGuids'; kind = 'bool'; param = '-IncludeQcPdfGuids'; default = 'Off' }
            )
        }
        @{
            id = 'repair-folder-paths'
            label = 'Repair folder path casing'
            script = (Join-Path $m 'Repair-QCDocumentsFolderPaths.ps1')
            danger = 'Low'
            flags = @(
                @{ name = 'Table'; kind = 'choice'; param = '-Table'; options = @('all', 'telemetry', 'sheet_index', 'processing_jobs', 'audit_events', 'document_activity'); default = 'all' }
                @{ name = 'WhatIf'; kind = 'bool'; param = '-WhatIf'; default = 'Off' }
            )
        }
        @{
            id = 'db-retention'
            label = 'Database retention (audit_events)'
            script = (Join-Path $m 'Invoke-QCDatabaseRetention.ps1')
            danger = 'Medium'
            flags = @(
                @{ name = 'AuditEventsDays'; kind = 'choice'; param = '-AuditEventsDays'; options = @('Default', '30', '60', '90', '180', '365'); default = 'Default' }
                @{ name = 'ConfirmDeletes'; kind = 'bool'; param = '-ConfirmDeletes'; default = 'Off' }
                @{ name = 'DryRun'; kind = 'bool'; param = '-DryRun'; default = 'On' }
                @{ name = 'AuditIncludeUnprocessed'; kind = 'bool'; param = '-AuditIncludeUnprocessed'; default = 'Off' }
            )
        }
        @{
            id = 'reset-folder-workflow'
            label = 'Reset folder workflow + telemetry'
            script = (Join-Path $m 'Reset-QCFolderWorkflow.ps1')
            danger = 'High'
            flags = @(
                @{ name = 'FolderPath'; kind = 'folder'; param = '-FolderPath'; required = $true }
                @{ name = 'DryRun'; kind = 'bool'; param = '-DryRun'; default = 'On' }
                @{ name = 'ConfirmReset'; kind = 'bool'; param = '-ConfirmReset'; default = 'Off' }
                @{ name = 'KeepLanePdfRegistry'; kind = 'bool'; param = '-KeepLanePdfRegistry'; default = 'Off' }
            )
        }
        @{
            id = 'rewind-watermark'
            label = 'Rewind audit watermark'
            script = '(module) Set-QCOpsAuditWatermark'
            danger = 'High'
            confirmText = 'REWIND WATERMARK'
            flags = @(
                @{ name = 'When'; kind = 'choice'; options = @('1 hour ago', '6 hours ago', '1 day ago', '7 days ago', 'Custom UTC'); default = '1 hour ago' }
                @{ name = 'CustomUtc'; kind = 'text'; placeholder = '2026-08-24T12:00:00Z' }
            )
        }
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
