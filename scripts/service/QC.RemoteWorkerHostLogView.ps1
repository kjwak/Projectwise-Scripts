# Dot-source helpers for the remote worker host console.
# Used by Start-QCRemoteWorkerHost.ps1 (supervisor) and Watch-QCRemoteWorkerHostConsole.ps1 (logon viewer).
# In-flight work is read from queue\running (this host's machineName). JSONL is stage detail only.

if (-not $script:QCRemoteWorkerLogView) {
    $script:QCRemoteWorkerLogView = @{
        ActiveJobs = @{}
        RunningIds = @{}
        ProcJsonLog = @{
            hour = ''
            offset = [int64]0
            tail = ''
            skipExisting = $true
        }
        LocalWorkersOnly = $false
        WorkerLabelPattern = '^RW'
    }
}

function Initialize-QCRemoteWorkerHostLogView {
    param(
        [switch]$SkipExisting,
        [switch]$LocalWorkersOnly,
        [string]$WorkerLabelPattern = '^RW'
    )
    $script:QCRemoteWorkerLogView.ActiveJobs = @{}
    $script:QCRemoteWorkerLogView.RunningIds = @{}
    $script:QCRemoteWorkerLogView.ProcJsonLog = @{
        hour = ''
        offset = [int64]0
        tail = ''
        skipExisting = [bool]$SkipExisting.IsPresent
    }
    $script:QCRemoteWorkerLogView.LocalWorkersOnly = [bool]$LocalWorkersOnly.IsPresent
    $script:QCRemoteWorkerLogView.WorkerLabelPattern = [string]$WorkerLabelPattern
}

function Get-QCNextRemoteWorkerLabel {
    param([string[]]$CurrentLabels = @())
    $used = @{}
    foreach ($l in @($CurrentLabels)) {
        if ([string]$l -match '^RW(\d+)$') {
            $n = 0
            try { $n = [int]$Matches[1] } catch { $n = 0 }
            if ($n -gt 0) { $used[$n] = $true }
        }
    }
    $idx = 1
    while ($used.ContainsKey($idx)) { $idx++ }
    return ('RW{0}' -f $idx)
}

function Get-QCRemoteWorkerHostRunningJobs {
    param(
        [Parameter(Mandatory)][string]$QueueRoot,
        [string]$HostName = '',
        [switch]$AllHosts
    )
    $out = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($QueueRoot)) { return @() }
    $runningDir = Join-Path $QueueRoot 'running'
    if (-not (Test-Path -LiteralPath $runningDir)) { return @() }
    $wantHost = ([string]$HostName).Trim()
    foreach ($f in @(Get-ChildItem -LiteralPath $runningDir -Filter '*.json' -File -ErrorAction SilentlyContinue)) {
        $j = $null
        try { $j = Get-Content -LiteralPath $f.FullName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop } catch { $j = $null }
        $jobId = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
        $type = ''; $src = ''; $machine = ''; $label = ''
        if ($j) {
            try { if ($j.id) { $jobId = [string]$j.id } } catch { }
            try { if ($j.type) { $type = [string]$j.type } } catch { }
            try {
                if ($j.sourceName) { $src = [string]$j.sourceName }
                elseif ($j.sourcePath) { $src = [System.IO.Path]::GetFileName([string]$j.sourcePath) }
            } catch { }
            try {
                if ($j.PSObject.Properties['machineName'] -and $j.machineName) { $machine = ([string]$j.machineName).Trim() }
            } catch { }
            try {
                if ($j.PSObject.Properties['workerLabel'] -and $j.workerLabel) { $label = [string]$j.workerLabel }
            } catch { }
        }
        if (-not $AllHosts.IsPresent) {
            if ([string]::IsNullOrWhiteSpace($wantHost)) { continue }
            if ([string]::IsNullOrWhiteSpace($machine)) { continue }
            if ($machine.ToUpperInvariant() -ne $wantHost.ToUpperInvariant()) { continue }
        }
        [void]$out.Add([pscustomobject]@{
            jobId = $jobId
            type = $type
            sourceName = $src
            machineName = $machine
            workerLabel = $label
        })
    }
    return @($out.ToArray())
}

function Sync-QCRemoteWorkerHostRunningView {
    param(
        [Parameter(Mandatory)][string]$QueueRoot,
        [string]$HostName = '',
        [switch]$AllHosts,
        [switch]$Quiet
    )
    if (-not $script:QCRemoteWorkerLogView.RunningIds) { $script:QCRemoteWorkerLogView.RunningIds = @{} }
    if (-not $script:QCRemoteWorkerLogView.ActiveJobs) { $script:QCRemoteWorkerLogView.ActiveJobs = @{} }
    $jobs = @(Get-QCRemoteWorkerHostRunningJobs -QueueRoot $QueueRoot -HostName $HostName -AllHosts:$AllHosts)
    $current = @{}
    foreach ($j in $jobs) {
        if ($j.jobId) { $current[[string]$j.jobId] = $j }
    }
    $prev = $script:QCRemoteWorkerLogView.RunningIds
    $activeJobs = $script:QCRemoteWorkerLogView.ActiveJobs
    $entered = New-Object System.Collections.Generic.List[object]
    foreach ($id in @($current.Keys)) {
        if ($prev.ContainsKey($id)) { continue }
        $j = $current[$id]
        $src = $(if ($j.sourceName) { [string]$j.sourceName } else { [string]$j.type })
        $alreadyLogged = $activeJobs.ContainsKey($id)
        $prev[$id] = $src
        $activeJobs[$id] = $src
        if (-not $Quiet.IsPresent -and -not $alreadyLogged) { [void]$entered.Add($j) }
    }
    foreach ($id in @($prev.Keys)) {
        if (-not $current.ContainsKey($id)) { $prev.Remove($id) | Out-Null }
    }
    if (-not $Quiet.IsPresent) {
        foreach ($j in $entered) {
            $parts = New-Object System.Collections.Generic.List[string]
            [void]$parts.Add(('[{0}]' -f (Get-Date -Format 'HH:mm:ss')))
            [void]$parts.Add('QUEUE_RUNNING')
            if ($j.type) { [void]$parts.Add([string]$j.type) }
            if ($j.sourceName) { [void]$parts.Add([string]$j.sourceName) }
            if ($j.jobId) { [void]$parts.Add([string]$j.jobId) }
            Write-Host ($parts -join ' ') -ForegroundColor Cyan
        }
    }
    return @{ Jobs = $jobs; Entered = @($entered.ToArray()) }
}

function Get-QCRemoteWorkerHostProcessorJsonlPath {
    param(
        [Parameter(Mandatory)][string]$LogDir,
        [Parameter(Mandatory)][string]$HourStamp
    )
    return (Join-Path $LogDir ("Run-QCProcessor_${HourStamp}.jsonl"))
}

function _QCRemoteWorkerReadLogChunkFromOffset {
    param(
        [string]$Path,
        [int64]$StartPos
    )
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) {
        return @{ Text = ''; NewPos = [int64]$StartPos }
    }
    try {
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        try {
            if ($StartPos -gt $fs.Length) { $StartPos = [int64]0 }
            $fs.Position = $StartPos
            $sr = New-Object System.IO.StreamReader($fs, [System.Text.UTF8Encoding]::new($false), $false, 4096, $true)
            $text = $sr.ReadToEnd()
            $newPos = $fs.Position
            $sr.Dispose()
            return @{ Text = [string]$text; NewPos = [int64]$newPos }
        } finally {
            $fs.Dispose()
        }
    } catch {
        return @{ Text = ''; NewPos = [int64]$StartPos }
    }
}

function Test-QCRemoteWorkerHostJobEvent {
    param([string]$Code)
    if ([string]::IsNullOrWhiteSpace($Code)) { return $false }
    if ($Code -in @(
            'WORKER_NO_JOB'
            'WORKER_START'
            'WORKER_BUDGET'
            'WORKER_LEASE'
            'WORKER_LOCK_RACE'
            'WORKER_DB_SCHEMA_INIT_FAILED'
            'WORKER_QUEUE_DUPLICATE_CLEANUP'
            'JOB_TELEMETRY_WRITTEN'
            'JOB_TELEMETRY_SKIPPED'
        )) { return $false }
    if ($Code -match '^(WORKER_|QC_PREPEND|STATUS_SET|REMOTE_HOST_)') { return $true }
    return $false
}

function Write-QCRemoteWorkerHostJobLine {
    param([object]$Obj)
    if (-not $Obj) { return $false }

    $code = [string]$Obj.code
    if (-not (Test-QCRemoteWorkerHostJobEvent $code)) { return $false }

    $d = $null
    try { $d = $Obj.data } catch { }
    $label = ''; $jobId = ''; $jobType = ''; $src = ''; $stage = ''
    if ($d) {
        try { if ($d.workerLabel) { $label = [string]$d.workerLabel } } catch { }
        try { if ($d.label) { $label = [string]$d.label } } catch { }
        try { if ($d.jobId) { $jobId = [string]$d.jobId } } catch { }
        try { if ($d.jobType) { $jobType = [string]$d.jobType } } catch { }
        try {
            if ($d.sourceName) { $src = [string]$d.sourceName }
            elseif ($d.sourcePath) { $src = [System.IO.Path]::GetFileName([string]$d.sourcePath) }
        } catch { }
        try { if ($d.stage) { $stage = [string]$d.stage } } catch { }
    }

    if ($script:QCRemoteWorkerLogView.LocalWorkersOnly -and $label) {
        $pat = [string]$script:QCRemoteWorkerLogView.WorkerLabelPattern
        if ($pat -and $label -notmatch $pat) { return $false }
    }

    $activeJobs = $script:QCRemoteWorkerLogView.ActiveJobs
    $runningIds = $script:QCRemoteWorkerLogView.RunningIds
    if (-not $src -and $jobId -and $activeJobs -and $activeJobs.ContainsKey($jobId)) {
        try { $src = [string]$activeJobs[$jobId] } catch { }
    }
    if ($code -eq 'WORKER_STAGE' -and -not $jobId) { return $false }

    # running\ already showed this claim; JSONL SELECTED is duplicate, not a second start.
    if ($code -eq 'WORKER_SELECTED' -and $jobId) {
        $already = $false
        if ($activeJobs -and $activeJobs.ContainsKey($jobId)) { $already = $true }
        if ($runningIds -and $runningIds.ContainsKey($jobId)) { $already = $true }
        if ($already) {
            if ($src) {
                if ($activeJobs) { $activeJobs[$jobId] = $src }
                if ($runningIds) { $runningIds[$jobId] = $src }
            }
            return $true
        }
    }

    $msg = [string]$Obj.message
    $parts = New-Object System.Collections.Generic.List[string]
    [void]$parts.Add(('[{0}]' -f (Get-Date -Format 'HH:mm:ss')))
    if ($label) { [void]$parts.Add($label) }
    [void]$parts.Add($code)
    if ($jobType) { [void]$parts.Add($jobType) }
    if ($src) { [void]$parts.Add($src) }
    if ($jobId) { [void]$parts.Add($jobId) }
    if ($stage) { [void]$parts.Add(("stage={0}" -f $stage)) }
    if ($msg) { [void]$parts.Add($msg) }

    $color = 'White'
    switch -Regex ($code) {
        '^(WORKER_SELECTED|WORKER_CLAIMING)$' { $color = 'Cyan' }
        '^WORKER_STAGE$' { $color = 'Yellow' }
        '^QC_PREPEND_(PW_CHILD|PW_RECONNECT|TAG_CLEAR|WORKFLOW_WRITEBACK)' { $color = 'Yellow' }
        '^(WORKER_SUCCEEDED|QC_PREPEND_OK)$' { $color = 'Green' }
        '(FAILED|UNHANDLED|ERROR|STAMPS_MISSING|STAMP_UNRESOLVED)' { $color = 'Red' }
        default { $color = 'Gray' }
    }
    Write-Host ($parts -join ' ') -ForegroundColor $color

    if ($jobId) {
        if ($code -eq 'WORKER_SELECTED') {
            $activeJobs[$jobId] = $(if ($src) { $src } else { $jobType })
        } elseif ($code -match '^(WORKER_SUCCEEDED|WORKER_FAILED|WORKER_JOB_UNHANDLED)$') {
            try { $activeJobs.Remove($jobId) | Out-Null } catch { }
        }
    }
    return $true
}

function _QCRemoteWorkerProcessJsonlBuffer {
    param([string]$Buffer)
    $nl = $Buffer.LastIndexOfAny(@([char]10, [char]13))
    if ($nl -lt 0) {
        return @{ Complete = ''; Remainder = $Buffer }
    }
    return @{
        Complete = $Buffer.Substring(0, $nl + 1)
        Remainder = $Buffer.Substring($nl + 1)
    }
}

function _QCRemoteWorkerDrainJsonlFile {
    param([string]$Path)
    $state = $script:QCRemoteWorkerLogView.ProcJsonLog
    $chunk = _QCRemoteWorkerReadLogChunkFromOffset -Path $Path -StartPos ([int64]$state.offset)
    if ($chunk.NewPos -lt $state.offset) {
        $state.offset = [int64]0
        $state.tail = ''
        $chunk = _QCRemoteWorkerReadLogChunkFromOffset -Path $Path -StartPos 0
    }
    $state.offset = [int64]$chunk.NewPos
    $split = _QCRemoteWorkerProcessJsonlBuffer -Buffer (([string]$state.tail) + ([string]$chunk.Text))
    $state.tail = [string]$split.Remainder
    foreach ($line in ([string]$split.Complete -split '\r?\n')) {
        $t = ([string]$line).Trim()
        if (-not $t -or -not $t.StartsWith('{')) { continue }
        try {
            $o = $t | ConvertFrom-Json -ErrorAction Stop
            Write-QCRemoteWorkerHostJobLine -Obj $o | Out-Null
        } catch { }
    }
}

function Drain-QCRemoteWorkerHostJsonLogs {
    param([Parameter(Mandatory)][string]$LogDir)
    $hour = Get-QCLogHourStamp
    $state = $script:QCRemoteWorkerLogView.ProcJsonLog
    if ($state.hour -and [string]$state.hour -ne $hour) {
        $prev = Get-QCRemoteWorkerHostProcessorJsonlPath -LogDir $LogDir -HourStamp ([string]$state.hour)
        _QCRemoteWorkerDrainJsonlFile -Path $prev
        $state.offset = [int64]0
        $state.tail = ''
        $state.skipExisting = $false
    }
    $state.hour = $hour
    $path = Get-QCRemoteWorkerHostProcessorJsonlPath -LogDir $LogDir -HourStamp $hour
    if ($state.skipExisting) {
        if (Test-Path -LiteralPath $path) {
            try { $state.offset = [int64](Get-Item -LiteralPath $path).Length } catch { $state.offset = [int64]0 }
        }
        $state.skipExisting = $false
        return
    }
    _QCRemoteWorkerDrainJsonlFile -Path $path
}

function Show-QCRemoteWorkerHostRecentJsonLogs {
    param(
        [Parameter(Mandatory)][string]$LogDir,
        [int]$MaxLines = 40
    )
    $hour = Get-QCLogHourStamp
    $path = Get-QCRemoteWorkerHostProcessorJsonlPath -LogDir $LogDir -HourStamp $hour
    if (-not (Test-Path -LiteralPath $path)) { return }
    $lines = @(Get-Content -LiteralPath $path -Tail ($MaxLines * 3) -ErrorAction SilentlyContinue)
    $shown = 0
    foreach ($line in $lines) {
        $t = ([string]$line).Trim()
        if (-not $t -or -not $t.StartsWith('{')) { continue }
        try {
            $o = $t | ConvertFrom-Json -ErrorAction Stop
            if (Write-QCRemoteWorkerHostJobLine -Obj $o) { $shown++ }
        } catch { }
        if ($shown -ge $MaxLines) { break }
    }
    $script:QCRemoteWorkerLogView.ProcJsonLog.offset = [int64]0
    $script:QCRemoteWorkerLogView.ProcJsonLog.tail = ''
    $script:QCRemoteWorkerLogView.ProcJsonLog.skipExisting = $false
    if (Test-Path -LiteralPath $path) {
        try {
            $script:QCRemoteWorkerLogView.ProcJsonLog.offset = [int64](Get-Item -LiteralPath $path).Length
        } catch { }
    }
}

function Get-QCRemoteWorkerHostBusySummary {
    param(
        [string]$QueueRoot = '',
        [string]$HostName = '',
        [switch]$AllHosts
    )
    if (-not [string]::IsNullOrWhiteSpace($QueueRoot)) {
        $jobs = @(Get-QCRemoteWorkerHostRunningJobs -QueueRoot $QueueRoot -HostName $HostName -AllHosts:$AllHosts)
        $busyN = @($jobs).Count
        if ($busyN -le 0) { return 'idle' }
        $names = @($jobs | ForEach-Object {
            if ($_.sourceName) { [string]$_.sourceName } elseif ($_.type) { [string]$_.type } else { [string]$_.jobId }
        } | Where-Object { $_ } | Select-Object -First 3)
        return 'busy=' + $busyN + ' ' + ($names -join ',')
    }
    $activeJobs = $script:QCRemoteWorkerLogView.ActiveJobs
    $busyN = @($activeJobs.Keys).Count
    if ($busyN -le 0) { return 'idle' }
    return 'busy=' + $busyN + ' ' + ((@($activeJobs.Values) | Select-Object -First 3) -join ',')
}
