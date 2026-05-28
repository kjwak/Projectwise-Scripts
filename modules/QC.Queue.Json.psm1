# QC.Queue.Json.psm1
# Responsibility: JSON-backed queue persistence, lifecycle transitions, and queue reporting.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Core.Runtime.psm1') -Force

function _QCQJ-IsNullOrWhiteSpace([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function _QCQJ-GetQueueRoot([hashtable]$Config) {
    if ($Config.ContainsKey('queue') -and $Config.queue) {
        if ($Config.queue.ContainsKey('rootDir') -and $Config.queue.rootDir) { return [string]$Config.queue.rootDir }
        if ($Config.queue.ContainsKey('root') -and $Config.queue.root) { return [string]$Config.queue.root }
        if ($Config.queue.ContainsKey('path') -and $Config.queue.path) { return [string]$Config.queue.path }
    }

    # Default: <repo-root>\queue (modules folder is one level down)
    $repoRoot = Split-Path -Parent $PSScriptRoot
    return (Join-Path $repoRoot 'queue')
}

function _QCQJ-GetLockAcquireSettings {
    <#
    .SYNOPSIS
    Timeouts for queue/job lock files. Defaults are conservative for AV scanners
    that briefly hold handles on the queue folder or lock files.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [hashtable]$Config
    )
    $timeoutMs = 30000
    $sleepMs = 150
    if ($Config -and $Config.ContainsKey('queue') -and $Config.queue) {
        $q = $Config.queue
        if ($q -is [hashtable]) {
            if ($q.ContainsKey('lockAcquireTimeoutMs') -and $null -ne $q['lockAcquireTimeoutMs']) {
                try { $timeoutMs = [int]$q['lockAcquireTimeoutMs'] } catch { }
            }
            if ($q.ContainsKey('lockAcquireSleepMs') -and $null -ne $q['lockAcquireSleepMs']) {
                try { $sleepMs = [int]$q['lockAcquireSleepMs'] } catch { }
            }
        } elseif ($q.PSObject -and $q.PSObject.Properties) {
            try {
                $p1 = $q.PSObject.Properties['lockAcquireTimeoutMs']
                if ($p1 -and $null -ne $p1.Value) { $timeoutMs = [int]$p1.Value }
            } catch { }
            try {
                $p2 = $q.PSObject.Properties['lockAcquireSleepMs']
                if ($p2 -and $null -ne $p2.Value) { $sleepMs = [int]$p2.Value }
            } catch { }
        }
    }
    if ($timeoutMs -lt 1000) { $timeoutMs = 1000 }
    if ($sleepMs -lt 20) { $sleepMs = 20 }
    if ($sleepMs -gt 2000) { $sleepMs = 2000 }
    return @{ TimeoutMs = $timeoutMs; SleepMs = $sleepMs }
}

function _QCQJ-NormalizeState([string]$State) {
    $s = ($State -as [string]).Trim().ToLowerInvariant()
    if (-not $s) { return $null }
    switch ($s) {
        'queued' { return 'pending' }
        'queue' { return 'pending' }
        'pending' { return 'pending' }
        'processing' { return 'running' } # backward-compatible alias (orchestrator currently uses "processing")
        'running' { return 'running' }
        'succeeded' { return 'succeeded' }
        'success' { return 'succeeded' }
        'failed' { return 'failed' }
        'failure' { return 'failed' }
        default { return $s }
    }
}

function _QCQJ-EnsureLayout([string]$Root) {
    foreach ($d in @('pending', 'running', 'succeeded', 'failed', 'locks')) {
        $p = Join-Path $Root $d
        if (-not (Test-Path -LiteralPath $p)) {
            New-Item -ItemType Directory -Path $p -Force | Out-Null
        }
    }
}

function _QCQJ-JobFilePath([string]$Root, [string]$State, [string]$JobId) {
    $s = _QCQJ-NormalizeState $State
    return (Join-Path (Join-Path $Root $s) ($JobId + '.json'))
}

function _QCQJ-LockFilePath([string]$Root, [string]$JobId) {
    return (Join-Path (Join-Path $Root 'locks') ($JobId + '.lock'))
}

function _QCQJ-QueueWriteLockPath([string]$Root) {
    return (Join-Path (Join-Path $Root 'locks') '_queue_write.lock')
}

function _QCQJ-DedupeIndexPath([string]$Root) {
    return (Join-Path (Join-Path $Root '_watcher') 'dedupe-index.json')
}

function _QCQJ-ReadDedupeIndex([string]$Path) {
    $empty = @{ version = 1; entries = @{}; loadedAtUtc = (Get-QCTimestamp) }
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return $empty }
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return $empty }
        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        $entries = @{}
        if ($obj -and $obj.entries) {
            foreach ($prop in @($obj.entries.PSObject.Properties)) {
                $entries[[string]$prop.Name] = $prop.Value
            }
        }
        return @{ version = 1; entries = $entries; loadedAtUtc = (Get-QCTimestamp) }
    } catch {
        return $empty
    }
}

function _QCQJ-WriteDedupeIndexAtomic([string]$Path, [hashtable]$Index) {
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $tmp = ($Path + '.tmp.' + ([guid]::NewGuid().ToString('N')))
    $safeEntries = if ($Index -and $Index.ContainsKey('entries')) { $Index.entries } else { @{} }
    $payload = @{ version = 1; writtenAtUtc = (Get-QCTimestamp); entries = $safeEntries } | ConvertTo-Json -Depth 20
    $enc = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($tmp, $payload, $enc)
    _QCQJ-MoveItemWithRetry -LiteralPath $tmp -Destination $Path
}

function _QCQJ-UpdateDedupeIndexForJob([hashtable]$Config, [string]$Root, [string]$DedupeKey, [string]$JobId, [string]$State) {
    if (_QCQJ-IsNullOrWhiteSpace $DedupeKey) { return }
    $path = _QCQJ-DedupeIndexPath -Root $Root
    $idx = _QCQJ-ReadDedupeIndex -Path $path
    if (-not $idx.entries) { $idx['entries'] = @{} }
    $idx.entries[[string]$DedupeKey] = @{
        jobId = $JobId
        state = $State
        updatedAtUtc = (Get-QCTimestamp)
    }
    _QCQJ-WriteDedupeIndexAtomic -Path $path -Index $idx
}

function _QCQJ-IsLockOwnerDead([string]$LockPath) {
    # Returns $true only if we can read the lock file, parse a non-zero owner PID,
    # AND that PID is no longer alive. Returns $false on any uncertainty so we
    # never steal a lock from a live owner.
    $payload = $null
    try { $payload = Get-Content -LiteralPath $LockPath -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json -ErrorAction SilentlyContinue } catch { return $false }
    if (-not $payload) { return $false }
    $ownerPid = 0
    try {
        if ($payload.PSObject -and $payload.PSObject.Properties['pid']) { $ownerPid = [int]$payload.pid }
    } catch { return $false }
    if ($ownerPid -le 0) { return $false }
    if ($ownerPid -eq $PID) { return $false }   # never steal from our own process
    $proc = Get-Process -Id $ownerPid -ErrorAction SilentlyContinue
    return (-not $proc)
}

function _QCQJ-AcquireLockFile([string]$LockPath, [int]$TimeoutMs = 30000, [int]$SleepMs = 150) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $stealAttempted = $false
    while ($sw.ElapsedMilliseconds -lt $TimeoutMs) {
        try {
            $fs = [System.IO.File]::Open($LockPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
            try {
                $payload = @{
                    pid = $PID
                    createdAtUtc = (Get-QCTimestamp)
                } | ConvertTo-Json -Depth 5
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
                $fs.Write($bytes, 0, $bytes.Length)
                $fs.Flush()
            } finally {
                $fs.Dispose()
            }
            return $true
        } catch {
            # CreateNew failed: someone else has the lock. Check whether the owner
            # is dead. If so, steal it (delete + retry CreateNew). Without this,
            # a worker killed by AV / Ctrl-C / OS leaves _queue_write.lock or a
            # per-job .lock stuck on disk forever and every subsequent worker
            # spins on QUEUE_LOCK_TIMEOUT until the dashboard restarts and
            # Recover-QCStaleJobs sweeps it. We only attempt the steal once per
            # call to keep the cost bounded.
            if (-not $stealAttempted) {
                $stealAttempted = $true
                if (_QCQJ-IsLockOwnerDead -LockPath $LockPath) {
                    try { Remove-Item -LiteralPath $LockPath -Force -ErrorAction Stop } catch { }
                    continue
                }
            }
            Start-Sleep -Milliseconds $SleepMs
        }
    }
    return $false
}

function _QCQJ-ReleaseLockFile([string]$LockPath) {
    Remove-Item -LiteralPath $LockPath -Force -ErrorAction SilentlyContinue
}

function _QCQJ-DeepToJsonSafeObject {
    <#
    .SYNOPSIS
    Recursively normalizes arbitrary PowerShell objects into JSON-friendly
    primitives, hashtables, and arrays. Used before serializing queue jobs so
    mixed PSCustomObject / Hashtable graphs (e.g. after a handler merges
    processor results into an in-memory job) cannot cause ConvertTo-Json to
    throw — which surfaced as QUEUE_MOVE_ERROR / WORKER_MOVE_FAILED even when
    the processor itself succeeded.
    #>
    param(
        [AllowNull()]
        [object]$Value,
        [int]$CurrentDepth = 0
    )

    if ($CurrentDepth -gt 64) { return '<qcq_json_max_depth>' }
    if ($null -eq $Value) { return $null }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [bool]) { return $Value }
    if ($Value -is [System.ValueType]) {
        if ($Value -is [datetime]) {
            try { return (ConvertTo-QCTimestamp -DateTime ([datetime]$Value)) } catch { return [string]$Value }
        }
        if ($Value -is [datetimeoffset]) {
            try { return (ConvertTo-QCTimestamp -DateTime ([datetimeoffset]$Value).UtcDateTime) } catch { return [string]$Value }
        }
        return $Value
    }

    if ($Value -is [System.Collections.IDictionary]) {
        $h = @{}
        foreach ($k in @($Value.Keys)) {
            $ks = [string]$k
            $h[$ks] = (_QCQJ-DeepToJsonSafeObject -Value $Value[$k] -CurrentDepth ($CurrentDepth + 1))
        }
        return $h
    }
    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $out = @()
        foreach ($i in $Value) {
            $out += _QCQJ-DeepToJsonSafeObject -Value $i -CurrentDepth ($CurrentDepth + 1)
        }
        return $out
    }
    if ($Value.PSObject -and $Value.PSObject.Properties) {
        $h = @{}
        foreach ($p in $Value.PSObject.Properties) {
            $h[$p.Name] = (_QCQJ-DeepToJsonSafeObject -Value $p.Value -CurrentDepth ($CurrentDepth + 1))
        }
        return $h
    }
    try { return [string]$Value } catch { return '<qcq_json_unserializable>' }
}

function _QCQJ-MoveItemWithRetry {
    param(
        [Parameter(Mandatory)]
        [string]$LiteralPath,
        [Parameter(Mandatory)]
        [string]$Destination,
        [int]$Attempts = 10,
        [int]$SleepMs = 200
    )
    $lastEx = $null
    for ($a = 1; $a -le $Attempts; $a++) {
        try {
            Move-Item -LiteralPath $LiteralPath -Destination $Destination -Force -ErrorAction Stop
            return
        } catch {
            $lastEx = $_
            if ($a -ge $Attempts) { throw $lastEx }
            Start-Sleep -Milliseconds $SleepMs
        }
    }
}

function _QCQJ-ReadJobFile([string]$Path) {
    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
    $obj = $raw | ConvertFrom-Json -ErrorAction Stop

    function _ToHashtable([object]$Value) {
        if ($null -eq $Value) { return $null }
        if ($Value -is [string]) { return $Value }
        if ($Value -is [System.ValueType]) { return $Value }
        if ($Value -is [System.Collections.IDictionary]) {
            $h = @{}
            foreach ($k in $Value.Keys) {
                $nk = [string]$k
                $h[$nk] = (_ToHashtable $Value[$k])
            }
            return $h
        }
        if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
            $out = @()
            foreach ($i in $Value) { $out += (_ToHashtable $i) }
            return $out
        }
        if ($Value.PSObject -and $Value.PSObject.Properties) {
            $h = @{}
            foreach ($p in $Value.PSObject.Properties) {
                $h[$p.Name] = (_ToHashtable $p.Value)
            }
            return $h
        }
        return $Value
    }

    return (_ToHashtable $obj)
}

function _QCQJ-WriteJobFileAtomic([string]$Path, [hashtable]$Job) {
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $tmp = ($Path + '.tmp.' + ([guid]::NewGuid().ToString('N')))
    try {
        $safe = _QCQJ-DeepToJsonSafeObject -Value $Job -CurrentDepth 0
        $json = $safe | ConvertTo-Json -Depth 100 -Compress -ErrorAction Stop
        # No BOM: pwsh/Set-Content utf8 adds BOM; Windows PowerShell utf8 adds BOM.
        # Use UTF8Encoding(false) for bytes the dashboard/tools expect.
        $enc = New-Object System.Text.UTF8Encoding $false
        [System.IO.File]::WriteAllText($tmp, $json, $enc)
        _QCQJ-MoveItemWithRetry -LiteralPath $tmp -Destination $Path
    } catch {
        try { if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue } } catch { }
        throw
    }
}

function _QCQJ-FindJobFile([string]$Root, [string]$JobId) {
    foreach ($s in @('pending', 'running', 'succeeded', 'failed')) {
        $p = _QCQJ-JobFilePath -Root $Root -State $s -JobId $JobId
        if (Test-Path -LiteralPath $p) { return @{ state = $s; path = $p } }
    }
    return $null
}

function Add-QCQueueJob {
    <#
    .SYNOPSIS
    Adds a job to the queue backend.
    .DESCRIPTION
    Persists a validated job payload into the queued state store.
    .PARAMETER Job
    Job payload to persist.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: local queue state writes.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    try {
        $root = _QCQJ-GetQueueRoot -Config $Config
        _QCQJ-EnsureLayout -Root $root

        $jobId = [string]$Job['id']
        if (_QCQJ-IsNullOrWhiteSpace $jobId) {
            return New-QCFailureResult -Code 'QUEUE_VALIDATION_MISSING_JOB_ID' -Message 'Job.id is required.' -Data @{ job = $Job }
        }
        if (_QCQJ-IsNullOrWhiteSpace ([string]$Job['dedupeKey'])) {
            return New-QCFailureResult -Code 'QUEUE_VALIDATION_MISSING_DEDUPE_KEY' -Message 'Job.dedupeKey is required.' -Data @{ job = $Job }
        }

        $lockPath = _QCQJ-QueueWriteLockPath -Root $root
        $lk = _QCQJ-GetLockAcquireSettings -Config $Config
        if (-not (_QCQJ-AcquireLockFile -LockPath $lockPath -TimeoutMs $lk.TimeoutMs -SleepMs $lk.SleepMs)) {
            return New-QCFailureResult -Code 'QUEUE_LOCK_TIMEOUT' -Message 'Timed out acquiring queue write lock.' -Data @{ lockPath = $lockPath }
        }
        try {
            $dest = _QCQJ-JobFilePath -Root $root -State 'pending' -JobId $jobId
            if (Test-Path -LiteralPath $dest) {
                return New-QCFailureResult -Code 'QUEUE_JOB_ALREADY_EXISTS' -Message 'Job already exists in pending state.' -Data @{ jobId = $jobId; path = $dest }
            }
            $Job.status = 'pending'
            $Job.enqueuedAtUtc = (Get-QCTimestamp)
            _QCQJ-WriteJobFileAtomic -Path $dest -Job $Job
            try {
                _QCQJ-UpdateDedupeIndexForJob -Config $Config -Root $root -DedupeKey ([string]$Job['dedupeKey']) -JobId $jobId -State 'pending'
            } catch { }
        } finally {
            _QCQJ-ReleaseLockFile -LockPath $lockPath
        }

        return New-QCSuccessResult -Code 'QUEUE_ENQUEUED' -Message 'Job enqueued to pending.' -Data @{ jobId = $jobId; state = 'pending'; root = $root }
    } catch {
        return New-QCFailureResult -Code 'QUEUE_ENQUEUE_ERROR' -Message 'Failed to enqueue job.' -Data @{ error = $_ }
    }
}

function Get-NextQCJob {
    <#
    .SYNOPSIS
    Retrieves the next eligible queued job.
    .DESCRIPTION
    Selects a queue item according to configured ordering and eligibility.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .PARAMETER ExcludeJobIds
    Optional list of job ids to skip during selection. Used by parallel workers
    that just lost a Lock-QCJob race so they don't immediately re-pick the same loser.
    .PARAMETER ExcludeJobTypes
    Optional list of job-type strings to skip during selection (e.g. 'STATUS_SET_GEN'
    while the watcher pass is still in progress).
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: none (selection only).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory = $false)]
        [string[]]$ExcludeJobIds = @(),
        [Parameter(Mandatory = $false)]
        [string[]]$ExcludeJobTypes = @()
    )

    try {
        $root = _QCQJ-GetQueueRoot -Config $Config
        _QCQJ-EnsureLayout -Root $root

        $excludeSet = @{}
        foreach ($x in @($ExcludeJobIds)) {
            if (-not [string]::IsNullOrWhiteSpace($x)) { $excludeSet[[string]$x] = $true }
        }
        $excludeTypeSet = @{}
        foreach ($t in @($ExcludeJobTypes)) {
            if (-not [string]::IsNullOrWhiteSpace($t)) { $excludeTypeSet[([string]$t).Trim()] = $true }
        }

        $pendingDir = Join-Path $root 'pending'
        $files = @(Get-ChildItem -LiteralPath $pendingDir -Filter '*.json' -File -ErrorAction SilentlyContinue | Sort-Object -Property LastWriteTimeUtc, Name)

        # Optional selection preferences (simple scheduler):
        # Config.queue.selection.preferJobTypes = ['STATUS_SET_GEN','QC_PREPEND']
        $preferTypes = @()
        try {
            if ($Config.ContainsKey('queue') -and $Config.queue -and $Config.queue.ContainsKey('selection') -and $Config.queue.selection) {
                $sel = $Config.queue.selection
                if ($sel -is [hashtable] -and $sel.ContainsKey('preferJobTypes') -and $sel.preferJobTypes) {
                    $preferTypes = @($sel.preferJobTypes | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
                }
            }
        } catch { $preferTypes = @() }

        # Helper: returns $true only if the per-job lock blocks selection.
        # A lock blocks selection only when the owner is still alive. If the
        # owner PID is dead we self-heal here by deleting the orphan lock so
        # the job becomes selectable again. Without this, a worker that
        # crashed mid-Lock-QCJob (or that wrote a per-job lock and then died)
        # would orphan the job in pending\ forever, since this function would
        # always skip it via Test-Path.
        $isJobLockedAlive = {
            param([string]$JobId)
            $lockPath = _QCQJ-LockFilePath -Root $root -JobId $JobId
            if (-not (Test-Path -LiteralPath $lockPath)) { return $false }
            if (_QCQJ-IsLockOwnerDead -LockPath $lockPath) {
                try { Remove-Item -LiteralPath $lockPath -Force -ErrorAction Stop } catch { }
                return $false
            }
            return $true
        }

        if ($preferTypes.Count -gt 0) {
            foreach ($pt in $preferTypes) {
                if ($excludeTypeSet.ContainsKey($pt)) { continue }
                foreach ($f in $files) {
                    $jobId = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
                    if ($excludeSet.ContainsKey($jobId)) { continue }
                    if (& $isJobLockedAlive $jobId) { continue }

                    $job = _QCQJ-ReadJobFile -Path $f.FullName
                    $jt = ''
                    try { $jt = [string]$job.type } catch { $jt = '' }
                    if ($excludeTypeSet.ContainsKey($jt)) { continue }
                    if ($jt -eq $pt) {
                        return New-QCSuccessResult -Code 'QUEUE_NEXT_JOB' -Message 'Next pending job selected (preferred type).' -Data @{ job = $job; jobId = $jobId; state = 'pending'; preferredType = $pt }
                    }
                }
            }
        }

        foreach ($f in $files) {
            $jobId = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
            if ($excludeSet.ContainsKey($jobId)) { continue }
            if (& $isJobLockedAlive $jobId) { continue }

            $job = _QCQJ-ReadJobFile -Path $f.FullName
            $jt = ''
            try { $jt = [string]$job.type } catch { $jt = '' }
            if ($excludeTypeSet.ContainsKey($jt)) { continue }
            return New-QCSuccessResult -Code 'QUEUE_NEXT_JOB' -Message 'Next pending job selected.' -Data @{ job = $job; jobId = $jobId; state = 'pending' }
        }

        return New-QCSuccessResult -Code 'QUEUE_EMPTY' -Message 'No eligible pending jobs.' -Data @{ job = $null }
    } catch {
        return New-QCFailureResult -Code 'QUEUE_GET_NEXT_ERROR' -Message 'Failed to select next job.' -Data @{ error = $_ }
    }
}

function Get-QCWatcherActiveFlagPath {
    <#
    .SYNOPSIS
    Returns the path of the watcher-active sentinel flag file used by the dashboard
    to gate STATUS_SET_GEN job processing while the watcher pass is in progress.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
    $root = _QCQJ-GetQueueRoot -Config $Config
    return (Join-Path $root '_watcher_active.flag')
}

function Test-QCWatcherActive {
    <#
    .SYNOPSIS
    Returns $true if a watcher pass is currently in progress (sentinel flag exists).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
    try {
        $flag = Get-QCWatcherActiveFlagPath -Config $Config
        return (Test-Path -LiteralPath $flag)
    } catch {
        return $false
    }
}

function Set-QCWatcherActive {
    <#
    .SYNOPSIS
    Marks the watcher pass as active by creating the sentinel flag file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
    try {
        $root = _QCQJ-GetQueueRoot -Config $Config
        _QCQJ-EnsureLayout -Root $root
        $flag = Join-Path $root '_watcher_active.flag'
        $payload = @{
            pid = $PID
            createdAtUtc = (Get-QCTimestamp)
        } | ConvertTo-Json -Depth 5
        Set-Content -LiteralPath $flag -Value $payload -Encoding utf8 -ErrorAction Stop
        return New-QCSuccessResult -Code 'QUEUE_WATCHER_ACTIVE_SET' -Message 'Watcher-active flag set.' -Data @{ path = $flag }
    } catch {
        return New-QCFailureResult -Code 'QUEUE_WATCHER_ACTIVE_SET_ERROR' -Message 'Failed to set watcher-active flag.' -Data @{ error = $_ }
    }
}

function Clear-QCWatcherActive {
    <#
    .SYNOPSIS
    Clears the watcher-active sentinel flag (no-op if not present).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
    try {
        $flag = Get-QCWatcherActiveFlagPath -Config $Config
        if (Test-Path -LiteralPath $flag) {
            Remove-Item -LiteralPath $flag -Force -ErrorAction SilentlyContinue
        }
        return New-QCSuccessResult -Code 'QUEUE_WATCHER_ACTIVE_CLEARED' -Message 'Watcher-active flag cleared.' -Data @{ path = $flag }
    } catch {
        return New-QCFailureResult -Code 'QUEUE_WATCHER_ACTIVE_CLEAR_ERROR' -Message 'Failed to clear watcher-active flag.' -Data @{ error = $_ }
    }
}

function _QCQJ-NormalizeFolderKey([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    return ([string]$Path).Trim().TrimEnd('\', '/').Replace('/', '\').ToLowerInvariant()
}

function Test-QCStatusSetJobInFlight {
    <#
    .SYNOPSIS
    Returns whether a STATUS_SET_GEN job is already pending or running for a Sheets folder.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,

        [Parameter(Mandatory)]
        [string]$SourceFolder
    )
    try {
        $want = _QCQJ-NormalizeFolderKey $SourceFolder
        if (-not $want) {
            return New-QCSuccessResult -Code 'QUEUE_STATUSSET_NOT_IN_FLIGHT' -Message 'No source folder to match.' -Data @{ inFlight = $false }
        }

        $root = _QCQJ-GetQueueRoot -Config $Config
        foreach ($state in @('pending', 'running')) {
            $dir = Join-Path $root $state
            if (-not (Test-Path -LiteralPath $dir)) { continue }
            $files = @(Get-ChildItem -LiteralPath $dir -Filter '*.json' -File -ErrorAction SilentlyContinue)
            foreach ($f in $files) {
                $job = _QCQJ-ReadJobFile -Path $f.FullName
                $jt = ''
                try { $jt = [string]$job.type } catch { $jt = '' }
                if ($jt -ne 'STATUS_SET_GEN') { continue }

                $sf = ''
                try { $sf = [string]$job.sourceFolder } catch { $sf = '' }
                if (-not $sf) {
                    try { $sf = [string]$job.sourcePath } catch { $sf = '' }
                }
                if ((_QCQJ-NormalizeFolderKey $sf) -ne $want) { continue }

                $jobId = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
                return New-QCSuccessResult -Code 'QUEUE_STATUSSET_IN_FLIGHT' -Message 'STATUS_SET_GEN already queued or running for folder.' -Data @{
                    inFlight = $true
                    queueState = $state
                    jobId = $jobId
                    sourceFolder = $sf
                }
            }
        }

        return New-QCSuccessResult -Code 'QUEUE_STATUSSET_NOT_IN_FLIGHT' -Message 'No in-flight STATUS_SET_GEN for folder.' -Data @{ inFlight = $false }
    } catch {
        return New-QCFailureResult -Code 'QUEUE_STATUSSET_IN_FLIGHT_ERROR' -Message 'Failed to check in-flight STATUS_SET_GEN.' -Data @{ error = $_ }
    }
}

function Get-QCJobById {
    <#
    .SYNOPSIS
    Retrieves a job by ID from queue storage.
    .DESCRIPTION
    Looks up job metadata/state by identifier.
    .PARAMETER JobId
    Job identifier.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: file read only.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$JobId,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    try {
        $root = _QCQJ-GetQueueRoot -Config $Config
        _QCQJ-EnsureLayout -Root $root

        $loc = _QCQJ-FindJobFile -Root $root -JobId $JobId
        if (-not $loc) {
            return New-QCSuccessResult -Code 'QUEUE_JOB_NOT_FOUND' -Message 'Job not found.' -Data @{ jobId = $JobId; found = $false }
        }

        $job = _QCQJ-ReadJobFile -Path $loc.path
        return New-QCSuccessResult -Code 'QUEUE_JOB_FOUND' -Message 'Job found.' -Data @{ jobId = $JobId; found = $true; state = $loc.state; job = $job }
    } catch {
        return New-QCFailureResult -Code 'QUEUE_GET_BY_ID_ERROR' -Message 'Failed to load job by id.' -Data @{ jobId = $JobId; error = $_ }
    }
}

function Set-QCJobStatus {
    <#
    .SYNOPSIS
    Updates job status metadata.
    .DESCRIPTION
    Applies a status change to an existing job record.
    .PARAMETER JobId
    Job identifier.
    .PARAMETER Status
    Target status value.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: local queue state writes.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$JobId,
        [Parameter(Mandatory)]
        [string]$Status,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    try {
        $root = _QCQJ-GetQueueRoot -Config $Config
        _QCQJ-EnsureLayout -Root $root

        $lockPath = _QCQJ-QueueWriteLockPath -Root $root
        $lk = _QCQJ-GetLockAcquireSettings -Config $Config
        if (-not (_QCQJ-AcquireLockFile -LockPath $lockPath -TimeoutMs $lk.TimeoutMs -SleepMs $lk.SleepMs)) {
            return New-QCFailureResult -Code 'QUEUE_LOCK_TIMEOUT' -Message 'Timed out acquiring queue write lock.' -Data @{ lockPath = $lockPath }
        }
        try {
            $loc = _QCQJ-FindJobFile -Root $root -JobId $JobId
            if (-not $loc) {
                return New-QCFailureResult -Code 'QUEUE_JOB_NOT_FOUND' -Message 'Job not found for status update.' -Data @{ jobId = $JobId }
            }
            $job = _QCQJ-ReadJobFile -Path $loc.path
            $job.status = $Status
            $job.updatedAtUtc = (Get-QCTimestamp)
            _QCQJ-WriteJobFileAtomic -Path $loc.path -Job $job
            return New-QCSuccessResult -Code 'QUEUE_STATUS_UPDATED' -Message 'Job status updated.' -Data @{ jobId = $JobId; state = $loc.state; status = $Status }
        } finally {
            _QCQJ-ReleaseLockFile -LockPath $lockPath
        }
    } catch {
        return New-QCFailureResult -Code 'QUEUE_SET_STATUS_ERROR' -Message 'Failed to update job status.' -Data @{ jobId = $JobId; error = $_ }
    }
}

function Update-QCJob {
    <#
    .SYNOPSIS
    Updates a job record in-place.
    .DESCRIPTION
    Writes the provided job hashtable over the existing job JSON (same state folder), without moving states.
    Use this to persist fields like attempts/lastError/result before calling Move-QCJob.
    .PARAMETER Job
    Job hashtable (must include id).
    .PARAMETER Config
    Loaded app configuration hashtable.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    $jobId = [string]$Job.id
    if ([string]::IsNullOrWhiteSpace($jobId)) {
        return New-QCFailureResult -Code 'QUEUE_UPDATE_MISSING_ID' -Message 'Job.id is required.' -Data @{ }
    }

    try {
        $root = _QCQJ-GetQueueRoot -Config $Config
        _QCQJ-EnsureLayout -Root $root

        $lockPath = _QCQJ-QueueWriteLockPath -Root $root
        $lk = _QCQJ-GetLockAcquireSettings -Config $Config
        if (-not (_QCQJ-AcquireLockFile -LockPath $lockPath -TimeoutMs $lk.TimeoutMs -SleepMs $lk.SleepMs)) {
            return New-QCFailureResult -Code 'QUEUE_LOCK_TIMEOUT' -Message 'Timed out acquiring queue write lock.' -Data @{ lockPath = $lockPath }
        }
        try {
            $loc = _QCQJ-FindJobFile -Root $root -JobId $jobId
            if (-not $loc) {
                return New-QCFailureResult -Code 'QUEUE_JOB_NOT_FOUND' -Message 'Job not found for update.' -Data @{ jobId = $jobId }
            }
            $Job.updatedAtUtc = (Get-QCTimestamp)
            _QCQJ-WriteJobFileAtomic -Path $loc.path -Job $Job
            return New-QCSuccessResult -Code 'QUEUE_JOB_UPDATED' -Message 'Job updated.' -Data @{ jobId = $jobId; state = $loc.state; path = $loc.path }
        } finally {
            _QCQJ-ReleaseLockFile -LockPath $lockPath
        }
    } catch {
        return New-QCFailureResult -Code 'QUEUE_UPDATE_ERROR' -Message 'Failed to update job.' -Data @{ jobId = $jobId; error = $_ }
    }
}

function Move-QCJob {
    <#
    .SYNOPSIS
    Moves a job between queue state buckets.
    .DESCRIPTION
    Performs state-folder transition for a job record under a single queue-write
    lock acquisition. When -Job is supplied, the in-memory hashtable replaces the
    on-disk job content (status/updatedAtUtc are stamped here), letting callers
    persist a result/lastError without a separate Update-QCJob + Move-QCJob round
    trip. That round trip used to acquire the global write lock twice in a row,
    which under concurrent workers was the dominant source of QUEUE_LOCK_TIMEOUT
    on the success/failure path.
    .PARAMETER JobId
    Job identifier.
    .PARAMETER FromState
    Current queue state.
    .PARAMETER ToState
    Destination queue state.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .PARAMETER Job
    Optional in-memory job hashtable. When provided, this is what gets written to
    disk in the destination state (with status/updatedAtUtc stamped). Job.id MUST
    match JobId. When omitted, the existing on-disk job file is read, its status
    updated, and rewritten - matching the historical Move-QCJob behavior.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: local queue file move/write.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$JobId,
        [Parameter(Mandatory)]
        [string]$FromState,
        [Parameter(Mandatory)]
        [string]$ToState,
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory = $false)]
        [hashtable]$Job
    )

    try {
        $root = _QCQJ-GetQueueRoot -Config $Config
        _QCQJ-EnsureLayout -Root $root

        $from = _QCQJ-NormalizeState $FromState
        $to = _QCQJ-NormalizeState $ToState
        if (@('pending', 'running', 'succeeded', 'failed') -notcontains $from) {
            return New-QCFailureResult -Code 'QUEUE_INVALID_STATE' -Message "Invalid FromState: $FromState" -Data @{ fromState = $FromState; normalized = $from }
        }
        if (@('pending', 'running', 'succeeded', 'failed') -notcontains $to) {
            return New-QCFailureResult -Code 'QUEUE_INVALID_STATE' -Message "Invalid ToState: $ToState" -Data @{ toState = $ToState; normalized = $to }
        }
        if ($PSBoundParameters.ContainsKey('Job') -and $Job) {
            $jobIdInJob = [string]$Job.id
            if (-not [string]::IsNullOrWhiteSpace($jobIdInJob) -and $jobIdInJob -ne $JobId) {
                return New-QCFailureResult -Code 'QUEUE_JOB_ID_MISMATCH' -Message 'Job.id does not match JobId.' -Data @{ jobId = $JobId; jobObjectId = $jobIdInJob }
            }
        }

        $lockPath = _QCQJ-QueueWriteLockPath -Root $root
        $lk = _QCQJ-GetLockAcquireSettings -Config $Config
        if (-not (_QCQJ-AcquireLockFile -LockPath $lockPath -TimeoutMs $lk.TimeoutMs -SleepMs $lk.SleepMs)) {
            return New-QCFailureResult -Code 'QUEUE_LOCK_TIMEOUT' -Message 'Timed out acquiring queue write lock.' -Data @{ lockPath = $lockPath }
        }
        try {
            $src = _QCQJ-JobFilePath -Root $root -State $from -JobId $JobId
            if (-not (Test-Path -LiteralPath $src)) {
                return New-QCFailureResult -Code 'QUEUE_JOB_NOT_FOUND' -Message 'Job file not found in source state.' -Data @{ jobId = $JobId; fromState = $from; path = $src }
            }
            $dst = _QCQJ-JobFilePath -Root $root -State $to -JobId $JobId

            if ($PSBoundParameters.ContainsKey('Job') -and $Job) {
                # Caller supplied the authoritative job content (e.g. with result/lastError).
                # Stamp status + updatedAtUtc and write to source path before the rename.
                $Job.status = $to
                $Job.updatedAtUtc = (Get-QCTimestamp)
                if (-not $Job.ContainsKey('id') -or [string]::IsNullOrWhiteSpace([string]$Job.id)) { $Job.id = $JobId }
                _QCQJ-WriteJobFileAtomic -Path $src -Job $Job
            } else {
                $existing = _QCQJ-ReadJobFile -Path $src
                $existing.status = $to
                $existing.updatedAtUtc = (Get-QCTimestamp)
                _QCQJ-WriteJobFileAtomic -Path $src -Job $existing
            }
            _QCQJ-MoveItemWithRetry -LiteralPath $src -Destination $dst
            try {
                $dedupeKey = $null
                if ($PSBoundParameters.ContainsKey('Job') -and $Job) { $dedupeKey = [string]$Job['dedupeKey'] }
                if (-not $dedupeKey) {
                    try {
                        $readBack = _QCQJ-ReadJobFile -Path $dst
                        if ($readBack -and $readBack.ContainsKey('dedupeKey')) { $dedupeKey = [string]$readBack['dedupeKey'] }
                    } catch { }
                }
                _QCQJ-UpdateDedupeIndexForJob -Config $Config -Root $root -DedupeKey $dedupeKey -JobId $JobId -State $to
            } catch { }
            return New-QCSuccessResult -Code 'QUEUE_JOB_MOVED' -Message 'Job moved between states.' -Data @{ jobId = $JobId; fromState = $from; toState = $to }
        } finally {
            _QCQJ-ReleaseLockFile -LockPath $lockPath
        }
    } catch {
        $inner = ''
        try { $inner = [string]$_.Exception.Message } catch { }
        if ([string]::IsNullOrWhiteSpace($inner)) { try { $inner = [string]$_ } catch { $inner = 'unknown error' } }
        return New-QCFailureResult -Code 'QUEUE_MOVE_ERROR' -Message ('Failed to move job between states: ' + $inner) -Data @{ jobId = $JobId; error = $_; innerMessage = $inner }
    }
}

function Lock-QCJob {
    <#
    .SYNOPSIS
    Acquires a processing lock for a job.
    .DESCRIPTION
    Creates/claims lock marker for exclusive job processing.
    .PARAMETER JobId
    Job identifier.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: local lock-file write.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$JobId,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    try {
        $root = _QCQJ-GetQueueRoot -Config $Config
        _QCQJ-EnsureLayout -Root $root

        $jobLock = _QCQJ-LockFilePath -Root $root -JobId $JobId
        $lk = _QCQJ-GetLockAcquireSettings -Config $Config
        # Do NOT short-circuit on Test-Path here. _QCQJ-AcquireLockFile already
        # detects orphan locks held by dead PIDs and steals them. A naive
        # Test-Path check would incorrectly mark such locks as held and stall
        # the queue (the original cause of "WORKER_LOCK_RACE" hot-spinning
        # after a crashed worker). Live owners are still respected because the
        # acquire loop only steals when _QCQJ-IsLockOwnerDead returns true.
        if (-not (_QCQJ-AcquireLockFile -LockPath $jobLock -TimeoutMs $lk.TimeoutMs -SleepMs $lk.SleepMs)) {
            return New-QCFailureResult -Code 'QUEUE_LOCK_TIMEOUT' -Message 'Timed out acquiring job lock.' -Data @{ jobId = $JobId; lockPath = $jobLock }
        }

        # Transition pending -> running on lock acquire (prevents repeated selection).
        $lockPath = _QCQJ-QueueWriteLockPath -Root $root
        if (-not (_QCQJ-AcquireLockFile -LockPath $lockPath -TimeoutMs $lk.TimeoutMs -SleepMs $lk.SleepMs)) {
            _QCQJ-ReleaseLockFile -LockPath $jobLock
            return New-QCFailureResult -Code 'QUEUE_LOCK_TIMEOUT' -Message 'Timed out acquiring queue write lock.' -Data @{ lockPath = $lockPath }
        }
        try {
            # IMPORTANT: a worker may have "selected" this job from pending\, then
            # lost the per-job lock race and only acquired the lock after the winner
            # already moved the job to succeeded\ or failed\. In that case, treat the
            # late lock acquisition as a no-op and release it immediately so we do
            # not double-process the job.
            $loc = _QCQJ-FindJobFile -Root $root -JobId $JobId
            if (-not $loc) {
                _QCQJ-ReleaseLockFile -LockPath $jobLock
                return New-QCFailureResult -Code 'QUEUE_JOB_NOT_FOUND' -Message 'Job not found while acquiring lock.' -Data @{ jobId = $JobId }
            }
            if ($loc.state -ne 'pending') {
                _QCQJ-ReleaseLockFile -LockPath $jobLock
                return New-QCFailureResult -Code 'QUEUE_JOB_ALREADY_MOVED' -Message 'Job is no longer pending; skipping lock.' -Data @{ jobId = $JobId; state = $loc.state; path = $loc.path }
            }

            $pending = _QCQJ-JobFilePath -Root $root -State 'pending' -JobId $JobId
            $running = _QCQJ-JobFilePath -Root $root -State 'running' -JobId $JobId

            $job = _QCQJ-ReadJobFile -Path $pending
            $job.status = 'running'
            $job.startedAtUtc = (Get-QCTimestamp)
            _QCQJ-WriteJobFileAtomic -Path $pending -Job $job
            Move-Item -LiteralPath $pending -Destination $running -Force -ErrorAction Stop
        } finally {
            _QCQJ-ReleaseLockFile -LockPath $lockPath
        }

        return New-QCSuccessResult -Code 'QUEUE_LOCK_ACQUIRED' -Message 'Job lock acquired.' -Data @{ jobId = $JobId; lockPath = $jobLock }
    } catch {
        return New-QCFailureResult -Code 'QUEUE_LOCK_ERROR' -Message 'Failed to lock job.' -Data @{ jobId = $JobId; error = $_ }
    }
}

function Unlock-QCJob {
    <#
    .SYNOPSIS
    Releases a processing lock for a job.
    .DESCRIPTION
    Removes/clears lock marker for a processed or abandoned job.
    .PARAMETER JobId
    Job identifier.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: local lock-file delete/write.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$JobId,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    try {
        $root = _QCQJ-GetQueueRoot -Config $Config
        _QCQJ-EnsureLayout -Root $root

        $jobLock = _QCQJ-LockFilePath -Root $root -JobId $JobId
        _QCQJ-ReleaseLockFile -LockPath $jobLock
        return New-QCSuccessResult -Code 'QUEUE_LOCK_RELEASED' -Message 'Job lock released.' -Data @{ jobId = $JobId; lockPath = $jobLock }
    } catch {
        return New-QCFailureResult -Code 'QUEUE_UNLOCK_ERROR' -Message 'Failed to unlock job.' -Data @{ jobId = $JobId; error = $_ }
    }
}

function Recover-QCStaleJobs {
    <#
    .SYNOPSIS
    Recovers stale processing jobs.
    .DESCRIPTION
    Identifies stale processing items and returns/reassigns them per policy.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: local queue state writes/moves.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    try {
        $root = _QCQJ-GetQueueRoot -Config $Config
        _QCQJ-EnsureLayout -Root $root

        $staleSeconds = 900
        $maxAttempts = 3
        if ($Config.ContainsKey('queue') -and $Config.queue) {
            if ($Config.queue.ContainsKey('recover') -and $Config.queue.recover) {
                if ($Config.queue.recover.ContainsKey('staleSeconds') -and $Config.queue.recover.staleSeconds) { $staleSeconds = [int]$Config.queue.recover.staleSeconds }
                if ($Config.queue.recover.ContainsKey('maxAttempts') -and $Config.queue.recover.maxAttempts) { $maxAttempts = [int]$Config.queue.recover.maxAttempts }
            }
        }

        $runningDir = Join-Path $root 'running'
        $files = @(Get-ChildItem -LiteralPath $runningDir -Filter '*.json' -File -ErrorAction SilentlyContinue)
        $now = [DateTime]::UtcNow

        $result = @{
            staleSeconds = $staleSeconds
            maxAttempts  = $maxAttempts
            scanned      = $files.Count
            recoveredToPending = 0
            recoveredToFailed  = 0
            recoveredOrphan    = 0
            skippedNotStale    = 0
            orphanLocksRemoved = 0
            writeLockAcquireFailures = 0
            details            = @()
        }

        $writeLock = _QCQJ-QueueWriteLockPath -Root $root
        $lk = _QCQJ-GetLockAcquireSettings -Config $Config

        # Phase 1 (no global write lock): classify candidates so workers are not starved
        # while we read every running\ payload / scan locks.
        $recoverCandidates = @()
        foreach ($f in $files) {
            $jobId = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
            $job = _QCQJ-ReadJobFile -Path $f.FullName

            $started = $null
            if ($job.ContainsKey('startedAtUtc') -and $job.startedAtUtc) {
                try { $started = [DateTime]::Parse([string]$job.startedAtUtc).ToUniversalTime() } catch { $started = $null }
            }
            if (-not $started) { $started = $f.LastWriteTimeUtc }

            $age = ($now - $started).TotalSeconds

            $lockPathJob = _QCQJ-LockFilePath -Root $root -JobId $jobId
            $isOrphan = $false
            $orphanReason = $null
            if (-not (Test-Path -LiteralPath $lockPathJob)) {
                $isOrphan = $true
                $orphanReason = 'NO_LOCK_FILE'
            } else {
                $payload = $null
                try { $payload = Get-Content -LiteralPath $lockPathJob -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json -ErrorAction SilentlyContinue } catch { $payload = $null }
                $ownerPid = 0
                if ($payload -and $payload.PSObject.Properties.Name -contains 'pid') {
                    try { $ownerPid = [int]$payload.pid } catch { $ownerPid = 0 }
                }
                if ($ownerPid -gt 0) {
                    $proc = Get-Process -Id $ownerPid -ErrorAction SilentlyContinue
                    if (-not $proc) { $isOrphan = $true; $orphanReason = "DEAD_PID($ownerPid)" }
                }
            }

            if (-not $isOrphan -and $age -lt $staleSeconds) {
                $result.skippedNotStale++
                continue
            }

            $recoverCandidates += ,@($f.FullName, $jobId)
        }

        # Phase 2: one short global write-lock window per job so Move-QCJob / workers can interleave.
        foreach ($pair in $recoverCandidates) {
            $jobId = [string]$pair[1]
            if (-not (_QCQJ-AcquireLockFile -LockPath $writeLock -TimeoutMs $lk.TimeoutMs -SleepMs $lk.SleepMs)) {
                $result.writeLockAcquireFailures++
                $result.details += @{ jobId = $jobId; action = 'skipped_lock_timeout' }
                continue
            }
            try {
                $src = _QCQJ-JobFilePath -Root $root -State 'running' -JobId $jobId
                if (-not (Test-Path -LiteralPath $src)) { continue }

                $job = _QCQJ-ReadJobFile -Path $src
                $f2 = Get-Item -LiteralPath $src -ErrorAction Stop

                $started2 = $null
                if ($job.ContainsKey('startedAtUtc') -and $job.startedAtUtc) {
                    try { $started2 = [DateTime]::Parse([string]$job.startedAtUtc).ToUniversalTime() } catch { $started2 = $null }
                }
                if (-not $started2) { $started2 = $f2.LastWriteTimeUtc }
                $age2 = ($now - $started2).TotalSeconds

                $lockPathJob2 = _QCQJ-LockFilePath -Root $root -JobId $jobId
                $isOrphan2 = $false
                $orphanReason2 = $null
                if (-not (Test-Path -LiteralPath $lockPathJob2)) {
                    $isOrphan2 = $true
                    $orphanReason2 = 'NO_LOCK_FILE'
                } else {
                    $payload2 = $null
                    try { $payload2 = Get-Content -LiteralPath $lockPathJob2 -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json -ErrorAction SilentlyContinue } catch { $payload2 = $null }
                    $ownerPid2 = 0
                    if ($payload2 -and $payload2.PSObject.Properties.Name -contains 'pid') {
                        try { $ownerPid2 = [int]$payload2.pid } catch { $ownerPid2 = 0 }
                    }
                    if ($ownerPid2 -gt 0) {
                        $proc2 = Get-Process -Id $ownerPid2 -ErrorAction SilentlyContinue
                        if (-not $proc2) { $isOrphan2 = $true; $orphanReason2 = "DEAD_PID($ownerPid2)" }
                    }
                }

                if (-not $isOrphan2 -and $age2 -lt $staleSeconds) { continue }

                $attempts = 0
                if ($job.ContainsKey('attempts') -and $job.attempts -ne $null) { $attempts = [int]$job.attempts }
                $attempts++
                $job.attempts = $attempts
                $job.updatedAtUtc = (ConvertTo-QCTimestamp -DateTime $now)

                _QCQJ-ReleaseLockFile -LockPath $lockPathJob2

                if ($attempts -ge $maxAttempts) {
                    $job.status = 'failed'
                    _QCQJ-WriteJobFileAtomic -Path $src -Job $job
                    $dst = _QCQJ-JobFilePath -Root $root -State 'failed' -JobId $jobId
                    Move-Item -LiteralPath $src -Destination $dst -Force -ErrorAction Stop
                    $result.recoveredToFailed++
                    $entry = @{ jobId = $jobId; action = 'failed'; attempts = $attempts; ageSeconds = [int]$age2 }
                    if ($isOrphan2) { $entry.orphan = $true; $entry.orphanReason = $orphanReason2; $result.recoveredOrphan++ }
                    $result.details += $entry
                } else {
                    $job.status = 'pending'
                    $job.startedAtUtc = $null
                    _QCQJ-WriteJobFileAtomic -Path $src -Job $job
                    $dst = _QCQJ-JobFilePath -Root $root -State 'pending' -JobId $jobId
                    Move-Item -LiteralPath $src -Destination $dst -Force -ErrorAction Stop
                    $result.recoveredToPending++
                    $entry = @{ jobId = $jobId; action = 'requeued'; attempts = $attempts; ageSeconds = [int]$age2 }
                    if ($isOrphan2) { $entry.orphan = $true; $entry.orphanReason = $orphanReason2; $result.recoveredOrphan++ }
                    $result.details += $entry
                }
            } catch {
                $result.details += @{ jobId = $jobId; action = 'error'; error = [string]$_.Exception.Message }
            } finally {
                _QCQJ-ReleaseLockFile -LockPath $writeLock
            }
        }

        # Janitor: remove orphan lock files (brief lock per cleanup op).
        $locksDir = Join-Path $root 'locks'
        if (Test-Path -LiteralPath $locksDir) {
            $lockFiles = @(Get-ChildItem -LiteralPath $locksDir -Filter '*.lock' -File -ErrorAction SilentlyContinue)
            foreach ($lf in $lockFiles) {
                $name = $lf.Name
                if ($name -ieq '_queue_write.lock') {
                    if (-not (_QCQJ-AcquireLockFile -LockPath $writeLock -TimeoutMs $lk.TimeoutMs -SleepMs $lk.SleepMs)) {
                        $result.writeLockAcquireFailures++
                        continue
                    }
                    try {
                        if (_QCQJ-IsLockOwnerDead -LockPath $lf.FullName) {
                            try { Remove-Item -LiteralPath $lf.FullName -Force -ErrorAction Stop; $result.orphanLocksRemoved++ } catch { }
                        }
                    } finally {
                        _QCQJ-ReleaseLockFile -LockPath $writeLock
                    }
                    continue
                }
                $jobIdL = [System.IO.Path]::GetFileNameWithoutExtension($name)
                $runningPath = _QCQJ-JobFilePath -Root $root -State 'running' -JobId $jobIdL
                if (Test-Path -LiteralPath $runningPath) { continue }
                if (-not (_QCQJ-IsLockOwnerDead -LockPath $lf.FullName)) { continue }
                if (-not (_QCQJ-AcquireLockFile -LockPath $writeLock -TimeoutMs $lk.TimeoutMs -SleepMs $lk.SleepMs)) {
                    $result.writeLockAcquireFailures++
                    continue
                }
                try {
                    Remove-Item -LiteralPath $lf.FullName -Force -ErrorAction Stop
                    $result.orphanLocksRemoved++
                } catch { }
                finally {
                    _QCQJ-ReleaseLockFile -LockPath $writeLock
                }
            }
        }

        return New-QCSuccessResult -Code 'QUEUE_RECOVERY_OK' -Message 'Stale running jobs recovery completed.' -Data $result
    } catch {
        return New-QCFailureResult -Code 'QUEUE_RECOVERY_ERROR' -Message 'Failed to recover stale jobs.' -Data @{ error = $_ }
    }
}

function Invoke-QCQueueStartupCheck {
    <#
    .SYNOPSIS
    Startup hygiene: queue stats, stale/orphan running-job recovery, optional watcher-active clear.
    .DESCRIPTION
    Call once before the first watcher tick or worker loop so stuck running\ jobs from a prior
    session are requeued/failed and operators can see pending backlog counts.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [switch]$ClearWatcherActive
    )

    $out = @{
        queueStats = $null
        recovery   = $null
        errors     = [System.Collections.Generic.List[string]]::new()
    }

    try {
        $stats = Get-QCQueueStats -Config $Config
        if ($stats.IsSuccess) { $out.queueStats = $stats.Data } else { [void]$out.errors.Add([string]$stats.Message) }
    } catch {
        [void]$out.errors.Add([string]$_.Exception.Message)
    }

    try {
        $rec = Recover-QCStaleJobs -Config $Config
        if ($rec.IsSuccess) { $out.recovery = $rec.Data } else { [void]$out.errors.Add([string]$rec.Message) }
    } catch {
        [void]$out.errors.Add([string]$_.Exception.Message)
    }

    if ($ClearWatcherActive) {
        try { Clear-QCWatcherActive -Config $Config | Out-Null } catch {
            [void]$out.errors.Add([string]$_.Exception.Message)
        }
    }

    return New-QCSuccessResult -Code 'QUEUE_STARTUP_CHECK_OK' -Message 'Queue startup check completed.' -Data $out
}

function Test-QCDuplicateJob {
    <#
    .SYNOPSIS
    Checks if a dedupe key already exists.
    .DESCRIPTION
    Looks for recently queued/processed jobs with matching dedupe key.
    .PARAMETER DedupeKey
    Computed dedupe key.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: file read only.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DedupeKey,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    try {
        if (_QCQJ-IsNullOrWhiteSpace $DedupeKey) {
            return New-QCFailureResult -Code 'QUEUE_VALIDATION_MISSING_DEDUPE_KEY' -Message 'DedupeKey is required.' -Data @{}
        }

        $root = _QCQJ-GetQueueRoot -Config $Config
        _QCQJ-EnsureLayout -Root $root

        $idxPath = _QCQJ-DedupeIndexPath -Root $root
        $idx = _QCQJ-ReadDedupeIndex -Path $idxPath
        if ($idx -and $idx.entries -and $idx.entries.ContainsKey($DedupeKey)) {
            $hit = $idx.entries[$DedupeKey]
            $state = ''
            $jobId = ''
            try { $state = [string]$hit.state } catch { $state = '' }
            try { $jobId = [string]$hit.jobId } catch { $jobId = '' }
            return New-QCSuccessResult -Code 'QUEUE_DEDUPE_INDEX_HIT' -Message 'Duplicate check complete (index hit).' -Data @{
                dedupeKey = $DedupeKey
                isDuplicate = $true
                matches = @(@{ jobId = $jobId; state = $state; path = '' })
                indexPath = $idxPath
            }
        }

        $matches = @()
        foreach ($s in @('pending', 'running', 'succeeded', 'failed')) {
            $dir = Join-Path $root $s
            $files = @(Get-ChildItem -LiteralPath $dir -Filter '*.json' -File -ErrorAction SilentlyContinue)
            foreach ($f in $files) {
                $job = $null
                try { $job = _QCQJ-ReadJobFile -Path $f.FullName } catch { continue }
                if ($job -and ([string]$job['dedupeKey'] -eq $DedupeKey)) {
                    $matches += @{ jobId = [string]$job['id']; state = $s; path = $f.FullName }
                }
            }
        }

        if ($matches.Count -gt 0) {
            try {
                # Best-effort: self-heal index from scan result
                $m0 = $matches | Select-Object -First 1
                _QCQJ-UpdateDedupeIndexForJob -Config $Config -Root $root -DedupeKey $DedupeKey -JobId ([string]$m0.jobId) -State ([string]$m0.state)
            } catch { }
        }

        return New-QCSuccessResult -Code 'QUEUE_DEDUPE_CHECKED' -Message 'Duplicate check complete.' -Data @{
            dedupeKey = $DedupeKey
            isDuplicate = ($matches.Count -gt 0)
            matches = $matches
        }
    } catch {
        return New-QCFailureResult -Code 'QUEUE_DEDUPE_ERROR' -Message 'Failed to check for duplicate jobs.' -Data @{ error = $_ }
    }
}

function Get-QCQueueStats {
    <#
    .SYNOPSIS
    Returns queue state counts and summary stats.
    .DESCRIPTION
    Produces lightweight queue metrics for health checks/dashboarding.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: file read only.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    try {
        $root = _QCQJ-GetQueueRoot -Config $Config
        _QCQJ-EnsureLayout -Root $root

        $counts = @{}
        foreach ($s in @('pending', 'running', 'succeeded', 'failed')) {
            $dir = Join-Path $root $s
            $counts[$s] = @(Get-ChildItem -LiteralPath $dir -Filter '*.json' -File -ErrorAction SilentlyContinue).Count
        }
        $locksDir = Join-Path $root 'locks'
        $lockCount = @(Get-ChildItem -LiteralPath $locksDir -Filter '*.lock' -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne '_queue_write.lock' }).Count

        return New-QCSuccessResult -Code 'QUEUE_STATS' -Message 'Queue stats collected.' -Data @{
            root = $root
            states = $counts
            locks = @{ count = $lockCount }
        }
    } catch {
        return New-QCFailureResult -Code 'QUEUE_STATS_ERROR' -Message 'Failed to collect queue stats.' -Data @{ error = $_ }
    }
}

function Get-QCRecentJobs {
    <#
    .SYNOPSIS
    Retrieves recent jobs across queue states.
    .DESCRIPTION
    Returns a bounded list of recent job records for operational visibility.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .PARAMETER Limit
    Maximum jobs to return.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: file read only.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [int]$Limit = 100
    )

    try {
        $root = _QCQJ-GetQueueRoot -Config $Config
        _QCQJ-EnsureLayout -Root $root

        $items = @()
        foreach ($s in @('pending', 'running', 'succeeded', 'failed')) {
            $dir = Join-Path $root $s
            $files = @(Get-ChildItem -LiteralPath $dir -Filter '*.json' -File -ErrorAction SilentlyContinue)
            foreach ($f in $files) {
                $items += [pscustomobject]@{
                    state = $s
                    path  = $f.FullName
                    lastWriteTimeUtc = $f.LastWriteTimeUtc
                }
            }
        }

        $recent = @($items | Sort-Object -Property lastWriteTimeUtc -Descending | Select-Object -First $Limit)
        $jobs = @()
        foreach ($r in $recent) {
            try {
                $job = _QCQJ-ReadJobFile -Path $r.path
                if ($job) {
                    $jobs += @{
                        state = $r.state
                        job = $job
                        lastWriteTimeUtc = $r.lastWriteTimeUtc.ToString('o')
                    }
                }
            } catch {
                continue
            }
        }

        return New-QCSuccessResult -Code 'QUEUE_RECENT_JOBS' -Message 'Recent jobs loaded.' -Data @{ jobs = $jobs; limit = $Limit }
    } catch {
        return New-QCFailureResult -Code 'QUEUE_RECENT_ERROR' -Message 'Failed to get recent jobs.' -Data @{ error = $_ }
    }
}
Export-ModuleMember -Function *
