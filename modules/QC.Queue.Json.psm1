# QC.Queue.Json.psm1
# Responsibility: JSON-backed queue persistence, lifecycle transitions, and queue reporting.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force

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

function _QCQJ-AcquireLockFile([string]$LockPath, [int]$TimeoutMs = 5000, [int]$SleepMs = 100) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $stealAttempted = $false
    while ($sw.ElapsedMilliseconds -lt $TimeoutMs) {
        try {
            $fs = [System.IO.File]::Open($LockPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
            try {
                $payload = @{
                    pid = $PID
                    createdAtUtc = ([DateTime]::UtcNow.ToString('o'))
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
            try { return ([datetime]$Value).ToUniversalTime().ToString('o') } catch { return [string]$Value }
        }
        if ($Value -is [datetimeoffset]) {
            try { return ([datetimeoffset]$Value).UtcDateTime.ToString('o') } catch { return [string]$Value }
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
        [int]$Attempts = 6,
        [int]$SleepMs = 120
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
        if (-not (_QCQJ-AcquireLockFile -LockPath $lockPath)) {
            return New-QCFailureResult -Code 'QUEUE_LOCK_TIMEOUT' -Message 'Timed out acquiring queue write lock.' -Data @{ lockPath = $lockPath }
        }
        try {
            $dest = _QCQJ-JobFilePath -Root $root -State 'pending' -JobId $jobId
            if (Test-Path -LiteralPath $dest) {
                return New-QCFailureResult -Code 'QUEUE_JOB_ALREADY_EXISTS' -Message 'Job already exists in pending state.' -Data @{ jobId = $jobId; path = $dest }
            }
            $Job.status = 'pending'
            $Job.enqueuedAtUtc = ([DateTime]::UtcNow.ToString('o'))
            _QCQJ-WriteJobFileAtomic -Path $dest -Job $Job
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
            createdAtUtc = ([DateTime]::UtcNow.ToString('o'))
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
        if (-not (_QCQJ-AcquireLockFile -LockPath $lockPath)) {
            return New-QCFailureResult -Code 'QUEUE_LOCK_TIMEOUT' -Message 'Timed out acquiring queue write lock.' -Data @{ lockPath = $lockPath }
        }
        try {
            $loc = _QCQJ-FindJobFile -Root $root -JobId $JobId
            if (-not $loc) {
                return New-QCFailureResult -Code 'QUEUE_JOB_NOT_FOUND' -Message 'Job not found for status update.' -Data @{ jobId = $JobId }
            }
            $job = _QCQJ-ReadJobFile -Path $loc.path
            $job.status = $Status
            $job.updatedAtUtc = ([DateTime]::UtcNow.ToString('o'))
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
        if (-not (_QCQJ-AcquireLockFile -LockPath $lockPath)) {
            return New-QCFailureResult -Code 'QUEUE_LOCK_TIMEOUT' -Message 'Timed out acquiring queue write lock.' -Data @{ lockPath = $lockPath }
        }
        try {
            $loc = _QCQJ-FindJobFile -Root $root -JobId $jobId
            if (-not $loc) {
                return New-QCFailureResult -Code 'QUEUE_JOB_NOT_FOUND' -Message 'Job not found for update.' -Data @{ jobId = $jobId }
            }
            $Job.updatedAtUtc = ([DateTime]::UtcNow.ToString('o'))
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
        if (-not (_QCQJ-AcquireLockFile -LockPath $lockPath)) {
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
                $Job.updatedAtUtc = ([DateTime]::UtcNow.ToString('o'))
                if (-not $Job.ContainsKey('id') -or [string]::IsNullOrWhiteSpace([string]$Job.id)) { $Job.id = $JobId }
                _QCQJ-WriteJobFileAtomic -Path $src -Job $Job
            } else {
                $existing = _QCQJ-ReadJobFile -Path $src
                $existing.status = $to
                $existing.updatedAtUtc = ([DateTime]::UtcNow.ToString('o'))
                _QCQJ-WriteJobFileAtomic -Path $src -Job $existing
            }
            _QCQJ-MoveItemWithRetry -LiteralPath $src -Destination $dst
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
        # Do NOT short-circuit on Test-Path here. _QCQJ-AcquireLockFile already
        # detects orphan locks held by dead PIDs and steals them. A naive
        # Test-Path check would incorrectly mark such locks as held and stall
        # the queue (the original cause of "WORKER_LOCK_RACE" hot-spinning
        # after a crashed worker). Live owners are still respected because the
        # acquire loop only steals when _QCQJ-IsLockOwnerDead returns true.
        if (-not (_QCQJ-AcquireLockFile -LockPath $jobLock)) {
            return New-QCFailureResult -Code 'QUEUE_LOCK_TIMEOUT' -Message 'Timed out acquiring job lock.' -Data @{ jobId = $JobId; lockPath = $jobLock }
        }

        # Transition pending -> running on lock acquire (prevents repeated selection).
        $lockPath = _QCQJ-QueueWriteLockPath -Root $root
        if (-not (_QCQJ-AcquireLockFile -LockPath $lockPath)) {
            _QCQJ-ReleaseLockFile -LockPath $jobLock
            return New-QCFailureResult -Code 'QUEUE_LOCK_TIMEOUT' -Message 'Timed out acquiring queue write lock.' -Data @{ lockPath = $lockPath }
        }
        try {
            $pending = _QCQJ-JobFilePath -Root $root -State 'pending' -JobId $JobId
            $running = _QCQJ-JobFilePath -Root $root -State 'running' -JobId $JobId

            if (Test-Path -LiteralPath $pending) {
                $job = _QCQJ-ReadJobFile -Path $pending
                $job.status = 'running'
                $job.startedAtUtc = ([DateTime]::UtcNow.ToString('o'))
                _QCQJ-WriteJobFileAtomic -Path $pending -Job $job
                Move-Item -LiteralPath $pending -Destination $running -Force -ErrorAction Stop
            }
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
            details            = @()
        }

        $writeLock = _QCQJ-QueueWriteLockPath -Root $root
        if (-not (_QCQJ-AcquireLockFile -LockPath $writeLock)) {
            return New-QCFailureResult -Code 'QUEUE_LOCK_TIMEOUT' -Message 'Timed out acquiring queue write lock.' -Data @{ lockPath = $writeLock }
        }
        try {
            foreach ($f in $files) {
                $jobId = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
                $job = _QCQJ-ReadJobFile -Path $f.FullName

                $started = $null
                if ($job.ContainsKey('startedAtUtc') -and $job.startedAtUtc) {
                    try { $started = [DateTime]::Parse([string]$job.startedAtUtc).ToUniversalTime() } catch { $started = $null }
                }
                if (-not $started) { $started = $f.LastWriteTimeUtc }

                $age = ($now - $started).TotalSeconds

                # Orphan detection: if the per-job lock file is missing OR the owner PID is dead,
                # immediately reclaim regardless of staleSeconds. Workers were killed/crashed.
                $lock = _QCQJ-LockFilePath -Root $root -JobId $jobId
                $isOrphan = $false
                $orphanReason = $null
                if (-not (Test-Path -LiteralPath $lock)) {
                    $isOrphan = $true
                    $orphanReason = 'NO_LOCK_FILE'
                } else {
                    $payload = $null
                    try { $payload = Get-Content -LiteralPath $lock -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json -ErrorAction SilentlyContinue } catch { $payload = $null }
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

                $attempts = 0
                if ($job.ContainsKey('attempts') -and $job.attempts -ne $null) { $attempts = [int]$job.attempts }
                $attempts++
                $job.attempts = $attempts
                $job.updatedAtUtc = $now.ToString('o')

                _QCQJ-ReleaseLockFile -LockPath $lock

                if ($attempts -ge $maxAttempts) {
                    $job.status = 'failed'
                    _QCQJ-WriteJobFileAtomic -Path $f.FullName -Job $job
                    $dst = _QCQJ-JobFilePath -Root $root -State 'failed' -JobId $jobId
                    Move-Item -LiteralPath $f.FullName -Destination $dst -Force -ErrorAction Stop
                    $result.recoveredToFailed++
                    $entry = @{ jobId = $jobId; action = 'failed'; attempts = $attempts; ageSeconds = [int]$age }
                    if ($isOrphan) { $entry.orphan = $true; $entry.orphanReason = $orphanReason; $result.recoveredOrphan++ }
                    $result.details += $entry
                } else {
                    $job.status = 'pending'
                    $job.startedAtUtc = $null
                    _QCQJ-WriteJobFileAtomic -Path $f.FullName -Job $job
                    $dst = _QCQJ-JobFilePath -Root $root -State 'pending' -JobId $jobId
                    Move-Item -LiteralPath $f.FullName -Destination $dst -Force -ErrorAction Stop
                    $result.recoveredToPending++
                    $entry = @{ jobId = $jobId; action = 'requeued'; attempts = $attempts; ageSeconds = [int]$age }
                    if ($isOrphan) { $entry.orphan = $true; $entry.orphanReason = $orphanReason; $result.recoveredOrphan++ }
                    $result.details += $entry
                }
            }

            # Janitor: also remove orphan lock files that have no corresponding
            # job in running\ but DO live alongside a pending\ job whose owner
            # PID is dead. Without this, Get-NextQCJob (and the live worker's
            # AcquireLockFile self-heal) handle the simple cases, but if the
            # job was never selected after the orphan was created, the lock
            # would persist until first attempted re-selection. Cleaning here
            # keeps the locks/ directory tidy and matches operator expectation
            # that a single Recover sweep returns the queue to a clean state.
            $locksDir = Join-Path $root 'locks'
            if (Test-Path -LiteralPath $locksDir) {
                $lockFiles = @(Get-ChildItem -LiteralPath $locksDir -Filter '*.lock' -File -ErrorAction SilentlyContinue)
                foreach ($lf in $lockFiles) {
                    $name = $lf.Name
                    if ($name -ieq '_queue_write.lock') {
                        # Global write lock: clean only if dead.
                        if (_QCQJ-IsLockOwnerDead -LockPath $lf.FullName) {
                            try { Remove-Item -LiteralPath $lf.FullName -Force -ErrorAction Stop; $result.orphanLocksRemoved++ } catch { }
                        }
                        continue
                    }
                    $jobId = [System.IO.Path]::GetFileNameWithoutExtension($name)
                    $runningPath = _QCQJ-JobFilePath -Root $root -State 'running' -JobId $jobId
                    if (Test-Path -LiteralPath $runningPath) { continue }   # handled by per-job loop above
                    if (-not (_QCQJ-IsLockOwnerDead -LockPath $lf.FullName)) { continue }
                    try {
                        Remove-Item -LiteralPath $lf.FullName -Force -ErrorAction Stop
                        $result.orphanLocksRemoved++
                    } catch { }
                }
            }
        } finally {
            _QCQJ-ReleaseLockFile -LockPath $writeLock
        }

        return New-QCSuccessResult -Code 'QUEUE_RECOVERY_OK' -Message 'Stale running jobs recovery completed.' -Data $result
    } catch {
        return New-QCFailureResult -Code 'QUEUE_RECOVERY_ERROR' -Message 'Failed to recover stale jobs.' -Data @{ error = $_ }
    }
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
