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

function _QCQJ-AcquireLockFile([string]$LockPath, [int]$TimeoutMs = 5000, [int]$SleepMs = 100) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
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
            Start-Sleep -Milliseconds $SleepMs
        }
    }
    return $false
}

function _QCQJ-ReleaseLockFile([string]$LockPath) {
    Remove-Item -LiteralPath $LockPath -Force -ErrorAction SilentlyContinue
}

function _QCQJ-ReadJobFile([string]$Path) {
    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
    $obj = $raw | ConvertFrom-Json -ErrorAction Stop

    function _ToHashtable([object]$Value) {
        if ($null -eq $Value) { return $null }
        if ($Value -is [string]) { return $Value }
        if ($Value -is [System.ValueType]) { return $Value }
        if ($Value -is [System.Collections.IDictionary]) { return $Value }
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
    $json = $Job | ConvertTo-Json -Depth 50
    Set-Content -LiteralPath $tmp -Value $json -Encoding utf8 -ErrorAction Stop
    Move-Item -LiteralPath $tmp -Destination $Path -Force -ErrorAction Stop
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
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: none (selection only).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    try {
        $root = _QCQJ-GetQueueRoot -Config $Config
        _QCQJ-EnsureLayout -Root $root

        $pendingDir = Join-Path $root 'pending'
        $files = @(Get-ChildItem -LiteralPath $pendingDir -Filter '*.json' -File -ErrorAction SilentlyContinue | Sort-Object -Property LastWriteTimeUtc, Name)

        foreach ($f in $files) {
            $jobId = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
            $lock = _QCQJ-LockFilePath -Root $root -JobId $jobId
            if (Test-Path -LiteralPath $lock) { continue }

            $job = _QCQJ-ReadJobFile -Path $f.FullName
            return New-QCSuccessResult -Code 'QUEUE_NEXT_JOB' -Message 'Next pending job selected.' -Data @{ job = $job; jobId = $jobId; state = 'pending' }
        }

        return New-QCSuccessResult -Code 'QUEUE_EMPTY' -Message 'No eligible pending jobs.' -Data @{ job = $null }
    } catch {
        return New-QCFailureResult -Code 'QUEUE_GET_NEXT_ERROR' -Message 'Failed to select next job.' -Data @{ error = $_ }
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

function Move-QCJob {
    <#
    .SYNOPSIS
    Moves a job between queue state buckets.
    .DESCRIPTION
    Performs state-folder transition for a job record.
    .PARAMETER JobId
    Job identifier.
    .PARAMETER FromState
    Current queue state.
    .PARAMETER ToState
    Destination queue state.
    .PARAMETER Config
    Loaded app configuration hashtable.
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
        [hashtable]$Config
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

            $job = _QCQJ-ReadJobFile -Path $src
            $job.status = $to
            $job.updatedAtUtc = ([DateTime]::UtcNow.ToString('o'))
            _QCQJ-WriteJobFileAtomic -Path $src -Job $job
            Move-Item -LiteralPath $src -Destination $dst -Force -ErrorAction Stop
            return New-QCSuccessResult -Code 'QUEUE_JOB_MOVED' -Message 'Job moved between states.' -Data @{ jobId = $JobId; fromState = $from; toState = $to }
        } finally {
            _QCQJ-ReleaseLockFile -LockPath $lockPath
        }
    } catch {
        return New-QCFailureResult -Code 'QUEUE_MOVE_ERROR' -Message 'Failed to move job between states.' -Data @{ jobId = $JobId; error = $_ }
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
        if (Test-Path -LiteralPath $jobLock) {
            return New-QCFailureResult -Code 'QUEUE_LOCK_EXISTS' -Message 'Job is already locked.' -Data @{ jobId = $JobId; lockPath = $jobLock }
        }
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
            skippedNotStale    = 0
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
                if ($age -lt $staleSeconds) {
                    $result.skippedNotStale++
                    continue
                }

                $attempts = 0
                if ($job.ContainsKey('attempts') -and $job.attempts -ne $null) { $attempts = [int]$job.attempts }
                $attempts++
                $job.attempts = $attempts
                $job.updatedAtUtc = $now.ToString('o')

                $lock = _QCQJ-LockFilePath -Root $root -JobId $jobId
                _QCQJ-ReleaseLockFile -LockPath $lock

                if ($attempts -ge $maxAttempts) {
                    $job.status = 'failed'
                    _QCQJ-WriteJobFileAtomic -Path $f.FullName -Job $job
                    $dst = _QCQJ-JobFilePath -Root $root -State 'failed' -JobId $jobId
                    Move-Item -LiteralPath $f.FullName -Destination $dst -Force -ErrorAction Stop
                    $result.recoveredToFailed++
                    $result.details += @{ jobId = $jobId; action = 'failed'; attempts = $attempts; ageSeconds = [int]$age }
                } else {
                    $job.status = 'pending'
                    $job.startedAtUtc = $null
                    _QCQJ-WriteJobFileAtomic -Path $f.FullName -Job $job
                    $dst = _QCQJ-JobFilePath -Root $root -State 'pending' -JobId $jobId
                    Move-Item -LiteralPath $f.FullName -Destination $dst -Force -ErrorAction Stop
                    $result.recoveredToPending++
                    $result.details += @{ jobId = $jobId; action = 'requeued'; attempts = $attempts; ageSeconds = [int]$age }
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
