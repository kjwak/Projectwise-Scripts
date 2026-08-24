# QC.Queue.Json.psm1
# Responsibility: JSON-backed queue persistence, lifecycle transitions, and queue reporting.

Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core/Core.Results.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core/Core.Runtime.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core/Core.Paths.psm1') -Force

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
    $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
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

function _QCQJ-GetMoveRetrySettings {
    <#
    Retries for queue JSON renames/deletes. AV scanners often hold handles for 5-30s on
    newly written files under queue\; defaults are higher than a quick 10x200ms burst.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [hashtable]$Config
    )
    $attempts = 30
    $sleepMs = 500
    if ($Config -and $Config.ContainsKey('queue') -and $Config.queue) {
        $q = $Config.queue
        if ($q -is [hashtable]) {
            if ($q.ContainsKey('moveRetryAttempts') -and $null -ne $q['moveRetryAttempts']) {
                try { $attempts = [int]$q['moveRetryAttempts'] } catch { }
            }
            if ($q.ContainsKey('moveRetrySleepMs') -and $null -ne $q['moveRetrySleepMs']) {
                try { $sleepMs = [int]$q['moveRetrySleepMs'] } catch { }
            }
        } elseif ($q.PSObject -and $q.PSObject.Properties) {
            try {
                $p1 = $q.PSObject.Properties['moveRetryAttempts']
                if ($p1 -and $null -ne $p1.Value) { $attempts = [int]$p1.Value }
            } catch { }
            try {
                $p2 = $q.PSObject.Properties['moveRetrySleepMs']
                if ($p2 -and $null -ne $p2.Value) { $sleepMs = [int]$p2.Value }
            } catch { }
        }
    }
    if ($attempts -lt 3) { $attempts = 3 }
    if ($attempts -gt 120) { $attempts = 120 }
    if ($sleepMs -lt 50) { $sleepMs = 50 }
    if ($sleepMs -gt 5000) { $sleepMs = 5000 }
    return @{ Attempts = $attempts; SleepMs = $sleepMs }
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

function _QCQJ-GetLocalMachineName {
    $n = [string]$env:COMPUTERNAME
    if ([string]::IsNullOrWhiteSpace($n)) { return '' }
    return $n.Trim()
}

function _QCQJ-GetCrossHostLockStaleSeconds([hashtable]$Config) {
    $sec = 90
    if ($Config -and $Config.ContainsKey('queue') -and $Config.queue) {
        $q = $Config.queue
        $recover = $null
        if ($q -is [hashtable] -and $q.ContainsKey('recover')) { $recover = $q.recover }
        elseif ($q.PSObject -and $q.PSObject.Properties['recover']) { $recover = $q.recover }
        if ($recover) {
            $val = $null
            if ($recover -is [hashtable] -and $recover.ContainsKey('crossHostLockStaleSeconds')) {
                $val = $recover['crossHostLockStaleSeconds']
            } elseif ($recover.PSObject -and $recover.PSObject.Properties['crossHostLockStaleSeconds']) {
                $val = $recover.crossHostLockStaleSeconds
            }
            if ($null -ne $val) {
                try { $sec = [int]$val } catch { }
            }
        }
    }
    if ($sec -lt 15) { $sec = 15 }
    if ($sec -gt 3600) { $sec = 3600 }
    return $sec
}

function _QCQJ-ReadLockPayload([string]$LockPath) {
    $raw = $null
    try { $raw = Get-Content -LiteralPath $LockPath -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json -ErrorAction SilentlyContinue } catch { return $null }
    if (-not $raw) { return $null }
    $pidVal = 0
    try {
        if ($raw.PSObject -and $raw.PSObject.Properties['pid']) { $pidVal = [int]$raw.pid }
    } catch { $pidVal = 0 }
    $machine = ''
    try {
        if ($raw.PSObject -and $raw.PSObject.Properties['machineName'] -and $null -ne $raw.machineName) {
            $machine = ([string]$raw.machineName).Trim()
        }
    } catch { $machine = '' }
    $created = $null
    try {
        if ($raw.PSObject -and $raw.PSObject.Properties['createdAtUtc'] -and $raw.createdAtUtc) {
            $created = [DateTime]::Parse([string]$raw.createdAtUtc).ToUniversalTime()
        }
    } catch { $created = $null }
    return @{
        pid = $pidVal
        machineName = $machine
        createdAtUtc = $created
        raw = $raw
    }
}

function _QCQJ-LockPayloadIsLocalHost($Payload) {
    # Legacy locks (pid only) are treated as local so single-server recovery stays immediate.
    if (-not $Payload) { return $true }
    $mn = [string]$Payload.machineName
    if ([string]::IsNullOrWhiteSpace($mn)) { return $true }
    $local = _QCQJ-GetLocalMachineName
    if ([string]::IsNullOrWhiteSpace($local)) { return $true }
    return ($mn.ToUpperInvariant() -eq $local.ToUpperInvariant())
}

function _QCQJ-LockCreatedAgeSeconds($Payload) {
    if (-not $Payload -or -not $Payload.createdAtUtc) { return $null }
    try { return ([DateTime]::UtcNow - [DateTime]$Payload.createdAtUtc).TotalSeconds } catch { return $null }
}

function _QCQJ-IsQueueWriteLockPath([string]$LockPath) {
    $leaf = [System.IO.Path]::GetFileName($LockPath)
    return ($leaf -ieq '_queue_write.lock')
}

function _QCQJ-NewLockPayloadJson {
    $ts = Get-QCTimestamp
    $payload = @{
        pid = $PID
        machineName = (_QCQJ-GetLocalMachineName)
        createdAtUtc = $ts
        heartbeatUtc = $ts
    }
    return ($payload | ConvertTo-Json -Depth 5 -Compress)
}

function _QCQJ-StampJobOwner {
    param([Parameter(Mandatory)][hashtable]$Job)
    $Job['machineName'] = (_QCQJ-GetLocalMachineName)
    $Job['pid'] = $PID
}

function _QCQJ-ClearJobOwner {
    param([Parameter(Mandatory)][hashtable]$Job)
    if ($Job.ContainsKey('machineName')) { $Job.Remove('machineName') }
    if ($Job.ContainsKey('pid')) { $Job.Remove('pid') }
}

function _QCQJ-IsLockOwnerDead {
    <#
    Same-host (or legacy pid-only) locks: dead when owner PID is gone.
    Other-host per-job locks: never stolen here (Recover uses job heartbeat).
    Other-host _queue_write.lock: dead when createdAtUtc is older than crossHostLockStaleSeconds.
    Returns $false on uncertainty so we never steal a live owner.
    #>
    param(
        [Parameter(Mandatory)][string]$LockPath,
        [hashtable]$Config = $null
    )
    $payload = _QCQJ-ReadLockPayload -LockPath $LockPath
    if (-not $payload) { return $false }
    $ownerPid = [int]$payload.pid
    $isLocal = _QCQJ-LockPayloadIsLocalHost -Payload $payload
    if ($isLocal) {
        if ($ownerPid -le 0) { return $false }
        if ($ownerPid -eq $PID) { return $false }
        $proc = Get-Process -Id $ownerPid -ErrorAction SilentlyContinue
        return (-not $proc)
    }
    if (-not (_QCQJ-IsQueueWriteLockPath -LockPath $LockPath)) { return $false }
    $age = _QCQJ-LockCreatedAgeSeconds -Payload $payload
    if ($null -eq $age) { return $false }
    $stale = _QCQJ-GetCrossHostLockStaleSeconds -Config $Config
    return ($age -ge $stale)
}

function _QCQJ-ShouldReclaimAbandonedLock {
    <#
    Pending / dangling per-job locks (job is not in running\): reclaim same-host
    dead PID immediately, or other-host locks whose createdAtUtc is stale.
    Never used by AcquireLockFile on a running job's lock.
    #>
    param(
        [Parameter(Mandatory)][string]$LockPath,
        [hashtable]$Config = $null
    )
    if (_QCQJ-IsLockOwnerDead -LockPath $LockPath -Config $Config) { return $true }
    $payload = _QCQJ-ReadLockPayload -LockPath $LockPath
    if (-not $payload) { return $false }
    if (_QCQJ-LockPayloadIsLocalHost -Payload $payload) { return $false }
    if (_QCQJ-IsQueueWriteLockPath -LockPath $LockPath) { return $false }
    $age = _QCQJ-LockCreatedAgeSeconds -Payload $payload
    if ($null -eq $age) { return $false }
    $stale = _QCQJ-GetCrossHostLockStaleSeconds -Config $Config
    return ($age -ge $stale)
}

function _QCQJ-GetRunningLockOrphanInfo {
    param(
        [Parameter(Mandatory)][string]$LockPath
    )
    if (-not (Test-Path -LiteralPath $LockPath)) {
        return @{ isOrphan = $true; reason = 'NO_LOCK_FILE' }
    }
    $payload = _QCQJ-ReadLockPayload -LockPath $LockPath
    if (-not $payload) {
        return @{ isOrphan = $false; reason = $null }
    }
    if (-not (_QCQJ-LockPayloadIsLocalHost -Payload $payload)) {
        return @{ isOrphan = $false; reason = $null }
    }
    $ownerPid = [int]$payload.pid
    if ($ownerPid -le 0) { return @{ isOrphan = $false; reason = $null } }
    if ($ownerPid -eq $PID) { return @{ isOrphan = $false; reason = $null } }
    $proc = Get-Process -Id $ownerPid -ErrorAction SilentlyContinue
    if (-not $proc) { return @{ isOrphan = $true; reason = ("DEAD_PID({0})" -f $ownerPid) } }
    return @{ isOrphan = $false; reason = $null }
}

function _QCQJ-AcquireLockFile([string]$LockPath, [int]$TimeoutMs = 30000, [int]$SleepMs = 150, [hashtable]$Config = $null) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $stealAttempted = $false
    while ($sw.ElapsedMilliseconds -lt $TimeoutMs) {
        try {
            $fs = [System.IO.File]::Open($LockPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
            try {
                $json = _QCQJ-NewLockPayloadJson
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
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
            # call to keep the cost bounded. Other-host per-job locks are never
            # stolen here (PID is meaningless across machines).
            if (-not $stealAttempted) {
                $stealAttempted = $true
                if (_QCQJ-IsLockOwnerDead -LockPath $LockPath -Config $Config) {
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
        # Avoid PowerShell array += here. Queue jobs can contain large paired-sheet
        # or processor result arrays; += reallocates/copies the array on every item
        # and turns serialization into an O(n^2) hotspot during enqueue/move.
        $out = [System.Collections.Generic.List[object]]::new()
        foreach ($i in $Value) {
            [void]$out.Add((_QCQJ-DeepToJsonSafeObject -Value $i -CurrentDepth ($CurrentDepth + 1)))
        }
        return $out.ToArray()
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

function _QCQJ-RemoveFileWithRetry {
    param(
        [Parameter(Mandatory)][string]$LiteralPath,
        [hashtable]$Config = $null,
        [int]$Attempts = 0,
        [int]$SleepMs = 0
    )
    if (-not (Test-Path -LiteralPath $LiteralPath)) { return $true }
    $mr = _QCQJ-GetMoveRetrySettings -Config $Config
    if ($Attempts -le 0) { $Attempts = $mr.Attempts }
    if ($SleepMs -le 0) { $SleepMs = $mr.SleepMs }
    $lastEx = $null
    for ($a = 1; $a -le $Attempts; $a++) {
        try {
            Remove-Item -LiteralPath $LiteralPath -Force -ErrorAction Stop
            return $true
        } catch {
            $lastEx = $_
            if ($a -ge $Attempts) { return $false }
            Start-Sleep -Milliseconds $SleepMs
        }
    }
    return $false
}

function _QCQJ-MoveItemWithRetry {
    param(
        [Parameter(Mandatory)]
        [string]$LiteralPath,
        [Parameter(Mandatory)]
        [string]$Destination,
        [hashtable]$Config = $null,
        [int]$Attempts = 0,
        [int]$SleepMs = 0
    )
    $mr = _QCQJ-GetMoveRetrySettings -Config $Config
    if ($Attempts -le 0) { $Attempts = $mr.Attempts }
    if ($SleepMs -le 0) { $SleepMs = $mr.SleepMs }
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

function _QCQJ-RemoveDuplicateJobFiles {
    <#
    After a successful transition, delete stray copies in other state folders (AV copy+delete
  failures can leave the same job id in pending, running, and succeeded).
    #>
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$JobId,
        [Parameter(Mandatory)][string]$KeepState,
        [hashtable]$Config
    )
    $keep = _QCQJ-NormalizeState $KeepState
    $removed = @()
    $failed = @()
    foreach ($s in @('pending', 'running', 'succeeded', 'failed')) {
        if ($s -eq $keep) { continue }
        $p = _QCQJ-JobFilePath -Root $Root -State $s -JobId $JobId
        if (-not (Test-Path -LiteralPath $p)) { continue }
        if (_QCQJ-RemoveFileWithRetry -LiteralPath $p -Config $Config) {
            $removed += $s
        } else {
            $failed += $s
        }
    }
    return @{ removed = $removed; failed = $failed }
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
            # Avoid += for large arrays when reading jobs back from JSON. This
            # keeps queue polling and duplicate/in-flight scans linear in the
            # number of values rather than repeatedly copying intermediate arrays.
            $out = [System.Collections.Generic.List[object]]::new()
            foreach ($i in $Value) { [void]$out.Add((_ToHashtable $i)) }
            return $out.ToArray()
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

function _QCQJ-WriteJobFileAtomic {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][hashtable]$Job,
        [hashtable]$Config = $null
    )
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
        _QCQJ-MoveItemWithRetry -LiteralPath $tmp -Destination $Path -Config $Config
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
            _QCQJ-WriteJobFileAtomic -Path $dest -Job $Job -Config $Config
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
    .PARAMETER IncludeJobTypes
    Optional allow-list of job types. Empty or omitted means all types are eligible
    (`workers.enabledJobTypes` unset). When set, jobs whose type is not in the list
    are skipped. Combined with ExcludeJobTypes: a job must be included and not excluded.
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
        [string[]]$ExcludeJobTypes = @(),
        [Parameter(Mandatory = $false)]
        [string[]]$IncludeJobTypes = @()
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
        $includeTypeSet = @{}
        foreach ($t in @($IncludeJobTypes)) {
            if (-not [string]::IsNullOrWhiteSpace($t)) { $includeTypeSet[([string]$t).Trim()] = $true }
        }
        $includeActive = $includeTypeSet.Count -gt 0

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
        # A lock blocks selection only when the owner is still alive on this host.
        # Dead local PIDs and stale other-host pending locks are self-healed here.
        $isJobLockedAlive = {
            param([string]$JobId)
            $lockPath = _QCQJ-LockFilePath -Root $root -JobId $JobId
            if (-not (Test-Path -LiteralPath $lockPath)) { return $false }
            if (_QCQJ-ShouldReclaimAbandonedLock -LockPath $lockPath -Config $Config) {
                try { Remove-Item -LiteralPath $lockPath -Force -ErrorAction Stop } catch { }
                return $false
            }
            return $true
        }

        # Read each eligible job at most once. The previous preferred-type scheduler
        # scanned and parsed the full pending folder once per preferred type, then
        # parsed it again for the fallback path. Under large backlogs that made each
        # worker poll repeatedly deserialize the same JSON files and increased lock
        # race windows. This single pass preserves the old priority semantics:
        # preferJobTypes order wins first, and LastWriteTimeUtc/Name order wins within
        # each priority tier; if no preferred job exists, the oldest eligible job wins.
        $preferRank = @{}
        for ($i = 0; $i -lt $preferTypes.Count; $i++) {
            $pt = [string]$preferTypes[$i]
            if (-not [string]::IsNullOrWhiteSpace($pt) -and -not $excludeTypeSet.ContainsKey($pt) -and -not $preferRank.ContainsKey($pt)) {
                $preferRank[$pt] = $i
            }
        }
        $firstEligible = $null
        $bestPreferred = $null
        $bestPreferredRank = [int]::MaxValue

        foreach ($f in $files) {
            $jobId = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
            if ($excludeSet.ContainsKey($jobId)) { continue }
            if (& $isJobLockedAlive $jobId) { continue }

            $job = _QCQJ-ReadJobFile -Path $f.FullName
            $jt = ''
            try { $jt = [string]$job.type } catch { $jt = '' }
            if ($excludeTypeSet.ContainsKey($jt)) { continue }
            if ($includeActive -and -not $includeTypeSet.ContainsKey($jt)) { continue }

            if (-not $firstEligible) {
                $firstEligible = @{ job = $job; jobId = $jobId; state = 'pending' }
            }

            if ($preferRank.ContainsKey($jt)) {
                $rank = [int]$preferRank[$jt]
                if ($rank -lt $bestPreferredRank) {
                    $bestPreferredRank = $rank
                    $bestPreferred = @{ job = $job; jobId = $jobId; state = 'pending'; preferredType = $jt }
                    if ($rank -eq 0) {
                        return New-QCSuccessResult -Code 'QUEUE_NEXT_JOB' -Message 'Next pending job selected (preferred type).' -Data $bestPreferred
                    }
                }
            }
        }

        if ($bestPreferred) {
            return New-QCSuccessResult -Code 'QUEUE_NEXT_JOB' -Message 'Next pending job selected (preferred type).' -Data $bestPreferred
        }
        if ($firstEligible) {
            return New-QCSuccessResult -Code 'QUEUE_NEXT_JOB' -Message 'Next pending job selected.' -Data $firstEligible
        }

        return New-QCSuccessResult -Code 'QUEUE_EMPTY' -Message 'No eligible pending jobs.' -Data @{ job = $null }
    } catch {
        return New-QCFailureResult -Code 'QUEUE_GET_NEXT_ERROR' -Message 'Failed to select next job.' -Data @{ error = $_ }
    }
}

function Get-QCEnabledJobTypes {
    <#
    .SYNOPSIS
    Returns workers.enabledJobTypes, or an empty array when unrestricted (all types).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
    $out = New-Object System.Collections.Generic.List[string]
    $raw = $null
    try {
        if ($Config.ContainsKey('workers') -and $Config.workers) {
            $w = $Config.workers
            if ($w -is [hashtable] -and $w.ContainsKey('enabledJobTypes')) { $raw = $w.enabledJobTypes }
            elseif ($w.PSObject -and $w.PSObject.Properties['enabledJobTypes']) { $raw = $w.enabledJobTypes }
        }
    } catch { $raw = $null }
    foreach ($t in @($raw)) {
        $s = ([string]$t).Trim()
        if (-not [string]::IsNullOrWhiteSpace($s) -and -not $out.Contains($s)) { [void]$out.Add($s) }
    }
    return @($out.ToArray())
}

function Test-QCQueueRootIsUnc {
    <#
    .SYNOPSIS
    True when a queue root is a Windows UNC path (\\server\share).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Path
    )
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    $p = $Path.Trim()
    if ($p.StartsWith('\\?\UNC\', [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    if ($p.StartsWith('\\') -and -not $p.StartsWith('\\?\')) { return $true }
    return $false
}

function Get-QCRemoteWorkerHostSettings {
    <#
    .SYNOPSIS
    Supervisor settings from workers / workers.remoteHost. Missing keys keep dashboard-compatible defaults.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
    $maxParallel = 1
    $maxJobsPerWorker = 25
    $leaseSeconds = 600
    $idleSleepMs = 400
    $spawnStaggerMs = 150
    $allowUncQueue = $false
    $throttleEnabled = $false
    $throttleSampleSeconds = 10
    $throttleCpuPercent = 0
    $throttleMemoryPercent = 0
    $throttleProcessCpuPercent = 0
    $throttleProcessMemoryMb = 0
    $throttleBusySlots = 1
    $throttlePatterns = New-Object System.Collections.Generic.List[string]
    try {
        if ($Config.ContainsKey('workers') -and $Config.workers) {
            $w = $Config.workers
            if ($w.maxParallel) { $maxParallel = [int]$w.maxParallel }
            if ($w.maxJobsPerWorker) { $maxJobsPerWorker = [int]$w.maxJobsPerWorker }
            if ($w.leaseSeconds) { $leaseSeconds = [int]$w.leaseSeconds }
            if ($w.idleSleepMs) { $idleSleepMs = [int]$w.idleSleepMs }
            if ($w.spawnStaggerMs) { $spawnStaggerMs = [int]$w.spawnStaggerMs }
            $rh = $null
            if ($w -is [hashtable] -and $w.ContainsKey('remoteHost')) { $rh = $w.remoteHost }
            elseif ($w.PSObject -and $w.PSObject.Properties['remoteHost']) { $rh = $w.remoteHost }
            if ($rh) {
                $flag = $null
                if ($rh -is [hashtable] -and $rh.ContainsKey('allowUncQueue')) { $flag = $rh.allowUncQueue }
                elseif ($rh.PSObject -and $rh.PSObject.Properties['allowUncQueue']) { $flag = $rh.allowUncQueue }
                if ($null -ne $flag) { $allowUncQueue = [bool]$flag }
                $th = $null
                if ($rh -is [hashtable] -and $rh.ContainsKey('throttle')) { $th = $rh.throttle }
                elseif ($rh.PSObject -and $rh.PSObject.Properties['throttle']) { $th = $rh.throttle }
                if ($th) {
                    if ($th -is [hashtable] -and $th.ContainsKey('enabled') -and $null -ne $th.enabled) {
                        try { $throttleEnabled = [bool]$th.enabled } catch { $throttleEnabled = $false }
                    } elseif ($th.PSObject -and $th.PSObject.Properties['enabled'] -and $null -ne $th.enabled) {
                        try { $throttleEnabled = [bool]$th.enabled } catch { $throttleEnabled = $false }
                    }
                    if ($null -ne $th.sampleSeconds) { try { $throttleSampleSeconds = [int]$th.sampleSeconds } catch { } }
                    if ($null -ne $th.cpuPercent) { try { $throttleCpuPercent = [double]$th.cpuPercent } catch { } }
                    if ($null -ne $th.memoryPercent) { try { $throttleMemoryPercent = [double]$th.memoryPercent } catch { } }
                    if ($null -ne $th.processCpuPercent) { try { $throttleProcessCpuPercent = [double]$th.processCpuPercent } catch { } }
                    if ($null -ne $th.processMemoryMb) { try { $throttleProcessMemoryMb = [double]$th.processMemoryMb } catch { } }
                    if ($null -ne $th.busyRecommendedSlots) { try { $throttleBusySlots = [int]$th.busyRecommendedSlots } catch { } }
                    $rawPats = $null
                    if ($th -is [hashtable] -and $th.ContainsKey('processNamePatterns')) { $rawPats = $th.processNamePatterns }
                    elseif ($th.PSObject -and $th.PSObject.Properties['processNamePatterns']) { $rawPats = $th.processNamePatterns }
                    foreach ($pat in @($rawPats)) {
                        $s = ([string]$pat).Trim()
                        if (-not [string]::IsNullOrWhiteSpace($s) -and -not $throttlePatterns.Contains($s)) {
                            [void]$throttlePatterns.Add($s)
                        }
                    }
                }
            }
        }
    } catch { }
    if ($maxParallel -lt 1) { $maxParallel = 1 }
    if ($maxJobsPerWorker -lt 1) { $maxJobsPerWorker = 1 }
    if ($leaseSeconds -lt 0) { $leaseSeconds = 0 }
    if ($idleSleepMs -lt 0) { $idleSleepMs = 0 }
    if ($spawnStaggerMs -lt 0) { $spawnStaggerMs = 0 }
    if ($throttleSampleSeconds -lt 1) { $throttleSampleSeconds = 10 }
    if ($throttleCpuPercent -lt 0) { $throttleCpuPercent = 0 }
    if ($throttleMemoryPercent -lt 0) { $throttleMemoryPercent = 0 }
    if ($throttleProcessCpuPercent -lt 0) { $throttleProcessCpuPercent = 0 }
    if ($throttleProcessMemoryMb -lt 0) { $throttleProcessMemoryMb = 0 }
    return @{
        maxParallel = $maxParallel
        maxJobsPerWorker = $maxJobsPerWorker
        leaseSeconds = $leaseSeconds
        idleSleepMs = $idleSleepMs
        spawnStaggerMs = $spawnStaggerMs
        allowUncQueue = $allowUncQueue
        enabledJobTypes = @(Get-QCEnabledJobTypes -Config $Config)
        throttle = @{
            enabled = [bool]$throttleEnabled
            sampleSeconds = [int]$throttleSampleSeconds
            cpuPercent = [double]$throttleCpuPercent
            memoryPercent = [double]$throttleMemoryPercent
            processCpuPercent = [double]$throttleProcessCpuPercent
            processMemoryMb = [double]$throttleProcessMemoryMb
            processNamePatterns = @($throttlePatterns.ToArray())
            busyRecommendedSlots = [int]$throttleBusySlots
        }
    }
}

function Test-QCUncQueueClaimAllowed {
    <#
    .SYNOPSIS
    False when queue.rootDir is UNC and the caller has not opted in.
    .DESCRIPTION
    Local queue paths always allowed. UNC claims require -AllowUncQueue or
    workers.remoteHost.allowUncQueue. Server dashboard workers use a local
    C:\ queue path, so this does not change production coordinator behavior.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [switch]$AllowUncQueue
    )
    $root = $null
    try {
        if ($Config.ContainsKey('queue') -and $Config.queue) {
            if ($Config.queue.rootDir) { $root = [string]$Config.queue.rootDir }
            elseif ($Config.queue.root) { $root = [string]$Config.queue.root }
            elseif ($Config.queue.path) { $root = [string]$Config.queue.path }
        }
    } catch { $root = $null }
    if (-not (Test-QCQueueRootIsUnc -Path $root)) { return $true }
    if ($AllowUncQueue.IsPresent) { return $true }
    $rh = Get-QCRemoteWorkerHostSettings -Config $Config
    return [bool]$rh.allowUncQueue
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
        if (-not (Test-Path -LiteralPath $flag)) { return $false }

        # Self-heal stale watcher-active flags left by a killed watcher. A bare
        # Test-Path keeps downstream gating stuck until a manual startup cleanup;
        # using the PID written by Set-QCWatcherActive lets callers recover on the
        # next status check while still treating unreadable/ambiguous flags as active.
        $payload = $null
        try { $payload = Get-Content -LiteralPath $flag -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop } catch { return $true }
        $ownerPid = 0
        try {
            if ($payload.PSObject -and $payload.PSObject.Properties['pid']) { $ownerPid = [int]$payload.pid }
        } catch { $ownerPid = 0 }
        if ($ownerPid -gt 0 -and $ownerPid -ne $PID) {
            $proc = Get-Process -Id $ownerPid -ErrorAction SilentlyContinue
            if (-not $proc) {
                try { Remove-Item -LiteralPath $flag -Force -ErrorAction Stop } catch { }
                return $false
            }
        }
        return $true
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
    if (Get-Command -Name 'Normalize-QCDocumentsFolderPath' -ErrorAction SilentlyContinue) {
        $r = Normalize-QCDocumentsFolderPath -Path $Path
        if ($r.IsSuccess) { return [string]$r.Data.path }
    }
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
            _QCQJ-WriteJobFileAtomic -Path $loc.path -Job $job -Config $Config
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
            _QCQJ-WriteJobFileAtomic -Path $loc.path -Job $Job -Config $Config
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
                $loc = _QCQJ-FindJobFile -Root $root -JobId $JobId
                if ($loc -and ([string]$loc.state) -eq $to) {
                    # Idempotent: a prior move (or retry) already placed the job in the target folder.
                    if ($PSBoundParameters.ContainsKey('Job') -and $Job) {
                        $Job.status = $to
                        $Job.updatedAtUtc = (Get-QCTimestamp)
                        if (-not $Job.ContainsKey('id') -or [string]::IsNullOrWhiteSpace([string]$Job.id)) { $Job.id = $JobId }
                        _QCQJ-WriteJobFileAtomic -Path $loc.path -Job $Job -Config $Config
                        try {
                            $dedupeKey = [string]$Job['dedupeKey']
                            if ($dedupeKey) {
                                _QCQJ-UpdateDedupeIndexForJob -Config $Config -Root $root -DedupeKey $dedupeKey -JobId $JobId -State $to
                            }
                        } catch { }
                    }
                    $dup = _QCQJ-RemoveDuplicateJobFiles -Root $root -JobId $JobId -KeepState $to -Config $Config
                    return New-QCSuccessResult -Code 'QUEUE_JOB_ALREADY_MOVED' -Message 'Job already in target state.' -Data @{
                        jobId = $JobId; fromState = $from; toState = $to; path = $loc.path; idempotent = $true
                        duplicatesRemoved = @($dup.removed); duplicatesRemoveFailed = @($dup.failed)
                    }
                }
                if ($loc) {
                    return New-QCFailureResult -Code 'QUEUE_JOB_WRONG_STATE' -Message "Job is in '$($loc.state)', not source '$from'." -Data @{
                        jobId = $JobId; fromState = $from; actualState = [string]$loc.state; toState = $to; path = $loc.path
                    }
                }
                return New-QCFailureResult -Code 'QUEUE_JOB_NOT_FOUND' -Message 'Job file not found in source state.' -Data @{ jobId = $JobId; fromState = $from; path = $src }
            }
            $dst = _QCQJ-JobFilePath -Root $root -State $to -JobId $JobId

            if ($PSBoundParameters.ContainsKey('Job') -and $Job) {
                # Caller supplied the authoritative job content (e.g. with result/lastError).
                # Stamp status + updatedAtUtc and write to source path before the rename.
                $Job.status = $to
                $Job.updatedAtUtc = (Get-QCTimestamp)
                if (-not $Job.ContainsKey('id') -or [string]::IsNullOrWhiteSpace([string]$Job.id)) { $Job.id = $JobId }
                _QCQJ-WriteJobFileAtomic -Path $src -Job $Job -Config $Config
            } else {
                $existing = _QCQJ-ReadJobFile -Path $src
                $existing.status = $to
                $existing.updatedAtUtc = (Get-QCTimestamp)
                _QCQJ-WriteJobFileAtomic -Path $src -Job $existing -Config $Config
            }
            _QCQJ-MoveItemWithRetry -LiteralPath $src -Destination $dst -Config $Config
            $dup = _QCQJ-RemoveDuplicateJobFiles -Root $root -JobId $JobId -KeepState $to -Config $Config
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
            return New-QCSuccessResult -Code 'QUEUE_JOB_MOVED' -Message 'Job moved between states.' -Data @{
                jobId = $JobId; fromState = $from; toState = $to
                duplicatesRemoved = @($dup.removed); duplicatesRemoveFailed = @($dup.failed)
            }
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

    $root = $null
    $jobLock = $null
    $jobLockHeld = $false
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
        if (-not (_QCQJ-AcquireLockFile -LockPath $jobLock -TimeoutMs $lk.TimeoutMs -SleepMs $lk.SleepMs -Config $Config)) {
            return New-QCFailureResult -Code 'QUEUE_LOCK_TIMEOUT' -Message 'Timed out acquiring job lock.' -Data @{ jobId = $JobId; lockPath = $jobLock }
        }
        $jobLockHeld = $true

        # Transition pending -> running on lock acquire (prevents repeated selection).
        $lockPath = _QCQJ-QueueWriteLockPath -Root $root
        if (-not (_QCQJ-AcquireLockFile -LockPath $lockPath -TimeoutMs $lk.TimeoutMs -SleepMs $lk.SleepMs -Config $Config)) {
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
            $dupLock = @{ removed = @(); failed = @() }

            $job = _QCQJ-ReadJobFile -Path $pending
            $job.status = 'running'
            $nowTs = Get-QCTimestamp
            $job.startedAtUtc = $nowTs
            $job.heartbeatUtc = $nowTs
            _QCQJ-StampJobOwner -Job $job
            if (-not $job.ContainsKey('recoveryCount')) { $job.recoveryCount = 0 }
            _QCQJ-WriteJobFileAtomic -Path $pending -Job $job -Config $Config
            _QCQJ-MoveItemWithRetry -LiteralPath $pending -Destination $running -Config $Config
            $dupLock = _QCQJ-RemoveDuplicateJobFiles -Root $root -JobId $JobId -KeepState 'running' -Config $Config
        } finally {
            _QCQJ-ReleaseLockFile -LockPath $lockPath
        }

        return New-QCSuccessResult -Code 'QUEUE_LOCK_ACQUIRED' -Message 'Job lock acquired.' -Data @{
            jobId = $JobId; lockPath = $jobLock
            duplicatesRemoved = @($dupLock.removed); duplicatesRemoveFailed = @($dupLock.failed)
        }
    } catch {
        if ($jobLockHeld -and $jobLock) {
            _QCQJ-ReleaseLockFile -LockPath $jobLock
        }
        if ($root) {
            try {
                $loc = _QCQJ-FindJobFile -Root $root -JobId $JobId
                if ($loc -and $loc.state -ne 'pending') {
                    return New-QCFailureResult -Code 'QUEUE_JOB_ALREADY_MOVED' -Message 'Job is no longer pending; skipping lock.' -Data @{
                        jobId = $JobId
                        state = $loc.state
                        path = $loc.path
                        lockRace = $true
                        error = $_
                    }
                }
            } catch { }
        }
        $inner = $null
        try { $inner = [string]$_.Exception.Message } catch { }
        if ([string]::IsNullOrWhiteSpace($inner)) { try { $inner = [string]$_ } catch { $inner = 'unknown error' } }
        return New-QCFailureResult -Code 'QUEUE_LOCK_ERROR' -Message 'Failed to lock job.' -Data @{ jobId = $JobId; error = $_; innerMessage = $inner }
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
            $heartbeat = $null
            if ($job.ContainsKey('heartbeatUtc') -and $job.heartbeatUtc) {
                try { $heartbeat = [DateTime]::Parse([string]$job.heartbeatUtc).ToUniversalTime() } catch { $heartbeat = $null }
            }
            $reference = if ($heartbeat) { $heartbeat } elseif ($started) { $started } else { $f.LastWriteTimeUtc }

            $age = ($now - $reference).TotalSeconds

            $lockPathJob = _QCQJ-LockFilePath -Root $root -JobId $jobId
            $orphanInfo = _QCQJ-GetRunningLockOrphanInfo -LockPath $lockPathJob
            $isOrphan = [bool]$orphanInfo.isOrphan
            $orphanReason = $orphanInfo.reason

            if (-not $isOrphan -and $age -lt $staleSeconds) {
                $result.skippedNotStale++
                continue
            }

            $recoverCandidates += ,@($f.FullName, $jobId)
        }

        # Phase 2: one short global write-lock window per job so Move-QCJob / workers can interleave.
        foreach ($pair in $recoverCandidates) {
            $jobId = [string]$pair[1]
            if (-not (_QCQJ-AcquireLockFile -LockPath $writeLock -TimeoutMs $lk.TimeoutMs -SleepMs $lk.SleepMs -Config $Config)) {
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
                $heartbeat2 = $null
                if ($job.ContainsKey('heartbeatUtc') -and $job.heartbeatUtc) {
                    try { $heartbeat2 = [DateTime]::Parse([string]$job.heartbeatUtc).ToUniversalTime() } catch { $heartbeat2 = $null }
                }
                $reference2 = if ($heartbeat2) { $heartbeat2 } elseif ($started2) { $started2 } else { $f2.LastWriteTimeUtc }
                $age2 = ($now - $reference2).TotalSeconds

                $lockPathJob2 = _QCQJ-LockFilePath -Root $root -JobId $jobId
                $orphanInfo2 = _QCQJ-GetRunningLockOrphanInfo -LockPath $lockPathJob2
                $isOrphan2 = [bool]$orphanInfo2.isOrphan
                $orphanReason2 = $orphanInfo2.reason

                if (-not $isOrphan2 -and $age2 -lt $staleSeconds) { continue }

                $attempts = 0
                if ($job.ContainsKey('attempts') -and $job.attempts -ne $null) { $attempts = [int]$job.attempts }
                $attempts++
                $job.attempts = $attempts
                $recoveryCount = 0
                if ($job.ContainsKey('recoveryCount') -and $null -ne $job.recoveryCount) {
                    try { $recoveryCount = [int]$job.recoveryCount } catch { $recoveryCount = 0 }
                }
                $recoveryCount++
                $job.recoveryCount = $recoveryCount
                $job.recoveryReason = if ($isOrphan2) { [string]$orphanReason2 } else { 'STALE_HEARTBEAT' }
                $job.updatedAtUtc = (ConvertTo-QCTimestamp -DateTime $now)
                $job.heartbeatUtc = $null

                _QCQJ-ReleaseLockFile -LockPath $lockPathJob2

                if ($attempts -ge $maxAttempts) {
                    $job.status = 'failed'
                    _QCQJ-WriteJobFileAtomic -Path $src -Job $job -Config $Config
                    $dst = _QCQJ-JobFilePath -Root $root -State 'failed' -JobId $jobId
                    _QCQJ-MoveItemWithRetry -LiteralPath $src -Destination $dst -Config $Config
                    _QCQJ-RemoveDuplicateJobFiles -Root $root -JobId $jobId -KeepState 'failed' -Config $Config | Out-Null
                    $result.recoveredToFailed++
                    $entry = @{ jobId = $jobId; action = 'failed'; attempts = $attempts; ageSeconds = [int]$age2 }
                    if ($isOrphan2) { $entry.orphan = $true; $entry.orphanReason = $orphanReason2; $result.recoveredOrphan++ }
                    $result.details += $entry
                } else {
                    $job.status = 'pending'
                    $job.startedAtUtc = $null
                    _QCQJ-ClearJobOwner -Job $job
                    _QCQJ-WriteJobFileAtomic -Path $src -Job $job -Config $Config
                    $dst = _QCQJ-JobFilePath -Root $root -State 'pending' -JobId $jobId
                    _QCQJ-MoveItemWithRetry -LiteralPath $src -Destination $dst -Config $Config
                    _QCQJ-RemoveDuplicateJobFiles -Root $root -JobId $jobId -KeepState 'pending' -Config $Config | Out-Null
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
                    if (-not (_QCQJ-AcquireLockFile -LockPath $writeLock -TimeoutMs $lk.TimeoutMs -SleepMs $lk.SleepMs -Config $Config)) {
                        $result.writeLockAcquireFailures++
                        continue
                    }
                    try {
                        if (_QCQJ-IsLockOwnerDead -LockPath $lf.FullName -Config $Config) {
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
                if (-not (_QCQJ-ShouldReclaimAbandonedLock -LockPath $lf.FullName -Config $Config)) { continue }
                if (-not (_QCQJ-AcquireLockFile -LockPath $writeLock -TimeoutMs $lk.TimeoutMs -SleepMs $lk.SleepMs -Config $Config)) {
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

function Repair-QCQueueDuplicateJobs {
    <#
    .SYNOPSIS
    Remove duplicate queue JSON files when the same job id exists in multiple state folders.
    .DESCRIPTION
    Keeps the canonical copy in the most advanced state (succeeded > running > pending > failed).
    Use after AV blocked deletes during moves, or enable repairDuplicateJobsOnStartup in appsettings.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    $root = _QCQJ-GetQueueRoot -Config $Config
    _QCQJ-EnsureLayout -Root $root
    $priority = @{ succeeded = 4; running = 3; pending = 2; failed = 1 }
    $byJob = @{}

    foreach ($s in @('pending', 'running', 'succeeded', 'failed')) {
        $dir = Join-Path $root $s
        if (-not (Test-Path -LiteralPath $dir)) { continue }
        foreach ($f in @(Get-ChildItem -LiteralPath $dir -Filter '*.json' -File -ErrorAction SilentlyContinue)) {
            $jid = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
            if ([string]::IsNullOrWhiteSpace($jid)) { continue }
            if (-not $byJob.ContainsKey($jid)) { $byJob[$jid] = @() }
            $byJob[$jid] += @{ state = $s; path = $f.FullName }
        }
    }

    $repaired = @()
    $removeFailed = @()
    foreach ($jid in @($byJob.Keys)) {
        $locs = @($byJob[$jid])
        if ($locs.Count -le 1) { continue }
        $canonical = ($locs | Sort-Object { $priority[[string]$_.state] } -Descending | Select-Object -First 1)
        $keepState = [string]$canonical.state
        foreach ($loc in $locs) {
            if ([string]$loc.state -eq $keepState) { continue }
            if ($PSCmdlet.ShouldProcess($loc.path, "Remove duplicate $($loc.state) copy (keep $keepState)")) {
                if (_QCQJ-RemoveFileWithRetry -LiteralPath $loc.path -Config $Config) {
                    $repaired += @{ jobId = $jid; removed = [string]$loc.state; kept = $keepState }
                } else {
                    $removeFailed += @{ jobId = $jid; path = $loc.path; state = [string]$loc.state; kept = $keepState }
                }
            }
        }
    }

    return New-QCSuccessResult -Code 'QUEUE_DUPLICATES_REPAIRED' -Message 'Duplicate queue job file sweep completed.' -Data @{
        root = $root
        repaired = $repaired
        removeFailed = $removeFailed
        duplicateJobCount = $repaired.Count
    }
}

function Update-QCJobHeartbeat {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$JobId,
        [Parameter(Mandatory)][hashtable]$Config,
        [hashtable]$Job = $null
    )
    $ts = Get-QCTimestamp
    if ($Job) {
        $Job.heartbeatUtc = $ts
        $Job.updatedAtUtc = $ts
        try {
            $root = _QCQJ-GetQueueRoot -Config $Config
            $runningPath = _QCQJ-JobFilePath -Root $root -State 'running' -JobId $JobId
            if (Test-Path -LiteralPath $runningPath) {
                _QCQJ-WriteJobFileAtomic -Path $runningPath -Job $Job -Config $Config | Out-Null
            }
        } catch { }
    }
    if (Get-Command -Name 'Update-QCProcessingJobHeartbeat' -ErrorAction SilentlyContinue) {
        try { Update-QCProcessingJobHeartbeat -Config $Config -JobId $JobId | Out-Null } catch { }
    }
}

function Set-QCJobCheckpoint {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$JobId,
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][hashtable]$Job,
        [Parameter(Mandatory)][string]$Checkpoint,
        [string]$CheckpointData = ''
    )
    $Job.checkpoint = [string]$Checkpoint
    if ($CheckpointData) { $Job.checkpointData = $CheckpointData }
    $Job.updatedAtUtc = Get-QCTimestamp
    try {
        $root = _QCQJ-GetQueueRoot -Config $Config
        $runningPath = _QCQJ-JobFilePath -Root $root -State 'running' -JobId $JobId
        if (Test-Path -LiteralPath $runningPath) {
            _QCQJ-WriteJobFileAtomic -Path $runningPath -Job $Job -Config $Config | Out-Null
        }
    } catch { }
    if (Get-Command -Name 'Update-QCProcessingJobCheckpoint' -ErrorAction SilentlyContinue) {
        try { Update-QCProcessingJobCheckpoint -Config $Config -JobId $JobId -Checkpoint $Checkpoint -CheckpointData $CheckpointData | Out-Null } catch { }
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

    $repairDup = $true
    if ($Config.ContainsKey('queue') -and $Config.queue) {
        $q = $Config.queue
        if ($q -is [hashtable] -and $q.ContainsKey('repairDuplicateJobsOnStartup')) {
            try { $repairDup = [bool]$q['repairDuplicateJobsOnStartup'] } catch { $repairDup = $true }
        } elseif ($q.PSObject -and $q.PSObject.Properties['repairDuplicateJobsOnStartup']) {
            try { $repairDup = [bool]$q.PSObject.Properties['repairDuplicateJobsOnStartup'].Value } catch { $repairDup = $true }
        }
    }
    if ($repairDup) {
        try {
            $dupRes = Repair-QCQueueDuplicateJobs -Config $Config
            if ($dupRes.IsSuccess) { $out.duplicateRepair = $dupRes.Data } else { [void]$out.errors.Add([string]$dupRes.Message) }
        } catch {
            [void]$out.errors.Add([string]$_.Exception.Message)
        }
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
