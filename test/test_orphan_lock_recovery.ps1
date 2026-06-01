$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$root = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $root 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $root 'modules\QC.Queue.Json.psm1') -Force

$failures = 0
function _Assert([bool]$ok, [string]$msg) {
    if (-not $ok) { Write-Host "  FAIL: $msg" -ForegroundColor Red; $script:failures++ }
    else          { Write-Host "  ok:   $msg" -ForegroundColor Green }
}

# --- Pick a known-dead PID by spawning, capturing its PID, and waiting for exit ---
function _NewDeadPid {
    $p = Start-Process -FilePath 'cmd.exe' -ArgumentList '/c','exit' -PassThru -WindowStyle Hidden
    $p.WaitForExit()
    Start-Sleep -Milliseconds 100
    return [int]$p.Id
}

function _NewQueueRoot {
    $r = Join-Path $env:TEMP ("qctest_orphan_" + [guid]::NewGuid().ToString('N').Substring(0,8))
    foreach ($s in 'pending','running','succeeded','failed','locks') {
        New-Item -ItemType Directory -Path (Join-Path $r $s) -Force | Out-Null
    }
    return $r
}

function _WriteJob([string]$Root, [string]$State, [string]$JobId, [hashtable]$Extra=@{}) {
    $body = @{
        id = $JobId
        type = 'STATUS_SET_GEN'
        dedupeKey = 'dq_' + $JobId
        status = $State
        sourceFolder = 'documents\test\sheets'
        attempts = 0
        enqueuedAtUtc = ([DateTime]::UtcNow.ToString('o'))
    }
    foreach ($k in $Extra.Keys) { $body[$k] = $Extra[$k] }
    $path = Join-Path (Join-Path $Root $State) ($JobId + '.json')
    ($body | ConvertTo-Json -Depth 10) | Set-Content -LiteralPath $path -Encoding utf8
    return $path
}

function _WriteLockFile([string]$Root, [string]$Name, [int]$OwnerPid) {
    $p = Join-Path (Join-Path $Root 'locks') ($Name)
    @{ pid = $OwnerPid; createdAtUtc = ([DateTime]::UtcNow.ToString('o')) } |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $p -Encoding utf8
    return $p
}

# --------------------------------------------------------------------------- #
# Test A: Get-NextQCJob ignores dead-owner per-job locks and self-heals them.
# --------------------------------------------------------------------------- #
Write-Host "Test A: Get-NextQCJob skips ALIVE-owner locks but ignores+cleans DEAD-owner locks" -ForegroundColor Cyan
$qroot = _NewQueueRoot
try {
    $cfg = @{ queue = @{ root = $qroot; backend = 'json' } }

    # Three pending jobs:
    #   alpha: no lock                          -> should be selectable
    #   beta:  lock owned by ALIVE pid ($PID)   -> should be skipped
    #   gamma: lock owned by DEAD pid           -> should be selectable AND lock should be removed
    $deadPid = _NewDeadPid
    _WriteJob $qroot 'pending' 'qc_a_alpha' | Out-Null
    _WriteJob $qroot 'pending' 'qc_b_beta'  | Out-Null
    _WriteJob $qroot 'pending' 'qc_c_gamma' | Out-Null
    _WriteLockFile $qroot 'qc_b_beta.lock'  $PID    | Out-Null
    $gammaLock = _WriteLockFile $qroot 'qc_c_gamma.lock' $deadPid

    # Simulate that beta was created last, so by file time alpha < beta < gamma
    # Selection order is by LastWriteTime ascending, so alpha will return first.
    $r1 = Get-NextQCJob -Config $cfg
    _Assert ($r1.IsSuccess -and $r1.Code -eq 'QUEUE_NEXT_JOB') "1st call returns a job"
    _Assert ($r1.Data.jobId -eq 'qc_a_alpha') "1st call returns alpha (oldest, unlocked)"

    # Skip alpha next time (worker already has it). Beta is locked by us (alive)
    # so it must be skipped. Gamma's lock is dead -> selectable.
    $r2 = Get-NextQCJob -Config $cfg -ExcludeJobIds @('qc_a_alpha')
    _Assert ($r2.IsSuccess -and $r2.Code -eq 'QUEUE_NEXT_JOB') "2nd call returns a job (gamma)"
    _Assert ($r2.Data.jobId -eq 'qc_c_gamma') "2nd call returns gamma (beta is alive-locked)"
    _Assert (-not (Test-Path -LiteralPath $gammaLock)) "gamma's dead-owner lock was removed by self-heal"

    # Final pass: only beta is left and its lock is alive, so empty result.
    $r3 = Get-NextQCJob -Config $cfg -ExcludeJobIds @('qc_a_alpha','qc_c_gamma')
    _Assert ($r3.IsSuccess -and $r3.Code -eq 'QUEUE_EMPTY') "3rd call returns empty (beta still alive-locked)"
}
finally { Remove-Item -LiteralPath $qroot -Recurse -Force -ErrorAction SilentlyContinue }

# --------------------------------------------------------------------------- #
# Test B: Lock-QCJob no longer short-circuits on Test-Path; it self-heals.
# --------------------------------------------------------------------------- #
Write-Host "`nTest B: Lock-QCJob steals a dead-owner lock instead of returning QUEUE_LOCK_EXISTS" -ForegroundColor Cyan
$qroot = _NewQueueRoot
try {
    $cfg = @{ queue = @{ root = $qroot; backend = 'json' } }
    $jobId = 'qc_b_lockjob'
    _WriteJob $qroot 'pending' $jobId | Out-Null
    $deadPid = _NewDeadPid
    $deadLock = _WriteLockFile $qroot ($jobId + '.lock') $deadPid

    $r = Lock-QCJob -JobId $jobId -Config $cfg
    _Assert ($r.IsSuccess) "Lock-QCJob succeeds against a dead-owner lock"
    _Assert ($r.Code -eq 'QUEUE_LOCK_ACQUIRED') "result code is QUEUE_LOCK_ACQUIRED"

    # The new lock should now belong to OUR pid.
    $payload = Get-Content -LiteralPath $deadLock -Raw | ConvertFrom-Json
    _Assert ([int]$payload.pid -eq $PID) ("new lock pid is ours ({0})" -f $PID)

    # Job should have been transitioned to running\.
    $runningPath = Join-Path (Join-Path $qroot 'running') ($jobId + '.json')
    _Assert (Test-Path -LiteralPath $runningPath) "job moved from pending\ to running\"
}
finally { Remove-Item -LiteralPath $qroot -Recurse -Force -ErrorAction SilentlyContinue }

# --------------------------------------------------------------------------- #
# Test C: Lock-QCJob still respects ALIVE owners.
# --------------------------------------------------------------------------- #
Write-Host "`nTest C: Lock-QCJob times out (does NOT steal) when owner is alive" -ForegroundColor Cyan
$qroot = _NewQueueRoot
try {
    $cfg = @{ queue = @{ root = $qroot; backend = 'json' } }
    $jobId = 'qc_c_alivelock'
    _WriteJob $qroot 'pending' $jobId | Out-Null
    # Owner = our own PID. _QCQJ-IsLockOwnerDead explicitly returns $false for $PID.
    _WriteLockFile $qroot ($jobId + '.lock') $PID | Out-Null

    $r = Lock-QCJob -JobId $jobId -Config $cfg
    _Assert (-not $r.IsSuccess) "Lock-QCJob fails when owner is alive"
    _Assert ($r.Code -eq 'QUEUE_LOCK_TIMEOUT') "fail code is QUEUE_LOCK_TIMEOUT (no theft of live owner)"
}
finally { Remove-Item -LiteralPath $qroot -Recurse -Force -ErrorAction SilentlyContinue }

# --------------------------------------------------------------------------- #
# Test D: Late Lock-QCJob after another worker moved pending -> running.
# --------------------------------------------------------------------------- #
Write-Host "`nTest D: Lock-QCJob returns QUEUE_JOB_ALREADY_MOVED when job left pending\" -ForegroundColor Cyan
$qroot = _NewQueueRoot
try {
    $cfg = @{ queue = @{ root = $qroot; backend = 'json' } }
    $jobId = 'qc_d_race'
    _WriteJob $qroot 'pending' $jobId | Out-Null

    $first = Lock-QCJob -JobId $jobId -Config $cfg
    _Assert ($first.IsSuccess) 'first Lock-QCJob succeeds'
    Unlock-QCJob -JobId $jobId -Config $cfg | Out-Null

    $second = Lock-QCJob -JobId $jobId -Config $cfg
    _Assert (-not $second.IsSuccess) 'second Lock-QCJob fails (job already running)'
    _Assert ($second.Code -eq 'QUEUE_JOB_ALREADY_MOVED') 'fail code is QUEUE_JOB_ALREADY_MOVED'
}
finally { Remove-Item -LiteralPath $qroot -Recurse -Force -ErrorAction SilentlyContinue }

# --------------------------------------------------------------------------- #
# Test E: Recover-QCStaleJobs janitors orphan locks for jobs in pending\ AND
# clears a stale _queue_write.lock.
# --------------------------------------------------------------------------- #
Write-Host "`nTest E: Recover-QCStaleJobs janitors orphan dead-PID locks (pending\ + global)" -ForegroundColor Cyan
$qroot = _NewQueueRoot
try {
    $cfg = @{ queue = @{ root = $qroot; backend = 'json' } }
    $deadPid = _NewDeadPid

    # 1) Job in pending\ with dead-owner per-job lock.
    _WriteJob $qroot 'pending' 'qc_d_pending' | Out-Null
    $perJobLock = _WriteLockFile $qroot 'qc_d_pending.lock' $deadPid

    # 2) Stale global write lock with dead pid.
    $writeLock = _WriteLockFile $qroot '_queue_write.lock' $deadPid

    # 3) Orphan lock with NO matching job file at all.
    $danglingLock = _WriteLockFile $qroot 'qc_d_dangling.lock' $deadPid

    # 4) Live-owner lock should NOT be removed.
    _WriteJob $qroot 'pending' 'qc_d_alive' | Out-Null
    $aliveLock = _WriteLockFile $qroot 'qc_d_alive.lock' $PID

    $r = Recover-QCStaleJobs -Config $cfg
    _Assert ($r.IsSuccess) "Recover-QCStaleJobs succeeds"
    # orphanLocksRemoved counts the janitor-pass removals (per-job + dangling).
    # _queue_write.lock is stolen by the recover function's own AcquireLockFile
    # at the start of the pass and then released (deleted) on exit, so it
    # doesn't go through the janitor counter, but the file IS gone.
    _Assert ($r.Data.orphanLocksRemoved -ge 2) ("orphanLocksRemoved >= 2 (got {0})" -f $r.Data.orphanLocksRemoved)
    _Assert (-not (Test-Path -LiteralPath $perJobLock))    "pending\ job's dead-owner lock removed"
    _Assert (-not (Test-Path -LiteralPath $writeLock))     "global _queue_write.lock removed"
    _Assert (-not (Test-Path -LiteralPath $danglingLock))  "dangling dead-owner lock removed"
    _Assert (Test-Path -LiteralPath $aliveLock)            "alive-owner lock preserved"
}
finally { Remove-Item -LiteralPath $qroot -Recurse -Force -ErrorAction SilentlyContinue }

if ($failures -gt 0) { Write-Host "`nFAILED ($failures)" -ForegroundColor Red; exit 1 }
Write-Host "`nPASSED" -ForegroundColor Green
