<#
.SYNOPSIS
Host-aware queue locks: never treat another machine's PID as local liveness.
#>
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $root 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $root 'modules\Queue\QC.Queue.Json.psm1') -Force

$failures = 0
function _Assert([bool]$ok, [string]$msg) {
    if (-not $ok) { Write-Host "  FAIL: $msg" -ForegroundColor Red; $script:failures++ }
    else { Write-Host "  ok:   $msg" -ForegroundColor Green }
}

function _NewDeadPid {
    $p = Start-Process -FilePath 'cmd.exe' -ArgumentList '/c','exit' -PassThru -WindowStyle Hidden
    $p.WaitForExit()
    Start-Sleep -Milliseconds 100
    return [int]$p.Id
}

function _NewQueueRoot {
    $r = Join-Path $env:TEMP ("qctest_hostlock_" + [guid]::NewGuid().ToString('N').Substring(0,8))
    foreach ($s in 'pending','running','succeeded','failed','locks') {
        New-Item -ItemType Directory -Path (Join-Path $r $s) -Force | Out-Null
    }
    return $r
}

function _WriteJob([string]$Root, [string]$State, [string]$JobId, [hashtable]$Extra = @{}) {
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

function _WriteHostLock([string]$Root, [string]$Name, [int]$OwnerPid, [string]$MachineName, [datetime]$CreatedUtc) {
    $p = Join-Path (Join-Path $Root 'locks') $Name
    @{
        pid = $OwnerPid
        machineName = $MachineName
        createdAtUtc = $CreatedUtc.ToString('o')
        heartbeatUtc = $CreatedUtc.ToString('o')
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $p -Encoding utf8
    return $p
}

$remoteHost = 'QC-REMOTE-TEST-HOST'
$deadPid = _NewDeadPid

Write-Host 'Test A: Get-NextQCJob skips fresh other-host pending lock; reclaims stale one' -ForegroundColor Cyan
$qroot = _NewQueueRoot
try {
    $cfg = @{ queue = @{ root = $qroot; recover = @{ crossHostLockStaleSeconds = 90 } } }
    _WriteJob $qroot 'pending' 'qc_fresh' | Out-Null
    _WriteJob $qroot 'pending' 'qc_stale' | Out-Null
    $freshLock = _WriteHostLock $qroot 'qc_fresh.lock' $deadPid $remoteHost ([DateTime]::UtcNow)
    $staleLock = _WriteHostLock $qroot 'qc_stale.lock' $deadPid $remoteHost ([DateTime]::UtcNow.AddMinutes(-5))

    $r1 = Get-NextQCJob -Config $cfg
    _Assert ($r1.IsSuccess -and $r1.Data.jobId -eq 'qc_stale') 'stale other-host pending lock is selectable'
    _Assert (-not (Test-Path -LiteralPath $staleLock)) 'stale other-host pending lock was reclaimed'
    _Assert (Test-Path -LiteralPath $freshLock) 'fresh other-host pending lock remains'

    $r2 = Get-NextQCJob -Config $cfg -ExcludeJobIds @('qc_stale')
    _Assert ($r2.IsSuccess -and $r2.Code -eq 'QUEUE_EMPTY') 'fresh other-host lock blocks selection'
}
finally { Remove-Item -LiteralPath $qroot -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host "`nTest B: Lock-QCJob does not steal a fresh other-host per-job lock" -ForegroundColor Cyan
$qroot = _NewQueueRoot
try {
    $cfg = @{ queue = @{ root = $qroot; lockAcquireTimeoutMs = 800; lockAcquireSleepMs = 50 } }
    $jobId = 'qc_remote_lock'
    _WriteJob $qroot 'pending' $jobId | Out-Null
    _WriteHostLock $qroot ($jobId + '.lock') $deadPid $remoteHost ([DateTime]::UtcNow) | Out-Null
    $r = Lock-QCJob -JobId $jobId -Config $cfg
    _Assert (-not $r.IsSuccess) 'Lock-QCJob fails against fresh other-host lock'
    _Assert ($r.Code -eq 'QUEUE_LOCK_TIMEOUT') 'fail code is QUEUE_LOCK_TIMEOUT'
}
finally { Remove-Item -LiteralPath $qroot -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host "`nTest C: Recover-QCStaleJobs does not reclaim a fresh other-host running job" -ForegroundColor Cyan
$qroot = _NewQueueRoot
try {
    $cfg = @{ queue = @{ rootDir = $qroot; recover = @{ staleSeconds = 900; maxAttempts = 5 } } }
    $now = [DateTime]::UtcNow
    _WriteJob $qroot 'running' 'qc_remote_live' @{
        startedAtUtc = $now.ToString('o')
        heartbeatUtc = $now.ToString('o')
        status = 'running'
    } | Out-Null
    _WriteHostLock $qroot 'qc_remote_live.lock' $deadPid $remoteHost $now | Out-Null

    $rec = Recover-QCStaleJobs -Config $cfg
    _Assert $rec.IsSuccess 'Recover succeeds'
    _Assert ([int]$rec.Data.recoveredToPending -eq 0) 'fresh remote running job not requeued'
    _Assert (Test-Path -LiteralPath (Join-Path $qroot 'running\qc_remote_live.json')) 'job remains in running\'
}
finally { Remove-Item -LiteralPath $qroot -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host "`nTest D: Recover-QCStaleJobs requeues other-host running job with stale heartbeat" -ForegroundColor Cyan
$qroot = _NewQueueRoot
try {
    $cfg = @{ queue = @{ rootDir = $qroot; recover = @{ staleSeconds = 900; maxAttempts = 5 } } }
    $old = [DateTime]::UtcNow.AddHours(-2)
    _WriteJob $qroot 'running' 'qc_remote_stale' @{
        startedAtUtc = $old.ToString('o')
        heartbeatUtc = $old.ToString('o')
        status = 'running'
    } | Out-Null
    _WriteHostLock $qroot 'qc_remote_stale.lock' $deadPid $remoteHost $old | Out-Null

    $rec = Recover-QCStaleJobs -Config $cfg
    _Assert $rec.IsSuccess 'Recover succeeds'
    _Assert ([int]$rec.Data.recoveredToPending -eq 1) 'stale remote running job requeued'
    _Assert (Test-Path -LiteralPath (Join-Path $qroot 'pending\qc_remote_stale.json')) 'job moved to pending\'
    _Assert (-not (Test-Path -LiteralPath (Join-Path $qroot 'locks\qc_remote_stale.lock'))) 'per-job lock released'
}
finally { Remove-Item -LiteralPath $qroot -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host "`nTest E: Recover janitor removes stale other-host pending lock" -ForegroundColor Cyan
$qroot = _NewQueueRoot
try {
    $cfg = @{ queue = @{ rootDir = $qroot; recover = @{ crossHostLockStaleSeconds = 90 } } }
    _WriteJob $qroot 'pending' 'qc_pending_remote' | Out-Null
    $lock = _WriteHostLock $qroot 'qc_pending_remote.lock' $deadPid $remoteHost ([DateTime]::UtcNow.AddMinutes(-5))
    $rec = Recover-QCStaleJobs -Config $cfg
    _Assert $rec.IsSuccess 'Recover succeeds'
    _Assert (-not (Test-Path -LiteralPath $lock)) 'stale other-host pending lock janitored'
}
finally { Remove-Item -LiteralPath $qroot -Recurse -Force -ErrorAction SilentlyContinue }

if ($failures -gt 0) { Write-Host "`nFAILED ($failures)" -ForegroundColor Red; exit 1 }
Write-Host "`nPASSED" -ForegroundColor Green
