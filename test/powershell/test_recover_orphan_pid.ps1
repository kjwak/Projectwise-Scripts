<#
.SYNOPSIS
Test: Recover-QCStaleJobs reclaims orphaned running\ jobs whose lock-owner PID is dead,
even when the job's startedAtUtc is younger than queue.recover.staleSeconds.
#>

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force -DisableNameChecking | Out-Null
Import-Module (Join-Path $repoRoot 'modules\Queue\QC.Queue.Json.psm1') -Force -DisableNameChecking | Out-Null

function New-StubJob([string]$Id, [string]$Type = 'STATUS_SET_GEN') {
    return @{
        id          = $Id
        type        = $Type
        sourcePath  = "documents\x\$Id.pdf"
        sourceName  = "$Id.pdf"
        triggerRule = @{ id = 'r1'; jobType = $Type }
        dedupeKey   = "dq_$Id"
        status      = 'pending'
        createdAt   = ([DateTime]::UtcNow.ToString('o'))
        attempts    = 0
        metadata    = @{}
    }
}

$tmpRoot = Join-Path $env:TEMP ("QC_Test_RecoverOrphan_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmpRoot -Force | Out-Null

try {
    $cfg = @{
        queue = @{
            rootDir = $tmpRoot
            recover = @{
                staleSeconds = 99999
                maxAttempts  = 5
            }
        }
    }

    $idA = 'orphan-deadpid'
    $idB = 'orphan-nolock'
    $idC = 'live-worker'

    foreach ($id in @($idA,$idB,$idC)) {
        $j = New-StubJob -Id $id
        $r = Add-QCQueueJob -Job $j -Config $cfg
        if (-not $r.IsSuccess) { throw "Enqueue $id failed: $($r.Message)" }
    }

    $lockA = Lock-QCJob -JobId $idA -Config $cfg
    if (-not $lockA.IsSuccess) { throw "Lock A failed: $($lockA.Message)" }
    $lockB = Lock-QCJob -JobId $idB -Config $cfg
    if (-not $lockB.IsSuccess) { throw "Lock B failed: $($lockB.Message)" }
    $lockC = Lock-QCJob -JobId $idC -Config $cfg
    if (-not $lockC.IsSuccess) { throw "Lock C failed: $($lockC.Message)" }

    # Job A: rewrite lock with a known-dead PID (use a PID we can't realistically have alive,
    # by spawning a sleeper, recording its PID, then killing it).
    $sleeper = Start-Process -FilePath 'powershell' -ArgumentList @('-NoProfile','-Command','Start-Sleep -Seconds 60') -PassThru -WindowStyle Hidden
    $deadPid = $sleeper.Id
    Stop-Process -Id $deadPid -Force
    Start-Sleep -Milliseconds 200

    $lockPathA = Join-Path $tmpRoot ("locks\$idA.lock")
    $payloadA = @{ pid = $deadPid; createdAtUtc = ([DateTime]::UtcNow.ToString('o')) } | ConvertTo-Json -Compress
    Set-Content -LiteralPath $lockPathA -Value $payloadA -Encoding utf8

    # Job B: remove the lock file entirely (simulates partial cleanup).
    $lockPathB = Join-Path $tmpRoot ("locks\$idB.lock")
    Remove-Item -LiteralPath $lockPathB -Force

    # Job C: rewrite lock with the current process PID (live owner -> must NOT be reclaimed).
    $lockPathC = Join-Path $tmpRoot ("locks\$idC.lock")
    $payloadC = @{ pid = $PID; createdAtUtc = ([DateTime]::UtcNow.ToString('o')) } | ConvertTo-Json -Compress
    Set-Content -LiteralPath $lockPathC -Value $payloadC -Encoding utf8

    $rec = Recover-QCStaleJobs -Config $cfg
    if (-not $rec.IsSuccess) { throw "Recover failed: $($rec.Message)" }
    if ([int]$rec.Data.recoveredOrphan -ne 2) { throw "Expected recoveredOrphan=2, got $($rec.Data.recoveredOrphan)" }
    if ([int]$rec.Data.recoveredToPending -ne 2) { throw "Expected recoveredToPending=2, got $($rec.Data.recoveredToPending)" }
    if ([int]$rec.Data.skippedNotStale -ne 1) { throw "Expected skippedNotStale=1 (live PID job C), got $($rec.Data.skippedNotStale)" }

    $pendDir = Join-Path $tmpRoot 'pending'
    $runDir  = Join-Path $tmpRoot 'running'
    if (-not (Test-Path -LiteralPath (Join-Path $pendDir "$idA.json"))) { throw "Job A not requeued to pending\" }
    if (-not (Test-Path -LiteralPath (Join-Path $pendDir "$idB.json"))) { throw "Job B not requeued to pending\" }
    if (-not (Test-Path -LiteralPath (Join-Path $runDir  "$idC.json"))) { throw "Job C should still be in running\ (live PID)" }

    Write-Host "test_recover_orphan_pid: PASS" -ForegroundColor Green
}
catch {
    Write-Host "test_recover_orphan_pid: FAIL" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    throw
}
finally {
    Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
}
