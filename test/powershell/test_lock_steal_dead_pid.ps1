<#
.SYNOPSIS
Regression: _QCQJ-AcquireLockFile must steal a stale lock file whose owner PID
is dead. Without this, _queue_write.lock or per-job .lock files left by
AV-killed workers pin the entire queue: every subsequent worker spins for the
TimeoutMs on QUEUE_LOCK_TIMEOUT until a dashboard restart.
#>

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $PSScriptRoot '_Resolve-ModuleImplPath.ps1')
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force -DisableNameChecking | Out-Null
Import-Module (Join-Path $repoRoot 'modules\Queue\QC.Queue.Json.psm1') -Force -DisableNameChecking | Out-Null

function Assert-True($Cond, $Msg) {
    if (-not $Cond) { throw "ASSERT FAILED: $Msg" }
}

$tmp = Join-Path $env:TEMP ("qc_test_locksteal_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null

try {
    # Manufacture a dead PID by spawning a sleeper, recording its PID, killing it.
    $sleeper = Start-Process -FilePath 'powershell' -ArgumentList @('-NoProfile','-Command','Start-Sleep -Seconds 60') -PassThru -WindowStyle Hidden
    $deadPid = $sleeper.Id
    Stop-Process -Id $deadPid -Force
    Start-Sleep -Milliseconds 200

    $module = Get-QCModuleImplementation -ModuleName 'QC.Queue.Json.psm1'
    if (-not $module) { throw 'QC.Queue.Json implementation module is not loaded' }
    $acquire = { param($Path,$ms) & $module { param($P,$M) _QCQJ-AcquireLockFile -LockPath $P -TimeoutMs $M -SleepMs 50 } $Path $ms }

    # 1) Lock file owned by dead PID -> AcquireLockFile must steal it within TimeoutMs.
    $lock1 = Join-Path $tmp 'dead.lock'
    $payload = @{ pid = $deadPid; createdAtUtc = ([DateTime]::UtcNow.ToString('o')) } | ConvertTo-Json -Compress
    Set-Content -LiteralPath $lock1 -Value $payload -Encoding utf8

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $ok = & $acquire $lock1 5000
    $sw.Stop()
    Assert-True $ok 'AcquireLockFile should steal a dead-PID lock'
    Assert-True ($sw.ElapsedMilliseconds -lt 1500) ("Steal should be near-instant; got $($sw.ElapsedMilliseconds) ms")
    # After steal, the lock should now belong to current PID.
    $after = Get-Content -LiteralPath $lock1 -Raw | ConvertFrom-Json
    Assert-True ([int]$after.pid -eq $PID) "Stolen lock should record current PID, got $($after.pid)"
    Assert-True ([string]$after.machineName -eq $env:COMPUTERNAME) "Stolen lock should record machineName"

    # cleanup
    Remove-Item -LiteralPath $lock1 -Force

    # 1b) Other-host lock with a dead PID must NOT be stolen (PID is not meaningful here).
    $lockRemote = Join-Path $tmp 'remote.lock'
    $payloadRemote = @{
        pid = $deadPid
        machineName = 'QC-REMOTE-TEST-HOST'
        createdAtUtc = ([DateTime]::UtcNow.ToString('o'))
    } | ConvertTo-Json -Compress
    Set-Content -LiteralPath $lockRemote -Value $payloadRemote -Encoding utf8
    $okRemote = & $acquire $lockRemote 800
    Assert-True (-not $okRemote) 'AcquireLockFile must not steal another host per-job lock based on local PID'
    $stillRemote = Get-Content -LiteralPath $lockRemote -Raw | ConvertFrom-Json
    Assert-True ([string]$stillRemote.machineName -eq 'QC-REMOTE-TEST-HOST') 'remote lock payload should be unchanged'
    Remove-Item -LiteralPath $lockRemote -Force

    # 2) Lock file owned by a LIVE PID (this process) -> must NOT be stolen, must time out.
    $lock2 = Join-Path $tmp 'live.lock'
    $payload2 = @{ pid = $PID; createdAtUtc = ([DateTime]::UtcNow.ToString('o')) } | ConvertTo-Json -Compress
    Set-Content -LiteralPath $lock2 -Value $payload2 -Encoding utf8

    $sw2 = [System.Diagnostics.Stopwatch]::StartNew()
    $ok2 = & $acquire $lock2 1500
    $sw2.Stop()
    Assert-True (-not $ok2) 'AcquireLockFile must NOT steal a live-PID lock'
    Assert-True ($sw2.ElapsedMilliseconds -ge 1400) ("Should wait the full TimeoutMs when owner is alive; got $($sw2.ElapsedMilliseconds) ms")

    # cleanup
    Remove-Item -LiteralPath $lock2 -Force

    # 3) Garbage payload (no parseable JSON / no pid) -> must NOT be stolen, must time out.
    # Conservative: we can't tell who owns it, so we leave it alone.
    $lock3 = Join-Path $tmp 'garbage.lock'
    Set-Content -LiteralPath $lock3 -Value 'not json' -Encoding utf8
    $ok3 = & $acquire $lock3 800
    Assert-True (-not $ok3) 'AcquireLockFile must not steal a lock with unparseable payload'

    Write-Host "test_lock_steal_dead_pid: PASS" -ForegroundColor Green
}
finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
