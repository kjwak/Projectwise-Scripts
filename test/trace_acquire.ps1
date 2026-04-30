$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $root 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $root 'modules\QC.Queue.Json.psm1') -Force

$lockPath = 'C:\QC_E2E_RealRun\queue\locks\_queue_write.lock'

if (-not (Test-Path -LiteralPath $lockPath)) {
    Write-Host "lock not present right now; nothing to test" -ForegroundColor Yellow
    exit 0
}

Write-Host "Attempting to acquire stale lock: $lockPath" -ForegroundColor Cyan
$mod = Get-Module QC.Queue.Json
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$got = & $mod { param($p) _QCQJ-AcquireLockFile -LockPath $p -TimeoutMs 6000 -SleepMs 100 } $lockPath
$sw.Stop()
Write-Host ("acquire returned: {0}    elapsed: {1} ms" -f $got, $sw.ElapsedMilliseconds) -ForegroundColor Yellow

if ($got) {
    Write-Host "Releasing immediately so we don't disturb the running queue."
    & $mod { param($p) _QCQJ-ReleaseLockFile -LockPath $p } $lockPath
}
