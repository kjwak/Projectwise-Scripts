$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module "$repoRoot/modules/QC.WatcherOrchestration.psm1" -Force

$config = @{
    watcher = @{
        stallRecovery = @{
            enabled = $true
            noLogActivitySeconds = 600
            auditScanMaxSeconds = 300
        }
    }
}

$settings = Get-QCWatcherStallRecoverySettings -Config $config
Assert-Eq $settings.auditScanMaxSeconds 300 'audit scan threshold from config'
Assert-Eq $settings.noLogActivitySeconds 600 'no-log threshold from config'

$now = [datetime]'2026-06-18T13:40:00Z'
$healthy = Test-QCWatcherChildStalled -Settings $settings -WatcherAlive $true `
    -LastLogActivityUtc $now.AddSeconds(-30) `
    -LastWatcherEventCode 'WATCH_TICK_SLEEP' `
    -LastWatcherEventUtc $now.AddSeconds(-30) `
    -NowUtc $now
Assert-True (-not $healthy.stalled) 'recent log activity should not be stalled'

$auditStall = Test-QCWatcherChildStalled -Settings $settings -WatcherAlive $true `
    -LastLogActivityUtc $now.AddSeconds(-120) `
    -LastWatcherEventCode 'WATCH_AUDIT_SCAN_START' `
    -LastWatcherEventUtc $now.AddSeconds(-400) `
    -NowUtc $now
Assert-True $auditStall.stalled 'audit scan start without completion should stall'
Assert-Eq $auditStall.reason 'audit_scan_timeout' 'audit stall reason'

$silent = Test-QCWatcherChildStalled -Settings $settings -WatcherAlive $true `
    -LastLogActivityUtc $now.AddSeconds(-900) `
    -LastWatcherEventCode 'WATCH_DONE' `
    -LastWatcherEventUtc $now.AddSeconds(-900) `
    -NowUtc $now
Assert-True $silent.stalled 'long silence should stall'
Assert-Eq $silent.reason 'no_log_activity' 'silent stall reason'

$disabled = @{ enabled = $false; noLogActivitySeconds = 60; auditScanMaxSeconds = 60 }
$off = Test-QCWatcherChildStalled -Settings $disabled -WatcherAlive $true `
    -LastLogActivityUtc $now.AddSeconds(-900) -NowUtc $now
Assert-True (-not $off.stalled) 'disabled stall recovery should not trigger'

$spawned = $now.AddSeconds(-60)
$replay = Test-QCWatcherChildStalled -Settings $settings -WatcherAlive $true `
    -LastLogActivityUtc $now.AddSeconds(-400) `
    -LastWatcherEventCode 'WATCH_AUDIT_SCAN_START' `
    -LastWatcherEventUtc $now.AddSeconds(-400) `
    -WatcherSpawnedAtUtc $spawned `
    -NowUtc $now
Assert-True (-not $replay.stalled) 'pre-spawn audit events should not trigger stall on respawn'

Write-Host 'OK: QC watcher stall recovery tests passed.' -ForegroundColor Green
