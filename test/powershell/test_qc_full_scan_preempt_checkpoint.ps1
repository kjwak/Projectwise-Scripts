$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module "$repoRoot/modules/Core/QC.WatcherOrchestration.psm1" -Force

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("qc-fullscan-test-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path (Join-Path $tempRoot '_watcher') -Force | Out-Null
$config = @{
    queue = @{ rootDir = $tempRoot }
    database = @{ enabled = $false }
    auditPoller = @{
        fullScanSchedule = @{
            times = @('06:00', '18:00')
            preempt = @{
                enabled = $true
                checkEveryNFolders = 2
            }
        }
    }
}

# --- preempt settings ---
$preempt = Get-QCFullScanPreemptSettings -Config $config
Assert-True $preempt.enabled 'preempt enabled from config'
Assert-Eq $preempt.checkEveryNFolders 2 'checkEveryNFolders from config'

$preemptOff = Get-QCFullScanPreemptSettings -Config @{ auditPoller = @{ fullScanSchedule = @{ preempt = @{ enabled = $false } } } }
Assert-True (-not $preemptOff.enabled) 'preempt can be disabled'
Assert-Eq $preemptOff.checkEveryNFolders 1 'default checkEveryNFolders is 1'

# --- progress checkpoint ---
$slotKey = 'full_scan_schedule|2026-07-17|06:00'
$queue = @(
    @{ FolderPath = 'Documents\A\Sheets'; EnableStatusSet = $true; EnableQcPrepend = $true; ScanReason = 'full_reconciliation' }
    @{ FolderPath = 'Documents\B\Sheets'; EnableStatusSet = $true; EnableQcPrepend = $false; ScanReason = 'full_reconciliation' }
    @{ FolderPath = 'Documents\C\Sheets'; EnableStatusSet = $false; EnableQcPrepend = $true; ScanReason = 'full_reconciliation' }
)
$ok = Set-QCFullScanProgress -Config $config -SlotKey $slotKey -CompletedFolders @('Documents\A\Sheets') -FolderQueue @($queue[1], $queue[2]) -QueueRoot $tempRoot
Assert-True $ok 'progress persist should succeed'

$loaded = Get-QCFullScanProgress -Config $config -QueueRoot $tempRoot -SlotKey $slotKey
Assert-Eq $loaded.slotKey $slotKey 'loaded slotKey'
Assert-Eq $loaded.completedFolders.Count 1 'one completed folder'
Assert-Eq $loaded.folderQueue.Count 2 'two remaining folders'
Assert-Eq ([string]$loaded.folderQueue[0].FolderPath) 'Documents\B\Sheets' 'first remaining path'

$wrongSlot = Get-QCFullScanProgress -Config $config -QueueRoot $tempRoot -SlotKey 'full_scan_schedule|2026-07-17|18:00'
Assert-Eq $wrongSlot.folderQueue.Count 0 'mismatched slot returns empty queue'
Assert-Eq $wrongSlot.completedFolders.Count 0 'mismatched slot returns empty completed'

[void](Clear-QCFullScanProgress -Config $config -SlotKey $slotKey -QueueRoot $tempRoot)
$cleared = Get-QCFullScanProgress -Config $config -QueueRoot $tempRoot -SlotKey $slotKey
Assert-Eq $cleared.folderQueue.Count 0 'cleared progress has empty queue'
Assert-True (-not (Test-Path -LiteralPath (Join-Path $tempRoot '_watcher\full-scan-progress.json'))) 'progress file removed'

# --- long-running work is timer-free and invokes same-runspace progress ---
$progressHits = [ref]0
$work = {
    param($Progress)
    Assert-True ($null -ne $Progress) 'Progress callback should be provided'
    & $Progress @{ step = 'mid' }
    $progressHits.Value++
    return 42
}
$result = Invoke-QCWatcherLongRunningWork -Phase 'unit_test_phase' -Data @{ folder = 'Documents\Test' } -HeartbeatIntervalSeconds 60 -Work $work
Assert-Eq $result 42 'work return value'
Assert-Eq $progressHits.Value 1 'progress callback invoked once from work'

# Pipeline pollution from incidental cmdlet output must not change the intended return value
$pollutedWork = {
    param($Progress)
    [pscustomobject]@{ IsSuccess = $true; Code = 'SIDE_EFFECT' }
    [pscustomobject]@{ IsSuccess = $true; Code = 'SIDE_EFFECT_2' }
    return 406
}
$pollutedResult = Invoke-QCWatcherLongRunningWork -Phase 'unit_test_pollution' -Data @{} -HeartbeatIntervalSeconds 60 -Work $pollutedWork
Assert-Eq $pollutedResult 406 'polluted pipeline still returns last (intended) value'
Assert-True ($pollutedResult -is [int] -or $pollutedResult -is [long]) 'polluted result is scalar numeric, not Object[]'

# Ensure no System.Threading.Timer is constructed by the helper
$fn = Get-Command Invoke-QCWatcherLongRunningWork
$def = $fn.ScriptBlock.ToString()
Assert-True ($def -notmatch 'New-Object\s+System\.Threading\.Timer') 'Invoke-QCWatcherLongRunningWork must not construct System.Threading.Timer'
Assert-True ($def -notmatch '\[System\.Threading\.TimerCallback\]') 'Invoke-QCWatcherLongRunningWork must not use TimerCallback'

Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
Write-Host 'OK: QC full-scan preempt/checkpoint/timer-free tests passed.' -ForegroundColor Green
