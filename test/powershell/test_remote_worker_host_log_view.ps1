# Remote host console: busy summary follows running\, labels reuse lowest RW* gap.
$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ("$Actual" -ne "$Expected") { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $repoRoot 'scripts\service\QC.RemoteWorkerHostLogView.ps1')

Initialize-QCRemoteWorkerHostLogView

Assert-Eq (Get-QCNextRemoteWorkerLabel -CurrentLabels @()) 'RW1' 'empty pool starts at RW1'
Assert-Eq (Get-QCNextRemoteWorkerLabel -CurrentLabels @('RW1', 'RW2', 'RW4')) 'RW3' 'reuses lowest gap'
Assert-Eq (Get-QCNextRemoteWorkerLabel -CurrentLabels @('RW1', 'RW2', 'RW102')) 'RW3' 'does not increment past extras'

$tempRoot = Join-Path $env:TEMP ('qc-rw-console-' + ([guid]::NewGuid().ToString('N')))
$queueDir = Join-Path $tempRoot 'queue'
$runningDir = Join-Path $queueDir 'running'
New-Item -ItemType Directory -Path $runningDir -Force | Out-Null
$hostName = [string]$env:COMPUTERNAME

function Write-RunningJob($Id, $Name, $Machine) {
    $obj = @{
        id = $Id
        type = 'QC_PREPEND'
        sourceName = $Name
        machineName = $Machine
        status = 'running'
    }
    ($obj | ConvertTo-Json -Compress) | Set-Content -LiteralPath (Join-Path $runningDir ($Id + '.json')) -Encoding utf8
}

try {
    Write-RunningJob 'job-local' 'ea003.pdf' $hostName
    Write-RunningJob 'job-server' 'notify.pdf' 'PXBENTLEY01'

    $mine = @(Get-QCRemoteWorkerHostRunningJobs -QueueRoot $queueDir -HostName $hostName)
    Assert-Eq $mine.Count 1 'this-host view excludes other machines'
    Assert-Eq $mine[0].sourceName 'ea003.pdf' 'this-host job name'

    $all = @(Get-QCRemoteWorkerHostRunningJobs -QueueRoot $queueDir -HostName $hostName -AllHosts)
    Assert-Eq $all.Count 2 'AllHosts includes every running job'

    $busy = Get-QCRemoteWorkerHostBusySummary -QueueRoot $queueDir -HostName $hostName
    Assert-True ($busy -eq 'busy=1 ea003.pdf') ("queue busy summary must match running\\, got: $busy")

    $idleOther = Get-QCRemoteWorkerHostBusySummary -QueueRoot $queueDir -HostName 'NO-SUCH-HOST'
    Assert-Eq $idleOther 'idle' 'unknown host with no matching running jobs is idle'

    Initialize-QCRemoteWorkerHostLogView
    $seed = Sync-QCRemoteWorkerHostRunningView -QueueRoot $queueDir -HostName $hostName -Quiet
    Assert-Eq @($seed.Jobs).Count 1 'seed sees the local running job'
    Assert-Eq @($seed.Entered).Count 0 'quiet seed does not treat existing jobs as new claims'

    $selected = Write-QCRemoteWorkerHostJobLine -Obj ([pscustomobject]@{
        code = 'WORKER_SELECTED'
        message = 'Job locked and running (exclusive owner).'
        data = [pscustomobject]@{
            workerLabel = 'RW2'
            jobId = 'job-local'
            jobType = 'QC_PREPEND'
            sourceName = 'ea003.pdf'
        }
    })
    Assert-True $selected 'duplicate SELECTED is consumed'
    Assert-Eq $script:QCRemoteWorkerLogView.ActiveJobs['job-local'] 'ea003.pdf' 'seeded job stays tracked'

    Write-RunningJob 'job-new' 'ea004.pdf' $hostName
    $live = Sync-QCRemoteWorkerHostRunningView -QueueRoot $queueDir -HostName $hostName
    Assert-Eq @($live.Entered).Count 1 'new running file is a claim'
    Assert-Eq $live.Entered[0].sourceName 'ea004.pdf' 'new claim name comes from running\\'
    Assert-Eq @($live.Jobs).Count 2 'two local jobs now running'

    $busy2 = Get-QCRemoteWorkerHostBusySummary -QueueRoot $queueDir -HostName $hostName
    Assert-True ($busy2 -match '^busy=2 ') ("parallel running jobs must show count 2, got: $busy2")
    Assert-True ($busy2 -match 'ea003') 'busy summary includes first sheet'
    Assert-True ($busy2 -match 'ea004') 'busy summary includes second sheet'
} finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

$hostScript = Get-Content -LiteralPath (Join-Path $repoRoot 'scripts\service\Start-QCRemoteWorkerHost.ps1') -Raw
$watchScript = Get-Content -LiteralPath (Join-Path $repoRoot 'scripts\service\Watch-QCRemoteWorkerHostConsole.ps1') -Raw
Assert-True ($hostScript -match 'Get-QCNextRemoteWorkerLabel') 'supervisor reuses lowest RW label'
Assert-True ($hostScript -match '_Test-WorkerSlotAlive') 'supervisor reaps via live PID not stale HasExited'
Assert-True ($hostScript -match 'Sync-QCRemoteWorkerHostRunningView') 'supervisor follows running\\'
Assert-True ($watchScript -match 'Sync-QCRemoteWorkerHostRunningView') 'logon console follows running\\'
Assert-True ($watchScript -match 'Get-QCRemoteWorkerHostBusySummary -QueueRoot') 'logon heartbeat uses queue snapshot'

Write-Host 'test_remote_worker_host_log_view: passed' -ForegroundColor Green
