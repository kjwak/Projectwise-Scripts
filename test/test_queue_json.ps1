# Unit tests for QC.Queue.Json (filesystem queue port).
$ErrorActionPreference = 'Stop'

Import-Module "$PSScriptRoot/../modules/Core.Results.psm1" -Force
Import-Module "$PSScriptRoot/../modules/QC.Queue.Json.psm1" -Force

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}

function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

function New-TestJob([string]$Id, [string]$DedupeKey, [string]$Type = 'QC_PREPEND') {
    return @{
        id = $Id
        type = $Type
        sourcePath = "documents\\x\\$Id.pdf"
        sourceName = "$Id.pdf"
        triggerRule = @{ id = 'r1'; jobType = $Type }
        dedupeKey = $DedupeKey
        status = 'pending'
        createdAt = '2026-01-01T00:00:00.0000000Z'
        attempts = 0
        metadata = @{}
    }
}

$root = Join-Path $env:TEMP ("qc-queue-test-" + ([guid]::NewGuid().ToString('N')))
New-Item -ItemType Directory -Path $root -Force | Out-Null

try {
    $config = @{
        queue = @{
            rootDir = $root
            recover = @{ staleSeconds = 1; maxAttempts = 2 }
        }
    }

    # enqueue
    $jobA = New-TestJob -Id 'job-a' -DedupeKey 'dq_a'
    $enqA = Add-QCQueueJob -Job $jobA -Config $config
    Assert-True $enqA.IsSuccess 'Add-QCQueueJob should succeed'

    $stats1 = Get-QCQueueStats -Config $config
    Assert-True $stats1.IsSuccess 'Get-QCQueueStats should succeed'
    Assert-Eq $stats1.Data.states.pending 1 'Pending count should be 1 after enqueue'

    # duplicate detection
    $dup = Test-QCDuplicateJob -DedupeKey 'dq_a' -Config $config
    Assert-True ($dup.IsSuccess -and $dup.Data.isDuplicate) 'Duplicate check should return true for existing dedupeKey'

    $noDup = Test-QCDuplicateJob -DedupeKey 'dq_missing' -Config $config
    Assert-True ($noDup.IsSuccess -and -not $noDup.Data.isDuplicate) 'Duplicate check should be false for missing dedupeKey'

    # get next job (oldest)
    Start-Sleep -Milliseconds 50
    $jobB = New-TestJob -Id 'job-b' -DedupeKey 'dq_b'
    $enqB = Add-QCQueueJob -Job $jobB -Config $config
    Assert-True $enqB.IsSuccess 'Second enqueue should succeed'

    $next1 = Get-NextQCJob -Config $config
    Assert-True ($next1.IsSuccess -and $next1.Data.job) 'Get-NextQCJob should return a job'
    Assert-Eq $next1.Data.job['id'] 'job-a' 'Oldest pending should be returned first'

    # lock/unlock + pending->running transition on lock
    $lockA = Lock-QCJob -JobId 'job-a' -Config $config
    Assert-True $lockA.IsSuccess 'Lock-QCJob should succeed'

    $stats2 = Get-QCQueueStats -Config $config
    Assert-Eq $stats2.Data.states.pending 1 'Lock should move job-a out of pending'
    Assert-Eq $stats2.Data.states.running 1 'Lock should move job-a into running'

    $next2 = Get-NextQCJob -Config $config
    Assert-True ($next2.IsSuccess -and $next2.Data.job) 'Get-NextQCJob should still return a job'
    Assert-Eq $next2.Data.job['id'] 'job-b' 'Locked/running job should not be returned as pending'

    $unlockA = Unlock-QCJob -JobId 'job-a' -Config $config
    Assert-True $unlockA.IsSuccess 'Unlock-QCJob should succeed'

    # status update
    $st = Set-QCJobStatus -JobId 'job-a' -Status 'running' -Config $config
    Assert-True $st.IsSuccess 'Set-QCJobStatus should succeed'
    $getA = Get-QCJobById -JobId 'job-a' -Config $config
    Assert-True ($getA.IsSuccess -and $getA.Data.found) 'Get-QCJobById should find job-a'
    Assert-Eq $getA.Data.state 'running' 'job-a should be in running after lock'
    Assert-Eq $getA.Data.job.status 'running' 'Status should be updated in job file'

    # move job
    $mv = Move-QCJob -JobId 'job-a' -FromState 'running' -ToState 'succeeded' -Config $config
    Assert-True $mv.IsSuccess 'Move-QCJob should succeed'
    $stats3 = Get-QCQueueStats -Config $config
    Assert-Eq $stats3.Data.states.running 0 'Running count should be 0 after move'
    Assert-Eq $stats3.Data.states.succeeded 1 'Succeeded count should be 1 after move'

    # stale recovery (running -> pending, then running -> failed based on attempts)
    $lockB = Lock-QCJob -JobId 'job-b' -Config $config
    Assert-True $lockB.IsSuccess 'Lock job-b should succeed'

    # Force startedAtUtc far in the past.
    $bLoc = Get-QCJobById -JobId 'job-b' -Config $config
    Assert-True ($bLoc.IsSuccess -and $bLoc.Data.state -eq 'running') 'job-b should be running'
    $jobB2 = $bLoc.Data.job
    $jobB2.startedAtUtc = ([DateTime]::UtcNow.AddSeconds(-10).ToString('o'))
    $runningPath = Join-Path (Join-Path $root 'running') 'job-b.json'
    Set-Content -LiteralPath $runningPath -Value ($jobB2 | ConvertTo-Json -Depth 50) -Encoding utf8

    $rec1 = Recover-QCStaleJobs -Config $config
    Assert-True $rec1.IsSuccess 'Recover-QCStaleJobs should succeed'
    $stats4 = Get-QCQueueStats -Config $config
    Assert-Eq $stats4.Data.states.pending 1 'Recovery should requeue stale job-b back to pending (attempt 1)'

    # Lock again, force stale again, expect failed because maxAttempts=2
    $lockB2 = Lock-QCJob -JobId 'job-b' -Config $config
    Assert-True $lockB2.IsSuccess 'Lock job-b again should succeed'
    $bLoc2 = Get-QCJobById -JobId 'job-b' -Config $config
    $jobB3 = $bLoc2.Data.job
    $jobB3.startedAtUtc = ([DateTime]::UtcNow.AddSeconds(-10).ToString('o'))
    Set-Content -LiteralPath $runningPath -Value ($jobB3 | ConvertTo-Json -Depth 50) -Encoding utf8

    $rec2 = Recover-QCStaleJobs -Config $config
    Assert-True $rec2.IsSuccess 'Second recovery should succeed'
    $stats5 = Get-QCQueueStats -Config $config
    Assert-Eq $stats5.Data.states.failed 1 'Recovery should move job-b to failed on max attempts'

    # recent jobs
    $recent = Get-QCRecentJobs -Config $config -Limit 10
    Assert-True ($recent.IsSuccess -and $recent.Data.jobs.Count -ge 2) 'Get-QCRecentJobs should return jobs across states'

    Write-Host 'All QC.Queue.Json tests passed.' -ForegroundColor Green
} finally {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}

