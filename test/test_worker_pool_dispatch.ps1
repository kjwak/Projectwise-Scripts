# Integration test: parallel workers drain the queue without duplicate processing.
# Spawns 3 workers in parallel against a queue of 6 STATUS_SET_GEN stub jobs and
# verifies all reach 'succeeded' exactly once with no leftover lock files.
$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module "$repoRoot/modules/Core.Results.psm1" -Force
Import-Module "$repoRoot/modules/QC.Queue.Json.psm1" -Force

$tempRoot = Join-Path $env:TEMP ("qc-pool-test-" + ([guid]::NewGuid().ToString('N')))
$queueDir = Join-Path $tempRoot 'queue'
New-Item -ItemType Directory -Path $queueDir -Force | Out-Null
$cfgPath = Join-Path $tempRoot 'appsettings.json'

try {
    # Stub-mode STATUS_SET_GEN: just needs Job.sourceFolder; succeeds without PW or qpdf.
    $cfg = @{
        dryRun = $false
        queue = @{ rootDir = $queueDir; recover = @{ maxAttempts = 3 } }
        processors = @{
            processorMap = @{ STATUS_SET_GEN = 'Invoke-StatusSetProcessor' }
            dryRun = @{ allowStateChange = $false; invokeHandler = $false }
        }
        statusSet = @{ mode = 'stub' }
    }
    ($cfg | ConvertTo-Json -Depth 50) | Set-Content -LiteralPath $cfgPath -Encoding utf8

    $cfgQueue = @{ queue = @{ rootDir = $queueDir } }

    $jobIds = @()
    for ($i = 1; $i -le 6; $i++) {
        $id = "pool-job-$i"
        $jobIds += $id
        $job = @{
            id = $id
            type = 'STATUS_SET_GEN'
            sourcePath = "c:\\work\\$id"
            sourceFolder = "c:\\work\\folder-$i"
            sourceName = $id
            dedupeKey = "dq_$id"
            status = 'pending'
            attempts = 0
            metadata = @{}
        }
        $r = Add-QCQueueJob -Job $job -Config $cfgQueue
        Assert-True $r.IsSuccess "enqueue $id"
    }

    $statsBefore = Get-QCQueueStats -Config $cfgQueue
    Assert-Eq $statsBefore.Data.states.pending 6 'all 6 enqueued'

    function _Quote([string]$s) {
        if ($s -match '[\s"]') { return ('"' + ($s -replace '"','\"') + '"') }
        return $s
    }

    $scriptPath = Join-Path $repoRoot 'scripts\Run-QCProcessor.ps1'
    $procs = @()
    for ($w = 1; $w -le 3; $w++) {
        $pArgs = @(
            '-NoProfile','-ExecutionPolicy','Bypass',
            '-File',(_Quote $scriptPath),
            '-AppSettingsPath',(_Quote $cfgPath),
            '-MaxJobs','10',
            '-LeaseSeconds','30',
            '-IdleSleepMs','200',
            '-WorkerLabel',"W$w"
        )
        $procs += Start-Process -FilePath 'powershell.exe' -ArgumentList $pArgs -NoNewWindow -PassThru
    }

    $deadline = (Get-Date).AddSeconds(60)
    foreach ($p in $procs) {
        $remainingMs = [Math]::Max(1000, [int]($deadline - (Get-Date)).TotalMilliseconds)
        if (-not $p.WaitForExit($remainingMs)) {
            try { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } catch { }
            throw "Worker pid $($p.Id) did not exit within deadline"
        }
    }

    $statsAfter = Get-QCQueueStats -Config $cfgQueue
    Assert-Eq $statsAfter.Data.states.pending   0 'no pending after pool drain'
    Assert-Eq $statsAfter.Data.states.running   0 'no running after pool drain'
    Assert-Eq $statsAfter.Data.states.failed    0 'no failed after pool drain'
    Assert-Eq $statsAfter.Data.states.succeeded 6 'all 6 succeeded'
    Assert-Eq $statsAfter.Data.locks.count      0 'no leftover per-job locks'

    # Each job processed exactly once (attempts == 0 since stub never failed).
    foreach ($id in $jobIds) {
        $row = Get-QCJobById -JobId $id -Config $cfgQueue
        Assert-True $row.IsSuccess "load $id"
        Assert-True ($row.Data.found) "$id found in queue"
        $job = $row.Data.job
        Assert-Eq ([string]$job.status) 'succeeded' "$id status == succeeded"
        $att = 0
        if ($job.ContainsKey('attempts') -and $null -ne $job.attempts) { $att = [int]$job.attempts }
        Assert-Eq $att 0 "$id processed once (attempts == 0)"
    }
} finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host 'test_worker_pool_dispatch: passed' -ForegroundColor Green
