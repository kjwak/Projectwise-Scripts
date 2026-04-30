# Integration tests for processor dispatch path (Run-QCProcessor.ps1 + QC.Processors dispatcher).
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

function New-TestJob([string]$Id, [string]$Type, [string]$DedupeKey) {
    return @{
        id = $Id
        type = $Type
        sourcePath = "c:\\work\\$Id.pdf"
        sourceName = "$Id.pdf"
        sourceFolder = "c:\\work"
        groupKey = ''
        triggerRule = @{ id = 'r1'; jobType = $Type; grouping = @{ enabled = $false; groupBy = 'file' } }
        dedupeKey = $DedupeKey
        status = 'queued'
        createdAt = '2026-01-01T00:00:00.0000000Z'
        attempts = 0
        metadata = @{}
    }
}

function New-TempEnv() {
    $tempRoot = Join-Path $env:TEMP ("qc-proc-test-" + ([guid]::NewGuid().ToString('N')))
    $queueDir = Join-Path $tempRoot "queue"
    New-Item -ItemType Directory -Path $queueDir -Force | Out-Null
    return @{ tempRoot = $tempRoot; queueDir = $queueDir }
}

function Write-AppSettings([string]$Path, [string]$QueueDir, [bool]$DryRun, [hashtable]$ProcessorMap) {
    $cfg = @{
        dryRun = $DryRun
        queue = @{ rootDir = $QueueDir; recover = @{ maxAttempts = 3 } }
        processors = @{
            processorMap = $ProcessorMap
            dryRun = @{ allowStateChange = $false; invokeHandler = $false }
        }
        qcPrepend = @{
            historyRoot = (Join-Path (Split-Path -Parent $Path) 'history')
            outputRoot = (Join-Path (Split-Path -Parent $Path) 'output')
            tempRoot = (Join-Path (Split-Path -Parent $Path) 'temp')
            enableOverlay = $false
            qpdfExePath = (Join-Path $repoRoot 'tools\\qpdf\\bin\\qpdf.exe')
        }
        # Opt-in to the no-op stub for STATUS_SET_GEN: this test only verifies that
        # the dispatcher routes the job to a handler and the queue moves to succeeded.
        # The native handler requires a live ProjectWise connection, which unit tests
        # don't have. Production config uses statusSet.mode = 'native'.
        statusSet = @{ mode = 'stub' }
    }
    ($cfg | ConvertTo-Json -Depth 50) | Set-Content -LiteralPath $Path -Encoding utf8
}

# QC_PREPEND dispatch succeeds
$e1 = New-TempEnv
try {
    $cfgPath = Join-Path $e1.tempRoot "appsettings.json"
    Write-AppSettings -Path $cfgPath -QueueDir $e1.queueDir -DryRun:$false -ProcessorMap @{ QC_PREPEND = 'Invoke-QCPrependProcessor'; STATUS_SET_GEN = 'Invoke-StatusSetProcessor' }
    $cfg = @{ queue = @{ rootDir = $e1.queueDir } }

    # Create minimal valid PDF (qpdf can validate/consume).
    $srcDir = Join-Path $e1.tempRoot 'src'
    New-Item -ItemType Directory -Path $srcDir -Force | Out-Null
    $sourcePdf = Join-Path $srcDir 'job-prepend.pdf'
    Set-Content -LiteralPath $sourcePdf -Value "%PDF-1.4`n1 0 obj<<>>endobj`ntrailer<<>>`n%%EOF`n" -Encoding ascii

    $job = New-TestJob -Id 'job-prepend' -Type 'QC_PREPEND' -DedupeKey 'dq_prepend'
    $job.sourcePath = $sourcePdf
    $job.sourceName = 'job-prepend.pdf'
    $job.sourceFolder = $srcDir
    $enq = Add-QCQueueJob -Job $job -Config $cfg
    Assert-True $enq.IsSuccess 'Enqueue should succeed'

    & powershell -NoProfile -ExecutionPolicy Bypass -File "$repoRoot\\Run-QCProcessor.ps1" -AppSettingsPath $cfgPath | Out-Null
    $stats = Get-QCQueueStats -Config $cfg
    Assert-Eq $stats.Data.states.succeeded 1 'QC_PREPEND should end in succeeded'
} finally {
    Remove-Item -LiteralPath $e1.tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# STATUS_SET_GEN dispatch succeeds
$e2 = New-TempEnv
try {
    $cfgPath = Join-Path $e2.tempRoot "appsettings.json"
    Write-AppSettings -Path $cfgPath -QueueDir $e2.queueDir -DryRun:$false -ProcessorMap @{ QC_PREPEND = 'Invoke-QCPrependProcessor'; STATUS_SET_GEN = 'Invoke-StatusSetProcessor' }
    $cfg = @{ queue = @{ rootDir = $e2.queueDir } }
    $job = New-TestJob -Id 'job-status' -Type 'STATUS_SET_GEN' -DedupeKey 'dq_status'
    $enq = Add-QCQueueJob -Job $job -Config $cfg
    Assert-True $enq.IsSuccess 'Enqueue should succeed'

    & powershell -NoProfile -ExecutionPolicy Bypass -File "$repoRoot\\Run-QCProcessor.ps1" -AppSettingsPath $cfgPath | Out-Null
    $stats = Get-QCQueueStats -Config $cfg
    Assert-Eq $stats.Data.states.succeeded 1 'STATUS_SET_GEN should end in succeeded'
} finally {
    Remove-Item -LiteralPath $e2.tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# Unknown jobType fails cleanly and is requeued (pending) on first failure
$e3 = New-TempEnv
try {
    $cfgPath = Join-Path $e3.tempRoot "appsettings.json"
    Write-AppSettings -Path $cfgPath -QueueDir $e3.queueDir -DryRun:$false -ProcessorMap @{}
    $cfg = @{ queue = @{ rootDir = $e3.queueDir; recover = @{ maxAttempts = 3 } } }
    $job = New-TestJob -Id 'job-unknown' -Type 'UNKNOWN_TYPE' -DedupeKey 'dq_unknown'
    $enq = Add-QCQueueJob -Job $job -Config $cfg
    Assert-True $enq.IsSuccess 'Enqueue should succeed'

    $ec = 0
    try { & powershell -NoProfile -ExecutionPolicy Bypass -File "$repoRoot\\Run-QCProcessor.ps1" -AppSettingsPath $cfgPath | Out-Null } catch { $ec = 1 }
    $stats = Get-QCQueueStats -Config $cfg
    Assert-Eq $stats.Data.states.pending 1 'Unknown job should be requeued to pending on first failure'
} finally {
    Remove-Item -LiteralPath $e3.tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# DryRun does not change queue state by default (net pending count unchanged)
$e4 = New-TempEnv
try {
    $cfgPath = Join-Path $e4.tempRoot "appsettings.json"
    Write-AppSettings -Path $cfgPath -QueueDir $e4.queueDir -DryRun:$true -ProcessorMap @{ QC_PREPEND = 'Invoke-QCPrependProcessor' }
    $cfg = @{ queue = @{ rootDir = $e4.queueDir } }
    $job = New-TestJob -Id 'job-dryrun' -Type 'QC_PREPEND' -DedupeKey 'dq_dryrun'
    Add-QCQueueJob -Job $job -Config $cfg | Out-Null

    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$repoRoot\\Run-QCProcessor.ps1" -AppSettingsPath $cfgPath -DryRun 2>&1
    $stats = Get-QCQueueStats -Config $cfg
    Assert-Eq $stats.Data.states.pending 1 'DryRun should leave job pending by default'
    Assert-Eq $stats.Data.states.running 0 'DryRun should not leave job running'
    Assert-Eq $stats.Data.states.succeeded 0 'DryRun should not succeed job by default'
    Assert-Eq $stats.Data.locks.count 0 'DryRun should not lock any jobs by default'

    # invokeHandler=false => should not log WORKER_DRYRUN_HANDLER
    $hasHandlerLog = $false
    foreach ($l in @($out | ForEach-Object { $_ -as [string] } | Where-Object { $_ })) {
        if ($l -match '\"code\"\s*:\s*\"WORKER_DRYRUN_HANDLER\"') { $hasHandlerLog = $true; break }
    }
    Assert-True (-not $hasHandlerLog) 'invokeHandler=false should not invoke handler'
} finally {
    Remove-Item -LiteralPath $e4.tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# DryRun invokeHandler=true calls handler but does not mutate queue
$e5 = New-TempEnv
try {
    $cfgPath = Join-Path $e5.tempRoot "appsettings.json"
    Write-AppSettings -Path $cfgPath -QueueDir $e5.queueDir -DryRun:$true -ProcessorMap @{ QC_PREPEND = 'Invoke-QCPrependProcessor' }
    # flip invokeHandler true
    $raw = Get-Content -LiteralPath $cfgPath -Raw | ConvertFrom-Json
    $raw.processors.dryRun.invokeHandler = $true
    ($raw | ConvertTo-Json -Depth 50) | Set-Content -LiteralPath $cfgPath -Encoding utf8

    $cfg = @{ queue = @{ rootDir = $e5.queueDir } }
    $job = New-TestJob -Id 'job-dryrun-handler' -Type 'QC_PREPEND' -DedupeKey 'dq_dryrun_handler'
    Add-QCQueueJob -Job $job -Config $cfg | Out-Null

    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File "$repoRoot\\Run-QCProcessor.ps1" -AppSettingsPath $cfgPath -DryRun 2>&1
    $stats = Get-QCQueueStats -Config $cfg
    Assert-Eq $stats.Data.states.pending 1 'DryRun invokeHandler should leave job pending'
    Assert-Eq $stats.Data.locks.count 0 'DryRun invokeHandler should not lock any jobs'

    $hasHandlerLog = $false
    foreach ($l in @($out | ForEach-Object { $_ -as [string] } | Where-Object { $_ })) {
        if ($l -match '\"code\"\s*:\s*\"WORKER_DRYRUN_HANDLER\"') { $hasHandlerLog = $true; break }
    }
    Assert-True $hasHandlerLog 'invokeHandler=true should log WORKER_DRYRUN_HANDLER'
} finally {
    Remove-Item -LiteralPath $e5.tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host 'All processor dispatch tests passed.' -ForegroundColor Green

