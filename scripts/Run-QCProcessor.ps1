<#
.SYNOPSIS
QC worker: lock -> dispatch -> transition. One-shot by default; long-running with -MaxJobs/-LeaseSeconds/-IdleSleepMs.

.DESCRIPTION
Loads appsettings.json, retrieves next queued job from JSON queue, locks it,
dispatches by job type using QC.Processors, then transitions queue state.

When -MaxJobs > 1 or -LeaseSeconds > 0 or -IdleSleepMs > 0 the script loops:
  - Picks pending jobs and processes them race-safely against other parallel workers.
  - On Lock-QCJob race loss (QUEUE_LOCK_TIMEOUT, QUEUE_JOB_ALREADY_MOVED, QUEUE_LOCK_ERROR, etc.)
    the loser excludes that id from subsequent selections in this process and tries the next job.
  - Exits when budget is exhausted (jobs processed >= MaxJobs, or elapsed >= LeaseSeconds).
  - When the queue is empty: exits if -IdleSleepMs <= 0; otherwise sleeps and re-polls.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$AppSettingsPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'appsettings.json'),

    [Parameter(Mandatory = $false)]
    [switch]$DryRun,

    [Parameter(Mandatory = $false)]
    [int]$MaxJobs = 1,

    [Parameter(Mandatory = $false)]
    [int]$LeaseSeconds = 0,

    [Parameter(Mandatory = $false)]
    [int]$IdleSleepMs = 0,

    [Parameter(Mandatory = $false)]
    [string]$WorkerLabel = ''
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.Queue.Json.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.Processors.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.Worker.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Database.psm1') -Force

$script:WorkerLabel = $WorkerLabel

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config

if (-not $config.ContainsKey('dryRun')) { $config['dryRun'] = $false }
if ($DryRun.IsPresent) { $config['dryRun'] = $true }
$isDryRun = [bool]$config['dryRun']

$dryRunAllowStateChange = $false
if ($config.ContainsKey('processors') -and $config.processors -and $config.processors.ContainsKey('dryRun') -and $config.processors.dryRun) {
    try { $dryRunAllowStateChange = [bool]$config.processors.dryRun.allowStateChange } catch { $dryRunAllowStateChange = $false }
}
$dryRunInvokeHandler = $false
if ($config.ContainsKey('processors') -and $config.processors -and $config.processors.ContainsKey('dryRun') -and $config.processors.dryRun) {
    try { $dryRunInvokeHandler = [bool]$config.processors.dryRun.invokeHandler } catch { $dryRunInvokeHandler = $false }
}

$maxAttempts = 3
if ($config.ContainsKey('queue') -and $config.queue -and $config.queue.ContainsKey('recover') -and $config.queue.recover -and $config.queue.recover.ContainsKey('maxAttempts')) {
    $maxAttempts = [int]$config.queue.recover.maxAttempts
}

$loopMode = ($MaxJobs -gt 1) -or ($LeaseSeconds -gt 0) -or ($IdleSleepMs -gt 0)

if (Test-QCDatabaseEnabled -Config $config) {
    try {
        $schemaRes = Initialize-QCDatabaseSchema -Config $config
        if (-not $schemaRes.IsSuccess) {
            Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Flush -Level 'Warning' -Code 'WORKER_DB_SCHEMA_INIT_FAILED' -Message 'Database schema initialization failed; job telemetry may not persist.' -Data @{
                code = [string]$schemaRes.Code
                message = [string]$schemaRes.Message
            }
        }
    } catch {
        Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Flush -Level 'Warning' -Code 'WORKER_DB_SCHEMA_INIT_FAILED' -Message ('Database schema initialization threw: ' + $_.Exception.Message) -Data @{}
    }
}

Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Level 'Information' -Code 'WORKER_START' -Message 'Worker run started.' -Data @{
    appSettingsPath = $AppSettingsPath
    dryRun = $isDryRun
    dryRunAllowStateChange = $dryRunAllowStateChange
    maxJobs = $MaxJobs
    leaseSeconds = $LeaseSeconds
    idleSleepMs = $IdleSleepMs
    loopMode = $loopMode
}

function _Resolve-Handler([hashtable]$Job, [hashtable]$Config) {
    $jobType = [string]$Job['type']
    $map = $null
    if ($Config.ContainsKey('processors') -and $Config.processors -and $Config.processors.ContainsKey('processorMap')) { $map = $Config.processors.processorMap }
    elseif ($Config.ContainsKey('processorMap')) { $map = $Config.processorMap }
    if ($map -is [hashtable] -and $map.ContainsKey($jobType)) { return [string]$map[$jobType] }
    if ($jobType -eq 'QC_PREPEND') { return 'Invoke-QCPrependProcessor' }
    if ($jobType -eq 'STATUS_SET_GEN') { return 'Invoke-StatusSetProcessor' }
    return ''
}

function Write-WorkerStage {
    param(
        [string]$Stage,
        [string]$JobId = '',
        [string]$JobType = '',
        [string]$Handler = ''
    )
    Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Level 'Information' -Code 'WORKER_STAGE' -Message $Stage -Data @{
        stage = $Stage
        jobId = $JobId
        jobType = $JobType
        handler = $Handler
    }
}

function Write-WorkerJobTelemetryLog {
    param(
        [string]$JobId,
        [string]$JobType,
        [object]$TelemetryResult
    )
    if (-not $TelemetryResult) { return }
    $code = [string]$TelemetryResult.Code
    $written = $false
    $rowsAffected = $null
    try {
        if ($TelemetryResult.Data) {
            if ($TelemetryResult.Data -is [hashtable] -and $TelemetryResult.Data.ContainsKey('written')) { $written = [bool]$TelemetryResult.Data.written }
            if ($TelemetryResult.Data -is [hashtable] -and $TelemetryResult.Data.ContainsKey('rowsAffected')) { $rowsAffected = $TelemetryResult.Data.rowsAffected }
        }
    } catch { }
    if ($code -eq 'JOB_TELEMETRY_WRITTEN') {
        Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Level 'Information' -Code 'JOB_TELEMETRY_WRITTEN' -Message 'Job outcome recorded in processing_jobs.' -Data @{
            jobId = $JobId; jobType = $JobType; written = $written; rowsAffected = $rowsAffected
        }
        return
    }
    if ($code -eq 'JOB_TELEMETRY_SKIPPED') {
        $reason = $null
        try { if ($TelemetryResult.Data.reason) { $reason = [string]$TelemetryResult.Data.reason } } catch { }
        Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Level 'Warning' -Code 'JOB_TELEMETRY_SKIPPED' -Message $TelemetryResult.Message -Data @{
            jobId = $JobId; jobType = $JobType; reason = $reason
        }
        return
    }
    Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Flush -Level 'Error' -Code $code -Message $TelemetryResult.Message -Data @{
        jobId = $JobId; jobType = $JobType; written = $written
    }
}

function Move-QCJobWithLockRetries {
    param(
        [string]$JobId,
        [string]$FromState,
        [string]$ToState,
        [hashtable]$Config,
        [hashtable]$Job,
        [int]$MaxTries = 8,
        [int]$SleepMs = 3000
    )
    $last = $null
    for ($t = 1; $t -le $MaxTries; $t++) {
        $last = Move-QCJob -JobId $JobId -FromState $FromState -ToState $ToState -Config $Config -Job $Job
        if ($last.IsSuccess) { return $last }
        if ([string]$last.Code -ne 'QUEUE_LOCK_TIMEOUT') { return $last }
        if ($t -ge $MaxTries) { return $last }
        Start-Sleep -Milliseconds $SleepMs
    }
    return $last
}

function _Process-OneJob([hashtable]$Job, [string]$Handler, [hashtable]$Config, [bool]$IsDryRun, [bool]$DryRunAllowStateChange, [bool]$DryRunInvokeHandler, [int]$MaxAttempts) {
    # Returns hashtable: @{ Outcome = 'succeeded'|'failed'|'requeued'|'skipped_locked'|'dryrun_noop'; ExitOk = [bool]; SkipId = [string] }
    $jobId = [string]$Job['id']
    $jobType = [string]$Job['type']

    if ($IsDryRun -and -not $DryRunAllowStateChange) {
        Write-WorkerStage -Stage 'dry-run: evaluating read-only dispatch' -JobId $jobId -JobType $jobType -Handler $Handler
        Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Level 'Information' -Code 'WORKER_DRYRUN' -Message 'Dry-run: would dispatch job (read-only, no lock/state changes).' -Data @{
            jobId = $jobId; jobType = $jobType; handler = $Handler
            wouldRun = $true; wouldLock = $false; wouldChangeState = $false
        }
        if ($DryRunInvokeHandler -and $Handler) {
            try {
                $cmd = Get-Command -Name $Handler -ErrorAction SilentlyContinue
                if (-not $cmd) {
                    Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Level 'Error' -Code 'WORKER_DRYRUN_HANDLER' -Message 'Dry-run: handler not found.' -Data @{ jobId = $jobId; handler = $Handler; isSuccess = $false }
                    return @{ Outcome = 'dryrun_noop'; ExitOk = $true; SkipId = $jobId }
                }
                $Config['dryRun'] = $true
                $hr = & $Handler -Job $Job -Config $Config
                if ($null -eq $hr -or -not ($hr.PSObject.Properties.Name -contains 'IsSuccess')) {
                    Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Level 'Error' -Code 'WORKER_DRYRUN_HANDLER' -Message 'Dry-run: handler returned invalid result.' -Data @{ jobId = $jobId; handler = $Handler; isSuccess = $false }
                    return @{ Outcome = 'dryrun_noop'; ExitOk = $true; SkipId = $jobId }
                }
                Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Level 'Information' -Code 'WORKER_DRYRUN_HANDLER' -Message 'Dry-run: handler invoked.' -Data @{
                    jobId = $jobId; handler = $Handler
                    isSuccess = [bool]$hr.IsSuccess
                    resultCode = [string]$hr.Code
                    resultMessage = [string]$hr.Message
                    resultData = $hr.Data
                }
            } catch {
                Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Level 'Error' -Code 'WORKER_DRYRUN_HANDLER' -Message 'Dry-run: handler threw.' -Data @{ jobId = $jobId; handler = $Handler; isSuccess = $false; errorMessage = $_.Exception.Message }
            }
        }
        return @{ Outcome = 'dryrun_noop'; ExitOk = $true; SkipId = $jobId }
    }

    Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Level 'Information' -Code 'WORKER_CLAIMING' -Message 'Attempting to lock job (not running yet).' -Data @{
        jobId = $jobId
        jobType = $jobType
        handler = $Handler
        sourceFolder = if ($Job.ContainsKey('sourceFolder')) { [string]$Job['sourceFolder'] } else { '' }
    }
    Write-WorkerStage -Stage 'locking queue job' -JobId $jobId -JobType $jobType -Handler $Handler
    $lock = Lock-QCJob -JobId $jobId -Config $Config
    if (-not $lock.IsSuccess) {
        $benignLockCodes = @(
            'QUEUE_LOCK_EXISTS', 'QUEUE_LOCK_TIMEOUT', 'QUEUE_JOB_ALREADY_MOVED', 'QUEUE_JOB_NOT_FOUND', 'QUEUE_LOCK_ERROR'
        )
        if ([string]$lock.Code -in $benignLockCodes) {
            $lockDetail = $null
            try {
                if ($lock.Data) {
                    if ($lock.Data -is [hashtable] -and $lock.Data.ContainsKey('innerMessage')) { $lockDetail = [string]$lock.Data.innerMessage }
                    elseif ($lock.Data.PSObject.Properties['innerMessage']) { $lockDetail = [string]$lock.Data.innerMessage }
                    elseif ($lock.Data -is [hashtable] -and $lock.Data.ContainsKey('error')) { $lockDetail = [string]$lock.Data.error }
                }
            } catch { }
            Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Level 'Information' -Code 'WORKER_LOCK_RACE' -Message 'Lost race for job lock; skipping.' -Data @{
                jobId = $jobId
                lockCode = [string]$lock.Code
                lockMessage = [string]$lock.Message
                lockDetail = $lockDetail
            }
            return @{ Outcome = 'skipped_locked'; ExitOk = $true; SkipId = $jobId }
        }
        Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Level 'Error' -Code 'WORKER_LOCK_FATAL' -Message 'Unexpected lock failure; skipping job without crashing worker.' -Data @{
            jobId = $jobId
            lockCode = [string]$lock.Code
            lockMessage = [string]$lock.Message
        }
        return @{ Outcome = 'skipped_locked'; ExitOk = $true; SkipId = $jobId }
    }

    try {
        Write-WorkerStage -Stage 'loading locked job' -JobId $jobId -JobType $jobType -Handler $Handler
        $loaded = Get-QCJobById -JobId $jobId -Config $Config
        if (-not $loaded.IsSuccess -or -not $loaded.Data.found) { throw "Failed to load locked job: $jobId" }
        # After a long wait to acquire the job lock, the job may have already
        # been completed by another worker. Only process jobs that are truly in
        # running\ after the lock/transition step.
        if ([string]$loaded.Data.state -ne 'running') {
            Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Level 'Information' -Code 'WORKER_LOCK_RACE' -Message 'Job was already moved to a terminal state; skipping.' -Data @{
                jobId = $jobId
                state = [string]$loaded.Data.state
            }
            return @{ Outcome = 'skipped_locked'; ExitOk = $true; SkipId = $jobId }
        }
        $Job = [hashtable]$loaded.Data.job

        Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Level 'Information' -Code 'WORKER_SELECTED' -Message 'Job locked and running (exclusive owner).' -Data @{
            jobId = $jobId
            jobType = $jobType
            dedupeKey = if ($Job.ContainsKey('dedupeKey')) { [string]$Job['dedupeKey'] } else { '' }
            handler = $Handler
            sourceFolder = if ($Job.ContainsKey('sourceFolder')) { [string]$Job['sourceFolder'] } else { '' }
            sourcePath = if ($Job.ContainsKey('sourcePath')) { [string]$Job['sourcePath'] } else { '' }
            sourceName = if ($Job.ContainsKey('sourceName')) { [string]$Job['sourceName'] } else { '' }
        }

        if ($IsDryRun) {
            Write-WorkerStage -Stage 'dry-run: locked job dispatch check' -JobId $jobId -JobType $jobType -Handler $Handler
            Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Level 'Information' -Code 'WORKER_DRYRUN' -Message 'Dry-run: would dispatch job.' -Data @{
                jobId = $jobId; jobType = [string]$Job['type']; handler = $Handler
                wouldRun = $true; wouldLock = $true; wouldChangeState = $DryRunAllowStateChange
            }
            return @{ Outcome = 'dryrun_noop'; ExitOk = $true; SkipId = $jobId }
        }

        Write-WorkerStage -Stage ("running processor: $Handler") -JobId $jobId -JobType $jobType -Handler $Handler
        $jobSw = [System.Diagnostics.Stopwatch]::StartNew()
        $proc = Invoke-QCProcessorByType -Job $Job -Config $Config
        $jobSw.Stop()
        $jobDurationMs = [int]$jobSw.ElapsedMilliseconds
        if ($proc.IsSuccess) {
            # Capture rich result data so the per-job JSON keeps everything the
            # processor reported (pwUpload, needsFullRebuild, changedCount, etc).
            $resultData = $null
            try {
                if ($proc.Data) {
                    if ($proc.Data -is [hashtable]) { $resultData = $proc.Data }
                    elseif ($proc.Data.PSObject) {
                        $tmp = @{}
                        foreach ($p in $proc.Data.PSObject.Properties) { $tmp[$p.Name] = $p.Value }
                        $resultData = $tmp
                    }
                }
            } catch { $resultData = $null }
            $Job['result'] = @{
                code           = [string]$proc.Code
                message        = [string]$proc.Message
                completedAtUtc = Get-QCTimestamp
                data           = $resultData
            }
            # Single Move-QCJob call writes the in-memory $Job (with result/data)
            # to disk and renames to succeeded\ under one queue-write-lock cycle.
            # The previous Update-QCJob + Move-QCJob sequence acquired the global
            # write lock twice in a row, doubling contention and producing
            # spurious WORKER_MOVE_FAILED / QUEUE_LOCK_TIMEOUT under load.
            Write-WorkerStage -Stage 'moving completed job to succeeded' -JobId $jobId -JobType $jobType -Handler $Handler
            $mv = Move-QCJobWithLockRetries -JobId $jobId -FromState 'running' -ToState 'succeeded' -Config $Config -Job $Job
            if (-not $mv.IsSuccess) {
                $innerErr = ''
                try { if ($mv.Data -and $mv.Data.error) { $innerErr = [string]$mv.Data.error } } catch { }
                $logLevel = 'Error'
                if ([string]$mv.Code -eq 'QUEUE_JOB_WRONG_STATE') { $logLevel = 'Warning' }
                Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Level $logLevel -Code 'WORKER_MOVE_FAILED' -Message 'Failed to move succeeded job to succeeded\.' -Data @{
                    jobId = $jobId; fromState = 'running'; toState = 'succeeded'
                    errorCode = [string]$mv.Code; errorMessage = [string]$mv.Message; errorData = $innerErr
                    actualState = if ($mv.Data -and $mv.Data.actualState) { [string]$mv.Data.actualState } else { $null }
                }
                # The job is still in running\. Treat as a transient failure so the
                # recovery sweep (or next worker tick) can retry the move; do NOT
                # emit a fake WORKER_SUCCEEDED.
                return @{ Outcome = 'failed'; ExitOk = $false; SkipId = $jobId }
            }
            if ($mv.Data -and @($mv.Data.duplicatesRemoveFailed).Count -gt 0) {
                Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Level 'Warning' -Code 'WORKER_QUEUE_DUPLICATE_CLEANUP' -Message 'Job succeeded but stale queue JSON copies could not be deleted (AV lock?).' -Data @{
                    jobId = $jobId; duplicatesRemoved = @($mv.Data.duplicatesRemoved); duplicatesRemoveFailed = @($mv.Data.duplicatesRemoveFailed)
                }
            }

            # Surface the most useful processor data fields directly in the
            # success log so the dashboard's recent-events panel shows e.g.
            # whether PW was actually updated.
            $logData = @{
                jobId       = $jobId
                jobType     = [string]$Job['type']
                resultCode  = [string]$proc.Code
            }
            if ($resultData -is [hashtable]) {
                foreach ($k in 'pwUpload','writeBackToPW','needsFullRebuild','changedCount','outPdf','docSearchPath') {
                    if ($resultData.ContainsKey($k)) { $logData[$k] = $resultData[$k] }
                }
            }
            Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Level 'Information' -Code 'WORKER_SUCCEEDED' -Message 'Job succeeded.' -Data $logData
            $rdJson = $null; try { if ($resultData) { $rdJson = ($resultData | ConvertTo-Json -Depth 4 -Compress) } } catch {}
            $triggerSource = $null
            try {
                if ($Job.metadata -and $Job.metadata.candidate -and $Job.metadata.candidate.triggerSource) {
                    $triggerSource = [string]$Job.metadata.candidate.triggerSource
                }
            } catch { }
            $startedAtUtc = $null
            try {
                if ($Job.ContainsKey('startedAtUtc') -and $Job['startedAtUtc']) { $startedAtUtc = [string]$Job['startedAtUtc'] }
            } catch { }
            $jobAttempts = 0
            try { if ($Job.ContainsKey('attempts') -and $null -ne $Job['attempts']) { $jobAttempts = [int]$Job['attempts'] } } catch { }
            $telRes = Write-QCJobTelemetry -Config $Config -JobId $jobId -JobType $jobType -Status 'succeeded' `
                -SourcePath ([string]$Job['sourcePath']) -SourceFolder ([string]$Job['sourceFolder']) `
                -DedupeKey ([string]$Job['dedupeKey']) -TriggerSource $triggerSource -StartedAtUtc $startedAtUtc `
                -AttemptCount $jobAttempts -DurationMs $jobDurationMs -ResultData $rdJson
            Write-WorkerJobTelemetryLog -JobId $jobId -JobType $jobType -TelemetryResult $telRes

            # Link -qc.pdf to source sheet in sheet_index (fire-and-forget)
            if ($jobType -eq 'QC_PREPEND' -and $resultData -is [hashtable] -and $resultData.ContainsKey('qcOutputPdf')) {
                try {
                    $srcName = [string]$Job['sourceName']
                    $srcFolder = [string]$Job['sourceFolder']
                    $qcPdfName = [System.IO.Path]::GetFileName([string]$resultData.qcOutputPdf)
                    if ($srcName -and $srcFolder -and $qcPdfName) {
                        Invoke-QCDatabaseNonQuery -Config $Config -Sql @"
UPDATE sheet_index SET qc_pdf_name = @qcPdfName, last_updated_at = SYSDATETIMEOFFSET()
WHERE document_name = @srcName AND folder_path = @srcFolder
"@ -Parameters @{ qcPdfName = $qcPdfName; srcName = $srcName; srcFolder = $srcFolder } | Out-Null
                    }
                } catch { }
            }

            return @{ Outcome = 'succeeded'; ExitOk = $true; SkipId = $jobId }
        }

        $attempts = 0
        if ($Job.ContainsKey('attempts') -and $Job.attempts -ne $null) { $attempts = [int]$Job.attempts }
        $attempts++
        $Job['attempts'] = $attempts
        $Job['lastError'] = @{ code = [string]$proc.Code; message = [string]$proc.Message; atUtc = Get-QCTimestamp }

        $target = if ($attempts -ge $MaxAttempts) { 'failed' } else { 'pending' }
        # Single Move-QCJob call (same rationale as the success path - one lock
        # cycle, in-memory job preserved with attempts/lastError stamped on disk).
        Write-WorkerStage -Stage ("moving failed job to $target") -JobId $jobId -JobType $jobType -Handler $Handler
        $fmv = Move-QCJobWithLockRetries -JobId $jobId -FromState 'running' -ToState $target -Config $Config -Job $Job
        if (-not $fmv.IsSuccess) {
            $innerErr = ''
            try { if ($fmv.Data -and $fmv.Data.error) { $innerErr = [string]$fmv.Data.error } } catch { }
            Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Level 'Error' -Code 'WORKER_MOVE_FAILED' -Message ('Failed to move failed job to ' + $target + '\.') -Data @{
                jobId = $jobId; fromState = 'running'; toState = $target
                errorCode = [string]$fmv.Code; errorMessage = [string]$fmv.Message; errorData = $innerErr
            }
        }

        $details = $null
        $stdoutPreview = $null
        $stderrPreview = $null
        if ($proc.Data) {
            try {
                if ($proc.Data -is [hashtable]) {
                    if ($proc.Data.ContainsKey('stdout')) { $stdoutPreview = [string]$proc.Data.stdout }
                    if ($proc.Data.ContainsKey('stderr')) { $stderrPreview = [string]$proc.Data.stderr }
                }
                if ($stdoutPreview -and $stdoutPreview.Length -gt 4000) { $stdoutPreview = $stdoutPreview.Substring(0, 4000) }
                if ($stderrPreview -and $stderrPreview.Length -gt 4000) { $stderrPreview = $stderrPreview.Substring(0, 4000) }
                $details = ($proc.Data | ConvertTo-Json -Depth 6 -Compress)
                if ($details.Length -gt 2000) { $details = $details.Substring(0, 2000) }
            } catch { }
        }
        Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Level 'Error' -Code 'WORKER_FAILED' -Message 'Job failed.' -Data @{
            jobId = $jobId; jobType = [string]$Job['type']
            attempts = $attempts; maxAttempts = $MaxAttempts; movedTo = $target
            errorCode = [string]$proc.Code; errorMessage = [string]$proc.Message
            processorData = $details; processorStdout = $stdoutPreview; processorStderr = $stderrPreview
        }
        $triggerSourceFail = $null
        try {
            if ($Job.metadata -and $Job.metadata.candidate -and $Job.metadata.candidate.triggerSource) {
                $triggerSourceFail = [string]$Job.metadata.candidate.triggerSource
            }
        } catch { }
        $startedAtUtcFail = $null
        try {
            if ($Job.ContainsKey('startedAtUtc') -and $Job['startedAtUtc']) { $startedAtUtcFail = [string]$Job['startedAtUtc'] }
        } catch { }
        $telResFail = Write-QCJobTelemetry -Config $Config -JobId $jobId -JobType $jobType -Status $target `
            -SourcePath ([string]$Job['sourcePath']) -SourceFolder ([string]$Job['sourceFolder']) `
            -DedupeKey ([string]$Job['dedupeKey']) -TriggerSource $triggerSourceFail -StartedAtUtc $startedAtUtcFail `
            -DurationMs $jobDurationMs -AttemptCount $attempts `
            -ErrorCode ([string]$proc.Code) -ErrorMessage ([string]$proc.Message) -ResultData $details
        Write-WorkerJobTelemetryLog -JobId $jobId -JobType $jobType -TelemetryResult $telResFail
        if ($target -eq 'failed') {
            return @{ Outcome = 'failed'; ExitOk = $false; SkipId = $jobId }
        } else {
            return @{ Outcome = 'requeued'; ExitOk = $false; SkipId = $jobId }
        }
    } finally {
        Unlock-QCJob -JobId $jobId -Config $Config | Out-Null
    }
}

$skip = New-Object System.Collections.Generic.List[string]
$processed = 0
$lastOutcomeOk = $true
$startedAt = [DateTime]::UtcNow

while ($true) {
    if ($MaxJobs -gt 0 -and $processed -ge $MaxJobs) {
        Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Level 'Information' -Code 'WORKER_BUDGET' -Message 'Worker reached MaxJobs budget; exiting.' -Data @{ processed = $processed; maxJobs = $MaxJobs }
        break
    }
    if ($LeaseSeconds -gt 0 -and ([DateTime]::UtcNow - $startedAt).TotalSeconds -ge $LeaseSeconds) {
        Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Level 'Information' -Code 'WORKER_LEASE' -Message 'Worker reached LeaseSeconds budget; exiting.' -Data @{ processed = $processed; leaseSeconds = $LeaseSeconds }
        break
    }

    # No global watcher gate: each STATUS_SET_GEN job is enqueued only after the
    # watcher has finished analyzing that specific folder (PW listing + manifest
    # comparison), so picking it up immediately is safe. Holding it until the entire
    # pass exited starved workers during long sweeps. QC_PREPEND was always
    # immediately eligible.
    Write-WorkerStage -Stage 'polling queue for pending job'
    $next = Get-NextQCJob -Config $config -ExcludeJobIds @($skip)
    if (-not $next.IsSuccess) { throw $next.Message }
    if (-not $next.Data.job) {
        Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Level 'Information' -Code 'WORKER_NO_JOB' -Message 'No pending jobs.' -Data @{
            processed = $processed
        }
        if (-not $loopMode) { break }
        if ($IdleSleepMs -le 0) { break }
        Start-Sleep -Milliseconds $IdleSleepMs
        if ($skip.Count -gt 0) { $skip.Clear() }
        continue
    }

    $job = [hashtable]$next.Data.job
    $jobId = [string]$job['id']
    $handler = _Resolve-Handler -Job $job -Config $config

    try {
        $res = _Process-OneJob -Job $job -Handler $handler -Config $config -IsDryRun:$isDryRun -DryRunAllowStateChange:$dryRunAllowStateChange -DryRunInvokeHandler:$dryRunInvokeHandler -MaxAttempts $maxAttempts
    } catch {
        Write-QCJsonLog -WorkerLabel $script:WorkerLabel -IncludeWorkerPid -Level 'Error' -Code 'WORKER_JOB_UNHANDLED' -Message 'Unhandled exception processing job; worker continues.' -Data @{
            jobId = $jobId
            jobType = [string]$job['type']
            handler = $handler
            errorMessage = [string]$_.Exception.Message
        }
        $res = @{ Outcome = 'failed'; ExitOk = $false; SkipId = $jobId }
    }

    if ($res.SkipId) { $skip.Add($res.SkipId) | Out-Null }
    if ($res.Outcome -in @('succeeded', 'failed', 'requeued', 'dryrun_noop')) { $processed++ }
    if (-not $res.ExitOk) { $lastOutcomeOk = $false }

    if (-not $loopMode) { break }
}

if ($lastOutcomeOk) { exit 0 } else { exit 1 }
