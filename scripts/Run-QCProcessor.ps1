<#
.SYNOPSIS
QC worker: lock -> dispatch -> transition. One-shot by default; long-running with -MaxJobs/-LeaseSeconds/-IdleSleepMs.

.DESCRIPTION
Loads appsettings.json, retrieves next queued job from JSON queue, locks it,
dispatches by job type using QC.Processors, then transitions queue state.

When -MaxJobs > 1 or -LeaseSeconds > 0 or -IdleSleepMs > 0 the script loops:
  - Picks pending jobs and processes them race-safely against other parallel workers.
  - On Lock-QCJob race loss (QUEUE_LOCK_EXISTS / QUEUE_LOCK_TIMEOUT) the loser excludes
    that id from subsequent selections in this process and tries the next pending file.
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
Import-Module (Join-Path $repoRoot 'modules\QC.Queue.Json.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.Processors.psm1') -Force

function _ToHashtable([object]$Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [System.ValueType]) { return $Value }
    if ($Value -is [System.Collections.IDictionary]) { return $Value }
    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $out = @()
        foreach ($i in $Value) { $out += (_ToHashtable $i) }
        return $out
    }
    if ($Value.PSObject -and $Value.PSObject.Properties) {
        $h = @{}
        foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = (_ToHashtable $p.Value) }
        return $h
    }
    return $Value
}

function _Read-AppSettingsJson([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        return New-QCFailureResult -Code 'CONFIG_MISSING_FILE' -Message "appsettings.json not found: $Path" -Data @{ path = $Path }
    }
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        return New-QCSuccessResult -Code 'CONFIG_LOADED' -Message 'Config loaded.' -Data @{ config = (_ToHashtable $obj); path = $Path }
    } catch {
        return New-QCFailureResult -Code 'CONFIG_PARSE_ERROR' -Message 'Failed to read/parse appsettings.json.' -Data @{ path = $Path; errorMessage = $_.Exception.Message }
    }
}

function _Log([string]$Level, [string]$Code, [string]$Message, [hashtable]$Data) {
    if (-not $Data) { $Data = @{} }
    if (-not [string]::IsNullOrWhiteSpace($script:WorkerLabel) -and -not $Data.ContainsKey('workerLabel')) {
        $Data['workerLabel'] = $script:WorkerLabel
    }
    if (-not $Data.ContainsKey('workerPid')) { $Data['workerPid'] = $PID }
    $ts = [DateTime]::UtcNow.ToString('o')
    $payload = @{ ts = $ts; level = $Level; code = $Code; message = $Message; data = $Data } | ConvertTo-Json -Depth 20 -Compress
    Write-Host $payload
}

$script:WorkerLabel = $WorkerLabel

$cfgRes = _Read-AppSettingsJson -Path $AppSettingsPath
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

_Log -Level 'Information' -Code 'WORKER_START' -Message 'Worker run started.' -Data @{
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

function _Move-QCJobWithLockRetries {
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
        _Log -Level 'Information' -Code 'WORKER_DRYRUN' -Message 'Dry-run: would dispatch job (read-only, no lock/state changes).' -Data @{
            jobId = $jobId; jobType = $jobType; handler = $Handler
            wouldRun = $true; wouldLock = $false; wouldChangeState = $false
        }
        if ($DryRunInvokeHandler -and $Handler) {
            try {
                $cmd = Get-Command -Name $Handler -ErrorAction SilentlyContinue
                if (-not $cmd) {
                    _Log -Level 'Error' -Code 'WORKER_DRYRUN_HANDLER' -Message 'Dry-run: handler not found.' -Data @{ jobId = $jobId; handler = $Handler; isSuccess = $false }
                    return @{ Outcome = 'dryrun_noop'; ExitOk = $true; SkipId = $jobId }
                }
                $Config['dryRun'] = $true
                $hr = & $Handler -Job $Job -Config $Config
                if ($null -eq $hr -or -not ($hr.PSObject.Properties.Name -contains 'IsSuccess')) {
                    _Log -Level 'Error' -Code 'WORKER_DRYRUN_HANDLER' -Message 'Dry-run: handler returned invalid result.' -Data @{ jobId = $jobId; handler = $Handler; isSuccess = $false }
                    return @{ Outcome = 'dryrun_noop'; ExitOk = $true; SkipId = $jobId }
                }
                _Log -Level 'Information' -Code 'WORKER_DRYRUN_HANDLER' -Message 'Dry-run: handler invoked.' -Data @{
                    jobId = $jobId; handler = $Handler
                    isSuccess = [bool]$hr.IsSuccess
                    resultCode = [string]$hr.Code
                    resultMessage = [string]$hr.Message
                    resultData = $hr.Data
                }
            } catch {
                _Log -Level 'Error' -Code 'WORKER_DRYRUN_HANDLER' -Message 'Dry-run: handler threw.' -Data @{ jobId = $jobId; handler = $Handler; isSuccess = $false; errorMessage = $_.Exception.Message }
            }
        }
        return @{ Outcome = 'dryrun_noop'; ExitOk = $true; SkipId = $jobId }
    }

    $lock = Lock-QCJob -JobId $jobId -Config $Config
    if (-not $lock.IsSuccess) {
        if ($lock.Code -in @('QUEUE_LOCK_EXISTS', 'QUEUE_LOCK_TIMEOUT')) {
            _Log -Level 'Information' -Code 'WORKER_LOCK_RACE' -Message 'Lost race for job lock; skipping.' -Data @{ jobId = $jobId; lockCode = [string]$lock.Code }
            return @{ Outcome = 'skipped_locked'; ExitOk = $true; SkipId = $jobId }
        }
        throw $lock.Message
    }

    try {
        $loaded = Get-QCJobById -JobId $jobId -Config $Config
        if (-not $loaded.IsSuccess -or -not $loaded.Data.found) { throw "Failed to load locked job: $jobId" }
        $Job = [hashtable]$loaded.Data.job

        if ($IsDryRun) {
            _Log -Level 'Information' -Code 'WORKER_DRYRUN' -Message 'Dry-run: would dispatch job.' -Data @{
                jobId = $jobId; jobType = [string]$Job['type']; handler = $Handler
                wouldRun = $true; wouldLock = $true; wouldChangeState = $DryRunAllowStateChange
            }
            return @{ Outcome = 'dryrun_noop'; ExitOk = $true; SkipId = $jobId }
        }

        $proc = Invoke-QCProcessorByType -Job $Job -Config $Config
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
                completedAtUtc = ([DateTime]::UtcNow.ToString('o'))
                data           = $resultData
            }
            # Single Move-QCJob call writes the in-memory $Job (with result/data)
            # to disk and renames to succeeded\ under one queue-write-lock cycle.
            # The previous Update-QCJob + Move-QCJob sequence acquired the global
            # write lock twice in a row, doubling contention and producing
            # spurious WORKER_MOVE_FAILED / QUEUE_LOCK_TIMEOUT under load.
            $mv = _Move-QCJobWithLockRetries -JobId $jobId -FromState 'running' -ToState 'succeeded' -Config $Config -Job $Job
            if (-not $mv.IsSuccess) {
                $innerErr = ''
                try { if ($mv.Data -and $mv.Data.error) { $innerErr = [string]$mv.Data.error } } catch { }
                _Log -Level 'Error' -Code 'WORKER_MOVE_FAILED' -Message 'Failed to move succeeded job to succeeded\.' -Data @{
                    jobId = $jobId; fromState = 'running'; toState = 'succeeded'
                    errorCode = [string]$mv.Code; errorMessage = [string]$mv.Message; errorData = $innerErr
                }
                # The job is still in running\. Treat as a transient failure so the
                # recovery sweep (or next worker tick) can retry the move; do NOT
                # emit a fake WORKER_SUCCEEDED.
                return @{ Outcome = 'failed'; ExitOk = $false; SkipId = $jobId }
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
            _Log -Level 'Information' -Code 'WORKER_SUCCEEDED' -Message 'Job succeeded.' -Data $logData
            return @{ Outcome = 'succeeded'; ExitOk = $true; SkipId = $jobId }
        }

        $attempts = 0
        if ($Job.ContainsKey('attempts') -and $Job.attempts -ne $null) { $attempts = [int]$Job.attempts }
        $attempts++
        $Job['attempts'] = $attempts
        $Job['lastError'] = @{ code = [string]$proc.Code; message = [string]$proc.Message; atUtc = ([DateTime]::UtcNow.ToString('o')) }

        $target = if ($attempts -ge $MaxAttempts) { 'failed' } else { 'pending' }
        # Single Move-QCJob call (same rationale as the success path - one lock
        # cycle, in-memory job preserved with attempts/lastError stamped on disk).
        $fmv = _Move-QCJobWithLockRetries -JobId $jobId -FromState 'running' -ToState $target -Config $Config -Job $Job
        if (-not $fmv.IsSuccess) {
            $innerErr = ''
            try { if ($fmv.Data -and $fmv.Data.error) { $innerErr = [string]$fmv.Data.error } } catch { }
            _Log -Level 'Error' -Code 'WORKER_MOVE_FAILED' -Message ('Failed to move failed job to ' + $target + '\.') -Data @{
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
        _Log -Level 'Error' -Code 'WORKER_FAILED' -Message 'Job failed.' -Data @{
            jobId = $jobId; jobType = [string]$Job['type']
            attempts = $attempts; maxAttempts = $MaxAttempts; movedTo = $target
            errorCode = [string]$proc.Code; errorMessage = [string]$proc.Message
            processorData = $details; processorStdout = $stdoutPreview; processorStderr = $stderrPreview
        }
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
        _Log -Level 'Information' -Code 'WORKER_BUDGET' -Message 'Worker reached MaxJobs budget; exiting.' -Data @{ processed = $processed; maxJobs = $MaxJobs }
        break
    }
    if ($LeaseSeconds -gt 0 -and ([DateTime]::UtcNow - $startedAt).TotalSeconds -ge $LeaseSeconds) {
        _Log -Level 'Information' -Code 'WORKER_LEASE' -Message 'Worker reached LeaseSeconds budget; exiting.' -Data @{ processed = $processed; leaseSeconds = $LeaseSeconds }
        break
    }

    # No global watcher gate: each STATUS_SET_GEN job is enqueued only after the
    # watcher has finished analyzing that specific folder (PW listing + manifest
    # comparison), so picking it up immediately is safe. Holding it until the entire
    # pass exited starved workers during long sweeps. QC_PREPEND was always
    # immediately eligible.
    $next = Get-NextQCJob -Config $config -ExcludeJobIds @($skip)
    if (-not $next.IsSuccess) { throw $next.Message }
    if (-not $next.Data.job) {
        _Log -Level 'Information' -Code 'WORKER_NO_JOB' -Message 'No pending jobs.' -Data @{
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
    _Log -Level 'Information' -Code 'WORKER_SELECTED' -Message 'Selected job.' -Data @{ jobId = $jobId; jobType = [string]$job['type']; dedupeKey = [string]$job['dedupeKey']; handler = $handler }

    $res = _Process-OneJob -Job $job -Handler $handler -Config $config -IsDryRun:$isDryRun -DryRunAllowStateChange:$dryRunAllowStateChange -DryRunInvokeHandler:$dryRunInvokeHandler -MaxAttempts $maxAttempts

    if ($res.SkipId) { $skip.Add($res.SkipId) | Out-Null }
    if ($res.Outcome -in @('succeeded', 'failed', 'requeued', 'dryrun_noop')) { $processed++ }
    if (-not $res.ExitOk) { $lastOutcomeOk = $false }

    if (-not $loopMode) { break }
}

if ($lastOutcomeOk) { exit 0 } else { exit 1 }
