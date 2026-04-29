# Orchestrator.Pipeline.psm1
# Responsibility: Compose the QC pipeline from existing module functions.
# Notes:
# - No business logic here; only flow and result propagation.
# - Designed to be testable by providing fake implementations of port functions
#   (PW.* and QC.Queue.*) in the calling session.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force

function _Assert-QCResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Result,
        [Parameter(Mandatory)]
        [string]$DependencyName
    )

    if ($null -eq $Result) {
        return New-QCFailureResult -Code 'ORCH_DEPENDENCY_NULL_RESULT' -Message "Dependency returned null result: $DependencyName" -Data @{ dependency = $DependencyName }
    }

    $props = @($Result.PSObject.Properties.Name)
    $missing = @(@('IsSuccess', 'Code', 'Message', 'Data') | Where-Object { $props -notcontains $_ })
    if ($missing.Count -gt 0) {
        return New-QCFailureResult -Code 'ORCH_DEPENDENCY_INVALID_RESULT' -Message "Dependency did not return a QCResult-shaped object: $DependencyName" -Data @{ dependency = $DependencyName; missing = $missing; resultType = ($Result.GetType().FullName) }
    }

    return New-QCSuccessResult -Code 'ORCH_RESULT_OK' -Message 'QCResult shape validated.' -Data @{ dependency = $DependencyName }
}

function _Invoke-Dependency {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,
        [Parameter(Mandatory)]
        [scriptblock]$Call
    )

    try {
        $r = & $Call
    } catch {
        return New-QCFailureResult -Code 'ORCH_DEPENDENCY_THROW' -Message "Dependency threw during call: $Name" -Data @{ dependency = $Name; error = $_ }
    }

    $shape = _Assert-QCResult -Result $r -DependencyName $Name
    if (-not $shape.IsSuccess) { return $shape }
    return $r
}

function Invoke-QCPipelineTick {
    <#
    .SYNOPSIS
    Performs one discovery-and-enqueue tick.
    .DESCRIPTION
    Resolve watch paths, discover candidates, evaluate filters/triggers, and enqueue jobs.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    QCResult object.
    .NOTES
    Side effects depend on port implementations (PW discovery, queue persistence).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    $resolveWatch = _Invoke-Dependency -Name 'Resolve-WatchPaths' -Call { Resolve-WatchPaths -Config $Config }
    if (-not $resolveWatch.IsSuccess) { return $resolveWatch }
    $watchPaths = @($resolveWatch.Data.watchPaths)

    $getCandidates = _Invoke-Dependency -Name 'Get-PWTriggerCandidates' -Call { Get-PWTriggerCandidates -Config $Config -WatchPaths $watchPaths }
    if (-not $getCandidates.IsSuccess) { return $getCandidates }
    $candidates = @($getCandidates.Data.candidates)

    $tick = @{
        watchPathCount      = $watchPaths.Count
        candidateCount      = $candidates.Count
        evaluatedCount      = 0
        filteredCount       = 0
        ignoredNoMatchCount = 0
        duplicateCount      = 0
        enqueuedCount       = 0
        enqueueResults      = @()
    }

    foreach ($cand in $candidates) {
        $tick.evaluatedCount++

        $meta = _Invoke-Dependency -Name 'Get-PWCandidateMetadata' -Call { Get-PWCandidateMetadata -Candidate $cand -Config $Config }
        if (-not $meta.IsSuccess) { return $meta }
        $candidate = $meta.Data.candidate

        $allowed = _Invoke-Dependency -Name 'Test-QCPathAllowed' -Call { Test-QCPathAllowed -CandidatePath ([string]$candidate.path) -Config $Config }
        if (-not $allowed.IsSuccess) { return $allowed }
        if (-not [bool]$allowed.Data.allowed) {
            $tick.filteredCount++
            continue
        }

        $match = _Invoke-Dependency -Name 'Resolve-QCTriggerMatch' -Call { Resolve-QCTriggerMatch -Candidate $candidate -Config $Config }
        if (-not $match.IsSuccess) { return $match }
        if (-not [bool]$match.Data.matched) {
            $tick.ignoredNoMatchCount++
            continue
        }

        $jobRes = _Invoke-Dependency -Name 'New-QCJobObject' -Call { New-QCJobObject -Candidate $candidate -Rule $match.Data.rule -Config $Config }
        if (-not $jobRes.IsSuccess) { return $jobRes }
        $job = $jobRes.Data.job

        $dedupe = _Invoke-Dependency -Name 'Test-QCDuplicateJob' -Call { Test-QCDuplicateJob -DedupeKey ([string]$job.dedupeKey) -Config $Config }
        if (-not $dedupe.IsSuccess) { return $dedupe }
        if ([bool]$dedupe.Data.isDuplicate) {
            $tick.duplicateCount++
            continue
        }

        $enq = _Invoke-Dependency -Name 'Add-QCQueueJob' -Call { Add-QCQueueJob -Job $job -Config $Config }
        if (-not $enq.IsSuccess) { return $enq }

        $tick.enqueuedCount++
        $tick.enqueueResults += $enq
    }

    return New-QCSuccessResult -Code 'PIPELINE_TICK_OK' -Message 'Pipeline tick completed.' -Data $tick
}

function Invoke-QCWorkerTick {
    <#
    .SYNOPSIS
    Performs one worker tick (lock → process → transition).
    .DESCRIPTION
    Recovers stale jobs, selects next job, processes it, and updates queue state + metrics.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    QCResult object.
    .NOTES
    Side effects depend on queue + processor implementations.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    $recovery = _Invoke-Dependency -Name 'Recover-QCStaleJobs' -Call { Recover-QCStaleJobs -Config $Config }
    if (-not $recovery.IsSuccess) { return $recovery }

    $next = _Invoke-Dependency -Name 'Get-NextQCJob' -Call { Get-NextQCJob -Config $Config }
    if (-not $next.IsSuccess) { return $next }

    if (-not $next.Data.job) {
        return New-QCSuccessResult -Code 'WORKER_NO_JOB' -Message 'No eligible job found.' -Data @{ recovered = $recovery.Data }
    }

    $job = $next.Data.job
    $jobId = [string]$job.id

    $lock = _Invoke-Dependency -Name 'Lock-QCJob' -Call { Lock-QCJob -JobId $jobId -Config $Config }
    if (-not $lock.IsSuccess) { return $lock }

    $startUtc = [DateTime]::UtcNow
    $final = $null
    $state = @{
        jobId       = $jobId
        recovered   = $recovery.Data
        selected    = $job
        readyResult = $null
        processResult = $null
        moveResult  = $null
        unlockResult = $null
        durationMs  = $null
    }

    try {
        $ready = _Invoke-Dependency -Name 'Test-QCJobReady' -Call { Test-QCJobReady -Job $job -Config $Config }
        $state.readyResult = $ready
        if (-not $ready.IsSuccess) { return $ready }

        $proc = _Invoke-Dependency -Name 'Invoke-QCProcessorByType' -Call { Invoke-QCProcessorByType -Job $job -Config $Config }
        $state.processResult = $proc
        if (-not $proc.IsSuccess) {
            $moveFail = _Invoke-Dependency -Name 'Move-QCJob(failure)' -Call { Move-QCJob -JobId $jobId -FromState 'processing' -ToState 'failed' -Config $Config }
            $state.moveResult = $moveFail
            if (-not $moveFail.IsSuccess) { return $moveFail }
            $final = New-QCFailureResult -Code 'WORKER_JOB_FAILED' -Message 'Job processing failed.' -Data $state
            return $final
        }

        $moveOk = _Invoke-Dependency -Name 'Move-QCJob(success)' -Call { Move-QCJob -JobId $jobId -FromState 'processing' -ToState 'succeeded' -Config $Config }
        $state.moveResult = $moveOk
        if (-not $moveOk.IsSuccess) { return $moveOk }

        $final = New-QCSuccessResult -Code 'WORKER_JOB_SUCCEEDED' -Message 'Job processed successfully.' -Data $state
        return $final
    } finally {
        $endUtc = [DateTime]::UtcNow
        $state.durationMs = [int][Math]::Max(0, ($endUtc - $startUtc).TotalMilliseconds)

        $unlock = _Invoke-Dependency -Name 'Unlock-QCJob' -Call { Unlock-QCJob -JobId $jobId -Config $Config }
        $state.unlockResult = $unlock
        if (-not $unlock.IsSuccess) { return $unlock }

        $metrics = _Invoke-Dependency -Name 'Update-QCMetricsOnTick' -Call {
            if ($final -and $final.IsSuccess) {
                Update-QCMetricsOnSuccess -Job $job -DurationMs $state.durationMs
            } elseif ($final) {
                Update-QCMetricsOnFailure -Job $job -ErrorCode ([string]$final.Code)
            } else {
                New-QCSuccessResult -Code 'METRICS_SKIPPED' -Message 'No final result; metrics skipped.' -Data @{ jobId = $jobId }
            }
        }
        if (-not $metrics.IsSuccess) { return $metrics }
    }
}

Export-ModuleMember -Function Invoke-QCPipelineTick, Invoke-QCWorkerTick

