# Core.Metrics.psm1
# Responsibility: Maintain queue/processing counters and timing metrics.

function Initialize-QCMetrics {
    <#
    .SYNOPSIS
    Initializes metrics state storage.
    .DESCRIPTION
    Prepares in-memory and persisted metrics structures.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: metrics file/state initialization.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
}

function Get-QCMetricsSnapshot {
    <#
    .SYNOPSIS
    Retrieves current metrics snapshot.
    .DESCRIPTION
    Returns current aggregate metrics for health/dashboard views.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: read-only metrics access.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
}

function Update-QCMetricsOnEnqueue {
    <#
    .SYNOPSIS
    Updates metrics after enqueue.
    .DESCRIPTION
    Increments queue counters for newly queued jobs.
    .PARAMETER Job
    Job payload.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: metrics state update.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job
    )
}

function Update-QCMetricsOnStart {
    <#
    .SYNOPSIS
    Updates metrics when processing starts.
    .DESCRIPTION
    Adjusts queued/processing counters when a job begins execution.
    .PARAMETER Job
    Job payload.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: metrics state update.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job
    )
}

function Update-QCMetricsOnSuccess {
    <#
    .SYNOPSIS
    Updates metrics when a job succeeds.
    .DESCRIPTION
    Records completion counts and duration aggregates.
    .PARAMETER Job
    Job payload.
    .PARAMETER DurationMs
    Processing duration in milliseconds.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: metrics state update.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job,
        [Parameter(Mandatory)]
        [int]$DurationMs
    )
}

function Update-QCMetricsOnFailure {
    <#
    .SYNOPSIS
    Updates metrics when a job fails.
    .DESCRIPTION
    Records failure counts and optional error-category counters.
    .PARAMETER Job
    Job payload.
    .PARAMETER ErrorCode
    Error category/code.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: metrics state update.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job,
        [string]$ErrorCode
    )
}

function Update-QCMetricsOnRetry {
    <#
    .SYNOPSIS
    Updates metrics when a job is retried.
    .DESCRIPTION
    Tracks retry counts and retry-related classifications.
    .PARAMETER Job
    Job payload.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: metrics state update.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job
    )
}

function Reset-QCMetricsDaily {
    <#
    .SYNOPSIS
    Performs daily metrics reset/rollover.
    .DESCRIPTION
    Rotates daily counters while preserving historical aggregates as configured.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: metrics state write/rotation.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
}

function Flush-QCMetrics {
    <#
    .SYNOPSIS
    Flushes buffered metrics to persistence.
    .DESCRIPTION
    Commits pending in-memory metric updates to the configured backend.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: metrics backend writes.
    #>
    [CmdletBinding()]
    param(
        [hashtable]$Config
    )
}
Export-ModuleMember -Function *
