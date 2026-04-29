# QC.JobFactory.psm1
# Responsibility: Build and validate queue job payloads.

function New-QCJobId {
    <#
    .SYNOPSIS
    Generates a deterministic or policy-compliant job identifier.
    .DESCRIPTION
    Produces a job ID using candidate/rule/config inputs and naming conventions.
    .PARAMETER Candidate
    Candidate metadata object.
    .PARAMETER Rule
    Trigger rule or match object.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: none.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Candidate,
        [Parameter(Mandatory)]
        [hashtable]$Rule,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
}

function New-QCJobObject {
    <#
    .SYNOPSIS
    Creates a standardized queue job object.
    .DESCRIPTION
    Builds a job payload from candidate metadata, trigger match, and app settings.
    .PARAMETER Candidate
    Candidate metadata object.
    .PARAMETER Rule
    Trigger rule or match object.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: none.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Candidate,
        [Parameter(Mandatory)]
        [hashtable]$Rule,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
}

function Test-QCJobRequiredFields {
    <#
    .SYNOPSIS
    Validates required fields in a job payload.
    .DESCRIPTION
    Confirms required keys/values exist before enqueue operations.
    .PARAMETER Job
    Job payload to validate.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: none.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job
    )
}

function Get-QCDedupeKey {
    <#
    .SYNOPSIS
    Computes dedupe key for a job.
    .DESCRIPTION
    Produces a deduplication key based on configured key parts.
    .PARAMETER Job
    Job payload.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: none.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
}
Export-ModuleMember -Function *
