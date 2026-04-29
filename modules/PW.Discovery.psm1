# PW.Discovery.psm1
# Responsibility: Read-only ProjectWise watch-path resolution and candidate discovery.

function Resolve-WatchPaths {
    <#
    .SYNOPSIS
    Resolves effective watch paths from configuration.
    .DESCRIPTION
    Expands configured roots/folders into a de-duplicated watch path list.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: read-only discovery operations.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
}

function Get-PWTriggerCandidates {
    <#
    .SYNOPSIS
    Retrieves trigger candidate files/events from watch paths.
    .DESCRIPTION
    Performs read-only lookup of candidate documents relevant for trigger evaluation.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .PARAMETER WatchPaths
    Effective watch paths to query.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: read-only ProjectWise/API calls.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory)]
        [string[]]$WatchPaths
    )
}

function Get-PWCandidateMetadata {
    <#
    .SYNOPSIS
    Enriches a trigger candidate with metadata.
    .DESCRIPTION
    Reads additional metadata fields required for filtering and trigger classification.
    .PARAMETER Candidate
    Candidate object from discovery.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: read-only ProjectWise/API calls.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Candidate,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
}
Export-ModuleMember -Function *
