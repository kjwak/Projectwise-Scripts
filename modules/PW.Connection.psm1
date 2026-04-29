# PW.Connection.psm1
# Responsibility: ProjectWise connection health and reconnect logic wrappers.

function Test-PWLoginHealth {
    <#
    .SYNOPSIS
    Checks ProjectWise login/session health.
    .DESCRIPTION
    Verifies connection/session viability for read-only operations.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: read-only connectivity checks.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
}

function Connect-PWIfNeeded {
    <#
    .SYNOPSIS
    Ensures a valid ProjectWise session exists.
    .DESCRIPTION
    Establishes or refreshes connection only as needed for subsequent read operations.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: session management calls; no document writes.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
}

Export-ModuleMember -Function *
