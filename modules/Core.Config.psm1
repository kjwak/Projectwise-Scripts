# Core.Config.psm1
# Responsibility: Load app settings, access values, and validate compatibility.

function Read-AppConfig {
    <#
    .SYNOPSIS
    Loads application settings from a file path.
    .DESCRIPTION
    Reads appsettings content and returns an in-memory configuration object.
    .PARAMETER Path
    Path to the appsettings file.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: file read only.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )
}

function Test-AppSettings {
    <#
    .SYNOPSIS
    Validates required configuration structure and values.
    .DESCRIPTION
    Performs structural checks for required sections and expected types.
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
        [hashtable]$Config
    )
}

function Get-AppSetting {
    <#
    .SYNOPSIS
    Retrieves a nested setting value by key path.
    .DESCRIPTION
    Resolves a configuration value from the loaded settings object.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .PARAMETER KeyPath
    Dot-delimited key path (for example: queue.backend).
    .PARAMETER DefaultValue
    Optional default returned when key is not found.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: none.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory)]
        [string]$KeyPath,
        [object]$DefaultValue
    )
}

function Test-AppSettingsCompatibility {
    <#
    .SYNOPSIS
    Verifies appsettings version compatibility.
    .DESCRIPTION
    Confirms the loaded configuration version/schema is supported by the scripts.
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
        [hashtable]$Config
    )
}
Export-ModuleMember -Function *
