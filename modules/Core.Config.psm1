# Core.Config.psm1
# Responsibility: Load app settings, access values, and validate compatibility.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force

function _CC-ToHashtableDeep {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return $null }
    if ($Value -is [string] -or $Value -is [System.ValueType]) { return $Value }
    if ($Value -is [hashtable]) { return $Value }

    if ($Value -is [System.Collections.IDictionary]) {
        $h = @{}
        foreach ($k in $Value.Keys) {
            $h[[string]$k] = _CC-ToHashtableDeep $Value[$k]
        }
        return $h
    }

    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $arr = @()
        foreach ($item in $Value) { $arr += (_CC-ToHashtableDeep $item) }
        return $arr
    }

    if ($Value.PSObject -and $Value.PSObject.Properties) {
        $h = @{}
        foreach ($p in $Value.PSObject.Properties) {
            $h[$p.Name] = _CC-ToHashtableDeep $p.Value
        }
        return $h
    }

    return $Value
}

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

    if (-not (Test-Path -LiteralPath $Path)) {
        return New-QCFailureResult -Code 'CONFIG_MISSING_FILE' -Message "Config file not found: $Path" -Data @{ path = $Path }
    }

    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    } catch {
        return New-QCFailureResult -Code 'CONFIG_READ_FAILED' -Message 'Failed to read config file.' -Data @{ path = $Path; errorMessage = $_.Exception.Message }
    }

    try {
        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        $cfg = _CC-ToHashtableDeep $obj
        if (-not ($cfg -is [hashtable])) { $cfg = @{} + $cfg }
        $valRes = Test-AppSettings -Config $cfg
        if (-not $valRes.IsSuccess) { return $valRes }
        return New-QCSuccessResult -Code 'CONFIG_LOADED' -Message 'Config loaded.' -Data @{ path = $Path; config = $cfg }
    } catch {
        return New-QCFailureResult -Code 'CONFIG_PARSE_FAILED' -Message 'Failed to parse config as JSON.' -Data @{ path = $Path; errorMessage = $_.Exception.Message }
    }
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

    if (-not $Config.ContainsKey('projectWise') -or -not ($Config.projectWise -is [hashtable])) {
        return New-QCFailureResult -Code 'CONFIG_VALIDATION_MISSING_KEY' -Message 'Missing required section: projectWise' -Data @{ key = 'projectWise' }
    }
    if (-not $Config.projectWise.ContainsKey('datasourceName') -or [string]::IsNullOrWhiteSpace([string]$Config.projectWise.datasourceName)) {
        return New-QCFailureResult -Code 'CONFIG_VALIDATION_MISSING_KEY' -Message 'Missing required setting: projectWise.datasourceName' -Data @{ key = 'projectWise.datasourceName' }
    }
    if (-not $Config.projectWise.ContainsKey('credentialPath') -or [string]::IsNullOrWhiteSpace([string]$Config.projectWise.credentialPath)) {
        return New-QCFailureResult -Code 'CONFIG_VALIDATION_MISSING_KEY' -Message 'Missing required setting: projectWise.credentialPath' -Data @{ key = 'projectWise.credentialPath' }
    }
    return New-QCSuccessResult -Code 'CONFIG_VALID' -Message 'Config validated.' -Data @{}
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

    if ([string]::IsNullOrWhiteSpace($KeyPath)) {
        return New-QCFailureResult -Code 'CONFIG_KEYPATH_EMPTY' -Message 'KeyPath is required.' -Data @{}
    }

    $cur = $Config
    foreach ($part in ($KeyPath -split '\.')) {
        if ($cur -is [hashtable]) {
            if (-not $cur.ContainsKey($part)) { return New-QCSuccessResult -Code 'CONFIG_VALUE_DEFAULT' -Message 'Setting not found; returning default.' -Data @{ value = $DefaultValue; isDefault = $true } }
            $cur = $cur[$part]
            continue
        }
        return New-QCSuccessResult -Code 'CONFIG_VALUE_DEFAULT' -Message 'Setting not found; returning default.' -Data @{ value = $DefaultValue; isDefault = $true }
    }

    if ($null -eq $cur) {
        return New-QCSuccessResult -Code 'CONFIG_VALUE_DEFAULT' -Message 'Setting is null; returning default.' -Data @{ value = $DefaultValue; isDefault = $true }
    }
    return New-QCSuccessResult -Code 'CONFIG_VALUE' -Message 'Setting resolved.' -Data @{ value = $cur; isDefault = $false }
}

function _CC-ResolveProjectNamingAnchor([string]$WatchRootPath) {
    if ([string]::IsNullOrWhiteSpace($WatchRootPath)) { return '' }
    $rootNorm = ([string]$WatchRootPath).Trim() -replace '/', '\'
    $rootNorm = $rootNorm.TrimEnd('\')
    $parts = @($rootNorm -split '\\' | Where-Object { $_ -ne '' })
    if ($parts.Count -ge 2 -and $parts[1].Equals('Caltrans', [StringComparison]::OrdinalIgnoreCase)) {
        return 'Documents\Caltrans'
    }
    return $rootNorm
}

function Get-QCProjectNameFromFolderPath {
    <#
    .SYNOPSIS
    Extracts the project name from a ProjectWise folder path using watchList.roots configuration.
    .DESCRIPTION
    Returns the path segment(s) between the environment folder (Documents\Caltrans, Documents\AZDOT 2024,
    Documents\AZDOT) and sheetsPathFromProject (typically CADD\Sheets), even when the document lives in
    subfolders under Sheets (e.g. Seg_1).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory)]
        [string]$FolderPath
    )

    if ([string]::IsNullOrWhiteSpace($FolderPath)) { return $null }
    if (-not $Config) { return $null }

    $norm = ([string]$FolderPath).Trim() -replace '/', '\'
    $norm = $norm.TrimEnd('\')
    if ($norm -match '^(?i)pw:\\?') {
        $norm = $norm -replace '^(?i)pw:\\?', ''
    }
    $docsIdx = $norm.IndexOf('Documents\', [StringComparison]::OrdinalIgnoreCase)
    if ($docsIdx -gt 0) {
        $norm = $norm.Substring($docsIdx)
    }

    $roots = $null
    try {
        if ($Config.ContainsKey('projectWise') -and $Config.projectWise -and
            $Config.projectWise.ContainsKey('watchList') -and $Config.projectWise.watchList -and
            $Config.projectWise.watchList.ContainsKey('roots') -and $Config.projectWise.watchList.roots) {
            $roots = @($Config.projectWise.watchList.roots)
        }
    } catch { $roots = $null }
    if (-not $roots) { return $null }

    $anchors = @{}
    foreach ($r in $roots) {
        $rootPath = $null
        $sheetsRel = 'CADD\Sheets'
        try { if ($r.path) { $rootPath = [string]$r.path } } catch { $rootPath = $null }
        try { if ($r.sheetsPathFromProject) { $sheetsRel = [string]$r.sheetsPathFromProject } } catch { }
        if (-not $rootPath) { continue }

        $anchor = _CC-ResolveProjectNamingAnchor -WatchRootPath $rootPath
        if ([string]::IsNullOrWhiteSpace($anchor)) { continue }
        $suf = ([string]$sheetsRel).Trim() -replace '/', '\'
        $suf = $suf.Trim('\')
        if (-not $suf) { $suf = 'CADD\Sheets' }
        if (-not $anchors.ContainsKey($anchor)) {
            $anchors[$anchor] = $suf
        }
    }

    $anchorKeys = @($anchors.Keys | Sort-Object { $_.Length } -Descending)
    foreach ($anchor in $anchorKeys) {
        $suf = [string]$anchors[$anchor]
        $anchorPrefix = $anchor.TrimEnd('\') + '\'
        if (-not ($norm.StartsWith($anchorPrefix, [StringComparison]::OrdinalIgnoreCase))) { continue }

        $marker = '\' + $suf
        $markerIdx = $norm.IndexOf($marker, $anchorPrefix.Length, [StringComparison]::OrdinalIgnoreCase)
        if ($markerIdx -lt 0) { continue }

        $project = $norm.Substring($anchorPrefix.Length, $markerIdx - $anchorPrefix.Length)
        $project = $project.Trim('\')
        if (-not [string]::IsNullOrWhiteSpace($project)) { return $project }
    }

    return $null
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

    # No schema versioning in appsettings yet; keep placeholder for future.
    return New-QCSuccessResult -Code 'CONFIG_COMPAT_OK' -Message 'Config compatibility OK.' -Data @{}
}
Export-ModuleMember -Function *
