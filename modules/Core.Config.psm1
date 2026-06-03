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

function Get-QCProjectNameFromFolderPath {
    <#
    .SYNOPSIS
    Extracts the project name from a ProjectWise folder path using watchList.roots configuration.
    .DESCRIPTION
    For paths like Documents\AZDOT 2024\<project>\CADD\Sheets\..., returns the segment(s) between
    the configured watch root and sheetsPathFromProject (per projectDepth).
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

    foreach ($r in $roots) {
        $rootPath = $null
        $sheetsRel = $null
        $depth = 1
        try { if ($r.path) { $rootPath = [string]$r.path } } catch { $rootPath = $null }
        try { if ($r.sheetsPathFromProject) { $sheetsRel = [string]$r.sheetsPathFromProject } } catch { $sheetsRel = $null }
        try { if ($r.projectDepth) { $depth = [int]$r.projectDepth } } catch { $depth = 1 }
        if (-not $rootPath) { continue }
        if ($depth -lt 1) { $depth = 1 }

        $rootNorm = ([string]$rootPath).Trim() -replace '/', '\'
        $rootNorm = $rootNorm.TrimEnd('\')
        $prefix = $rootNorm + '\'
        if (-not ($norm.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase))) { continue }

        $rel = $norm.Substring($prefix.Length)
        if ([string]::IsNullOrWhiteSpace($rel)) { continue }

        if ($sheetsRel) {
            $suf = ([string]$sheetsRel).Trim() -replace '/', '\'
            $suf = $suf.Trim('\')
            if ($suf) {
                $suffix = '\' + $suf
                if ($rel.EndsWith($suffix, [StringComparison]::OrdinalIgnoreCase)) {
                    $rel = $rel.Substring(0, $rel.Length - $suffix.Length)
                }
            }
        }

        $rel = $rel.Trim('\')
        if ([string]::IsNullOrWhiteSpace($rel)) { continue }

        $parts = @($rel -split '\\' | Where-Object { $_ -ne '' })
        if ($parts.Count -le 0) { continue }
        $take = [Math]::Min([int]$depth, $parts.Count)
        $proj = ($parts | Select-Object -First $take) -join '\'
        if (-not [string]::IsNullOrWhiteSpace($proj)) { return $proj }
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
