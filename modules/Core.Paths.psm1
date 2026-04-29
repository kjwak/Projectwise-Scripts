# Core.Paths.psm1
# Responsibility: Normalize and inspect local/ProjectWise-style paths.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force

function Normalize-QCPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return New-QCFailureResult -Code 'PATH_EMPTY' -Message 'Path is null, empty, or whitespace.' -Data @{ input = $Path }
    }

    $normalized = $Path.Trim()
    $normalized = $normalized -replace '/', '\\'
    $normalized = $normalized -replace '\\{2,}', '\'

    if ($normalized.Length -gt 1 -and $normalized.EndsWith('\')) {
        $normalized = $normalized.TrimEnd('\\')
    }

    $normalized = $normalized.ToLowerInvariant()

    return New-QCSuccessResult -Code 'PATH_NORMALIZED' -Message 'Path normalized.' -Data @{ path = $normalized }
}

function Normalize-QCPaths {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$Paths
    )

    $items = @()
    foreach ($p in $Paths) {
        $r = Normalize-QCPath -Path $p
        if (-not $r.IsSuccess) {
            return New-QCFailureResult -Code 'PATHS_NORMALIZE_FAILED' -Message 'At least one path failed normalization.' -Data @{ failedPath = $p; result = $r }
        }
        $items += [string]$r.Data.path
    }

    return New-QCSuccessResult -Code 'PATHS_NORMALIZED' -Message 'Paths normalized.' -Data @{ paths = $items }
}

function Test-PathUnderRoot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [string]$Root
    )

    $pathRes = Normalize-QCPath -Path $Path
    if (-not $pathRes.IsSuccess) { return $pathRes }
    $rootRes = Normalize-QCPath -Path $Root
    if (-not $rootRes.IsSuccess) { return $rootRes }

    $p = [string]$pathRes.Data.path
    $r = [string]$rootRes.Data.path

    $isUnder = ($p -eq $r) -or $p.StartsWith($r + '\')

    return New-QCSuccessResult -Code 'PATH_UNDER_ROOT_EVALUATED' -Message 'Path-under-root evaluated.' -Data @{ isUnderRoot = $isUnder; path = $p; root = $r }
}

function Split-QCPathParts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $pathRes = Normalize-QCPath -Path $Path
    if (-not $pathRes.IsSuccess) { return $pathRes }

    $normalized = [string]$pathRes.Data.path
    $parts = @($normalized -split '\\' | Where-Object { $_ -ne '' })

    return New-QCSuccessResult -Code 'PATH_SPLIT' -Message 'Path split into parts.' -Data @{
        normalizedPath = $normalized
        parts = $parts
        partCount = $parts.Count
        leaf = if ($parts.Count -gt 0) { $parts[-1] } else { '' }
    }
}

Export-ModuleMember -Function *
