function Resolve-ModuleImplPath {
    <#
    .SYNOPSIS
    Resolves a flat modules/*.psm1 shim to its Phase 4E implementation path.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,

        [string]$ModulesDir
    )

    if (-not $ModulesDir) {
        $ModulesDir = Join-Path (Split-Path -Parent $PSScriptRoot) 'modules'
    }

    $shim = Join-Path $ModulesDir $ModuleName
    if (-not (Test-Path -LiteralPath $shim)) {
        throw "Module shim not found: $shim"
    }

    $text = Get-Content -LiteralPath $shim -Raw
    if ($text -match "Join-Path\s+\`$PSScriptRoot\s+'([^']+\.psm1)'") {
        $rel = $Matches[1] -replace '/', '\'
        $impl = Join-Path $ModulesDir $rel
        if (Test-Path -LiteralPath $impl) {
            return $impl
        }
    }

    return $shim
}

function Get-QCModuleImplementation {
    <#
    .SYNOPSIS
    Returns the loaded implementation module (not the flat compatibility shim) when both are present.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,

        [string]$ModulesDir
    )

    $implPath = [System.IO.Path]::GetFullPath((Resolve-ModuleImplPath -ModuleName $ModuleName -ModulesDir $ModulesDir))
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ModuleName)
    foreach ($m in @(Get-Module -Name $baseName -All)) {
        if ($m.Path -and ([System.IO.Path]::GetFullPath($m.Path) -eq $implPath)) {
            return $m
        }
    }

    return (Get-Module -Name $baseName | Select-Object -First 1)
}

function Remove-QCModuleFlatShims {
    <#
    .SYNOPSIS
    Removes flat-path compatibility shim modules, leaving folder implementations loaded.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName,

        [string]$ModulesDir
    )

    $implPath = [System.IO.Path]::GetFullPath((Resolve-ModuleImplPath -ModuleName $ModuleName -ModulesDir $ModulesDir))
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ModuleName)
    foreach ($m in @(Get-Module -Name $baseName -All)) {
        if ($m.Path -and ([System.IO.Path]::GetFullPath($m.Path) -ne $implPath)) {
            Remove-Module -Module $m -Force -ErrorAction SilentlyContinue
        }
    }
}
