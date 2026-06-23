# Core.Results.psm1
# Responsibility: Standardized result object constructors for all module functions.

function New-QCResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [bool]$IsSuccess,
        [Parameter(Mandatory)]
        [string]$Code,
        [Parameter(Mandatory)]
        [string]$Message,
        [object]$Data
    )

    return [pscustomobject]@{
        IsSuccess = $IsSuccess
        Code      = $Code
        Message   = $Message
        Data      = $Data
    }
}

function New-QCSuccessResult {
    [CmdletBinding()]
    param(
        [string]$Code = 'OK',
        [string]$Message = 'Success',
        [object]$Data
    )

    return New-QCResult -IsSuccess $true -Code $Code -Message $Message -Data $Data
}

function New-QCFailureResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Code,
        [Parameter(Mandatory)]
        [string]$Message,
        [object]$Data
    )

    return New-QCResult -IsSuccess $false -Code $Code -Message $Message -Data $Data
}

function New-QCErrorResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Code,
        [Parameter(Mandatory)]
        [string]$Message,
        [object]$Data
    )

    return New-QCFailureResult -Code $Code -Message $Message -Data $Data
}

function Ensure-QCJsonLogAvailable {
    <#
    .SYNOPSIS
    Re-imports Core.Runtime when nested Import-Module -Force dropped Write-QCJsonLog from the session.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ModulesRoot
    )

    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) { return $true }
    $runtimePath = Join-Path $ModulesRoot 'Core\Core.Runtime.psm1'
    if (-not (Test-Path -LiteralPath $runtimePath)) { return $false }
    Import-Module $runtimePath -Force -WarningAction SilentlyContinue | Out-Null
    return [bool](Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)
}

Export-ModuleMember -Function *
