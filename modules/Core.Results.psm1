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

Export-ModuleMember -Function *
