# PW.Connection.psm1
# Responsibility: ProjectWise connection port (read-only safe by default).

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force

function _PWC-IsNullOrWhiteSpace([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function Get-PWCredentialFromFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$CredentialPath
    )

    if (-not (Test-Path -LiteralPath $CredentialPath)) {
        return New-QCFailureResult -Code 'PW_CRED_MISSING_FILE' -Message "Credential file not found: $CredentialPath" -Data @{ path = $CredentialPath }
    }
    try {
        $lines = Get-Content -LiteralPath $CredentialPath -ErrorAction Stop
        $uLine = $lines | Where-Object { $_ -match '^\s*username\s*=' } | Select-Object -First 1
        $pLine = $lines | Where-Object { $_ -match '^\s*password\s*=' } | Select-Object -First 1
        if (-not $uLine -or -not $pLine) {
            return New-QCFailureResult -Code 'PW_CRED_BAD_FORMAT' -Message "Invalid format in credential file: $CredentialPath" -Data @{ path = $CredentialPath }
        }
        $user = ($uLine -split '=', 2)[1].Trim()
        $pass = ($pLine -split '=', 2)[1].Trim()
        if ([string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($pass)) {
            return New-QCFailureResult -Code 'PW_CRED_BAD_FORMAT' -Message "Credential file missing username/password values: $CredentialPath" -Data @{ path = $CredentialPath }
        }
        $sec = ConvertTo-SecureString $pass -AsPlainText -Force
        $cred = [pscredential]::new($user, $sec)
        return New-QCSuccessResult -Code 'PW_CRED_LOADED' -Message 'Credential loaded.' -Data @{ credential = $cred; userName = $user }
    } catch {
        return New-QCFailureResult -Code 'PW_CRED_READ_FAILED' -Message 'Failed to read credential file.' -Data @{ path = $CredentialPath; errorMessage = $_.Exception.Message }
    }
}

function Connect-PW {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DatasourceName,
        [Parameter(Mandatory)]
        [pscredential]$Credential
    )

    $cmd = Get-Command -Name Open-PWConnection -ErrorAction SilentlyContinue
    if (-not $cmd) {
        return New-QCFailureResult -Code 'PW_MISSING_MODULE' -Message 'Open-PWConnection not found. Run from ProjectWise PowerShell (pwps) so pwps_dab is loaded.' -Data @{}
    }

    try {
        Open-PWConnection -DatasourceName $DatasourceName -UserName $Credential.UserName -Password $Credential.Password -WarningAction SilentlyContinue | Out-Null
        return New-QCSuccessResult -Code 'PW_CONNECTED' -Message 'ProjectWise connected.' -Data @{ datasourceName = $DatasourceName; userName = $Credential.UserName }
    } catch {
        return New-QCFailureResult -Code 'PW_CONNECT_FAILED' -Message 'ProjectWise connection failed.' -Data @{ datasourceName = $DatasourceName; userName = $Credential.UserName; errorMessage = $_.Exception.Message }
    }
}

function Disconnect-PW {
    [CmdletBinding()]
    param()

    $cmd = Get-Command -Name Close-PWConnection -ErrorAction SilentlyContinue
    if (-not $cmd) {
        return New-QCFailureResult -Code 'PW_MISSING_MODULE' -Message 'Close-PWConnection not found. Run from ProjectWise PowerShell (pwps) so pwps_dab is loaded.' -Data @{}
    }

    try {
        Close-PWConnection -ErrorAction SilentlyContinue | Out-Null
        return New-QCSuccessResult -Code 'PW_DISCONNECTED' -Message 'ProjectWise disconnected.' -Data @{}
    } catch {
        return New-QCFailureResult -Code 'PW_DISCONNECT_FAILED' -Message 'ProjectWise disconnect failed.' -Data @{ errorMessage = $_.Exception.Message }
    }
}

Export-ModuleMember -Function *

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
