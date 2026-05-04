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

function Get-PWImmediateChildFolders {
    <#
    .SYNOPSIS
    Returns immediate child folder objects under a ProjectWise folder (read-only).

    .DESCRIPTION
    pwps_dab varies by path: Get-PWFoldersImmediateChildren often returns nothing under CADD\Sheets while
    Get-PWFolderView lists discipline subfolders. This aligns oneLevelDeep expansion with Sheets discovery.

    FolderPath MUST use the pw cmdlet convention: without a leading Documents\ segment.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FolderPath
    )

    function _PwcPwProp([object]$Obj, [string]$Name) {
        try {
            if ($null -eq $Obj -or -not $Obj.PSObject -or -not $Obj.PSObject.Properties[$Name]) { return $null }
            return $Obj.$Name
        } catch { return $null }
    }

    $normalized = (($FolderPath -as [string]).Trim()).TrimEnd('\')
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return @()
    }

    $collected = [System.Collections.Generic.List[object]]::new()

    $view = $null
    try {
        $view = Get-PWFolderView -FolderPath $normalized -ErrorAction Stop
    } catch {
        try {
            $folderObj = Get-PWFolders -FolderPath $normalized -JustOne -ErrorAction SilentlyContinue
            if ($folderObj) {
                $view = $folderObj | Get-PWFolderView -ErrorAction SilentlyContinue
            }
        } catch { }
    }

    if ($view -and $view.Folders) {
        foreach ($f in @($view.Folders)) {
            if ($f) { [void]$collected.Add($f) }
        }
    }

    if ($collected.Count -eq 0 -and $view -and $view.Children) {
        foreach ($c in @($view.Children)) {
            if (-not $c) { continue }
            $docId = _PwcPwProp $c 'DocumentID'
            if (-not ([string]::IsNullOrWhiteSpace([string]$docId))) {
                continue
            }
            if (-not (_PwcPwProp $c 'FolderPath')) {
                continue
            }
            [void]$collected.Add($c)
        }
    }

    if ($collected.Count -gt 0) {
        return @($collected.ToArray())
    }

    try {
        return @(Get-PWFoldersImmediateChildren -FolderPath $normalized -WarningAction SilentlyContinue -ErrorAction Stop)
    } catch {
        try {
            return @(Get-PWFoldersImmediateChildren -FolderPath $normalized -WarningAction SilentlyContinue -ErrorAction SilentlyContinue)
        } catch {
            return @()
        }
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
