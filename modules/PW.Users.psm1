# PW.Users.psm1
# Responsibility: Resolve ProjectWise user numbers to username/email and persist in pw_users.

function _PWU-GetPropertyValue {
    param(
        [Parameter(Mandatory)][object]$Object,
        [Parameter(Mandatory)][string[]]$Names
    )
    foreach ($name in $Names) {
        if ($null -eq $Object) { return $null }
        if ($Object -is [hashtable] -and $Object.ContainsKey($name)) {
            $v = $Object[$name]
            if ($null -ne $v -and -not ($v -is [DBNull])) { return [string]$v }
        }
        $prop = $Object.PSObject.Properties[$name]
        if ($prop -and $null -ne $prop.Value -and -not ($prop.Value -is [DBNull])) {
            return [string]$prop.Value
        }
    }
    return $null
}

function _PWU-MapUserFromPwCmdlet {
    param(
        [Parameter(Mandatory)][int]$UserNumber,
        [Parameter(Mandatory)][object]$User
    )
    $username = _PWU-GetPropertyValue -Object $User -Names @('UserName', 'Name', 'LoginName')
    $email = _PWU-GetPropertyValue -Object $User -Names @('EMail', 'Email', 'EMailAddress', 'EmailAddress')
    $display = _PWU-GetPropertyValue -Object $User -Names @('Description', 'FullName', 'DisplayName')
    return @{
        pw_userno     = $UserNumber
        pw_username   = $username
        pw_user_email = $email
        display_name  = $display
    }
}

function Sync-PWUserDirectory {
    <#
    .SYNOPSIS
    Resolves pw_userno values via Get-PWUser and upserts rows into dbo.pw_users.
    .DESCRIPTION
    Best-effort: no-ops when PW cmdlets are unavailable or database writes are disabled.
    Use -ResolveFromAuditEvents to backfill users referenced in audit_events but missing from pw_users.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [int[]]$UserNumbers = @(),
        [switch]$ResolveFromAuditEvents,
        [int]$MaxUsers = 50
    )

    if (-not (Get-Command -Name 'Write-QCPWUserDirectory' -ErrorAction SilentlyContinue)) {
        return New-QCFailureResult -Code 'PW_USER_DB_MODULE_MISSING' -Message 'Core.Database.psm1 must be imported (Write-QCPWUserDirectory).' -Data @{}
    }
    if (-not (Test-QCDatabaseEnabled -Config $Config)) {
        return New-QCSuccessResult -Code 'PW_USER_SYNC_SKIPPED' -Message 'Database disabled.' -Data @{ synced = 0; skipped = 0 }
    }
    if (-not (Test-QCDatabaseWritesAllowed -Config $Config)) {
        return New-QCSuccessResult -Code 'PW_USER_SYNC_SKIPPED' -Message 'Database writes not allowed (dry-run).' -Data @{ synced = 0; skipped = 0 }
    }

    $numbers = @($UserNumbers | Where-Object { $_ -gt 0 } | Select-Object -Unique)
    if ($ResolveFromAuditEvents) {
        $missingRes = Get-QCPWUnresolvedUserNumbers -Config $Config -MaxCount $MaxUsers
        if ($missingRes.IsSuccess -and $missingRes.Data.numbers) {
            $numbers = @($numbers + @($missingRes.Data.numbers) | Select-Object -Unique)
        }
    }
    if ($numbers.Count -gt $MaxUsers) {
        $numbers = @($numbers | Select-Object -First $MaxUsers)
    }
    if ($numbers.Count -eq 0) {
        return New-QCSuccessResult -Code 'PW_USER_SYNC_NONE' -Message 'No user numbers to resolve.' -Data @{ synced = 0; skipped = 0 }
    }

    if (-not (Get-Command -Name 'Get-PWUser' -ErrorAction SilentlyContinue)) {
        return New-QCFailureResult -Code 'PW_USER_CMDLET_MISSING' -Message 'Get-PWUser is not available (ProjectWise module not loaded).' -Data @{ requested = $numbers.Count }
    }

    $rows = @()
    $skipped = 0
    foreach ($n in $numbers) {
        try {
            $user = Get-PWUser -UserID $n -ErrorAction Stop
            if (-not $user) { $skipped++; continue }
            if ($user -is [System.Array]) { $user = $user | Select-Object -First 1 }
            $rows += _PWU-MapUserFromPwCmdlet -UserNumber $n -User $user
        } catch {
            $skipped++
        }
    }

    if ($rows.Count -eq 0) {
        return New-QCSuccessResult -Code 'PW_USER_SYNC_EMPTY' -Message 'No users resolved from ProjectWise.' -Data @{ synced = 0; skipped = $skipped; requested = $numbers.Count }
    }

    $writeRes = Write-QCPWUserDirectory -Config $Config -Users $rows
    if (-not $writeRes.IsSuccess) { return $writeRes }

    return New-QCSuccessResult -Code 'PW_USER_SYNC_OK' -Message "Synced $($rows.Count) user(s) into pw_users." -Data @{
        synced    = $rows.Count
        skipped   = $skipped
        requested = $numbers.Count
    }
}

Export-ModuleMember -Function Sync-PWUserDirectory
