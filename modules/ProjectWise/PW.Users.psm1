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

function _PWU-MapUserFromPwObject {
    param(
        [Parameter(Mandatory)][int]$UserNumber,
        [Parameter(Mandatory)][object]$User
    )
    $resolvedNo = $UserNumber
    $fromObj = _PWU-GetPropertyValue -Object $User -Names @('UserID', 'ID', 'o_userno', 'pw_userno')
    if ($fromObj) { try { $resolvedNo = [int]$fromObj } catch { } }

    return @{
        pw_userno     = $resolvedNo
        pw_username   = (_PWU-GetPropertyValue -Object $User -Names @('UserName', 'Name', 'LoginName', 'o_username'))
        pw_user_email = (_PWU-GetPropertyValue -Object $User -Names @('Email', 'EMail', 'EMailAddress', 'EmailAddress', 'o_email'))
        display_name  = (_PWU-GetPropertyValue -Object $User -Names @('Description', 'FullName', 'DisplayName'))
    }
}

function _PWU-MapUserFromSqlRow {
    param([Parameter(Mandatory)][object]$Row)
    $userno = 0
    foreach ($name in @('o_userno', 'pw_userno')) {
        $v = _PWU-GetPropertyValue -Object $Row -Names @($name)
        if ($v) { try { $userno = [int]$v; break } catch { } }
    }
    if ($userno -le 0) { return $null }
    return @{
        pw_userno     = $userno
        pw_username   = (_PWU-GetPropertyValue -Object $Row -Names @('o_username', 'UserName'))
        pw_user_email = (_PWU-GetPropertyValue -Object $Row -Names @('o_email', 'Email', 'EMail'))
        display_name  = $null
    }
}

function _PWU-ResolveUsersViaSql {
    param([Parameter(Mandatory)][int[]]$UserNumbers)

    if (-not (Get-Command -Name 'Select-PWSQL' -ErrorAction SilentlyContinue)) { return @() }
    $byNumber = @{}
    $chunkSize = 100
    for ($i = 0; $i -lt $UserNumbers.Count; $i += $chunkSize) {
        $chunk = @($UserNumbers[$i..[Math]::Min($i + $chunkSize - 1, $UserNumbers.Count - 1)] | Where-Object { $_ -gt 0 })
        if ($chunk.Count -eq 0) { continue }
        $inList = ($chunk | ForEach-Object { [string][int]$_ }) -join ','
        $sql = "SELECT o_userno, o_username, o_email FROM dms_user WHERE o_userno IN ($inList)"
        try {
            $result = Select-PWSQL -SQLSelectStatement $sql -ErrorAction Stop
            $sqlRows = @()
            if ($result -and $result.Rows) { $sqlRows = @($result.Rows) }
            elseif ($result) { $sqlRows = @($result) }
            foreach ($row in $sqlRows) {
                $mapped = _PWU-MapUserFromSqlRow -Row $row
                if ($mapped -and $mapped.pw_userno -gt 0) {
                    $byNumber[$mapped.pw_userno] = $mapped
                }
            }
        } catch {
            Write-Verbose "[PWU] dms_user batch query failed: $($_.Exception.Message)"
        }
    }
    return @($byNumber.Values)
}

function _PWU-ResolveUserViaCmdlet {
    param([Parameter(Mandatory)][int]$UserNumber)

    if (Get-Command -Name 'Get-PWUsersByMatch' -ErrorAction SilentlyContinue) {
        foreach ($paramName in @('UserId', 'UserID')) {
            try {
                $user = Get-PWUsersByMatch @{$paramName = $UserNumber} -ErrorAction Stop
                if ($user) {
                    if ($user -is [System.Array]) { $user = $user | Select-Object -First 1 }
                    return _PWU-MapUserFromPwObject -UserNumber $UserNumber -User $user
                }
            } catch {
                Write-Verbose "[PWU] Get-PWUsersByMatch -$paramName $UserNumber failed: $($_.Exception.Message)"
            }
        }
    }

    if (Get-Command -Name 'Get-PWUser' -ErrorAction SilentlyContinue) {
        try {
            $user = Get-PWUser -UserID $UserNumber -ErrorAction Stop
            if ($user) {
                if ($user -is [System.Array]) { $user = $user | Select-Object -First 1 }
                return _PWU-MapUserFromPwObject -UserNumber $UserNumber -User $user
            }
        } catch {
            Write-Verbose "[PWU] Get-PWUser -UserID $UserNumber failed: $($_.Exception.Message)"
        }
    }

    return $null
}

function Sync-PWUserDirectory {
    <#
    .SYNOPSIS
    Resolves pw_userno values and upserts rows into dbo.pw_users.
    .DESCRIPTION
    Uses Select-PWSQL against dms_user (primary), then Get-PWUsersByMatch / Get-PWUser for any remaining IDs.
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
        if (-not $missingRes.IsSuccess) {
            return New-QCFailureResult -Code 'PW_USER_NUMBERS_FAILED' -Message $missingRes.Message -Data @{ error = $missingRes.Data }
        }
        if ($missingRes.Data.numbers) {
            $numbers = @($numbers + @($missingRes.Data.numbers) | Select-Object -Unique)
        }
    }
    if ($numbers.Count -gt $MaxUsers) {
        $numbers = @($numbers | Select-Object -First $MaxUsers)
    }
    if ($numbers.Count -eq 0) {
        return New-QCSuccessResult -Code 'PW_USER_SYNC_NONE' -Message 'No user numbers to resolve (none in audit_events or all already in pw_users).' -Data @{ synced = 0; skipped = 0 }
    }

    Write-Verbose "Resolving $($numbers.Count) user number(s): $($numbers -join ', ')"

    $byNumber = @{}
    foreach ($mapped in @(_PWU-ResolveUsersViaSql -UserNumbers $numbers)) {
        if ($mapped.pw_userno -gt 0) { $byNumber[[int]$mapped.pw_userno] = $mapped }
    }
    $sqlResolved = $byNumber.Count

    $cmdletFailures = [System.Collections.Generic.List[string]]::new()
    foreach ($n in $numbers) {
        if ($byNumber.ContainsKey($n)) { continue }
        $mapped = _PWU-ResolveUserViaCmdlet -UserNumber $n
        if ($mapped) {
            $byNumber[$n] = $mapped
        } else {
            $cmdletFailures.Add([string]$n)
        }
    }

    $rows = @($byNumber.Values)
    if ($rows.Count -eq 0) {
        $hint = 'Ensure Select-PWSQL can read dms_user, or that Get-PWUsersByMatch is permitted for this login.'
        return New-QCSuccessResult -Code 'PW_USER_SYNC_EMPTY' -Message "No users resolved from ProjectWise for $($numbers.Count) id(s). $hint" -Data @{
            synced          = 0
            skipped         = $cmdletFailures.Count
            requested       = $numbers.Count
            sqlResolved     = $sqlResolved
            unresolvedIds   = @($cmdletFailures)
            sampleIds       = @($numbers | Select-Object -First 10)
        }
    }

    $writeRes = Write-QCPWUserDirectory -Config $Config -Users $rows
    if (-not $writeRes.IsSuccess) { return $writeRes }

    return New-QCSuccessResult -Code 'PW_USER_SYNC_OK' -Message "Synced $($rows.Count) user(s) into pw_users ($sqlResolved via dms_user SQL)." -Data @{
        synced        = $rows.Count
        skipped       = $cmdletFailures.Count
        requested     = $numbers.Count
        sqlResolved   = $sqlResolved
        unresolvedIds = @($cmdletFailures)
    }
}

Export-ModuleMember -Function Sync-PWUserDirectory
