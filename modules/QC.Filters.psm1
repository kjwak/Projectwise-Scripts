# QC.Filters.psm1
# Responsibility: Whitelist/blacklist filtering decisions before job creation.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Core.Paths.psm1') -Force

function Test-QCPathAllowed {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$CandidatePath,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    $c = Normalize-QCPath -Path $CandidatePath
    if (-not $c.IsSuccess) { return $c }
    $candidate = [string]$c.Data.path

    $filters = $Config.filters
    if (-not $filters) {
        return New-QCSuccessResult -Code 'FILTERS_NOT_CONFIGURED' -Message 'No filters configured; path allowed.' -Data @{ allowed = $true; reason = 'no_filters'; matchedRule = $null; candidatePath = $candidate }
    }

    $whitelistEnabled = [bool]($filters.whitelist.enabled)
    $whitelistPaths = @($filters.whitelist.paths)
    $blacklistPaths = @($filters.blacklist.paths)
    $blacklistPatterns = @($filters.blacklist.patterns)

    function Test-BlacklistHit([string]$PathToCheck, [object[]]$Paths, [object[]]$Patterns) {
        foreach ($bp in $Paths) {
            $t = Test-PathUnderRoot -Path $PathToCheck -Root ([string]$bp)
            if ($t.IsSuccess -and $t.Data.isUnderRoot) { return @{ hit=$true; rule=[string]$bp; type='blacklist_path' } }
        }
        foreach ($rx in $Patterns) {
            if ($PathToCheck -match [string]$rx) { return @{ hit=$true; rule=[string]$rx; type='blacklist_pattern' } }
        }
        return @{ hit=$false; rule=$null; type=$null }
    }

    if ($whitelistEnabled) {
        $whitelistMatch = $null
        foreach ($wp in $whitelistPaths) {
            $t = Test-PathUnderRoot -Path $candidate -Root ([string]$wp)
            if ($t.IsSuccess -and $t.Data.isUnderRoot) { $whitelistMatch = [string]$wp; break }
        }

        if (-not $whitelistMatch) {
            return New-QCSuccessResult -Code 'FILTERED_NOT_WHITELISTED' -Message 'Path is not under whitelist.' -Data @{ allowed = $false; reason = 'not_whitelisted'; matchedRule = $null; candidatePath = $candidate }
        }

        $blackHit = Test-BlacklistHit -PathToCheck $candidate -Paths $blacklistPaths -Patterns $blacklistPatterns
        if ($blackHit.hit) {
            return New-QCSuccessResult -Code 'FILTERED_BLACKLIST_OVERRIDE' -Message 'Path matched blacklist rule; blacklist overrides whitelist.' -Data @{ allowed = $false; reason = $blackHit.type; matchedRule = $blackHit.rule; candidatePath = $candidate }
        }

        return New-QCSuccessResult -Code 'ALLOWED_WHITELIST' -Message 'Path allowed by whitelist and not blacklisted.' -Data @{ allowed = $true; reason = 'whitelist_match'; matchedRule = $whitelistMatch; candidatePath = $candidate }
    }

    $blackHit = Test-BlacklistHit -PathToCheck $candidate -Paths $blacklistPaths -Patterns $blacklistPatterns
    if ($blackHit.hit) {
        return New-QCSuccessResult -Code 'FILTERED_BLACKLIST' -Message 'Path matched blacklist rule.' -Data @{ allowed = $false; reason = $blackHit.type; matchedRule = $blackHit.rule; candidatePath = $candidate }
    }

    return New-QCSuccessResult -Code 'ALLOWED_NO_WHITELIST' -Message 'Path allowed; whitelist disabled and no blacklist match.' -Data @{ allowed = $true; reason = 'allowed_no_whitelist'; matchedRule = $null; candidatePath = $candidate }
}

Export-ModuleMember -Function *
