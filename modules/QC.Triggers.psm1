# QC.Triggers.psm1
# Responsibility: Trigger rule ordering, evaluation, and job-type classification.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Core.Paths.psm1') -Force

function Get-OrderedTriggerRules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    $rules = @($Config.triggers.rules)
    $enabled = @($rules | Where-Object { [bool]$_.enabled })
    $ordered = @($enabled | Sort-Object -Property @{Expression={ [int]$_.priority }; Descending=$true}, @{Expression={ [string]$_.id }; Descending=$false})

    return New-QCSuccessResult -Code 'TRIGGER_RULES_ORDERED' -Message 'Enabled trigger rules ordered by priority.' -Data @{ rules = $ordered }
}

function Test-TriggerRule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Candidate,
        [Parameter(Mandatory)]
        [hashtable]$Rule
    )

    $pathRes = Normalize-QCPath -Path ([string]$Candidate.path)
    if (-not $pathRes.IsSuccess) { return $pathRes }
    $pathNorm = [string]$pathRes.Data.path
    $name = [string]$Candidate.fileName
    $ext = ([System.IO.Path]::GetExtension($name)).ToLowerInvariant()
    $desc = [string]$Candidate.description

    $when = $Rule.when
    if (-not $when) { $when = @{} }

    $groupResults = @{}
    $groupResults.extensions = $true
    if ($when.extensions) { $groupResults.extensions = @($when.extensions) -contains $ext }

    $groupResults.descriptionContainsAny = $true
    if ($when.descriptionContainsAny -and @($when.descriptionContainsAny).Count -gt 0) {
        $groupResults.descriptionContainsAny = $false
        foreach ($needle in @($when.descriptionContainsAny)) {
            if ($desc.IndexOf([string]$needle, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $groupResults.descriptionContainsAny = $true; break
            }
        }
    }

    $groupResults.pathRegexAny = $true
    if ($when.pathRegexAny -and @($when.pathRegexAny).Count -gt 0) {
        $groupResults.pathRegexAny = $false
        foreach ($rx in @($when.pathRegexAny)) { if ($pathNorm -match [string]$rx) { $groupResults.pathRegexAny = $true; break } }
    }

    $groupResults.fileNameRegexAny = $true
    if ($when.fileNameRegexAny -and @($when.fileNameRegexAny).Count -gt 0) {
        $groupResults.fileNameRegexAny = $false
        foreach ($rx in @($when.fileNameRegexAny)) { if ($name -match [string]$rx) { $groupResults.fileNameRegexAny = $true; break } }
    }

    $requireAll = @($Rule.requireAll)
    foreach ($key in $requireAll) {
        if (-not $groupResults.ContainsKey([string]$key) -or -not [bool]$groupResults[[string]$key]) {
            return New-QCSuccessResult -Code 'TRIGGER_NO_MATCH' -Message 'Required trigger criteria not met.' -Data @{ isMatch = $false; reason = "require_all_failed:$key"; groupResults = $groupResults; ruleId = $Rule.id }
        }
    }

    $exclude = $Rule.exclude
    if ($exclude) {
        foreach ($rx in @($exclude.pathRegexAny)) {
            if ($pathNorm -match [string]$rx) {
                return New-QCSuccessResult -Code 'TRIGGER_EXCLUDED' -Message 'Candidate excluded by path regex.' -Data @{ isMatch = $false; reason = 'exclude_path_regex'; matchedRule = [string]$rx; ruleId = $Rule.id }
            }
        }
        foreach ($rx in @($exclude.fileNameRegexAny)) {
            if ($name -match [string]$rx) {
                return New-QCSuccessResult -Code 'TRIGGER_EXCLUDED' -Message 'Candidate excluded by file name regex.' -Data @{ isMatch = $false; reason = 'exclude_name_regex'; matchedRule = [string]$rx; ruleId = $Rule.id }
            }
        }
    }

    return New-QCSuccessResult -Code 'TRIGGER_MATCH' -Message 'Candidate matched trigger rule.' -Data @{ isMatch = $true; ruleId = $Rule.id; jobType = $Rule.jobType; triggerType = $Rule.triggerType; groupResults = $groupResults }
}

function Test-QCTriggerCandidate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Candidate,
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory = $false)]
        [object[]]$OrderedRules = $null
    )

    if ($null -eq $OrderedRules) {
        $orderedRes = Get-OrderedTriggerRules -Config $Config
        if (-not $orderedRes.IsSuccess) { return $orderedRes }
        $OrderedRules = @($orderedRes.Data.rules)
    }

    foreach ($rule in @($OrderedRules)) {
        $r = Test-TriggerRule -Candidate $Candidate -Rule $rule
        if (-not $r.IsSuccess) { return $r }
        if ($r.Data.isMatch) {
            return New-QCSuccessResult -Code 'TRIGGER_MATCHED' -Message 'Candidate matched a trigger rule.' -Data @{ matched = $true; rule = $rule; evaluation = $r.Data }
        }
    }

    return New-QCSuccessResult -Code 'TRIGGER_NO_MATCH' -Message 'No enabled trigger rule matched candidate.' -Data @{ matched = $false; reason = 'no_match' }
}

function Resolve-QCTriggerMatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Candidate,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    $testRes = Test-QCTriggerCandidate -Candidate $Candidate -Config $Config
    if (-not $testRes.IsSuccess) { return $testRes }

    if (-not $testRes.Data.matched) {
        return New-QCSuccessResult -Code 'IGNORED_NO_MATCH' -Message 'Candidate ignored because no trigger matched.' -Data @{ matched = $false; action = 'ignore'; reason = 'no_match' }
    }

    $eval = $testRes.Data.evaluation
    return New-QCSuccessResult -Code 'TRIGGER_RESOLVED' -Message 'Trigger match resolved.' -Data @{
        matched = $true
        action = 'enqueue'
        ruleId = $eval.ruleId
        jobType = $eval.jobType
        triggerType = $eval.triggerType
        candidate = $Candidate
    }
}

Export-ModuleMember -Function *
