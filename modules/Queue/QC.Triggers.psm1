# QC.Triggers.psm1
# Responsibility: Trigger rule ordering, evaluation, and job-type classification.

Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core.Results.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core.Paths.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'QC.ProcessType.psm1') -Force -ErrorAction SilentlyContinue

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

function Test-QCIsStatusSetOutputPdfName {
    <#
    .SYNOPSIS
    True for aggregate status-set PDFs that must never receive QC_PREPEND.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$FileName)

    if ([string]::IsNullOrWhiteSpace($FileName)) { return $false }
    if ($FileName -match '(?i)^_StatusSet\.pdf$') { return $true }
    if ($FileName -match '(?i)^status_set_replace_.*\.pdf$') { return $true }
    return $false
}

function Test-QCStatusSetSourceDocument {
    <#
    .SYNOPSIS
    True when an audit document is eligible to trigger folder-level STATUS_SET_GEN.
    .DESCRIPTION
    Status-set sources are DGN sheet files and normal (non-QC) sheet PDFs under CADD\Sheets.
    QC artifacts, reference/working/template paths, and non-sheet extensions are excluded.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DocumentName,
        [Parameter(Mandatory)][string]$FolderPath
    )

    if ([string]::IsNullOrWhiteSpace($DocumentName)) { return $false }

    $normPath = ([string]$FolderPath).Trim().Replace('\', '/').TrimEnd('/').ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($normPath)) { return $false }
    if ($normPath.IndexOf('cadd/sheets', [StringComparison]::Ordinal) -lt 0) { return $false }

    $pathForFrag = $normPath + '/'
    $excludedFolderFragments = @(
        '/cadd/ref-ord/'
        '/cadd/ref-noord/'
        '/cadd/working/'
        '/templates/'
        '/template/'
        '/workspace/'
    )
    foreach ($frag in $excludedFolderFragments) {
        if ($pathForFrag.IndexOf($frag, [StringComparison]::Ordinal) -ge 0) { return $false }
    }

    $name = ([string]$DocumentName).Trim()
    if ([string]::IsNullOrWhiteSpace($name)) { return $false }
    if (Test-QCIsStatusSetOutputPdfName -FileName $name) { return $false }

    $ext = ([System.IO.Path]::GetExtension($name)).ToLowerInvariant()
    if ($ext -eq '.pdf') {
        if (Test-QCLegacyQcPdfDocumentName -DocumentName $name) { return $false }
        if ($name -match '(?i)_qc\.pdf$') { return $false }
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($name)
        if ($stem -and $stem.EndsWith('_QC', [StringComparison]::OrdinalIgnoreCase)) { return $false }
        return $true
    }
    if ($ext -eq '.dgn') { return $true }

    return $false
}

function Test-QCTriggerCandidate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Candidate,
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory = $false)]
        [object[]]$OrderedRules = $null,
        [Parameter(Mandatory = $false)]
        [string]$TriggerType = ''
    )

    if ($null -eq $OrderedRules) {
        $orderedRes = Get-OrderedTriggerRules -Config $Config
        if (-not $orderedRes.IsSuccess) { return $orderedRes }
        $OrderedRules = @($orderedRes.Data.rules)
    }

    $triggerTypeFilter = if ($TriggerType) { ([string]$TriggerType).Trim().ToLowerInvariant() } else { '' }

    foreach ($rule in @($OrderedRules)) {
        if ($triggerTypeFilter) {
            $ruleTriggerType = ''
            try {
                if ($rule -is [hashtable] -and $rule.ContainsKey('triggerType') -and $rule.triggerType) {
                    $ruleTriggerType = ([string]$rule.triggerType).Trim().ToLowerInvariant()
                } elseif ($rule.triggerType) {
                    $ruleTriggerType = ([string]$rule.triggerType).Trim().ToLowerInvariant()
                }
            } catch { $ruleTriggerType = '' }
            if ($ruleTriggerType -and ($ruleTriggerType -ne $triggerTypeFilter)) { continue }
        }
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
