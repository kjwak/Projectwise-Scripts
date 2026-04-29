# QC.JobFactory.psm1
# Responsibility: Build and validate queue job payloads.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Core.Paths.psm1') -Force

function _QCJF-IsNullOrWhiteSpace([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function _QCJF-ToStableString([object]$Value) {
    if ($null -eq $Value) { return '' }
    if ($Value -is [datetime]) { return ($Value.ToUniversalTime().ToString('o')) }
    return [string]$Value
}

function _QCJF-Sha256Hex([string]$Text) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = $sha.ComputeHash($bytes)
    } finally {
        $sha.Dispose()
    }
    return -join ($hash | ForEach-Object { $_.ToString('x2') })
}

function New-QCJobId {
    <#
    .SYNOPSIS
    Generates a deterministic or policy-compliant job identifier.
    .DESCRIPTION
    Produces a job ID using candidate/rule/config inputs and naming conventions.
    .PARAMETER Candidate
    Candidate metadata object.
    .PARAMETER Rule
    Trigger rule or match object.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: none.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Candidate,
        [Parameter(Mandatory)]
        [hashtable]$Rule,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    $jobType = [string]$Rule.jobType
    if (_QCJF-IsNullOrWhiteSpace $jobType) {
        return New-QCFailureResult -Code 'JOB_VALIDATION_MISSING_TYPE' -Message 'Rule.jobType is required to generate a job id.' -Data @{ rule = $Rule }
    }

    $pathRes = Normalize-QCPath -Path ([string]$Candidate.path)
    if (-not $pathRes.IsSuccess) { return $pathRes }
    $sourcePath = [string]$pathRes.Data.path
    if (_QCJF-IsNullOrWhiteSpace $sourcePath) {
        return New-QCFailureResult -Code 'JOB_VALIDATION_MISSING_SOURCE_PATH' -Message 'Candidate.path is required to generate a job id.' -Data @{ candidate = $Candidate }
    }

    $ruleId = [string]$Rule.id
    if (_QCJF-IsNullOrWhiteSpace $ruleId) {
        return New-QCFailureResult -Code 'JOB_VALIDATION_MISSING_TRIGGER_RULE' -Message 'Rule.id is required to generate a job id.' -Data @{ rule = $Rule }
    }

    $policy = 'deterministic'
    if ($Config.ContainsKey('jobFactory') -and $Config.jobFactory -and $Config.jobFactory.ContainsKey('idPolicy')) {
        $policy = [string]$Config.jobFactory.idPolicy
    } elseif ($Config.ContainsKey('jobs') -and $Config.jobs -and $Config.jobs.ContainsKey('idPolicy')) {
        $policy = [string]$Config.jobs.idPolicy
    }
    if (_QCJF-IsNullOrWhiteSpace $policy) { $policy = 'deterministic' }

    if ($policy -ieq 'guid') {
        $id = ([guid]::NewGuid().ToString('N'))
        return New-QCSuccessResult -Code 'JOB_ID_GUID' -Message 'Non-deterministic GUID job id generated per configured policy.' -Data @{ id = $id; policy = 'guid' }
    }

    $sourceName = [string]$Candidate.fileName
    if (_QCJF-IsNullOrWhiteSpace $sourceName) { $sourceName = [System.IO.Path]::GetFileName($sourcePath) }

    $stable = @(
        'jobIdV1'
        ('type=' + $jobType)
        ('rule=' + $ruleId)
        ('path=' + $sourcePath)
        ('name=' + $sourceName)
    ) -join '|'

    $hex = _QCJF-Sha256Hex -Text $stable
    $short = $hex.Substring(0, 16)
    $jobId = ("qc_" + ($jobType -replace '[^a-zA-Z0-9]+', '').ToLowerInvariant() + "_" + $short)

    return New-QCSuccessResult -Code 'JOB_ID_DETERMINISTIC' -Message 'Deterministic job id generated.' -Data @{
        id = $jobId
        policy = 'deterministic'
        stableInput = $stable
        hash = $hex
    }
}

function New-QCJobObject {
    <#
    .SYNOPSIS
    Creates a standardized queue job object.
    .DESCRIPTION
    Builds a job payload from candidate metadata, trigger match, and app settings.
    .PARAMETER Candidate
    Candidate metadata object.
    .PARAMETER Rule
    Trigger rule or match object.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: none.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Candidate,
        [Parameter(Mandatory)]
        [hashtable]$Rule,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    $jobType = [string]$Rule.jobType
    if (_QCJF-IsNullOrWhiteSpace $jobType) {
        return New-QCFailureResult -Code 'JOB_VALIDATION_MISSING_TYPE' -Message 'Job type is required (Rule.jobType).' -Data @{ rule = $Rule }
    }

    $pathRes = Normalize-QCPath -Path ([string]$Candidate.path)
    if (-not $pathRes.IsSuccess) { return $pathRes }
    $sourcePath = [string]$pathRes.Data.path
    if (_QCJF-IsNullOrWhiteSpace $sourcePath) {
        return New-QCFailureResult -Code 'JOB_VALIDATION_MISSING_SOURCE_PATH' -Message 'SourcePath is required (Candidate.path).' -Data @{ candidate = $Candidate }
    }

    $ruleId = [string]$Rule.id
    if (_QCJF-IsNullOrWhiteSpace $ruleId) {
        return New-QCFailureResult -Code 'JOB_VALIDATION_MISSING_TRIGGER_RULE' -Message 'TriggerRule is required (Rule.id).' -Data @{ rule = $Rule }
    }

    $sourceName = [string]$Candidate.fileName
    if (_QCJF-IsNullOrWhiteSpace $sourceName) { $sourceName = [System.IO.Path]::GetFileName($sourcePath) }

    $createdAt = $null
    if ($Candidate.ContainsKey('createdAtUtc') -and $Candidate.createdAtUtc) {
        $createdAt = [string]$Candidate.createdAtUtc
    } elseif ($Candidate.ContainsKey('detectedAtUtc') -and $Candidate.detectedAtUtc) {
        $createdAt = [string]$Candidate.detectedAtUtc
    } elseif ($Config.ContainsKey('jobFactory') -and $Config.jobFactory -and $Config.jobFactory.ContainsKey('createdAtUtc') -and $Config.jobFactory.createdAtUtc) {
        $createdAt = [string]$Config.jobFactory.createdAtUtc
    }
    if (_QCJF-IsNullOrWhiteSpace $createdAt) {
        # Keep pure/deterministic: no clock calls. Use a stable default and mark it.
        $createdAt = '2000-01-01T00:00:00.0000000Z'
        $createdAtDefaulted = $true
    } else {
        $createdAtDefaulted = $false
    }

    $jobIdRes = New-QCJobId -Candidate $Candidate -Rule $Rule -Config $Config
    if (-not $jobIdRes.IsSuccess) { return $jobIdRes }
    $jobId = [string]$jobIdRes.Data.id

    $triggerRule = @{
        id = $ruleId
        jobType = $jobType
        triggerType = [string]$Rule.triggerType
    }

    $job = @{
        id          = $jobId
        type        = $jobType
        sourcePath  = $sourcePath
        sourceName  = $sourceName
        triggerRule = $triggerRule
        dedupeKey   = $null
        status      = 'queued'
        createdAt   = $createdAt
        attempts    = 0
        metadata    = @{
            candidate = $Candidate
            rule = $Rule
            createdAtDefaulted = $createdAtDefaulted
            jobIdPolicy = [string]$jobIdRes.Data.policy
        }
    }

    $dedupeRes = Get-QCDedupeKey -Job $job -Config $Config
    if (-not $dedupeRes.IsSuccess) { return $dedupeRes }
    $job.dedupeKey = [string]$dedupeRes.Data.dedupeKey

    $req = Test-QCJobRequiredFields -Job $job
    if (-not $req.IsSuccess) { return $req }

    return New-QCSuccessResult -Code 'JOB_OBJECT_CREATED' -Message 'Job object created.' -Data @{ job = $job }
}

function Test-QCJobRequiredFields {
    <#
    .SYNOPSIS
    Validates required fields in a job payload.
    .DESCRIPTION
    Confirms required keys/values exist before enqueue operations.
    .PARAMETER Job
    Job payload to validate.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: none.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job
    )

    $missing = @()
    foreach ($k in @('type', 'sourcePath', 'triggerRule')) {
        if (-not $Job.ContainsKey($k) -or $(_QCJF-IsNullOrWhiteSpace $Job[$k])) {
            $missing += $k
        }
    }

    if ($missing.Count -gt 0) {
        return New-QCFailureResult -Code 'JOB_VALIDATION_MISSING_REQUIRED_FIELDS' -Message 'Job is missing required fields.' -Data @{ missing = $missing; job = $Job }
    }

    if (-not ($Job.triggerRule -is [hashtable]) -or $(_QCJF-IsNullOrWhiteSpace $Job.triggerRule.id)) {
        return New-QCFailureResult -Code 'JOB_VALIDATION_MISSING_TRIGGER_RULE' -Message 'Job.triggerRule.id is required.' -Data @{ job = $Job }
    }

    if (-not $Job.ContainsKey('dedupeKey') -or $(_QCJF-IsNullOrWhiteSpace $Job.dedupeKey)) {
        return New-QCFailureResult -Code 'JOB_VALIDATION_MISSING_DEDUPE_KEY' -Message 'Job.dedupeKey is required.' -Data @{ job = $Job }
    }

    return New-QCSuccessResult -Code 'JOB_REQUIRED_FIELDS_OK' -Message 'Job required fields validated.' -Data @{ ok = $true }
}

function Get-QCDedupeKey {
    <#
    .SYNOPSIS
    Computes dedupe key for a job.
    .DESCRIPTION
    Produces a deduplication key based on configured key parts.
    .PARAMETER Job
    Job payload.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: none.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    $type = [string]$Job.type
    if (_QCJF-IsNullOrWhiteSpace $type) {
        return New-QCFailureResult -Code 'JOB_VALIDATION_MISSING_TYPE' -Message 'Job.type is required to compute dedupe key.' -Data @{ job = $Job }
    }

    $pathRes = Normalize-QCPath -Path ([string]$Job.sourcePath)
    if (-not $pathRes.IsSuccess) { return $pathRes }
    $path = [string]$pathRes.Data.path
    if (_QCJF-IsNullOrWhiteSpace $path) {
        return New-QCFailureResult -Code 'JOB_VALIDATION_MISSING_SOURCE_PATH' -Message 'Job.sourcePath is required to compute dedupe key.' -Data @{ job = $Job }
    }

    $ruleId = $null
    if ($Job.ContainsKey('triggerRule') -and $Job.triggerRule -is [hashtable]) {
        $ruleId = [string]$Job.triggerRule.id
    }
    if (_QCJF-IsNullOrWhiteSpace $ruleId) {
        return New-QCFailureResult -Code 'JOB_VALIDATION_MISSING_TRIGGER_RULE' -Message 'Job.triggerRule.id is required to compute dedupe key.' -Data @{ job = $Job }
    }

    # Deterministic "same job" identity: type + normalized source path + trigger rule id.
    $stable = @(
        'dedupeV1'
        ('type=' + $type)
        ('rule=' + $ruleId)
        ('path=' + $path)
    ) -join '|'

    $hex = _QCJF-Sha256Hex -Text $stable
    $short = $hex.Substring(0, 24)
    $dedupeKey = ("dq_" + ($type -replace '[^a-zA-Z0-9]+', '').ToLowerInvariant() + "_" + $short)

    return New-QCSuccessResult -Code 'DEDUPE_KEY_COMPUTED' -Message 'Dedupe key computed.' -Data @{
        dedupeKey = $dedupeKey
        stableInput = $stable
        hash = $hex
    }
}
Export-ModuleMember -Function *
