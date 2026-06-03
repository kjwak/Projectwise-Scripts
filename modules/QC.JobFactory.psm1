# QC.JobFactory.psm1
# Responsibility: Build and validate queue job payloads.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Core.Runtime.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Core.Paths.psm1') -Force

function _QCJF-IsNullOrWhiteSpace([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function _QCJF-ToStableString([object]$Value) {
    if ($null -eq $Value) { return '' }
    if ($Value -is [datetime]) { return (ConvertTo-QCTimestamp $Value) }
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

function _QCJF-GetCandidateStateTransitionKey {
    param(
        [Parameter(Mandatory)][hashtable]$Candidate,
        [string]$RuleId = ''
    )
    if ($Candidate.ContainsKey('stateTransitionKey') -and -not (_QCJF-IsNullOrWhiteSpace $Candidate.stateTransitionKey)) {
        return [string]$Candidate.stateTransitionKey
    }
    if ($RuleId -notin @('qc-prepend-qc-initiated', 'qc-prepend-qc-finalizing')) { return $null }
    return $null
}

function _QCJF-GetLocalRootId([string]$Path) {
    try {
        $root = [System.IO.Path]::GetPathRoot($Path)
        if ([string]::IsNullOrWhiteSpace($root)) { return '' }
        return $root.TrimEnd('\').ToLowerInvariant()
    } catch {
        return ''
    }
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

    $stableParts = @(
        'jobIdV1'
        ('type=' + $jobType)
        ('rule=' + $ruleId)
        ('path=' + $sourcePath)
        ('name=' + $sourceName)
    )

    # For folder-level STATUS_SET_GEN, job ids must change when the folder content changes.
    # The watcher provides Candidate.folderStateHash; include it when present.
    if ($jobType -eq 'STATUS_SET_GEN' -and $Candidate.ContainsKey('folderStateHash') -and -not (_QCJF-IsNullOrWhiteSpace $Candidate.folderStateHash)) {
        $stableParts += ('folderStateHash=' + [string]$Candidate.folderStateHash)
    }
    $stateTransitionKey = _QCJF-GetCandidateStateTransitionKey -Candidate $Candidate -RuleId $ruleId
    if (-not (_QCJF-IsNullOrWhiteSpace $stateTransitionKey)) {
        $stableParts += ('stateTransition=' + $stateTransitionKey)
    }

    $stable = ($stableParts -join '|')

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
        grouping = $Rule.grouping
    }

    $sourceFolder = $null
    if ($Candidate.ContainsKey('sourceFolder') -and $Candidate.sourceFolder) {
        $sourceFolder = [string]$Candidate.sourceFolder
    } else {
        try { $sourceFolder = [System.IO.Path]::GetDirectoryName($sourcePath) } catch { $sourceFolder = $null }
    }
    if ($sourceFolder) {
        $sf = Normalize-QCPath -Path $sourceFolder
        if ($sf.IsSuccess) { $sourceFolder = [string]$sf.Data.path }
    }

    $groupKey = $null
    if ($Candidate.ContainsKey('groupKey') -and $Candidate.groupKey) {
        $groupKey = [string]$Candidate.groupKey
    }

    $job = @{
        id          = $jobId
        type        = $jobType
        sourcePath  = $sourcePath
        sourceName  = $sourceName
        sourceFolder = $sourceFolder
        groupKey    = $groupKey
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


function New-QCPackageJobDedupeKey {
    <#
    .SYNOPSIS
    Computes package-level dedupe key for ProjectWise QC package jobs.
    .DESCRIPTION
    Uses PackageId + JobType + CanonicalState + ReviewType so DGN/PDF/QC PDF audit bursts collapse into one job.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$PackageId,
        [Parameter(Mandatory)][string]$JobType,
        [string]$CanonicalState,
        [string]$ReviewType
    )
    if (_QCJF-IsNullOrWhiteSpace $PackageId) { return New-QCFailureResult -Code 'PACKAGE_DEDUPE_MISSING_PACKAGE_ID' -Message 'PackageId is required for package-level dedupe.' -Data @{} }
    if (_QCJF-IsNullOrWhiteSpace $JobType) { return New-QCFailureResult -Code 'PACKAGE_DEDUPE_MISSING_JOB_TYPE' -Message 'JobType is required for package-level dedupe.' -Data @{ packageId=$PackageId } }
    $stable = (@(
        'dedupeV1_package',
        ('packageId=' + $PackageId),
        ('jobType=' + $JobType),
        ('canonicalState=' + ($CanonicalState -as [string])),
        ('reviewType=' + ($ReviewType -as [string]))
    ) -join '|')
    $hex = _QCJF-Sha256Hex -Text $stable
    $dedupeKey = 'dq_pkg_' + (($JobType -replace '[^a-zA-Z0-9]+','').ToLowerInvariant()) + '_' + $hex.Substring(0,24)
    return New-QCSuccessResult -Code 'PACKAGE_DEDUPE_KEY_COMPUTED' -Message 'Package-level dedupe key computed.' -Data @{ dedupeKey=$dedupeKey; stableInput=$stable; hash=$hex }
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

    if ($Job.ContainsKey('metadata') -and $Job.metadata -is [hashtable] -and $Job.metadata.ContainsKey('package') -and $Job.metadata.package -is [hashtable]) {
        $pkg = $Job.metadata.package
        $pkgId = [string]$pkg.PackageId
        if (_QCJF-IsNullOrWhiteSpace $pkgId -and $pkg.ContainsKey('packageId')) { $pkgId = [string]$pkg.packageId }
        $canonicalState = [string]$pkg.CanonicalState
        if (_QCJF-IsNullOrWhiteSpace $canonicalState -and $pkg.ContainsKey('WorkflowState')) { $canonicalState = [string]$pkg.WorkflowState }
        $reviewType = [string]$pkg.ReviewType
        $pkgKey = New-QCPackageJobDedupeKey -PackageId $pkgId -JobType $type -CanonicalState $canonicalState -ReviewType $reviewType
        if ($pkgKey.IsSuccess) { return $pkgKey }
    }

    $ruleId = $null
    if ($Job.ContainsKey('triggerRule') -and $Job.triggerRule -is [hashtable]) {
        $ruleId = [string]$Job.triggerRule.id
    }
    if (_QCJF-IsNullOrWhiteSpace $ruleId) {
        return New-QCFailureResult -Code 'JOB_VALIDATION_MISSING_TRIGGER_RULE' -Message 'Job.triggerRule.id is required to compute dedupe key.' -Data @{ job = $Job }
    }

    $groupingEnabled = $false
    $groupBy = $null
    if ($Job.ContainsKey('triggerRule') -and $Job.triggerRule -is [hashtable] -and $Job.triggerRule.ContainsKey('grouping') -and $Job.triggerRule.grouping) {
        $g = $Job.triggerRule.grouping
        try { $groupingEnabled = [bool]$g.enabled } catch { $groupingEnabled = $false }
        if ($g.ContainsKey('groupBy')) { $groupBy = [string]$g.groupBy }
    }
    if ($groupBy) { $groupBy = $groupBy.Trim().ToLowerInvariant() }

    if ($groupingEnabled -and $groupBy -eq 'folder' -and $type -eq 'STATUS_SET_GEN') {
        $sourceFolder = $null
        if ($Job.ContainsKey('sourceFolder') -and $Job.sourceFolder) { $sourceFolder = [string]$Job.sourceFolder }
        if (_QCJF-IsNullOrWhiteSpace $sourceFolder) {
            try { $sourceFolder = [System.IO.Path]::GetDirectoryName($path) } catch { $sourceFolder = $null }
        }
        if (_QCJF-IsNullOrWhiteSpace $sourceFolder) {
            return New-QCFailureResult -Code 'JOB_VALIDATION_MISSING_SOURCE_FOLDER' -Message 'Job.sourceFolder is required for grouped folder dedupe.' -Data @{ job = $Job }
        }

        $rootId = ''
        if ($Job.ContainsKey('datasource') -and $Job.datasource) { $rootId = [string]$Job.datasource }
        if (_QCJF-IsNullOrWhiteSpace $rootId -and $Job.ContainsKey('metadata') -and $Job.metadata -and $Job.metadata.ContainsKey('candidate') -and $Job.metadata.candidate) {
            $cand = $Job.metadata.candidate
            if ($cand -is [hashtable] -and $cand.ContainsKey('datasourceName') -and $cand.datasourceName) { $rootId = [string]$cand.datasourceName }
        }
        if (_QCJF-IsNullOrWhiteSpace $rootId) {
            $rootId = _QCJF-GetLocalRootId -Path $path
        }

        $stableParts = @(
            'dedupeV2_group_folder'
            ('type=' + $type)
            ('root=' + $rootId)
            ('folder=' + $sourceFolder)
        )

        $folderStateHash = $null
        if ($Job.ContainsKey('metadata') -and $Job.metadata -and $Job.metadata.ContainsKey('candidate') -and $Job.metadata.candidate) {
            $cand = $Job.metadata.candidate
            if ($cand -is [hashtable] -and $cand.ContainsKey('folderStateHash') -and $cand.folderStateHash) {
                $folderStateHash = [string]$cand.folderStateHash
            }
        }
        if (-not (_QCJF-IsNullOrWhiteSpace $folderStateHash)) {
            $stableParts += ('folderStateHash=' + $folderStateHash)
        }

        $stable = ($stableParts -join '|')
    } else {
        # File-level identity:
        # - QC_PREPEND: include file hash when available.
        # - default: type + rule + path.
        $fileHash = $null
        if ($Job.ContainsKey('metadata') -and $Job.metadata -and $Job.metadata.ContainsKey('fileHash') -and $Job.metadata.fileHash) {
            $fileHash = [string]$Job.metadata.fileHash
        } elseif ($Job.ContainsKey('metadata') -and $Job.metadata -and $Job.metadata.ContainsKey('candidate') -and $Job.metadata.candidate) {
            $cand = $Job.metadata.candidate
            if ($cand -is [hashtable] -and $cand.ContainsKey('file') -and $cand.file -is [hashtable] -and $cand.file.ContainsKey('sha256') -and $cand.file.sha256) {
                $fileHash = [string]$cand.file.sha256
            }
        }

        $stableParts = @(
            'dedupeV2_file'
            ('type=' + $type)
            ('rule=' + $ruleId)
            ('path=' + $path)
        )
        if ($type -in @('QC_PREPEND', 'QC_COMMENT_STATUS_SYNC') -and -not (_QCJF-IsNullOrWhiteSpace $fileHash)) {
            $stableParts += ('fileHash=' + $fileHash)
        }
        $stateTransitionKey = $null
        if ($Job.ContainsKey('metadata') -and $Job.metadata -is [hashtable] -and $Job.metadata.ContainsKey('candidate') -and $Job.metadata.candidate -is [hashtable]) {
            $stateTransitionKey = _QCJF-GetCandidateStateTransitionKey -Candidate $Job.metadata.candidate -RuleId $ruleId
        }
        if (-not (_QCJF-IsNullOrWhiteSpace $stateTransitionKey)) {
            $stableParts += ('stateTransition=' + $stateTransitionKey)
        }
        $stable = ($stableParts -join '|')
    }

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
