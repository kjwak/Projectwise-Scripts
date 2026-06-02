# QC.Processors.psm1
# Responsibility: Processor readiness checks and job-type-based dispatch.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Core.Runtime.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'QC.Workflow.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path $PSScriptRoot 'QC.Reporting.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path $PSScriptRoot 'QC.CommentStatusProcessor.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path $PSScriptRoot 'QC.ReviewStamp.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path $PSScriptRoot 'QC.CommentSync.Job.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path $PSScriptRoot 'QC.Rendition.psm1') -Force -ErrorAction SilentlyContinue

# Per-process throttle (milliseconds) inserted between PDF/cache file ops
# (Move/Remove/Copy) so AV scanners (Fortinet, etc.) don't flag rapid temp
# churn. Default 2000 ms; override via top-level Config.fileOpThrottleMs.
# Set to 0 to disable.
$script:_QCP_FsThrottleMs = 2000

function _QCP-FsThrottle {
    if ($script:_QCP_FsThrottleMs -and $script:_QCP_FsThrottleMs -gt 0) {
        Start-Sleep -Milliseconds $script:_QCP_FsThrottleMs
    }
}

function _QCP-ApplyFsThrottleConfig([hashtable]$Config) {
    if (-not $Config) { return }
    try {
        if ($Config.ContainsKey('fileOpThrottleMs') -and $null -ne $Config['fileOpThrottleMs']) {
            $script:_QCP_FsThrottleMs = [int]$Config['fileOpThrottleMs']
        }
    } catch { }
}

function _QCP-IsNullOrWhiteSpace([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function _QCP-ToHashtable([object]$Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [hashtable]) { return $Value }
    if ($Value -is [System.Collections.IDictionary]) { return $Value }
    if ($Value -is [string] -or $Value -is [System.ValueType]) { return @{ value = $Value } }
    if ($Value.PSObject -and $Value.PSObject.Properties) {
        $h = @{}
        foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = $p.Value }
        return $h
    }
    return $null
}

function _QCP-TruncateTextForStorage {
    param(
        [string]$Text,
        [int]$MaxLen = 8192
    )
    if ([string]::IsNullOrEmpty($Text)) { return $null }
    if ($Text.Length -le $MaxLen) { return $Text }
    $omitted = $Text.Length - $MaxLen
    return ('...(truncated ' + $omitted + ' chars)...' + $Text.Substring($Text.Length - $MaxLen))
}

function _QCP-SanitizeProcessorDataForStorage([object]$Data) {
    $h = _QCP-ToHashtable $Data
    if (-not $h) { return $null }
    $out = @{}
    foreach ($k in $h.Keys) {
        $v = $h[$k]
        if ($k -in @('stdout', 'stderr', 'argLine')) {
            if ($null -ne $v) { $out[$k] = _QCP-TruncateTextForStorage -Text ([string]$v) }
            continue
        }
        if ($k -eq 'args' -and $v -is [array]) {
            $out[$k] = @($v | ForEach-Object { [string]$_ })
            continue
        }
        $out[$k] = $v
    }
    return $out
}

function _QCP-GetRepoRoot() {
    return (Split-Path -Parent $PSScriptRoot)
}

function _QCP-GetReviewStampRoleFieldsFromJob {
    param(
        [Parameter(Mandatory)][hashtable]$Job,
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$FolderPath = '',
        [string]$SourceDocumentName = ''
    )

    $roles = @{
        designerEmail = _QCP-GetJobMetadataValue -Job $Job -Keys @('designerEmail', 'qcDesignerEmail')
        reviewerEmail = _QCP-GetJobMetadataValue -Job $Job -Keys @('reviewerEmail', 'qcReviewerEmail')
        checkerEmail = _QCP-GetJobMetadataValue -Job $Job -Keys @('checkerEmail', 'qcCheckerEmail')
        qcReviewType = _QCP-GetJobMetadataValue -Job $Job -Keys @('reviewType', 'qcReviewType')
    }

    if (-not (_QCP-IsNullOrWhiteSpace $FolderPath) -and -not (_QCP-IsNullOrWhiteSpace $SourceDocumentName)) {
        if (Get-Command -Name 'Get-PWQcPrependRoleFieldsFromSourcePdf' -ErrorAction SilentlyContinue) {
            $pw = Get-PWQcPrependRoleFieldsFromSourcePdf -FolderPath $FolderPath -SourceDocumentName $SourceDocumentName -Config $Config
            if ($pw.found) {
                $roles.designerEmail = [string]$pw.designerEmail
                $roles.reviewerEmail = [string]$pw.reviewerEmail
                $roles.checkerEmail = [string]$pw.checkerEmail
                $roles.qcReviewType = [string]$pw.qcReviewType
            }
        }
    }

    return $roles
}

function _QCP-ReviewStampRequiredForReviewType {
    param(
        [hashtable]$StampSettings,
        [string]$ReviewType
    )

    if (-not $StampSettings -or (_QCP-IsNullOrWhiteSpace $ReviewType)) { return $false }
    foreach ($p in @($StampSettings.profiles)) {
        if ([string]$p.reviewType -eq [string]$ReviewType) { return $true }
    }
    return $false
}

function _QCP-TryApplyReviewStampFromJob {
    param(
        [Parameter(Mandatory)][string]$PdfPath,
        [Parameter(Mandatory)][hashtable]$Job,
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$FolderPath = '',
        [string]$SourceDocumentName = '',
        [string]$OverlayExe = ''
    )

    if (-not (Get-Command -Name 'Invoke-QCReviewStampForReviewType' -ErrorAction SilentlyContinue)) {
        return @{ applied = $false; reason = 'QC.ReviewStamp module not loaded' }
    }

    $roles = _QCP-GetReviewStampRoleFieldsFromJob -Job $Job -Config $Config -FolderPath $FolderPath -SourceDocumentName $SourceDocumentName

    if (_QCP-IsNullOrWhiteSpace $OverlayExe) {
        $qc = _QCP-ToHashtable $Config.qcPrepend
        if ($qc -and $qc.overlayExePath) { $OverlayExe = [string]$qc.overlayExePath }
    }

    return Invoke-QCReviewStampForReviewType -PdfPath $PdfPath -Config $Config -RoleFields $roles -OverlayExe $OverlayExe
}

function _QCP-TryApplyPeerReviewStampFromJob {
    param(
        [Parameter(Mandatory)][string]$PdfPath,
        [Parameter(Mandatory)][hashtable]$Job,
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$FolderPath = '',
        [string]$SourceDocumentName = '',
        [string]$OverlayExe = ''
    )

    return _QCP-TryApplyReviewStampFromJob -PdfPath $PdfPath -Job $Job -Config $Config `
        -FolderPath $FolderPath -SourceDocumentName $SourceDocumentName -OverlayExe $OverlayExe
}

function _QCP-EnsureDir([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function _QCP-Sha256Hex([string]$Text) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = $sha.ComputeHash($bytes)
    } finally {
        $sha.Dispose()
    }
    return -join ($hash | ForEach-Object { $_.ToString('x2') })
}

function _QCP-GetJobMetadataValue([hashtable]$Job, [string[]]$Keys) {
    foreach ($k in $Keys) {
        if ($Job.ContainsKey($k) -and $null -ne $Job[$k]) { return $Job[$k] }
    }
    try {
        if ($Job.ContainsKey('metadata') -and $Job.metadata) {
            $md = _QCP-ToHashtable $Job.metadata
            if ($md) {
                foreach ($k in $Keys) {
                    if ($md.ContainsKey($k) -and $null -ne $md[$k]) { return $md[$k] }
                }
                if ($md.ContainsKey('candidate') -and $md.candidate) {
                    $cand = _QCP-ToHashtable $md.candidate
                    if ($cand) {
                        foreach ($k in $Keys) {
                            if ($cand.ContainsKey($k) -and $null -ne $cand[$k]) { return $cand[$k] }
                        }
                    }
                }
            }
        }
    } catch { }
    return $null
}

function _QCP-ResolvePrependTrigger([hashtable]$Job) {
    $t = _QCP-GetJobMetadataValue -Job $Job -Keys @('prependTrigger','qcPrependTrigger','workflowPrependTrigger')
    if (-not (_QCP-IsNullOrWhiteSpace $t)) { return ([string]$t).Trim() }

    $corr = _QCP-GetJobMetadataValue -Job $Job -Keys @('correctionComplete','designerCorrectionComplete','qcCorrectionComplete')
    if (-not (_QCP-IsNullOrWhiteSpace $corr)) {
        $cv = ([string]$corr).Trim().ToLowerInvariant()
        if ($cv -in @('true','1','yes','y')) { return 'designerCorrectionComplete' }
    }

    $redline = _QCP-GetJobMetadataValue -Job $Job -Keys @('reviewerRedlineUpdate','redlineUpdate','qcRedlineUpdate')
    if (-not (_QCP-IsNullOrWhiteSpace $redline)) {
        $rv = ([string]$redline).Trim().ToLowerInvariant()
        if ($rv -in @('true','1','yes','y')) { return 'reviewerRedlineUpdate' }
    }

    try {
        if ($Job -and $Job.ContainsKey('metadata') -and $Job.metadata) {
            $md = _QCP-ToHashtable $Job.metadata
            if ($md -and $md.ContainsKey('candidate') -and $md.candidate) {
                $cand = _QCP-ToHashtable $md.candidate
                if ($cand) {
                    foreach ($dk in @('description','documentDescription','desc')) {
                        if ($cand.ContainsKey($dk) -and $cand[$dk] -and ([string]$cand[$dk]).IndexOf('QC_Archivist', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                            return 'initialQcPdf'
                        }
                    }
                }
            }
        }
    } catch { }

    return $null
}

function _QCP-NewWorkflowContext([hashtable]$Job, [hashtable]$Config, [string]$SourcePath, [string]$OutputPath, [string]$HistoryPath, [string]$ResultStatus, [string]$ErrorMessage, [object]$Document) {
    $now = Get-QCTimestamp
    $designerEmail = _QCP-GetJobMetadataValue -Job $Job -Keys @('designerEmail','qcDesignerEmail')
    $reviewerEmail = _QCP-GetJobMetadataValue -Job $Job -Keys @('reviewerEmail','qcReviewerEmail','reviewer','qcReviewer')
    $checkerEmail = _QCP-GetJobMetadataValue -Job $Job -Keys @('checkerEmail','qcCheckerEmail')
    $reviewType = _QCP-GetJobMetadataValue -Job $Job -Keys @('reviewType','qcReviewType')
    if (_QCP-IsNullOrWhiteSpace $reviewType) {
        $pwFolder = _QCP-GetJobMetadataValue -Job $Job -Keys @('folderPath','sourceFolder','incomingFolderPath')
        $pwDoc = _QCP-GetJobMetadataValue -Job $Job -Keys @('sourceName','incomingDocName','sourceDocumentName')
        if (_QCP-IsNullOrWhiteSpace $pwDoc) { $pwDoc = _QCP-GetJobMetadataValue -Job $Job -Keys @('sourcePath') }
        if (-not (_QCP-IsNullOrWhiteSpace $pwDoc) -and $pwDoc -match '\\') { $pwDoc = [System.IO.Path]::GetFileName([string]$pwDoc) }
        $qcRtEnabled = $true
        if (-not (_QCP-IsNullOrWhiteSpace $pwFolder) -and (Get-Command -Name 'Test-PWQcReviewTypeAttributesEnabled' -ErrorAction SilentlyContinue)) {
            $qcRtEnabled = Test-PWQcReviewTypeAttributesEnabled -Config $Config -FolderPath ([string]$pwFolder)
        }
        if ($qcRtEnabled -and -not (_QCP-IsNullOrWhiteSpace $pwFolder) -and -not (_QCP-IsNullOrWhiteSpace $pwDoc)) {
            if (Get-Command -Name 'Get-PWQcPrependRoleFieldsFromSourcePdf' -ErrorAction SilentlyContinue) {
                $pwRoles = Get-PWQcPrependRoleFieldsFromSourcePdf -FolderPath ([string]$pwFolder) -SourceDocumentName ([string]$pwDoc) -Config $Config
                if ($pwRoles.found -and -not (_QCP-IsNullOrWhiteSpace $pwRoles.qcReviewType)) {
                    $reviewType = [string]$pwRoles.qcReviewType
                }
            }
        }
    }
    $lastActionBy = _QCP-GetJobMetadataValue -Job $Job -Keys @('lastActionBy','userName','triggeredBy')
    $cycleId = _QCP-GetJobMetadataValue -Job $Job -Keys @('cycleId','qcCycleId')
    $cycleNumber = _QCP-GetJobMetadataValue -Job $Job -Keys @('cycleNumber','qcCycleNumber')
    if (_QCP-IsNullOrWhiteSpace $cycleId -and $Job.ContainsKey('id')) { $cycleId = [string]$Job.id }

    $lifecycleState = $null
    if ($Document) {
        try {
            if ($Document.PSObject.Properties['StateName'] -and $Document.StateName) { $lifecycleState = [string]$Document.StateName }
            elseif ($Document.PSObject.Properties['WorkflowState'] -and $Document.WorkflowState) { $lifecycleState = [string]$Document.WorkflowState }
        } catch { }
    }
    if (_QCP-IsNullOrWhiteSpace $lifecycleState) {
        $lifecycleState = _QCP-GetJobMetadataValue -Job $Job -Keys @('pwStateName','stateName','workflowState','currentState')
    }

    $settings = $null
    $assignedTo = $null
    $statusValue = $lifecycleState
    try {
        $settings = Get-QCWorkflowSettings -Config $Config
        if (_QCP-IsNullOrWhiteSpace $reviewType) {
            $pwFolderForRt = _QCP-GetJobMetadataValue -Job $Job -Keys @('folderPath','sourceFolder','incomingFolderPath')
            $applyDefaultRt = $true
            if (-not (_QCP-IsNullOrWhiteSpace $pwFolderForRt) -and (Get-Command -Name 'Test-PWQcReviewTypeAttributesEnabled' -ErrorAction SilentlyContinue)) {
                $applyDefaultRt = Test-PWQcReviewTypeAttributesEnabled -Config $Config -FolderPath ([string]$pwFolderForRt)
            }
            if ($applyDefaultRt) { $reviewType = [string]$settings.defaultReviewType }
        }
        $assignedTo = Resolve-QCWorkflowAssignee -Settings $settings -StateName $lifecycleState -ReviewType $reviewType `
            -DesignerEmail $designerEmail -ReviewerEmail $reviewerEmail -CheckerEmail $checkerEmail
        if (_QCP-IsNullOrWhiteSpace $assignedTo) {
            $assignedTo = _QCP-GetJobMetadataValue -Job $Job -Keys @('assignedTo','qcAssignedTo')
        }
        if (_QCP-IsNullOrWhiteSpace $statusValue -and -not (_QCP-IsNullOrWhiteSpace $lifecycleState)) { $statusValue = $lifecycleState }
        elseif (_QCP-IsNullOrWhiteSpace $statusValue) { $statusValue = $ResultStatus }
    } catch {
        if (_QCP-IsNullOrWhiteSpace $statusValue) { $statusValue = $ResultStatus }
        $assignedTo = _QCP-GetJobMetadataValue -Job $Job -Keys @('assignedTo','qcAssignedTo')
    }

    $docPath = if (-not (_QCP-IsNullOrWhiteSpace $SourcePath)) { $SourcePath } else { [string](_QCP-GetJobMetadataValue -Job $Job -Keys @('sourceDocumentPath','sourcePath')) }
    return @{
        jobId = [string]$Job.id
        jobType = [string]$Job.type
        resultStatus = $ResultStatus
        lifecycleState = $lifecycleState
        documentPath = $docPath
        sourcePath = $SourcePath
        outputPath = $OutputPath
        historyPath = $HistoryPath
        timestamp = $now
        attributes = @{
            qcActive = $true
            reviewType = $reviewType
            cycleId = $cycleId
            cycleNumber = $cycleNumber
            designerEmail = $designerEmail
            reviewerEmail = $reviewerEmail
            checkerEmail = $checkerEmail
            assignedTo = $assignedTo
            lastActionBy = $lastActionBy
            lastActionDate = $now
            status = $statusValue
            historyPdfPath = $HistoryPath
            latestOverlayPdfPath = $OutputPath
            sourceDocumentPath = $docPath
            automationLastRun = $now
            automationResult = $ResultStatus
            automationError = $ErrorMessage
        }
    }
}

function _QCP-ResolvePwDocumentForJob([hashtable]$Job, [hashtable]$Config) {
    if (-not $Job) { return $null }
    if ($Job.ContainsKey('document') -and $Job.document) { return $Job.document }

    if (Get-Command -Name 'Get-QCCommentSyncPwDocument' -ErrorAction SilentlyContinue) {
        try {
            $res = Get-QCCommentSyncPwDocument -Job $Job -Config $Config
            if ($res.IsSuccess -and $res.Data -and $res.Data.document) {
                $Job['document'] = $res.Data.document
                return $res.Data.document
            }
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Level 'Warning' -Code 'QC_PREPEND_PW_DOC_RESOLVE_FAILED' `
                    -Message 'Could not resolve ProjectWise document for workflow writeback.' -Data @{
                    jobId = if ($Job.id) { [string]$Job.id } else { $null }
                    code = [string]$res.Code
                    message = [string]$res.Message
                    sourceFolder = if ($Job.sourceFolder) { [string]$Job.sourceFolder } else { $null }
                    sourceName = if ($Job.sourceName) { [string]$Job.sourceName } else { $null }
                } | Out-Null
            }
        } catch {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Level 'Warning' -Code 'QC_PREPEND_PW_DOC_RESOLVE_FAILED' `
                    -Message 'ProjectWise document resolution threw.' -Data @{
                    jobId = if ($Job.id) { [string]$Job.id } else { $null }
                    errorMessage = [string]$_.Exception.Message
                } | Out-Null
            }
        }
    }
    return $null
}

function _QCP-LogPrependWorkflowWriteback([hashtable]$Job, [hashtable]$Config, [object]$Writeback, [string]$PrependTrigger) {
    if (-not (Get-Command -Name Write-QCJsonLog -ErrorAction SilentlyContinue)) { return }

    $jobId = $null
    if ($Job) {
        if ($Job.ContainsKey('id') -and $Job.id) { $jobId = [string]$Job.id }
        elseif ($Job.ContainsKey('jobId') -and $Job.jobId) { $jobId = [string]$Job.jobId }
    }

    $settings = $null
    try { $settings = Get-QCWorkflowSettings -Config $Config } catch { }

    $targetState = $null
    $previousState = $null
    $stateCode = $null
    $planned = $false
    $changed = $false
    $dryRun = $false
    $warnings = @()

    $transitionDetail = $null
    if ($Writeback -and $Writeback.Data) {
        $wbData = _QCP-ToHashtable $Writeback.Data
        if ($wbData) {
            if ($wbData.ContainsKey('dryRun')) { try { $dryRun = [bool]$wbData.dryRun } catch { } }
            if ($wbData.warnings) { $warnings = @($wbData.warnings) }
            if ($wbData.actions) {
                foreach ($a in @($wbData.actions)) {
                    if (-not $a) { continue }
                    $ac = [string]$a.Code
                    if ($ac -notmatch '^QC_WORKFLOW_STATE') { continue }
                    $stateCode = $ac
                    $ad = _QCP-ToHashtable $a.Data
                    if ($ad) {
                        if ($ad.stateName) { $targetState = [string]$ad.stateName }
                        if ($ad.currentStateName) { $previousState = [string]$ad.currentStateName }
                        if ($ad.ContainsKey('planned')) { try { $planned = [bool]$ad.planned } catch { } }
                        if ($ad.ContainsKey('changed')) { try { $changed = [bool]$ad.changed } catch { } }
                        if ($ad.transition) {
                            $tr = _QCP-ToHashtable $ad.transition
                            if ($tr -and $tr.Data) { $transitionDetail = _QCP-ToHashtable $tr.Data }
                        }
                    }
                    break
                }
            }
        }
    }

    if ((_QCP-IsNullOrWhiteSpace $targetState) -and $settings) {
        try {
            $resolveCtx = @{ resultStatus = 'Succeeded' }
            if (-not (_QCP-IsNullOrWhiteSpace $PrependTrigger)) { $resolveCtx['prependTrigger'] = $PrependTrigger }
            $targetState = Resolve-QCWorkflowStateAfterPrepend -Settings $settings -Context $resolveCtx
        } catch { }
    }

    $level = 'Information'
    if ($Writeback -and -not $Writeback.IsSuccess) { $level = 'Warning' }
    elseif ($stateCode -match 'FAILED|INVALID|UNAVAILABLE|MISSING') { $level = 'Warning' }

    $msg = if ($changed) {
        "QC_PREPEND set workflow state to '$targetState'."
    } elseif ($planned -or $dryRun) {
        "QC_PREPEND planned workflow state '$targetState' (dry-run)."
    } else {
        'QC_PREPEND workflow state writeback finished (no ProjectWise state change).'
    }

    Write-QCJsonLog -Level $level -Code 'QC_PREPEND_WORKFLOW_STATE' -Message $msg -Data @{
        jobId             = $jobId
        prependTrigger    = $PrependTrigger
        targetState       = $targetState
        previousState     = $previousState
        stateActionCode   = $stateCode
        planned           = $planned
        changed           = $changed
        dryRun            = $dryRun
        writebackCode     = if ($Writeback) { [string]$Writeback.Code } else { $null }
        writebackSuccess  = if ($Writeback) { [bool]$Writeback.IsSuccess } else { $false }
        autoSetState      = if ($settings) { [bool]$settings.autoSetState } else { $null }
        mode              = if ($settings) { [string]$settings.mode } else { $null }
        workflowName      = if ($settings) { [string]$settings.workflowName } else { $null }
        warnings          = @($warnings)
        targetStateExists = if ($transitionDetail -and $transitionDetail.ContainsKey('targetStateExists')) { $transitionDetail.targetStateExists } else { $null }
        transitionValid   = if ($transitionDetail -and $transitionDetail.ContainsKey('transitionValid')) { $transitionDetail.transitionValid } else { $null }
        workflowStates    = if ($transitionDetail -and $transitionDetail.states) { @($transitionDetail.states) } else { @() }
    } | Out-Null
}

function _QCP-AppendWorkflowWriteback([object]$Result, [hashtable]$Job, [hashtable]$Config, [string]$SourcePath, [string]$OutputPath, [string]$HistoryPath) {
    if ($null -eq $Result -or -not $Result.IsSuccess) { return $Result }
    $document = _QCP-ResolvePwDocumentForJob -Job $Job -Config $Config
    $ctx = _QCP-NewWorkflowContext -Job $Job -Config $Config -SourcePath $SourcePath -OutputPath $OutputPath -HistoryPath $HistoryPath -ResultStatus 'Succeeded' -ErrorMessage $null -Document $document
    $ctx['config'] = $Config
    $ctx['job'] = $Job
    if ($document) { $ctx['document'] = $document }
    $prependTrigger = _QCP-ResolvePrependTrigger -Job $Job
    if (-not (_QCP-IsNullOrWhiteSpace $prependTrigger)) { $ctx['prependTrigger'] = $prependTrigger }
    $writeback = Invoke-QCWorkflowWriteback -Config $Config -Context $ctx
    _QCP-LogPrependWorkflowWriteback -Job $Job -Config $Config -Writeback $writeback -PrependTrigger $prependTrigger
    $strict = $false
    try {
        $settings = Get-QCWorkflowSettings -Config $Config
        $strict = [bool]$settings.strictMode
    } catch { }
    $data = @{}
    if ($Result.Data) {
        $rd = _QCP-ToHashtable $Result.Data
        if ($rd) { foreach ($k in $rd.Keys) { $data[$k] = $rd[$k] } }
    }
    $data['workflowWriteback'] = $writeback
    if (-not $writeback.IsSuccess -and $strict) {
        return New-QCFailureResult -Code 'QC_PREPEND_WORKFLOW_WRITEBACK_FAILED' -Message 'QC_PREPEND succeeded but strict QC workflow writeback failed.' -Data $data
    }

    return New-QCSuccessResult -Code $Result.Code -Message $Result.Message -Data $data
}

function _QCP-ResolveSourcePdf([hashtable]$Job) {
    if ($Job.ContainsKey('source') -and $Job.source -is [hashtable] -and $Job.source.ContainsKey('localPath') -and $Job.source.localPath) {
        return [string]$Job.source.localPath
    }
    if ($Job.ContainsKey('sourcePath') -and $Job.sourcePath) {
        return [string]$Job.sourcePath
    }
    return $null
}

function _QCP-GetQpdfExePath([hashtable]$Config) {
    if ($Config.ContainsKey('qcPrepend') -and $Config.qcPrepend) {
        $qc = _QCP-ToHashtable $Config.qcPrepend
        if ($qc -and $qc.ContainsKey('qpdfExePath') -and $qc.qpdfExePath) {
            $p = [string]$qc.qpdfExePath
            try {
                if (-not [System.IO.Path]::IsPathRooted($p)) {
                    return (Join-Path (_QCP-GetRepoRoot) $p)
                }
            } catch { }
            return $p
        }
    }
    $default = Join-Path (_QCP-GetRepoRoot) 'tools\qpdf\bin\qpdf.exe'
    return $default
}

function _QCP-ResolveHistoryPath([hashtable]$Job, [hashtable]$Config) {
    $qc = $null
    if ($Config.ContainsKey('qcPrepend') -and $Config.qcPrepend) { $qc = _QCP-ToHashtable $Config.qcPrepend }
    $historyRoot = if ($qc -and $qc.historyRoot) { [string]$qc.historyRoot } else { $null }
    if (_QCP-IsNullOrWhiteSpace $historyRoot) { return New-QCFailureResult -Code 'QC_PREPEND_CONFIG_MISSING_HISTORY_ROOT' -Message 'qcPrepend.historyRoot is required.' -Data @{} }

    if ($Job.ContainsKey('metadata') -and $Job.metadata -is [hashtable] -and $Job.metadata.ContainsKey('historyPdfPath') -and $Job.metadata.historyPdfPath) {
        return New-QCSuccessResult -Code 'QC_PREPEND_HISTORY_PATH' -Message 'History path resolved from job metadata.' -Data @{ historyPdf = [string]$Job.metadata.historyPdfPath }
    }

    $sourceName = if ($Job.ContainsKey('sourceName') -and $Job.sourceName) { [string]$Job.sourceName } else { ([System.IO.Path]::GetFileName([string]$Job.sourcePath)) }
    if (_QCP-IsNullOrWhiteSpace $sourceName) { $sourceName = ([string]$Job.id + '.pdf') }

    $sourceFolder = if ($Job.ContainsKey('sourceFolder') -and $Job.sourceFolder) { [string]$Job.sourceFolder } else { '' }
    $folderKey = if (_QCP-IsNullOrWhiteSpace $sourceFolder) { 'root' } else { (_QCP-Sha256Hex -Text $sourceFolder).Substring(0, 8) }

    $destDir = Join-Path $historyRoot $folderKey
    $historyPdf = Join-Path $destDir $sourceName
    return New-QCSuccessResult -Code 'QC_PREPEND_HISTORY_PATH' -Message 'History path resolved from config.' -Data @{ historyPdf = $historyPdf; historyDir = $destDir; folderKey = $folderKey }
}

function _QCP-RunQpdfPrepend([string]$QpdfExe, [string]$SourcePdf, [string]$HistoryPdf, [string]$OutPdf) {
    # Prepend by concatenating pages: source + history
    $args = @('--empty', '--pages', $SourcePdf, $HistoryPdf, '--', $OutPdf)
    $stdoutPath = [System.IO.Path]::GetTempFileName()
    $stderrPath = [System.IO.Path]::GetTempFileName()
    try {
        $p = Start-Process -FilePath $QpdfExe -ArgumentList $args -NoNewWindow -Wait -PassThru -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
        $stdout = Get-Content -LiteralPath $stdoutPath -Raw -ErrorAction SilentlyContinue
        $stderr = Get-Content -LiteralPath $stderrPath -Raw -ErrorAction SilentlyContinue
        return @{ exitCode = $p.ExitCode; stdout = $stdout; stderr = $stderr; args = $args }
    } finally {
        Remove-Item -LiteralPath $stdoutPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $stderrPath -Force -ErrorAction SilentlyContinue
    }
}

function _QCP-RunOverlayExe([string]$ExePath, [string[]]$OverlayArgs, [string]$WorkingDirectory) {
    $stdoutPath = [System.IO.Path]::GetTempFileName()
    $stderrPath = [System.IO.Path]::GetTempFileName()
    try {
        # Start-Process rejects null/empty items in -ArgumentList (and some callers can accidentally pass them).
        $cleanArgs = @()
        foreach ($a in @($OverlayArgs)) {
            if ($null -eq $a) { continue }
            $s = [string]$a
            if ($s.Trim().Length -eq 0) { continue }
            $cleanArgs += $s
        }

        # Windows PowerShell can be picky about string[] ArgumentList. Use a single arg line string.
        $argLine = (($cleanArgs | ForEach-Object {
            $t = [string]$_
            if ($t -match '[\\s\"]') { return ('"' + ($t -replace '"', '\\"') + '"') }
            return $t
        }) -join ' ')

        if ($null -ne $WorkingDirectory -and ([string]$WorkingDirectory).Trim().Length -gt 0) {
            $p = Start-Process -FilePath $ExePath -ArgumentList $argLine -WorkingDirectory $WorkingDirectory -NoNewWindow -Wait -PassThru -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
        } else {
            $p = Start-Process -FilePath $ExePath -ArgumentList $argLine -NoNewWindow -Wait -PassThru -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
        }
        $stdout = Get-Content -LiteralPath $stdoutPath -Raw -ErrorAction SilentlyContinue
        $stderr = Get-Content -LiteralPath $stderrPath -Raw -ErrorAction SilentlyContinue
        return @{ exitCode = $p.ExitCode; stdout = $stdout; stderr = $stderr; args = $cleanArgs; argLine = $argLine }
    } finally {
        Remove-Item -LiteralPath $stdoutPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $stderrPath -Force -ErrorAction SilentlyContinue
    }
}

function _QCP-TokenizeArgs([string]$CommandLine) {
    # Minimal quoted-string tokenizer.
    $args = @()
    $cur = ''
    $inQuote = $false
    for ($i = 0; $i -lt $CommandLine.Length; $i++) {
        $ch = $CommandLine[$i]
        if ($ch -eq '"') { $inQuote = -not $inQuote; continue }
        if (-not $inQuote -and [char]::IsWhiteSpace($ch)) {
            if ($cur.Length -gt 0) { $args += $cur; $cur = '' }
            continue
        }
        $cur += $ch
    }
    if ($cur.Length -gt 0) { $args += $cur }
    return $args
}

function Test-QCJobReady {
    <#
    .SYNOPSIS
    Validates job readiness for processing.
    .DESCRIPTION
    Confirms required job fields, state, and preconditions before dispatch.
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

    $missing = @()
    foreach ($k in @('id', 'type', 'sourcePath', 'dedupeKey')) {
        if (-not $Job.ContainsKey($k) -or (_QCP-IsNullOrWhiteSpace $Job[$k])) { $missing += $k }
    }
    if ($missing.Count -gt 0) {
        return New-QCFailureResult -Code 'PROCESSOR_JOB_NOT_READY' -Message 'Job is missing required fields for processing.' -Data @{ missing = $missing; jobId = [string]$Job.id }
    }

    return New-QCSuccessResult -Code 'PROCESSOR_JOB_READY' -Message 'Job is ready for processing.' -Data @{ jobId = [string]$Job.id }
}

function Invoke-QCProcessorByType {
    <#
    .SYNOPSIS
    Dispatches a job to the mapped processor by job type.
    .DESCRIPTION
    Resolves processor mapping and invokes the appropriate processor entry point.
    .PARAMETER Job
    Job payload.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: may trigger downstream local processing workflows.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    $ready = Test-QCJobReady -Job $Job -Config $Config
    if (-not $ready.IsSuccess) { return $ready }

    $jobType = [string]$Job.type
    if (_QCP-IsNullOrWhiteSpace $jobType) {
        return New-QCFailureResult -Code 'PROCESSOR_MISSING_JOB_TYPE' -Message 'Job.type is required for dispatch.' -Data @{ jobId = [string]$Job.id }
    }

    $map = $null
    if ($Config.ContainsKey('processors') -and $Config.processors -and $Config.processors.ContainsKey('processorMap')) {
        $map = $Config.processors.processorMap
    } elseif ($Config.ContainsKey('processorMap')) {
        $map = $Config.processorMap
    }

    $handlerName = $null
    if ($map -is [hashtable] -and $map.ContainsKey($jobType)) {
        $handlerName = [string]$map[$jobType]
    }
    if (_QCP-IsNullOrWhiteSpace $handlerName) {
        # Default mapping for initial rollout.
        if ($jobType -eq 'QC_PREPEND') { $handlerName = 'Invoke-QCPrependProcessor' }
        elseif ($jobType -eq 'STATUS_SET_GEN') { $handlerName = 'Invoke-StatusSetProcessor' }
        elseif ($jobType -eq 'QC_REPORTING_SCAN') { $handlerName = 'Invoke-QCReportingScanProcessor' }
        elseif ($jobType -eq 'QC_COMMENT_STATUS_SYNC') { $handlerName = 'Invoke-QCCommentStatusSyncProcessor' }
        elseif ($jobType -eq 'QC_RENDITION') { $handlerName = 'Invoke-QCRenditionProcessor' }
    }

    if (_QCP-IsNullOrWhiteSpace $handlerName) {
        return New-QCFailureResult -Code 'PROCESSOR_NO_HANDLER' -Message "No processor handler mapped for job type: $jobType" -Data @{ jobId = [string]$Job.id; jobType = $jobType }
    }

    $cmd = Get-Command -Name $handlerName -ErrorAction SilentlyContinue
    if (-not $cmd) {
        return New-QCFailureResult -Code 'PROCESSOR_HANDLER_NOT_FOUND' -Message "Mapped processor handler not found: $handlerName" -Data @{ jobId = [string]$Job.id; jobType = $jobType; handler = $handlerName }
    }

    try {
        $r = & $handlerName -Job $Job -Config $Config
        # Handlers must return a single QCResult. In Windows PowerShell, any stray pipeline output
        # turns the return into an object[]; recover by selecting the last QCResult-like item.
        if ($r -is [object[]]) {
            $qcLike = @($r | Where-Object { $_ -ne $null -and $_.PSObject -and ($_.PSObject.Properties.Name -contains 'IsSuccess') })
            if ($qcLike.Count -gt 0) {
                $r = $qcLike[-1]
            }
        }
        if ($null -eq $r -or -not ($r.PSObject.Properties.Name -contains 'IsSuccess')) {
            $types = @()
            try {
                if ($r -is [object[]]) { $types = @($r | ForEach-Object { if ($_ -eq $null) { '<null>' } else { $_.GetType().FullName } } | Select-Object -First 10) }
                elseif ($null -ne $r) { $types = @($r.GetType().FullName) }
            } catch { }
            return New-QCFailureResult -Code 'PROCESSOR_INVALID_RESULT' -Message "Handler did not return a QCResult: $handlerName" -Data @{ jobId = [string]$Job.id; jobType = $jobType; handler = $handlerName; returnTypes = $types }
        }
        return $r
    } catch {
        return New-QCFailureResult -Code 'PROCESSOR_HANDLER_THROW' -Message "Handler threw: $handlerName" -Data @{ jobId = [string]$Job.id; jobType = $jobType; handler = $handlerName; error = $_.Exception.Message }
    }
}

function Invoke-QCPrependProcessor {
    <#
    .SYNOPSIS
    Entry point for QC_PREPEND jobs.
    .DESCRIPTION
    Handles QC prepend workflow orchestration for eligible jobs.
    .PARAMETER Job
    Job payload.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: local processing actions only (no ProjectWise write operations in this stub).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    _QCP-ApplyFsThrottleConfig -Config $Config

    $qc = @{}
    if ($Config.ContainsKey('qcPrepend') -and $Config.qcPrepend) {
        $qcNorm = _QCP-ToHashtable $Config.qcPrepend
        if ($qcNorm) { $qc = $qcNorm }
    }

    $mode = 'local'
    if ($qc.ContainsKey('mode') -and $qc.mode) { $mode = ([string]$qc.mode).Trim().ToLowerInvariant() }

    $isDryRun = $false
    if ($Config.ContainsKey('dryRun')) { $isDryRun = [bool]$Config.dryRun }

    # ProjectWise-triggered QC_PREPEND: use legacy prepend_qc.ps1 adapter (does PW export + overlay + history).
    if ($mode -eq 'legacypw') {
        $repoRoot = _QCP-GetRepoRoot
        $legacyPrepend = if ($qc.ContainsKey('legacyScriptPath') -and $qc.legacyScriptPath) { [string]$qc.legacyScriptPath } else { (Join-Path $repoRoot 'legacy\prepend_qc.ps1') }
        if (-not (Test-Path -LiteralPath $legacyPrepend)) {
            return New-QCFailureResult -Code 'QC_PREPEND_LEGACY_SCRIPT_MISSING' -Message "Legacy prepend_qc.ps1 not found: $legacyPrepend" -Data @{ script = $legacyPrepend }
        }

        $pwCfg = @{}
        if ($Config.ContainsKey('projectWise') -and $Config.projectWise) {
            $pwNorm = _QCP-ToHashtable $Config.projectWise
            if ($pwNorm) { $pwCfg = $pwNorm }
        }
        $ds = if ($pwCfg.ContainsKey('datasourceName') -and $pwCfg.datasourceName) { [string]$pwCfg.datasourceName } else { '' }
        if (_QCP-IsNullOrWhiteSpace $ds) { $ds = 'typsa-us-pw.bentley.com:typsa-us-pw-03' }

        $incomingFolder = ''
        # Prefer the original PW folder path from watcher metadata (preserves casing / Documents\ prefix).
        try {
            if ($Job.ContainsKey('metadata') -and $Job.metadata -and $Job.metadata.ContainsKey('candidate') -and $Job.metadata.candidate) {
                $cand = _QCP-ToHashtable $Job.metadata.candidate
                if ($cand -is [hashtable] -and $cand.ContainsKey('sourceFolder') -and $cand.sourceFolder) {
                    $incomingFolder = [string]$cand.sourceFolder
                }
            }
        } catch { }
        if (_QCP-IsNullOrWhiteSpace $incomingFolder) {
            if ($Job.ContainsKey('sourceFolder') -and $Job.sourceFolder) { $incomingFolder = [string]$Job.sourceFolder }
        }
        # pwps_dab cmdlets generally expect FolderPath without the leading "Documents\".
        if ($incomingFolder -match '^(?i)Documents\\') {
            $incomingFolder = ($incomingFolder -replace '^(?i)Documents\\', '')
        }
        $incomingDocName = if ($Job.ContainsKey('sourceName') -and $Job.sourceName) { [string]$Job.sourceName } else { '' }
        if (_QCP-IsNullOrWhiteSpace $incomingDocName) {
            $incomingDocName = [System.IO.Path]::GetFileName([string]$Job.sourcePath)
        }

        if ((_QCP-IsNullOrWhiteSpace $incomingFolder) -or (_QCP-IsNullOrWhiteSpace $incomingDocName)) {
            return New-QCFailureResult -Code 'QC_PREPEND_LEGACY_MISSING_INPUTS' -Message 'Legacy PW prepend requires Job.sourceFolder and Job.sourceName.' -Data @{ jobId = [string]$Job.id; sourceFolder = [string]$Job.sourceFolder; sourceName = [string]$Job.sourceName; sourcePath = [string]$Job.sourcePath }
        }

        $localRoot = if ($qc.ContainsKey('localRoot') -and $qc.localRoot) { [string]$qc.localRoot } else { 'C:\PW_QC_LOCAL' }
        $logDir = if ($qc.ContainsKey('logDir') -and $qc.logDir) { [string]$qc.logDir } else { '' }

        $qpdfExe = $null
        if ($qc.ContainsKey('qpdfExePath') -and $qc.qpdfExePath) { $qpdfExe = [string]$qc.qpdfExePath }
        if (-not (_QCP-IsNullOrWhiteSpace $qpdfExe)) {
            try {
                if (-not [System.IO.Path]::IsPathRooted($qpdfExe)) {
                    $qpdfExe = Join-Path $repoRoot $qpdfExe
                }
            } catch { }
        }
        if (_QCP-IsNullOrWhiteSpace $qpdfExe) { $qpdfExe = Join-Path $repoRoot 'tools\qpdf\bin\qpdf.exe' }
        $overlayExe = $null
        if ($qc.ContainsKey('overlayExePath') -and $qc.overlayExePath) { $overlayExe = [string]$qc.overlayExePath }
        if (_QCP-IsNullOrWhiteSpace $overlayExe) { $overlayExe = Join-Path $repoRoot 'dist\qc_overlay_prepend\qc_overlay_prepend.exe' }

        $ovOldFromHistoryOnly = $true
        if ($qc.ContainsKey('overlayOldFromHistoryOnly')) { try { $ovOldFromHistoryOnly = [bool]$qc.overlayOldFromHistoryOnly } catch { $ovOldFromHistoryOnly = $true } }
        $ovSheetWorkDir = $true
        if ($qc.ContainsKey('overlaySheetWorkDir')) { try { $ovSheetWorkDir = [bool]$qc.overlaySheetWorkDir } catch { $ovSheetWorkDir = $true } }

        if ($isDryRun) {
            return New-QCSuccessResult -Code 'QC_PREPEND_DRYRUN' -Message 'Dry-run: would run legacy prepend_qc.ps1 for ProjectWise job.' -Data @{
                jobId = [string]$Job.id
                datasourceName = $ds
                incomingFolderPath = $incomingFolder
                incomingDocName = $incomingDocName
                legacyScript = $legacyPrepend
                localRoot = $localRoot
                qpdfExe = $qpdfExe
                overlayExe = $overlayExe
            }
        }

        $args = @(
            '-NoProfile',
            '-ExecutionPolicy', 'Bypass',
            '-File', $legacyPrepend,
            '-IncomingFolderPath', $incomingFolder,
            '-IncomingDocName', $incomingDocName,
            '-DatasourceName', $ds,
            '-LocalRoot', $localRoot
        )
        if (-not (_QCP-IsNullOrWhiteSpace $logDir)) { $args += @('-LogDir', $logDir) }
        if (Test-Path -LiteralPath $qpdfExe) { $args += @('-QpdfExe', $qpdfExe) }
        if (Test-Path -LiteralPath $overlayExe) { $args += @('-QcOverlayExe', $overlayExe) }
        $args += @('-OverlayOldFromHistoryOnly', ([string]$ovOldFromHistoryOnly), '-OverlaySheetWorkDir', ([string]$ovSheetWorkDir))
        $appsettingsPath = Join-Path $repoRoot 'appsettings.json'
        if ($Config.ContainsKey('appsettingsPath') -and $Config.appsettingsPath) {
            $appsettingsPath = [string]$Config.appsettingsPath
        }
        if (Test-Path -LiteralPath $appsettingsPath) {
            $args += @('-AppsettingsPath', $appsettingsPath)
        }

        $stdoutPath = [System.IO.Path]::GetTempFileName()
        $stderrPath = [System.IO.Path]::GetTempFileName()
        try {
            # Windows PowerShell can mis-handle string[] -ArgumentList when paths contain spaces.
            # Use a single quoted arg line (same approach as _QCP-RunOverlayExe).
            $cleanArgs = @('-MTA') + @($args)
            $argLine = (($cleanArgs | ForEach-Object {
                $t = [string]$_
                if ($t -match '[\s"]') { return ('"' + ($t -replace '"', '\\"') + '"') }
                return $t
            }) -join ' ')

            $p = Start-Process -FilePath 'powershell.exe' -ArgumentList $argLine -NoNewWindow -Wait -PassThru -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
            $stdout = [string](Get-Content -LiteralPath $stdoutPath -Raw -ErrorAction SilentlyContinue)
            $stderr = [string](Get-Content -LiteralPath $stderrPath -Raw -ErrorAction SilentlyContinue)
            if ($p.ExitCode -ne 0) {
                $failData = _QCP-SanitizeProcessorDataForStorage @{
                    exitCode = [int]$p.ExitCode
                    stdout = $stdout
                    stderr = $stderr
                    args = $args
                    argLine = $argLine
                    legacyScript = $legacyPrepend
                }
                return New-QCFailureResult -Code 'QC_PREPEND_LEGACY_FAILED' -Message 'Legacy prepend_qc.ps1 failed.' -Data $failData
            }

            $tagCleared = $null
            $tagClearError = $null
            $writebackWhileConnected = $null
            $clearTag = $true
            if ($pwCfg.ContainsKey('clearTriggerTagOnSuccess')) { try { $clearTag = [bool]$pwCfg.clearTriggerTagOnSuccess } catch { $clearTag = $true } }
            $success = New-QCSuccessResult -Code 'QC_PREPEND_OK' -Message 'QC_PREPEND completed via legacy prepend_qc.ps1.' -Data (_QCP-SanitizeProcessorDataForStorage @{
                exitCode = [int]$p.ExitCode
                stdout = $stdout
                stderr = $stderr
                legacyScript = $legacyPrepend
                incomingFolderPath = $incomingFolder
                incomingDocName = $incomingDocName
                triggerTagCleared = $tagCleared
                triggerTagClearError = $tagClearError
            })
            if ($clearTag) {
                try {
                    $repoRoot = _QCP-GetRepoRoot
                    $pwConnMod = Join-Path $repoRoot 'modules\PW.Connection.psm1'
                    if (Test-Path -LiteralPath $pwConnMod) { Import-Module $pwConnMod -Force -ErrorAction SilentlyContinue | Out-Null }

                    $credPath = if ($pwCfg.ContainsKey('credentialPath') -and $pwCfg.credentialPath) { [string]$pwCfg.credentialPath } else { 'C:\PW_QC_LOCAL\pw_cred.txt' }
                    $credRes = Get-PWCredentialFromFile -CredentialPath $credPath
                    if (-not $credRes.IsSuccess) { throw ($credRes.Code + ': ' + $credRes.Message) }
                    $connRes2 = Connect-PW -DatasourceName $ds -Credential ([pscredential]$credRes.Data.credential)
                    if (-not $connRes2.IsSuccess) { throw ($connRes2.Code + ': ' + $connRes2.Message) }

                    $discRes = Ensure-PWDiscoveryCmdlets
                    if (-not $discRes.IsSuccess -or -not (Get-Command -Name Get-PWDocumentsBySearch -ErrorAction SilentlyContinue)) {
                        throw ('QC_PREPEND_PW_DISCOVERY_INCOMPLETE: ProjectWise document search cmdlet Get-PWDocumentsBySearch is missing after connect/re-import; cannot clear QC_Archivist trigger tag.')
                    }

                    $doc = Get-PWDocumentsBySearch -FolderPath $incomingFolder -JustThisFolder -DocumentName $incomingDocName -PopulatePath -ErrorAction Stop
                    if ($doc) {
                        $currentDesc = $doc.Description
                        if ($null -eq $currentDesc -and $doc.PSObject.Properties['Description']) { $currentDesc = $doc.Description }
                        if ($null -eq $currentDesc) { $currentDesc = '' }
                        $newDesc = ([string]$currentDesc -replace [regex]::Escape('QC_Archivist'), '').Trim()
                        $doc.Description = $newDesc
                        Update-PWDocumentProperties $doc -ErrorAction Stop
                        $tagCleared = $true
                    } else {
                        $tagCleared = $false
                        $tagClearError = 'Could not re-find document to clear trigger tag.'
                    }
                    if ($doc) {
                        $Job['document'] = $doc
                        $success.Data.triggerTagCleared = $tagCleared
                        $success.Data.triggerTagClearError = $tagClearError
                        $writebackWhileConnected = _QCP-AppendWorkflowWriteback -Result $success -Job $Job -Config $Config -SourcePath $incomingDocName -OutputPath $null -HistoryPath $null
                    }
                } catch {
                    $tagCleared = $false
                    $tagClearError = [string]$_.Exception.Message
                    try {
                        if ($success.Data) {
                            $success.Data.triggerTagCleared = $tagCleared
                            $success.Data.triggerTagClearError = $tagClearError
                        }
                    } catch { }
                } finally {
                    try { Disconnect-PW | Out-Null } catch { }
                    if ($Job.ContainsKey('document')) { $Job.Remove('document') }
                }
            }
            if ($writebackWhileConnected) { return $writebackWhileConnected }
            return (_QCP-AppendWorkflowWriteback -Result $success -Job $Job -Config $Config -SourcePath $incomingDocName -OutputPath $null -HistoryPath $null)
        } finally {
            Remove-Item -LiteralPath $stdoutPath -Force -ErrorAction SilentlyContinue
            Remove-Item -LiteralPath $stderrPath -Force -ErrorAction SilentlyContinue
        }
    }

    $sourcePdf = _QCP-ResolveSourcePdf -Job $Job
    if (_QCP-IsNullOrWhiteSpace $sourcePdf) {
        return New-QCFailureResult -Code 'QC_PREPEND_SOURCE_MISSING' -Message 'Job does not include a source PDF path.' -Data @{ jobId = [string]$Job.id }
    }
    if (-not (Test-Path -LiteralPath $sourcePdf)) {
        return New-QCFailureResult -Code 'QC_PREPEND_SOURCE_NOT_FOUND' -Message "Source PDF not found: $sourcePdf" -Data @{ jobId = [string]$Job.id; sourcePdf = $sourcePdf }
    }

    $histRes = _QCP-ResolveHistoryPath -Job $Job -Config $Config
    if (-not $histRes.IsSuccess) { return $histRes }
    $historyPdf = [string]$histRes.Data.historyPdf
    $historyDir = Split-Path -Parent $historyPdf

    $enableOverlay = $false
    if ($qc.ContainsKey('enableOverlay')) { $enableOverlay = [bool]$qc.enableOverlay }
    $overlayExePath = if ($qc.ContainsKey('overlayExePath') -and $qc.overlayExePath) { [string]$qc.overlayExePath } else { (Join-Path (_QCP-GetRepoRoot) 'dist\qc_overlay_prepend\qc_overlay_prepend.exe') }

    # Output path for overlay results (optional)
    $outputRoot = if ($qc.ContainsKey('outputRoot') -and $qc.outputRoot) { [string]$qc.outputRoot } else { '' }
    $tempRoot = if ($qc.ContainsKey('tempRoot') -and $qc.tempRoot) { [string]$qc.tempRoot } else { (Join-Path ([System.IO.Path]::GetTempPath()) 'qc_prepend') }
    _QCP-EnsureDir -Path $tempRoot

    $overlayOutPdf = $null
    if (-not (_QCP-IsNullOrWhiteSpace $outputRoot)) {
        _QCP-EnsureDir -Path $outputRoot
        $base = [System.IO.Path]::GetFileNameWithoutExtension([string]$Job.sourceName)
        if (_QCP-IsNullOrWhiteSpace $base) { $base = [string]$Job.id }
        $folderKey = 'root'
        if ($histRes.Data -and ($histRes.Data -is [hashtable]) -and $histRes.Data.ContainsKey('folderKey') -and $histRes.Data.folderKey) {
            $folderKey = [string]$histRes.Data.folderKey
        }
        $outDir = Join-Path $outputRoot $folderKey
        _QCP-EnsureDir -Path $outDir
        # Primary QC deliverable naming: <base>-qc.pdf
        $overlayOutPdf = Join-Path $outDir ($base + '-qc.pdf')
    }

    $overlayTemplate = if ($qc.ContainsKey('overlayArgumentsTemplate') -and $qc.overlayArgumentsTemplate) { [string]$qc.overlayArgumentsTemplate } else { '' }
    $overlayCmdWouldRun = $null
    if ($enableOverlay) {
        $overlayCmdWouldRun = @{
            exe = $overlayExePath
            argsTemplate = $overlayTemplate
            outputPdf = $overlayOutPdf
            historyPdf = $historyPdf
            sourcePdf = $sourcePdf
        }
    }

    if ($isDryRun) {
        return New-QCSuccessResult -Code 'QC_PREPEND_DRYRUN' -Message 'Dry-run: QC_PREPEND would update history and optionally run overlay.' -Data @{
            jobId = [string]$Job.id
            sourcePdf = $sourcePdf
            targetHistoryPdf = $historyPdf
            historyExists = (Test-Path -LiteralPath $historyPdf)
            overlayEnabled = $enableOverlay
            overlayCommandWouldRun = $overlayCmdWouldRun
            qcOutputPdf = $overlayOutPdf
        }
    }

    $qpdfExe = _QCP-GetQpdfExePath -Config $Config
    if (-not (Test-Path -LiteralPath $qpdfExe)) {
        return New-QCFailureResult -Code 'PDF_MERGE_TOOL_MISSING' -Message "qpdf.exe not found: $qpdfExe" -Data @{ qpdfExe = $qpdfExe }
    }

    _QCP-EnsureDir -Path $historyDir

    $tmpHistory = Join-Path $tempRoot ([string]$Job.id + '.history.tmp.' + ([guid]::NewGuid().ToString('N')) + '.pdf')
    $tmpQcOut = Join-Path $tempRoot ([string]$Job.id + '.qc.tmp.' + ([guid]::NewGuid().ToString('N')) + '.pdf')

    try {
        $historyExists = Test-Path -LiteralPath $historyPdf

        # 1) If overlay enabled, build QC output first:
        #    qc.pdf = overlay(source vs page1(history)) + history tail pages
        #    Use overlay/qc_overlay_prepend's contract directly: incoming + qc_history -> merged QC output.
        if ($enableOverlay) {
            if (-not (Test-Path -LiteralPath $overlayExePath)) {
                return New-QCFailureResult -Code 'QC_OVERLAY_EXE_MISSING' -Message "Overlay exe missing: $overlayExePath" -Data @{ overlayExePath = $overlayExePath }
            }
            if (_QCP-IsNullOrWhiteSpace $overlayOutPdf) {
                return New-QCFailureResult -Code 'QC_OVERLAY_OUTPUT_MISSING' -Message 'qcPrepend.outputRoot is required to produce QC output when overlay is enabled.' -Data @{}
            }

            $overlayArgs = @(
                $sourcePdf,
                $historyPdf,
                '-o',
                $tmpQcOut
            )

            # Pass through selected overlay engine options if present in config.
            # Mirrors overlay/qc_overlay_prepend.py CLI.
            if ($qc.ContainsKey('overlayAlpha') -and $qc.overlayAlpha -ne $null -and ([string]$qc.overlayAlpha).Trim().Length -gt 0) {
                $overlayArgs += @('--alpha', [string]$qc.overlayAlpha)
            }
            if ($qc.ContainsKey('overlayFit') -and [bool]$qc.overlayFit) {
                $overlayArgs += @('--fit')
            }
            if ($qc.ContainsKey('overlayVerbose') -and [bool]$qc.overlayVerbose) {
                $overlayArgs += @('--verbose')
            }
            if ($qc.ContainsKey('overlayKeepTemp') -and [bool]$qc.overlayKeepTemp) {
                $overlayArgs += @('--keep-temp')
            }
            if ($qc.ContainsKey('overlayNoFlattenSources') -and [bool]$qc.overlayNoFlattenSources) {
                $overlayArgs += @('--no-flatten-sources')
            }
            if ($qc.ContainsKey('overlayFlattenRaster') -and [bool]$qc.overlayFlattenRaster) {
                $overlayArgs += @('--flatten-raster')
            }
            if ($qc.ContainsKey('overlayFlattenDpi') -and $qc.overlayFlattenDpi -ne $null -and ([string]$qc.overlayFlattenDpi).Trim().Length -gt 0) {
                $overlayArgs += @('--flatten-dpi', [string]$qc.overlayFlattenDpi)
            }
            if ($qc.ContainsKey('overlayCurrentMasterPath') -and $qc.overlayCurrentMasterPath) {
                $overlayArgs += @('--current-master', [string]$qc.overlayCurrentMasterPath)
            }
            if ($qc.ContainsKey('overlaySheetWorkDir') -and $qc.overlaySheetWorkDir) {
                $overlayArgs += @('--sheet-work-dir', [string]$qc.overlaySheetWorkDir)
            }

            $overlayRes = $null
            try {
                $overlayRes = _QCP-RunOverlayExe -ExePath $overlayExePath -OverlayArgs $overlayArgs -WorkingDirectory (Split-Path -Parent $sourcePdf)
            } catch {
                return New-QCFailureResult -Code 'QC_OVERLAY_LAUNCH_FAILED' -Message 'Failed to launch overlay exe.' -Data @{ exe = $overlayExePath; errorMessage = $_.Exception.Message; args = $overlayArgs }
            }
            if ([int]$overlayRes.exitCode -ne 0) {
                return New-QCFailureResult -Code 'QC_OVERLAY_FAILED' -Message 'Overlay exe failed.' -Data @{ exitCode = $overlayRes.exitCode; stderr = $overlayRes.stderr; stdout = $overlayRes.stdout; args = $overlayRes.args; argLine = $overlayRes.argLine; exe = $overlayExePath }
            }

            $stampCfg = if (Get-Command -Name 'Get-QCReviewStampSettings' -ErrorAction SilentlyContinue) {
                Get-QCReviewStampSettings -Config $Config
            } else { $null }
            if ($stampCfg) {
                $stampRes = _QCP-TryApplyReviewStampFromJob -PdfPath $tmpQcOut -Job $Job -Config $Config `
                    -FolderPath ([string]$Job.sourceFolder) -SourceDocumentName ([string]$Job.sourceName) `
                    -OverlayExe $overlayExePath
                $jobRt = _QCP-GetJobMetadataValue -Job $Job -Keys @('reviewType', 'qcReviewType')
                if (_QCP-IsNullOrWhiteSpace $jobRt) { $jobRt = [string]$stampRes.reviewType }
                if (_QCP-ReviewStampRequiredForReviewType -StampSettings $stampCfg -ReviewType $jobRt -and -not $stampRes.skipped -and -not $stampRes.applied) {
                    return New-QCFailureResult -Code 'QC_REVIEW_STAMP_FAILED' -Message 'Review stamp failed on QC PDF.' -Data $stampRes
                }
            }

            try {
                Move-Item -LiteralPath $tmpQcOut -Destination $overlayOutPdf -Force -ErrorAction Stop
                _QCP-FsThrottle
            } catch {
                return New-QCFailureResult -Code 'QC_PREPEND_QC_WRITE_FAILED' -Message 'Failed to write QC output PDF.' -Data @{ tmpQcOut = $tmpQcOut; qcOutputPdf = $overlayOutPdf; errorMessage = $_.Exception.Message }
            }
        }

        # 2) Update history:
        #    history.pdf = source + oldHistory  (or init from source)
        if (-not $historyExists) {
            try {
                Copy-Item -LiteralPath $sourcePdf -Destination $tmpHistory -Force -ErrorAction Stop
                _QCP-FsThrottle
            } catch {
                return New-QCFailureResult -Code 'QC_PREPEND_HISTORY_INIT_FAILED' -Message 'Failed to create initial history PDF (copy failed).' -Data @{ sourcePdf = $sourcePdf; tmpHistory = $tmpHistory; errorMessage = $_.Exception.Message }
            }
        } else {
            $merge = _QCP-RunQpdfPrepend -QpdfExe $qpdfExe -SourcePdf $sourcePdf -HistoryPdf $historyPdf -OutPdf $tmpHistory
            if ([int]$merge.exitCode -ne 0) {
                return New-QCFailureResult -Code 'QC_PREPEND_MERGE_FAILED' -Message 'qpdf prepend/merge failed.' -Data @{ exitCode = $merge.exitCode; stderr = $merge.stderr; stdout = $merge.stdout; args = $merge.args }
            }
        }

        if (-not $enableOverlay) {
            $stampCfg = if (Get-Command -Name 'Get-QCReviewStampSettings' -ErrorAction SilentlyContinue) {
                Get-QCReviewStampSettings -Config $Config
            } else { $null }
            if ($stampCfg) {
                $stampRes = _QCP-TryApplyReviewStampFromJob -PdfPath $tmpHistory -Job $Job -Config $Config `
                    -FolderPath ([string]$Job.sourceFolder) -SourceDocumentName ([string]$Job.sourceName) `
                    -OverlayExe $overlayExePath
                $rt = _QCP-GetJobMetadataValue -Job $Job -Keys @('reviewType', 'qcReviewType')
                if (_QCP-ReviewStampRequiredForReviewType -StampSettings $stampCfg -ReviewType $rt -and -not $stampRes.skipped -and -not $stampRes.applied) {
                    return New-QCFailureResult -Code 'QC_REVIEW_STAMP_FAILED' -Message 'Review stamp failed on history PDF.' -Data $stampRes
                }
            }
        }

        try {
            Move-Item -LiteralPath $tmpHistory -Destination $historyPdf -Force -ErrorAction Stop
            _QCP-FsThrottle
        } catch {
            return New-QCFailureResult -Code 'QC_PREPEND_HISTORY_WRITE_FAILED' -Message 'Failed to write updated history PDF (move/replace failed).' -Data @{ tmpHistory = $tmpHistory; targetHistoryPdf = $historyPdf; errorMessage = $_.Exception.Message }
        }

        $success = New-QCSuccessResult -Code 'QC_PREPEND_OK' -Message 'QC_PREPEND history updated.' -Data @{
            jobId = [string]$Job.id
            sourcePdf = $sourcePdf
            targetHistoryPdf = $historyPdf
            overlayEnabled = $enableOverlay
            overlayExe = if ($enableOverlay) { $overlayExePath } else { $null }
            qcOutputPdf = $overlayOutPdf
        }
        return (_QCP-AppendWorkflowWriteback -Result $success -Job $Job -Config $Config -SourcePath $sourcePdf -OutputPath $overlayOutPdf -HistoryPath $historyPdf)
    } finally {
        if (Test-Path -LiteralPath $tmpHistory) {
            Remove-Item -LiteralPath $tmpHistory -Force -ErrorAction SilentlyContinue
            _QCP-FsThrottle
        }
        if (Test-Path -LiteralPath $tmpQcOut) {
            Remove-Item -LiteralPath $tmpQcOut -Force -ErrorAction SilentlyContinue
            _QCP-FsThrottle
        }
    }
}

function Invoke-StatusSetProcessor {
    <#
    .SYNOPSIS
    Entry point for STATUS_SET_GEN jobs.
    .DESCRIPTION
    Handles status set generation workflow orchestration for eligible jobs.
    .PARAMETER Job
    Job payload.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: depends on statusSet.mode (stub, native: manifest + PW, legacy: external script).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    $isDryRun = $false
    if ($Config.ContainsKey('dryRun')) { $isDryRun = [bool]$Config.dryRun }

    $sourceFolder = if ($Job.ContainsKey('sourceFolder') -and $Job.sourceFolder) { [string]$Job.sourceFolder } else { '' }
    if (_QCP-IsNullOrWhiteSpace $sourceFolder) {
        return New-QCFailureResult -Code 'STATUS_SET_MISSING_SOURCE_FOLDER' -Message 'STATUS_SET_GEN requires Job.sourceFolder (target folder / sheets folder).' -Data @{ jobId = [string]$Job.id }
    }

    $ss = @{}
    if ($Config.ContainsKey('statusSet') -and $Config.statusSet) {
        $ssNorm = _QCP-ToHashtable $Config.statusSet
        if ($ssNorm) { $ss = $ssNorm }
    }

    # Default to the supported native implementation (QC.StatusSet.psm1). The previous
    # default was 'stub', which returned SUCCESS without doing any work whenever the
    # config lacked a statusSet block - jobs appeared to succeed in <1s while PW was
    # never updated and no PDF was written. Misconfigured/unknown modes now fail loudly
    # instead of silently succeeding.
    $mode = 'native'
    if ($ss.ContainsKey('mode') -and $ss.mode) { $mode = ([string]$ss.mode).Trim().ToLowerInvariant() }

    if ($mode -eq 'native') {
        $ssMod = Join-Path (_QCP-GetRepoRoot) 'modules\QC.StatusSet.psm1'
        if (-not (Test-Path -LiteralPath $ssMod)) {
            return New-QCFailureResult -Code 'STATUS_SET_MODULE_MISSING' -Message "QC.StatusSet.psm1 not found: $ssMod" -Data @{ path = $ssMod }
        }
        Import-Module $ssMod -Force
        return Invoke-StatusSetNativeJob -Job $Job -Config $Config
    }

    if ($mode -eq 'stub') {
        return New-QCSuccessResult -Code 'STATUS_SET_STUB_OK' -Message 'STATUS_SET_GEN explicitly configured as stub (statusSet.mode = "stub"); no work performed.' -Data @{
            jobId = [string]$Job.id
            jobType = [string]$Job.type
            sourceFolder = $sourceFolder
            mode = $mode
        }
    }

    if ($mode -ne 'legacy') {
        return New-QCFailureResult -Code 'STATUS_SET_UNKNOWN_MODE' -Message ("statusSet.mode '{0}' is not recognized. Valid values: native, legacy, stub." -f $mode) -Data @{
            jobId = [string]$Job.id
            jobType = [string]$Job.type
            sourceFolder = $sourceFolder
            mode = $mode
        }
    }

    if ($isDryRun) {
        return New-QCSuccessResult -Code 'STATUS_SET_DRYRUN' -Message 'Dry-run: would invoke legacy combine_status_set.ps1.' -Data @{
            jobId = [string]$Job.id
            jobType = [string]$Job.type
            sourceFolder = $sourceFolder
            mode = $mode
        }
    }

    $repoRoot = _QCP-GetRepoRoot
    $combineScript = if ($ss.ContainsKey('legacyScriptPath') -and $ss.legacyScriptPath) { [string]$ss.legacyScriptPath } else { (Join-Path $repoRoot 'legacy\combine_status_set.ps1') }
    if (-not (Test-Path -LiteralPath $combineScript)) {
        return New-QCFailureResult -Code 'STATUS_SET_LEGACY_SCRIPT_MISSING' -Message "Legacy combine_status_set.ps1 not found: $combineScript" -Data @{ script = $combineScript }
    }

    $datasource = if ($ss.ContainsKey('datasourceName') -and $ss.datasourceName) { [string]$ss.datasourceName } else { 'typsa-us-pw.bentley.com:typsa-us-pw-03' }
    $localRoot = if ($ss.ContainsKey('localRoot') -and $ss.localRoot) { [string]$ss.localRoot } else { 'C:\PW_QC_LOCAL' }
    $logDir = if ($ss.ContainsKey('logDir') -and $ss.logDir) { [string]$ss.logDir } else { '' }
    $qpdfExe = if ($ss.ContainsKey('qpdfExe') -and $ss.qpdfExe) { [string]$ss.qpdfExe } else { 'qpdf' }
    $forceRebuild = $false
    if ($ss.ContainsKey('forceRebuild')) { try { $forceRebuild = [bool]$ss.forceRebuild } catch { $forceRebuild = $false } }
    $promptForCredential = $false
    if ($ss.ContainsKey('promptForCredential')) { try { $promptForCredential = [bool]$ss.promptForCredential } catch { $promptForCredential = $false } }
    $writeBackToPW = $false
    if ($ss.ContainsKey('writeBackToPW')) { try { $writeBackToPW = [bool]$ss.writeBackToPW } catch { $writeBackToPW = $false } }

    $args = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $combineScript,
        '-SheetsFolderPath', $sourceFolder,
        '-DatasourceName', $datasource,
        '-LocalRoot', $localRoot,
        '-QpdfExe', $qpdfExe,
        '-PollIntervalSeconds', '0',
        '-RunOnce'
    )
    if (-not (_QCP-IsNullOrWhiteSpace $logDir)) { $args += @('-LogDir', $logDir) }
    if ($writeBackToPW) { $args += @('-WriteBackToPW') }
    if ($forceRebuild) { $args += @('-ForceRebuild') }
    if ($promptForCredential) { $args += @('-PromptForCredential') }

    $stdoutPath = [System.IO.Path]::GetTempFileName()
    $stderrPath = [System.IO.Path]::GetTempFileName()
    try {
        # combine_status_set.ps1 re-launches itself in MTA if needed; but call powershell.exe -MTA anyway.
        $p = Start-Process -FilePath 'powershell.exe' -ArgumentList (@('-MTA') + $args) -NoNewWindow -Wait -PassThru -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
        $stdout = Get-Content -LiteralPath $stdoutPath -Raw -ErrorAction SilentlyContinue
        $stderr = Get-Content -LiteralPath $stderrPath -Raw -ErrorAction SilentlyContinue

        if ($p.ExitCode -ne 0) {
            return New-QCFailureResult -Code 'STATUS_SET_LEGACY_FAILED' -Message 'Legacy combine_status_set.ps1 failed.' -Data @{
                jobId = [string]$Job.id
                exitCode = [int]$p.ExitCode
                stdout = $stdout
                stderr = $stderr
                script = $combineScript
                args = $args
                writeBackToPW = $writeBackToPW
            }
        }

        return New-QCSuccessResult -Code 'STATUS_SET_OK' -Message 'STATUS_SET_GEN completed via legacy combine_status_set.ps1.' -Data @{
            jobId = [string]$Job.id
            exitCode = [int]$p.ExitCode
            stdout = $stdout
            stderr = $stderr
            script = $combineScript
            args = $args
            writeBackToPW = $writeBackToPW
        }
    } finally {
        Remove-Item -LiteralPath $stdoutPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $stderrPath -Force -ErrorAction SilentlyContinue
    }
}

function Invoke-QCReportingScanProcessor {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    $mod = Join-Path (_QCP-GetRepoRoot) 'modules\QC.Reporting.psm1'
    if (-not (Test-Path -LiteralPath $mod)) {
        return New-QCFailureResult -Code 'QC_REPORTING_MODULE_MISSING' -Message "QC.Reporting.psm1 not found: $mod" -Data @{ path = $mod }
    }
    Import-Module $mod -Force
    return Invoke-QCReportingScan -Job $Job -Config $Config
}

Export-ModuleMember -Function *
