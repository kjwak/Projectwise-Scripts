# QC.CommentStatusProcessor.psm1

# Responsibility: Thin orchestrator for QC_COMMENT_STATUS_SYNC jobs.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force

Import-Module (Join-Path $PSScriptRoot 'Core.Runtime.psm1') -Force -ErrorAction SilentlyContinue

Import-Module (Join-Path $PSScriptRoot 'Core.Hashing.psm1') -Force -ErrorAction SilentlyContinue

Import-Module (Join-Path $PSScriptRoot 'QC.PdfExport.psm1') -Force

Import-Module (Join-Path $PSScriptRoot 'QC.CommentExtract.psm1') -Force

Import-Module (Join-Path $PSScriptRoot 'QC.CommentStatusDecision.psm1') -Force

Import-Module (Join-Path $PSScriptRoot 'QC.CommentSync.Job.psm1') -Force

Import-Module (Join-Path $PSScriptRoot 'QC.CommentSync.State.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'QC.AuditTriggers.psm1') -Force -ErrorAction SilentlyContinue

Import-Module (Join-Path $PSScriptRoot 'QC.CommentSync.Database.psm1') -Force

Import-Module (Join-Path $PSScriptRoot 'QC.CommentSync.Notifications.psm1') -Force

Import-Module (Join-Path $PSScriptRoot 'QC.Workflow.psm1') -Force -ErrorAction SilentlyContinue

function _QCSP-Log([string]$Code, [string]$Message, [hashtable]$Data) {

    if (Get-Command -Name Write-QCJsonLog -ErrorAction SilentlyContinue) {

        Write-QCJsonLog -Level 'Information' -Code $Code -Message $Message -Data $Data

    }

}

function Test-QCCommentSyncStateApplySucceeded {

    param(

        [object]$StateRes,

        [hashtable]$Settings

    )

    if (-not $StateRes -or -not $StateRes.IsSuccess) { return $false }

    $failOnApply = $true

    if ($Settings -and $null -ne $Settings.failOnStateApply) {

        try { $failOnApply = [bool]$Settings.failOnStateApply } catch { }

    }

    if (-not $failOnApply) { return $true }

    $blockedCodes = @(

        'QC_WORKFLOW_STATE_TRANSITION_INVALID',

        'QC_WORKFLOW_STATE_UNAVAILABLE',

        'QC_WORKFLOW_STATE_MISSING',

        'QC_WORKFLOW_STATE_FAILED',

        'QC_COMMENT_STATE_MODULE_MISSING'

    )

    return -not ($blockedCodes -contains [string]$StateRes.Code)

}

function _QCSP-WriteWorkflowEvent {

    param(

        [hashtable]$Config,

        [long]$RunId = 0,

        [string]$JobId = '',

        [string]$DocumentId = '',

        [string]$EventType = '',

        [string]$PreviousPwState = '',

        [string]$TargetPwState = '',

        [string]$DecisionCode = '',

        [string]$ProcessorVersion = '',

        [string]$PayloadJson = '',

        [switch]$PlannedOnly

    )

    Write-QCWorkflowEvent -Config $Config -RunId $RunId -JobId $JobId -DocumentId $DocumentId `
        -EventType $EventType -PreviousPwState $PreviousPwState -TargetPwState $TargetPwState `
        -DecisionCode $DecisionCode -ProcessorVersion $ProcessorVersion -PayloadJson $PayloadJson `
        -PlannedOnly:$PlannedOnly | Out-Null

}

function Invoke-QCCommentStatusSyncProcessor {

    [CmdletBinding()]

    param(

        [Parameter(Mandatory)]

        [hashtable]$Job,

        [Parameter(Mandatory)]

        [hashtable]$Config,

        $InputDocument = $null,

        [scriptblock]$MockExporter = $null,

        [object[]]$MockAnnotations = $null

    )

    $settings = Get-QCCommentSyncSettings -Config $Config

    if (-not [bool]$settings.enabled) {

        return New-QCSuccessResult -Code 'QC_COMMENT_SYNC_DISABLED' -Message 'qcCommentSync.enabled is false.' -Data @{ jobId = [string]$Job.id }

    }

    $procVer = if ($settings.processorVersion) { [string]$settings.processorVersion } else { '1.0.0' }

    $dryRun = Test-QCCommentSyncDryRun -Config $Config

    $dbPlanned = -not (Test-QCCommentSyncDatabaseWritesAllowed -Config $Config)

    $eventPlanned = $dryRun -or $dbPlanned

    _QCSP-Log -Code 'QC_COMMENT_SYNC_START' -Message 'Comment status sync started.' -Data @{

        jobId = [string]$Job.id

        sourcePath = [string]$Job.sourcePath

        dryRun = $dryRun

        processorVersion = $procVer

    }

    $previousState = ''

    $document = $InputDocument

    if ($document -and (Get-Command -Name 'Get-PWDocumentWorkflowInfo' -ErrorAction SilentlyContinue)) {

        $info = Get-PWDocumentWorkflowInfo -Document $document -Context @{ job = $Job }

        if ($info.Data.stateName) { $previousState = [string]$info.Data.stateName }

    }

    $meta = New-QCCommentSyncJobMetadata -Job $Job -SourcePwState $previousState

    # 1. export

    if (-not $document) {

        $docRes = Get-QCCommentSyncPwDocument -Job $Job -Config $Config

        if (-not $docRes.IsSuccess) { return $docRes }

        $document = $docRes.Data.document

    }

    $exportRes = Export-QCPdfToStaging -InputDocument $document -Config $Config -JobId ([string]$Job.id) -MockExporter $MockExporter

    if (-not $exportRes.IsSuccess) {

        _QCSP-Log -Code 'QC_COMMENT_SYNC_EXPORT_FAILED' -Message $exportRes.Message -Data @{ jobId = [string]$Job.id }

        return $exportRes

    }

    $localPdf = [string]$exportRes.Data.localPath

    $staging = [string]$exportRes.Data.stagingFolder

    $meta.localDownloadPath = $localPdf

    _QCSP-Log -Code 'QC_COMMENT_SYNC_EXPORT_OK' -Message 'PDF exported.' -Data @{ localPath = $localPdf }

    if ((Get-Command -Name 'Get-Sha256FileHex' -ErrorAction SilentlyContinue) -and (Test-Path -LiteralPath $localPdf)) {

        try { $meta.fileHash = Get-Sha256FileHex -Path $localPdf } catch { }

    }

    # 2. extract

    $extractRes = Invoke-QCCommentExtract -LocalPdfPath $localPdf -Config $Config -MockAnnotations $MockAnnotations

    if (-not $extractRes.IsSuccess) {

        _QCSP-Log -Code 'QC_COMMENT_SYNC_PARSE_FAILED' -Message $extractRes.Message -Data @{ jobId = [string]$Job.id }

        _QCSP-WriteWorkflowEvent -Config $Config -JobId ([string]$Job.id) -DocumentId ([string]$meta.documentId) `
            -EventType 'PARSE_FAILED' -PreviousPwState $previousState -ProcessorVersion $procVer `
            -PayloadJson (@{ code = [string]$extractRes.Code; message = [string]$extractRes.Message } | ConvertTo-Json -Compress) `
            -PlannedOnly:$eventPlanned

        Remove-QCCommentSyncStaging -Config $Config -StagingFolder $staging | Out-Null

        return $extractRes

    }

    $annotations = @($extractRes.Data.annotations)

    $parserStatus = [string]$extractRes.Data.parserStatus

    if ($parserStatus -eq 'error') {

        $parseFail = New-QCFailureResult -Code 'QC_COMMENT_EXTRACT_PARSER_ERROR' -Message 'Parser reported error status.' -Data @{

            parserStatus = $parserStatus

            warnings = @($extractRes.Data.warnings)

        }

        _QCSP-Log -Code 'QC_COMMENT_SYNC_PARSE_FAILED' -Message $parseFail.Message -Data @{ jobId = [string]$Job.id; parserStatus = $parserStatus }

        _QCSP-WriteWorkflowEvent -Config $Config -JobId ([string]$Job.id) -DocumentId ([string]$meta.documentId) `
            -EventType 'PARSE_FAILED' -PreviousPwState $previousState -ProcessorVersion $procVer `
            -PayloadJson (@{ parserStatus = $parserStatus } | ConvertTo-Json -Compress) -PlannedOnly:$eventPlanned

        Remove-QCCommentSyncStaging -Config $Config -StagingFolder $staging | Out-Null

        return $parseFail

    }

    _QCSP-Log -Code 'QC_COMMENT_SYNC_PARSE_OK' -Message 'Comments parsed.' -Data @{

        count = $annotations.Count

        parserStatus = $parserStatus

    }

    # 3. decide (pure)

    $decision = Resolve-QCCommentWorkflowState -Annotations $annotations -Config $Config -ParserStatus $parserStatus

    _QCSP-Log -Code 'QC_COMMENT_SYNC_DECISION' -Message 'Workflow state decided.' -Data @{

        targetState = [string]$decision.targetState

        decisionCode = [string]$decision.decisionCode

        summary = [string]$decision.summary

    }

    # 4. persist (decision + comment rows; state result updated after apply)

    $stateResultCode = 'pending'

    $persist = Invoke-QCCommentSyncPersist -Config $Config -Job $Job -JobMetadata $meta -Decision $decision `
        -Annotations $annotations -ParserStatus $parserStatus -StateUpdateResult $stateResultCode `
        -ProcessorVersion $procVer -PreviousPwState $previousState -DryRun:$dryRun

    _QCSP-Log -Code $(if ($dbPlanned) { 'QC_COMMENT_SYNC_DB_PLANNED' } else { 'QC_COMMENT_SYNC_DB_OK' }) `
        -Message 'Database persist step completed.' -Data @{

            planned = $persist.planned

            runId = $persist.runId

        }

    # 5. update state

    $stateRes = Set-QCPdfCommentSyncWorkflowState -Config $Config -TargetStateName ([string]$decision.targetState) `
        -Document $document -Job $Job -PreviousStateName $previousState -DryRun:$dryRun

    if ($stateRes.IsSuccess) {

        $stateResultCode = [string]$stateRes.Code

        _QCSP-Log -Code $(if ($dryRun) { 'QC_COMMENT_SYNC_STATE_PLANNED' } else { 'QC_COMMENT_SYNC_STATE_OK' }) `
            -Message $stateRes.Message -Data @{

                targetState = [string]$decision.targetState

                previousState = $previousState

                code = $stateResultCode

            }

    } else {

        $stateResultCode = [string]$stateRes.Code

        _QCSP-Log -Code 'QC_COMMENT_SYNC_STATE_FAILED' -Message $stateRes.Message -Data @{ code = $stateResultCode }

    }

    $stateApplyOk = Test-QCCommentSyncStateApplySucceeded -StateRes $stateRes -Settings $settings

    if ($stateApplyOk -and -not $dryRun -and (Get-Command -Name 'Invoke-QCProcessorWorkflowStateTelemetry' -ErrorAction SilentlyContinue)) {
        $telemetryCtx = @{
            config       = $Config
            job          = $Job
            document     = $document
            documentGuid = [string]$meta.documentId
            folderPath   = if ($meta.folderPath) { [string]$meta.folderPath } else { '' }
        }
        Invoke-QCProcessorWorkflowStateTelemetry -Config $Config -Context $telemetryCtx `
            -PreviousState $previousState -CurrentState ([string]$decision.targetState) `
            -JobType 'QC_COMMENT_STATUS_SYNC' | Out-Null
    }

    if ($persist.runId -gt 0) {

        Update-QCCommentSyncRunStateResult -Config $Config -RunId $persist.runId -StateUpdateResult $stateResultCode -PlannedOnly:$eventPlanned | Out-Null

    }

    $stateEventType = if ($dryRun) { 'STATE_SKIPPED_DRYRUN' } elseif ($stateApplyOk) { 'STATE_APPLIED' } else { 'STATE_APPLY_FAILED' }

    _QCSP-WriteWorkflowEvent -Config $Config -RunId $persist.runId -JobId ([string]$Job.id) -DocumentId ([string]$meta.documentId) `
        -EventType $stateEventType -PreviousPwState $previousState -TargetPwState ([string]$decision.targetState) `
        -DecisionCode ([string]$decision.decisionCode) -ProcessorVersion $procVer -PlannedOnly:$eventPlanned

    # 6. notify

    $notifyRes = Send-QCCommentSyncNotification -Config $Config -Decision $decision -JobMetadata $meta `
        -Document $document -Job $Job -PreviousState $previousState

    $notifyEventType = if ($notifyRes.Data.planned) { 'NOTIFY_PLANNED' } elseif ($notifyRes.IsSuccess) { 'NOTIFY_OK' } else { 'NOTIFY_FAILED' }

    _QCSP-WriteWorkflowEvent -Config $Config -RunId $persist.runId -JobId ([string]$Job.id) -DocumentId ([string]$meta.documentId) `
        -EventType $notifyEventType -PreviousPwState $previousState -TargetPwState ([string]$decision.targetState) `
        -DecisionCode ([string]$decision.decisionCode) -ProcessorVersion $procVer `
        -PayloadJson (@{ code = [string]$notifyRes.Code; message = [string]$notifyRes.Message } | ConvertTo-Json -Compress) `
        -PlannedOnly:($eventPlanned -or $notifyRes.Data.planned)

    _QCSP-Log -Code $(if ($notifyRes.Data.planned) { 'QC_COMMENT_SYNC_NOTIFY_PLANNED' } elseif ($notifyRes.IsSuccess) { 'QC_COMMENT_SYNC_NOTIFY_OK' } else { 'QC_COMMENT_SYNC_NOTIFY_FAILED' }) `
        -Message $notifyRes.Message -Data @{ code = [string]$notifyRes.Code }

    Remove-QCCommentSyncStaging -Config $Config -StagingFolder $staging | Out-Null

    if (-not $stateApplyOk) {

        return New-QCFailureResult -Code 'QC_COMMENT_SYNC_STATE_FAILED' -Message $stateRes.Message -Data @{

            jobId = [string]$Job.id

            decision = $decision

            parserStatus = $parserStatus

            annotationCount = $annotations.Count

            runId = $persist.runId

            stateResult = $stateRes

            notifyResult = $notifyRes

            dryRun = $dryRun

            processorVersion = $procVer

        }

    }

    return New-QCSuccessResult -Code 'QC_COMMENT_SYNC_COMPLETE' -Message 'Comment status sync completed.' -Data @{

        jobId = [string]$Job.id

        decision = $decision

        parserStatus = $parserStatus

        annotationCount = $annotations.Count

        runId = $persist.runId

        stateResult = $stateRes

        notifyResult = $notifyRes

        dryRun = $dryRun

        processorVersion = $procVer

    }

}

Export-ModuleMember -Function Invoke-QCCommentStatusSyncProcessor, Test-QCCommentSyncStateApplySucceeded

