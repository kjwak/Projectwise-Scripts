# Unit checks for QC.AuditTriggers (no PW / SQL).
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Database.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.AuditTriggers.psm1') -Force

function Assert-Eq($a, $b, $msg) {
    if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" }
}

function Assert-True($cond, $msg) {
    if (-not $cond) { throw "ASSERT FAILED: $msg" }
}

Assert-True (Test-QCIsQcPdfDocumentName -DocumentName 'sheet-qc.pdf') 'qc pdf suffix'
Assert-True (-not (Test-QCIsQcPdfDocumentName -DocumentName 'sheet.pdf')) 'plain pdf is not qc pdf'

$cfg = @{
    auditPoller = @{
        workflowTriggers = @{
            enabled = $false
            recordStateHistory = $true
        }
    }
}
$s = Get-QCAuditWorkflowTriggerSettings -Config $cfg
Assert-Eq $s.enabled $false 'workflowTriggers.enabled respected'
Assert-Eq $s.recordStateHistory $true 'other flags still parsed'

$cfgDefault = @{ auditPoller = @{} }
$s2 = Get-QCAuditWorkflowTriggerSettings -Config $cfgDefault
Assert-Eq $s2.enabled $true 'defaults enable workflow triggers'
Assert-Eq $s2.recordFromProcessor $true 'defaults enable processor telemetry'

Assert-Eq (Resolve-QCWorkflowEventQcReviewType -Context @{ attributes = @{ reviewType = 'Independent Check' } }) `
    'Independent Check' 'context attributes reviewType'
Assert-Eq (Resolve-QCWorkflowEventQcReviewType -Context @{ attributes = @{ qcReviewType = 'Production QC' } }) `
    'Production QC' 'context attributes qcReviewType'

Import-Module (Join-Path $repoRoot 'modules\QC.WatcherOrchestration.psm1') -Force
$prependActions = Get-QCPrependAuditActions -Config $cfgDefault
Assert-True ($prependActions -contains 'DOCUMENT_ATTR') 'default prepend actions include DOCUMENT_ATTR'
Assert-True ($prependActions -contains 'DOCUMENT_STATE') 'default prepend actions include DOCUMENT_STATE'

Assert-Eq (Get-QCAuditStateTransitionKey -AuditEventId 99 -TriggerDocumentGuid 'g') 'audit:99' 'audit event id wins'
Assert-Eq (Get-QCAuditStateTransitionKey -TransitionId 12 -TriggerDocumentGuid 'g') 'transition:12' 'transition id when no audit id'

Assert-Eq (Get-QCPrependStateTransitionDedupeKey -AuditEventId 9001 -SheetStem '080J082001ab001' -PreviousSheetState 'In Production' -TargetStateName 'QC Initiated') `
    'audit:9001' 'prepend dedupe prefers audit id (sibling echoes share one key)'
Assert-Eq (Get-QCPrependStateTransitionDedupeKey -AuditEventId 9002 -SheetStem '080J082001ab001' -PreviousSheetState 'In Production' -TargetStateName 'QC Initiated') `
    'audit:9002' 'new audit event allows another QC cycle prepend'
Assert-Eq (Get-QCPrependStateTransitionDedupeKey -SheetStem '080J082001ab001' -PreviousSheetState 'Ready for QC' -TargetStateName 'QC Initiated' -PrependTrigger 'initialQcPdf') `
    'sheet:080j082001ab001|trigger:initialqcpdf|from:ready for qc|to:qc initiated' 'prepend dedupe falls back to sheet transition without audit id (initiated)'
Assert-Eq (Get-QCPrependStateTransitionDedupeKey -AuditEventId 9101 -SheetStem '080J082001ab001' -PreviousSheetState 'Corrections Received' -TargetStateName 'QC Finalizing' -PrependTrigger 'finalQcComplete') `
    'audit:9101' 'QC Finalizing prepend dedupe prefers audit id (sibling echoes share one key)'
Assert-Eq (Get-QCPrependStateTransitionDedupeKey -AuditEventId 9102 -SheetStem '080J082001ab001' -PreviousSheetState 'Corrections Received' -TargetStateName 'QC Finalizing' -PrependTrigger 'finalQcComplete') `
    'audit:9102' 'another QC Finalizing transition gets a new prepend job'
Assert-Eq (Get-QCInitiatedWorkflowStateName -Config $cfgDefault) 'QC Initiated' 'default initiated state name'
Assert-True (Test-QCWorkflowStateIsQcInitiated -StateName 'QC Initiated' -Config $cfgDefault) 'qc initiated match'
Assert-True (-not (Test-QCWorkflowStateIsQcInitiated -StateName 'In Production' -Config $cfgDefault)) 'non-initiated state'

$cfgDbOff = @{ database = @{ enabled = $false }; auditPoller = @{ workflowTriggers = @{ enabled = $true } } }
Invoke-QCAuditWorkflowStateChangeTriggers -Config $cfgDbOff -DocumentGuid 'g1' -DocumentName 'a-qc.pdf' `
    -FolderPath 'Documents\X\CADD\Sheets' -PreviousState 'In Production' -CurrentState 'QC Received' | Out-Null
Invoke-QCAuditWorkflowAttributeChangeTriggers -Config $cfgDbOff -DocumentGuid 'g1' -DocumentName 'a.pdf' `
    -FolderPath 'Documents\X\CADD\Sheets' -FieldChanges @{ designer_email = @{ oldValue = 'a@x.com'; newValue = 'b@x.com' } } | Out-Null

$cfgAuto = @{
    auditPoller = @{
        workflowTriggers = @{
            enabled = $true
            ignoreStateChangeFromAutomation = $true
            automationPwUsernames = @('srv_typsa_archivist')
            automationPwUserNumbers = @(42)
            notifyOnStateChange = $true
        }
    }
}
Assert-True (Test-QCIsAutomationPwActor -Config $cfgAuto -ChangedByUser 42) 'userno match'
Assert-True (Test-QCIsAutomationPwActor -Config $cfgAuto -ChangedByUsername 'srv_typsa_archivist') 'username match'
Assert-True (-not (Test-QCIsAutomationPwActor -Config $cfgAuto -ChangedByUser 99 -ChangedByUsername 'human')) 'non-automation'
Assert-True (-not (Test-QCShouldSuppressAuditSheetStateSync -Config $cfgAuto -DocumentName 'sheet-qc.pdf' -ChangedByUser 42)) 'allow sync on qc pdf'
Assert-True (Test-QCShouldSuppressAuditSheetStateSync -Config $cfgAuto -DocumentName 'sheet.dgn' -ChangedByUser 42) 'suppress sync echo on dgn'

$cfgAutoOff = @{ auditPoller = @{ workflowTriggers = @{ ignoreStateChangeFromAutomation = $false; automationPwUsernames = @('srv_typsa_archivist') } } }
Assert-True (-not (Test-QCIsAutomationPwActor -Config $cfgAutoOff -ChangedByUsername 'srv_typsa_archivist')) 'ignore flag off'

# Final QC prepend success must record QC Finalizing -> QC Complete in transition_events.
$script:finalPrependTransitions = [System.Collections.Generic.List[object]]::new()
$script:finalPrependWorkflowEvents = [System.Collections.Generic.List[object]]::new()
function Write-QCTransitionEvent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [Parameter(Mandatory)][string]$TransitionType,
        [string]$DocumentName = '',
        [string]$FolderPath = '',
        [string]$FromValue = '',
        [string]$ToValue = '',
        [string]$JobId = '',
        [string]$JobType = '',
        [Nullable[long]]$TriggerAuditId = $null
    )
    $script:finalPrependTransitions.Add(@{
        documentGuid = $DocumentGuid
        transitionType = $TransitionType
        fromValue = $FromValue
        toValue = $ToValue
        jobId = $JobId
        jobType = $JobType
    }) | Out-Null
    return [pscustomobject]@{
        IsSuccess = $true
        Code = 'TRANSITION_EVENT_WRITTEN'
        Message = 'Captured for test.'
        Data = @{ written = $true; transitionId = 99 }
    }
}

function Write-QCWorkflowEventRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$DocumentId = '',
        [string]$JobId = '',
        [string]$EventType = '',
        [string]$PreviousPwState = '',
        [string]$TargetPwState = '',
        [string]$DecisionCode = '',
        [string]$PayloadJson = '',
        [string]$QcReviewType = ''
    )
    $payload = $null
    try { if ($PayloadJson) { $payload = $PayloadJson | ConvertFrom-Json } } catch { }
    $script:finalPrependWorkflowEvents.Add(@{
        documentId = $DocumentId
        eventType = $EventType
        previousPwState = $PreviousPwState
        targetPwState = $TargetPwState
        decisionCode = $DecisionCode
        qcReviewType = $QcReviewType
        payload = $payload
    }) | Out-Null
    return [pscustomobject]@{
        IsSuccess = $true
        Code = 'QC_WORKFLOW_EVENT_WRITTEN'
        Message = 'Captured for test.'
        Data = @{ written = $true }
    }
}

$cfgFinalPrepend = @{
    database = @{ enabled = $true; allowWritesInDryRun = $true }
    auditPoller = @{
        workflowTriggers = @{
            enabled = $true
            recordFromProcessor = $true
            recordTransitions = $true
            recordStateHistory = $false
            recordProcessingJobs = $false
        }
    }
}
$finalCtx = @{
    job = @{
        id = 'qc_prepend_final_test'
        sourceFolder = 'Documents\X\CADD\Sheets'
        metadata = @{
            triggerDocumentGuid = 'guid-final-prepend'
            triggerDocumentName = 'A101.pdf'
            pwStateName = 'QC Finalizing'
            prependTrigger = 'finalQcComplete'
        }
    }
    previousState = 'QC Finalizing'
    lifecycleState = 'QC Finalizing'
    documentGuid = 'guid-final-prepend'
    documentName = 'A101.pdf'
    attributes = @{
        reviewType = 'Independent Check'
    }
}
Invoke-QCProcessorWorkflowStateTelemetry -Config $cfgFinalPrepend -Context $finalCtx `
    -PreviousState 'QC Finalizing' -CurrentState 'QC Complete' -JobType 'QC_PREPEND'
Assert-Eq $script:finalPrependTransitions.Count 1 'final prepend should write one transition event'
Assert-Eq $script:finalPrependTransitions[0].fromValue 'QC Finalizing' 'transition from QC Finalizing'
Assert-Eq $script:finalPrependTransitions[0].toValue 'QC Complete' 'transition to QC Complete'
Assert-Eq $script:finalPrependTransitions[0].jobType 'QC_PREPEND' 'transition job type QC_PREPEND'
Assert-Eq $script:finalPrependTransitions[0].documentGuid 'guid-final-prepend' 'transition uses trigger document guid'
Assert-Eq $script:finalPrependWorkflowEvents.Count 1 'final prepend should mirror one workflow event'
Assert-Eq $script:finalPrependWorkflowEvents[0].qcReviewType 'Independent Check' 'workflow event includes qc_review_type column value'

Write-Host 'test_audit_workflow_triggers.ps1 passed'
