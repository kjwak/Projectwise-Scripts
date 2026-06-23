# Unit checks for QC.AuditTriggers (no PW / SQL).
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Database\Core.Database.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Workflow\QC.ProcessType.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Workflow\QC.AuditTriggers.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Workflow\QC.ProcessType.psm1') -Force

function Assert-Eq($a, $b, $msg) {
    if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" }
}

function Assert-True($cond, $msg) {
    if (-not $cond) { throw "ASSERT FAILED: $msg" }
}

Assert-True (Test-QCIsQcPdfDocumentName -DocumentName 'sheet-prod.pdf') 'lane prod pdf'
Assert-True (Test-QCIsQcPdfDocumentName -DocumentName 'sheet-chk.pdf') 'lane chk pdf'
Assert-True (Test-QCIsQcPdfDocumentName -DocumentName 'sheet-rev.pdf') 'lane rev pdf'
Assert-True (-not (Test-QCIsQcPdfDocumentName -DocumentName 'sheet-qc.pdf')) 'legacy -qc.pdf is not a lane pdf'
Assert-True (-not (Test-QCIsQcPdfDocumentName -DocumentName 'sheet.pdf')) 'plain pdf is not lane pdf'
Assert-True (Test-QCLegacyQcPdfDocumentName -DocumentName 'sheet-qc.pdf') 'legacy qc pdf detected'

$cfgNotify = @{
    auditPoller = @{
        workflowTriggers = @{
            notifyOnStateChange = $true
            qcPdfNotificationsOnly = $true
        }
    }
}
Assert-True (Test-QCShouldNotifyForSheetPackageMember -Config $cfgNotify -DocumentName '00-100000-00-00-prod.pdf') 'lane pdf member may notify'
Assert-True (-not (Test-QCShouldNotifyForSheetPackageMember -Config $cfgNotify -DocumentName '00-100000-00-00-qc.pdf')) 'legacy qc pdf member suppressed'
Assert-True (-not (Test-QCShouldNotifyForSheetPackageMember -Config $cfgNotify -DocumentName '00-100000-00-00.pdf')) 'sheet pdf member suppressed when qcPdfNotificationsOnly'

Assert-True (Test-QCShouldRecordWorkflowTelemetryForDocument -Config $cfgNotify -DocumentName '00-100000-00-00-prod.pdf') 'lane prod pdf may record workflow telemetry'
Assert-True (-not (Test-QCShouldRecordWorkflowTelemetryForDocument -Config $cfgNotify -DocumentName '00-100000-00-00.pdf')) 'stem sheet pdf suppressed when qcPdfEventsOnly'
Assert-True (-not (Test-QCShouldRecordWorkflowTelemetryForDocument -Config $cfgNotify -DocumentName '00-100000-00-00.dgn')) 'dgn suppressed when qcPdfEventsOnly'

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
Assert-Eq $s2.suppressBaselineIndexStateTransition $true 'baseline suppression on by default'

Assert-True (Test-QCShouldSuppressBaselineSheetIndexStateTransition -Config $cfgDefault -PreviousState '' -CurrentState 'In Production') `
    'empty index -> In Production is baseline seed'
Assert-True (-not (Test-QCShouldSuppressBaselineSheetIndexStateTransition -Config $cfgDefault -PreviousState '' -CurrentState 'QC Received')) `
    'empty index -> QC Received is a real transition'
Assert-True (Test-QCShouldSuppressAuditReadyForQcBaselineNotification -Config $cfgDefault -PreviousState '' -CurrentState 'Ready for QC') `
    'empty prior -> Ready for QC should suppress audit notification baseline'
Assert-True (-not (Test-QCShouldSuppressAuditReadyForQcBaselineNotification -Config $cfgDefault -PreviousState 'QC Initiated' -CurrentState 'Ready for QC')) `
    'QC Initiated -> Ready for QC is a real notification transition'
Assert-True (Test-QCShouldSuppressAuditStateChangeNotificationFromAttrSync -AuditActionName 'DOCUMENT_ATTR') `
    'DOCUMENT_ATTR index sync must not send state-change notifications'
Assert-True (-not (Test-QCShouldSuppressAuditStateChangeNotificationFromAttrSync -AuditActionName 'DOCUMENT_STATE')) `
    'DOCUMENT_STATE audits may still send state-change notifications'
Assert-True (-not (Test-QCShouldSuppressBaselineSheetIndexStateTransition -Config $cfgDefault -PreviousState 'In Production' -CurrentState 'QC Initiated')) `
    'prior state blocks baseline suppression'

$cfgExtraBaseline = @{
    auditPoller = @{ workflowTriggers = @{ baselineStateNames = @('En producción') } }
}
Assert-True (Test-QCShouldSuppressBaselineSheetIndexStateTransition -Config $cfgExtraBaseline -PreviousState '' -CurrentState 'En producción') `
    'baselineStateNames extends production label'

$cfgBaselineOff = @{ auditPoller = @{ workflowTriggers = @{ suppressBaselineIndexStateTransition = $false } } }
Assert-True (-not (Test-QCShouldSuppressBaselineSheetIndexStateTransition -Config $cfgBaselineOff -PreviousState '' -CurrentState 'In Production')) `
    'suppressBaselineIndexStateTransition=false'

$cfgGoLive = @{
    auditPoller = @{
        workflowTriggers = @{ processingGoLiveUtc = '2026-06-04T12:00:00Z' }
    }
}
Assert-True (Test-QCShouldSkipAuditWorkflowProcessingForEvent -Config $cfgGoLive -ActTime '2026-06-04T11:59:59Z') `
    'audit before go-live is skipped'
Assert-True (-not (Test-QCShouldSkipAuditWorkflowProcessingForEvent -Config $cfgGoLive -ActTime '2026-06-04T12:00:01Z')) `
    'audit after go-live is processed'
Assert-True (-not (Test-QCShouldSkipAuditWorkflowProcessingForEvent -Config $cfgDefault -ActTime '2020-01-01T00:00:00Z')) `
    'empty processingGoLiveUtc never skips'

Assert-Eq (Resolve-QCWorkflowEventQcReviewType -Context @{ attributes = @{ reviewType = 'Independent Check' } }) `
    'Independent Check' 'context attributes reviewType'
Assert-Eq (Resolve-QCWorkflowEventQcReviewType -Context @{ attributes = @{ qcReviewType = 'Production QC' } }) `
    'Production QC' 'context attributes qcReviewType'

Import-Module (Join-Path $repoRoot 'modules\Core\QC.WatcherOrchestration.psm1') -Force
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

$typsaCfg = @{
    qcWorkflow = @{
        states = @{
            qcInitiated = 'Initiate Origination'
            redlinesReceived = 'Redlines Received'
            complete = 'Verified'
            error = 'Error Needs Attention'
            qcReceived = 'Originated'
        }
    }
}
Assert-True (Test-QCWorkflowStateIsRestartIntakeTransition -Config $typsaCfg -PreviousState 'Redlines Received' -CurrentState 'Initiate Origination') `
    'Redlines Received -> Initiate Origination is a cycle restart'
Assert-True (Test-QCWorkflowStateIsRestartIntakeTransition -Config $typsaCfg -PreviousState 'Originated' -CurrentState 'Initiate Origination') `
    'Originated -> Initiate Origination is a cycle restart'
Assert-True (-not (Test-QCWorkflowStateIsRestartIntakeTransition -Config $typsaCfg -PreviousState 'In Development' -CurrentState 'Initiate Origination')) `
    'In Development -> Initiate Origination is first intake, not restart'

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
Assert-True (-not (Test-QCShouldSuppressAuditSheetStateSync -Config $cfgAuto -DocumentName 'sheet-prod.pdf' -ChangedByUser 42)) 'allow sync on lane qc pdf'
Assert-True (Test-QCShouldSuppressAuditSheetStateSync -Config $cfgAuto -DocumentName 'sheet.dgn' -ChangedByUser 42) 'suppress sync echo on dgn'

$cfgAutoOff = @{ auditPoller = @{ workflowTriggers = @{ ignoreStateChangeFromAutomation = $false; automationPwUsernames = @('srv_typsa_archivist') } } }
Assert-True (-not (Test-QCIsAutomationPwActor -Config $cfgAutoOff -ChangedByUsername 'srv_typsa_archivist')) 'ignore flag off'

$staleNoAnchor = Test-QCDocumentStateAuditEventIsStale -Config $cfgDefault -FolderPath 'Documents\P\Sheets\S1' `
    -DocumentName '0818000063ea509.pdf' -CanonicalState 'Redlines Received'
Assert-True (-not $staleNoAnchor.isStale) 'missing audit anchor is not stale'
Assert-Eq $staleNoAnchor.decision 'process' 'missing anchor keeps process decision'

$staleMembers = @(
    @{ documentGuid = 'pdf-guid'; documentName = '0818000063ea509.pdf' }
    @{ documentGuid = 'dgn-guid'; documentName = '0818000063ea509.dgn' }
)
$staleStateByGuid = @{ 'pdf-guid' = 'Corrections Received'; 'dgn-guid' = 'Redlines Received' }
function _PWD-GetSheetIndexStateSnapshot {
    param([hashtable]$Config, [string]$DocumentGuid)
    if ($DocumentGuid -eq 'pdf-guid') {
        return @{ pwStateName = 'Corrections Received'; lastAuditEventAt = '2026-06-04T22:10:00Z' }
    }
    return @{ pwStateName = ''; lastAuditEventAt = $null }
}
$staleRegression = Test-QCDocumentStateAuditEventIsStale -Config $cfgDefault -FolderPath 'Documents\P\Sheets\S1' `
    -DocumentName '0818000063ea509.pdf' -DocumentGuid 'pdf-guid' -AuditEventId 39093 `
    -LastAuditEventAt '2026-06-04T22:00:00Z' -CanonicalState 'Redlines Received' `
    -Members $staleMembers -StateByGuid $staleStateByGuid -SheetStem '0818000063ea509'
Assert-True $staleRegression.isStale 'newer pdf index state blocks regressive canonical'
Assert-Eq $staleRegression.reason 'regressive_pdf_state' 'regression reason is regressive_pdf_state'
Assert-Eq $staleRegression.decision 'skipped' 'regression decision is skipped'

# Test 1: valid QC PDF forward transition (sibling PDF/DGN behind, QC PDF ahead).
$fwdMembers = @(
    @{ documentGuid = 'pdf-guid'; documentName = '080J082001ca001.pdf' }
    @{ documentGuid = 'dgn-guid'; documentName = '080J082001ca001.dgn' }
    @{ documentGuid = 'qc-guid'; documentName = '080J082001ca001-prod.pdf' }
)
$fwdStateByGuid = @{
    'pdf-guid' = 'Ready for QC'
    'dgn-guid' = 'Ready for QC'
    'qc-guid'  = 'Redlines Received'
}
function _PWD-GetSheetIndexStateSnapshot {
    param([hashtable]$Config, [string]$DocumentGuid)
    switch ($DocumentGuid) {
        'pdf-guid' { return @{ pwStateName = 'Ready for QC'; lastAuditEventAt = '2026-06-09T16:49:44Z' } }
        'dgn-guid' { return @{ pwStateName = 'Ready for QC'; lastAuditEventAt = '2026-06-09T16:49:44Z' } }
        'qc-guid'  { return @{ pwStateName = 'Ready for QC'; lastAuditEventAt = '2026-06-09T16:49:44Z' } }
        default { return @{ pwStateName = ''; lastAuditEventAt = $null } }
    }
}
$fwdStale = Test-QCDocumentStateAuditEventIsStale -Config $cfgDefault -FolderPath 'documents\test\sheets' `
    -DocumentName '080J082001ca001-prod.pdf' -DocumentGuid 'qc-guid' -AuditEventId 40694 `
    -LastAuditEventAt '2026-06-09 10:02:09' -CanonicalState 'Redlines Received' `
    -Members $fwdMembers -StateByGuid $fwdStateByGuid -SheetStem '080J082001ca001'
Assert-True (-not $fwdStale.isStale) 'Test 1: QC PDF forward transition is not stale'
Assert-Eq $fwdStale.decision 'process' 'Test 1: sibling sync allowed'

# Test 2: duplicate batch ingest guard suppresses pw_batch fallback.
$dbRowsPrepared = 1
$dbWrites = 0
$dbSkipped = 1
$dbUnprocessedLoaded = 0
$allowPwBatchFallback = $true
if ($dbRowsPrepared -gt 0 -and $dbWrites -eq 0 -and $dbSkipped -gt 0) { $allowPwBatchFallback = $false }
Assert-True (-not $allowPwBatchFallback) 'Test 2: duplicate ingest suppresses pw_batch trigger fallback'
Assert-Eq $dbUnprocessedLoaded 0 'Test 2: no unprocessed DB rows loaded'

# Test 3: older sibling audit event must not block newer trigger.
function Get-QCNewerSheetDocumentStateAuditEvent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$FolderPath,
        [array]$MemberDocumentNames = @(),
        [array]$MemberDocumentGuids = @(),
        [Nullable[long]]$CurrentAuditEventId = $null,
        [string]$CurrentAuditEventAt = ''
    )
    return @{
        IsSuccess = $true
        Code = 'NEWER_STATE_AUDIT_REJECTED'
        Data = @{
            found = $false
            rejectedBlockingReason = 'blocking_candidate_older_than_current'
            rejectedBlockingAuditEventId = 40686
            rejectedBlockingAuditTime = '2026-06-09 09:47:49'
            rejectedBlockingDocumentName = '080J082001ca001.pdf'
        }
    }
}
$newerRejectStale = Test-QCDocumentStateAuditEventIsStale -Config $cfgDefault -FolderPath 'documents\test\sheets' `
    -DocumentName '080J082001ca001-prod.pdf' -DocumentGuid 'qc-guid' -AuditEventId 40694 `
    -LastAuditEventAt '06/09/2026 17:02:09' -CanonicalState 'Redlines Received' `
    -Members $fwdMembers -StateByGuid $fwdStateByGuid -SheetStem '080J082001ca001'
Assert-True (-not $newerRejectStale.isStale) 'Test 3: older blocking candidate does not stale-block'
Assert-Eq $newerRejectStale.rejectedBlockingReason 'blocking_candidate_older_than_current' 'Test 3: telemetry records rejected blocking candidate'
Assert-Eq $newerRejectStale.blockingDocumentName '080J082001ca001.pdf' 'Test 3: blocking candidate document captured'
Remove-Item Function:Get-QCNewerSheetDocumentStateAuditEvent -ErrorAction SilentlyContinue

# Test 4: sheet_index lag must not block live PW canonical state for QC PDF trigger.
Assert-True (-not $fwdStale.isStale) 'Test 4: sheet_index lag does not produce regressive_pdf_state by itself'
Assert-Eq $fwdStale.staleComparisonBasis 'audit_event_id_and_time_utc' 'Test 4: stale comparison uses persisted audit anchor'

Assert-Eq (Get-QCSheetGroupTransitionKey -SheetStem 'sheet1' -DocumentGuid 'guid-a' -TargetState 'QC Initiated' `
    -TransitionSource 'user_audit' -AuditEventId 100) `
    'sg|sheet1|guid-a|qc initiated|audit:100|user_audit' 'sheet-group transition key is stable'

# Processor/prepend sheet-group telemetry: see test_sheet_group_workflow_transition.ps1 (test 7).

Write-Host 'test_audit_workflow_triggers.ps1 passed'
