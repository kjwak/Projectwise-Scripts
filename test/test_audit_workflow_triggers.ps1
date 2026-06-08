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

$cfgNotify = @{
    auditPoller = @{
        workflowTriggers = @{
            notifyOnStateChange = $true
            qcPdfNotificationsOnly = $true
        }
    }
}
Assert-True (Test-QCShouldNotifyForSheetPackageMember -Config $cfgNotify -DocumentName '00-100000-00-00-qc.pdf') 'qc pdf member may notify'
Assert-True (-not (Test-QCShouldNotifyForSheetPackageMember -Config $cfgNotify -DocumentName '00-100000-00-00.pdf')) 'sheet pdf member suppressed when qcPdfNotificationsOnly'

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

Assert-Eq (Get-QCSheetGroupTransitionKey -SheetStem 'sheet1' -DocumentGuid 'guid-a' -TargetState 'QC Initiated' `
    -TransitionSource 'user_audit' -AuditEventId 100) `
    'sg|sheet1|guid-a|qc initiated|audit:100|user_audit' 'sheet-group transition key is stable'

# Processor/prepend sheet-group telemetry: see test_sheet_group_workflow_transition.ps1 (test 7).

Write-Host 'test_audit_workflow_triggers.ps1 passed'
