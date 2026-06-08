# Sheet-group workflow transition telemetry (no PW / SQL runtime).
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

$script:transitionCalls = [System.Collections.Generic.List[object]]::new()
$script:historyCalls = [System.Collections.Generic.List[object]]::new()
$script:workflowEvents = [System.Collections.Generic.List[object]]::new()
$script:notificationCalls = 0
$script:transitionReuseCount = 0

function Write-QCDocumentStateHistoryRow {
    param(
        [hashtable]$Config, [string]$DocumentGuid, [string]$EventType,
        [string]$DocumentName = '', [string]$FolderPath = '',
        [string]$OldValue = '', [string]$NewValue = '', [string]$FieldName = '',
        [Nullable[int]]$ChangedByUser = $null, [string]$ChangedByUsername = '',
        [Nullable[long]]$SourceAuditId = $null
    )
    $script:historyCalls.Add(@{
        documentGuid = $DocumentGuid
        documentName = $DocumentName
        oldValue = $OldValue
        newValue = $NewValue
        sourceAuditId = $SourceAuditId
    }) | Out-Null
    return [pscustomobject]@{ IsSuccess = $true; Data = @{ written = $true } }
}

function Ensure-QCTransitionEvent {
    param(
        [hashtable]$Config, [string]$DocumentGuid, [string]$TransitionType,
        [string]$DocumentName = '', [string]$FolderPath = '',
        [string]$FromValue = '', [string]$ToValue = '',
        [string]$JobId = '', [string]$JobType = '',
        [Nullable[long]]$TriggerAuditId = $null
    )
    $key = ($DocumentGuid + '|' + $FromValue + '|' + $ToValue + '|' + [string]$TriggerAuditId + '|' + $JobId)
    $existing = @($script:transitionCalls | Where-Object {
        ($_.documentGuid + '|' + $_.fromValue + '|' + $_.toValue + '|' + [string]$_.triggerAuditId + '|' + $_.jobId) -eq $key
    })
    if ($existing.Count -gt 0) {
        $script:transitionReuseCount++
        return [pscustomobject]@{
            IsSuccess = $true
            Data = @{ written = $false; reused = $true; transitionId = 1000 + $script:transitionCalls.Count }
        }
    }
    $script:transitionCalls.Add(@{
        documentGuid = $DocumentGuid
        documentName = $DocumentName
        transitionType = $TransitionType
        fromValue = $FromValue
        toValue = $ToValue
        jobId = $JobId
        jobType = $JobType
        triggerAuditId = $TriggerAuditId
    }) | Out-Null
    return [pscustomobject]@{
        IsSuccess = $true
        Data = @{ written = $true; reused = $false; transitionId = 1000 + $script:transitionCalls.Count }
    }
}

function Write-QCWorkflowEventRow {
    param(
        [hashtable]$Config, [string]$DocumentId = '', [string]$JobId = '',
        [string]$EventType = '', [string]$PreviousPwState = '', [string]$TargetPwState = '',
        [string]$DecisionCode = '', [string]$PayloadJson = '', [string]$QcReviewType = '',
        [Nullable[int]]$TransitionEventId = $null
    )
    $script:workflowEvents.Add(@{
        documentId = $DocumentId
        eventType = $EventType
        previousPwState = $PreviousPwState
        targetPwState = $TargetPwState
        decisionCode = $DecisionCode
        transitionEventId = $TransitionEventId
    }) | Out-Null
    return [pscustomobject]@{ IsSuccess = $true; Data = @{ written = $true } }
}

function Invoke-QCWorkflowStateChangeNotification {
    param([hashtable]$Config, [hashtable]$Context, [string]$PreviousState, [string]$CurrentState, $Document)
    $script:notificationCalls++
    return [pscustomobject]@{ IsSuccess = $true; Code = 'QC_NOTIFICATION_SENT' }
}

function _QCAT-BuildNotificationDocument {
    param([hashtable]$Config, [string]$FolderPath, [string]$DocumentName, [string]$DocumentGuid, [hashtable]$Attributes)
    return [pscustomobject]@{ Name = $DocumentName; DocumentGUID = $DocumentGuid }
}

function _PWD-GetSheetIndexPwStateName {
    param([hashtable]$Config, [string]$DocumentGuid)
    return $script:sheetIndexStates[$DocumentGuid]
}

$cfg = @{
    database = @{ enabled = $true; allowWritesInDryRun = $true }
    auditPoller = @{
        workflowTriggers = @{
            enabled = $true
            recordStateHistory = $true
            recordTransitions = $true
            recordFromProcessor = $true
            notifyOnStateChange = $true
            qcPdfNotificationsOnly = $true
        }
    }
}

$sheetStem = '080J082001ab001'
$folder = 'Documents\X\CADD\Sheets'
$dgnGuid = 'guid-dgn'
$pdfGuid = 'guid-pdf'
$qcGuid = 'guid-qc'
$members = @(
    @{ documentGuid = $dgnGuid; documentName = ($sheetStem + '.dgn'); document = $null }
    @{ documentGuid = $pdfGuid; documentName = ($sheetStem + '.pdf'); document = $null }
    @{ documentGuid = $qcGuid; documentName = ($sheetStem + '-qc.pdf'); document = $null }
)

function Reset-TestState {
    $script:transitionCalls.Clear()
    $script:historyCalls.Clear()
    $script:workflowEvents.Clear()
    $script:notificationCalls = 0
    $script:transitionReuseCount = 0
    $script:sheetIndexStates = @{
        $dgnGuid = 'In Production'
        $pdfGuid = 'In Production'
        $qcGuid = 'In Production'
    }
}

# 1. User changes DGN -> QC Initiated; all siblings get events after prepend target Ready for QC
Reset-TestState
$prevMap = @{
    ($dgnGuid.ToLowerInvariant()) = 'In Production'
    ($pdfGuid.ToLowerInvariant()) = 'In Production'
    ($qcGuid.ToLowerInvariant()) = 'In Production'
}
$r1 = Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $dgnGuid `
    -TriggerDocumentName ($sheetStem + '.dgn') -FolderPath $folder -SourceState 'In Production' `
    -TargetState 'QC Initiated' -TransitionSource 'user_audit' -Members $members `
    -PreviousStateByGuid $prevMap -AuditEventId 5001 -ChangedByUser 7 -ChangedByUsername 'human.user'
Assert-Eq $script:historyCalls.Count 3 'DGN trigger: history for all siblings'
Assert-Eq $script:transitionCalls.Count 3 'DGN trigger: transitions for all siblings'
Assert-Eq $script:workflowEvents.Count 3 'DGN trigger: workflow mirror per sibling'
Assert-True (@($script:workflowEvents | Where-Object { $null -ne $_.transitionEventId -and $_.transitionEventId -gt 0 }).Count -eq 3) `
    'Workflow mirror rows should link to transition_events.id'
Assert-Eq $script:notificationCalls 1 'DGN trigger: one notification per sheet group'

# 2. User changes Sheet PDF
Reset-TestState
$r2 = Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $pdfGuid `
    -TriggerDocumentName ($sheetStem + '.pdf') -FolderPath $folder -SourceState 'In Production' `
    -TargetState 'QC Initiated' -TransitionSource 'user_audit' -Members $members `
    -PreviousStateByGuid $prevMap -AuditEventId 5002
Assert-Eq $script:transitionCalls.Count 3 'PDF trigger: all siblings receive transitions'

# 3. User changes QC PDF
Reset-TestState
$r3 = Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $qcGuid `
    -TriggerDocumentName ($sheetStem + '-qc.pdf') -FolderPath $folder -SourceState 'In Production' `
    -TargetState 'QC Initiated' -TransitionSource 'user_audit' -Members $members `
    -PreviousStateByGuid $prevMap -AuditEventId 5003
Assert-Eq $script:transitionCalls.Count 3 'QC PDF trigger: all siblings receive transitions'
Assert-Eq $script:notificationCalls 1 'QC PDF trigger: one notification'

# 4. Notification once per transition (already asserted above; explicit recount)
Reset-TestState
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $pdfGuid `
    -TriggerDocumentName ($sheetStem + '.pdf') -FolderPath $folder -TargetState 'QC Initiated' `
    -TransitionSource 'user_audit' -Members $members -PreviousStateByGuid $prevMap -AuditEventId 5004 | Out-Null
Assert-Eq $script:notificationCalls 1 'notification is not sent per sibling'

# 5. Duplicate audit rows do not create duplicate lifecycle events
Reset-TestState
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $pdfGuid `
    -TriggerDocumentName ($sheetStem + '.pdf') -FolderPath $folder -TargetState 'QC Initiated' `
    -TransitionSource 'user_audit' -Members $members -PreviousStateByGuid $prevMap -AuditEventId 5005 | Out-Null
$firstCount = $script:transitionCalls.Count
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $pdfGuid `
    -TriggerDocumentName ($sheetStem + '.pdf') -FolderPath $folder -TargetState 'QC Initiated' `
    -TransitionSource 'user_audit' -Members $members -PreviousStateByGuid $prevMap -AuditEventId 5005 | Out-Null
Assert-Eq $script:transitionCalls.Count $firstCount 'duplicate audit id does not insert duplicate transitions'
Assert-True ($script:transitionReuseCount -ge 3) 'duplicate audit id reuses transition rows'

# 6. Missing sibling skipped with telemetry
Reset-TestState
$partialMembers = @(
    @{ documentGuid = $dgnGuid; documentName = ($sheetStem + '.dgn'); document = $null }
    @{ documentGuid = $pdfGuid; documentName = ($sheetStem + '.pdf'); document = $null }
)
$r6 = Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $dgnGuid `
    -TriggerDocumentName ($sheetStem + '.dgn') -FolderPath $folder -TargetState 'QC Initiated' `
    -TransitionSource 'user_audit' -Members $partialMembers -PreviousStateByGuid $prevMap -AuditEventId 5006
$missing = @($r6.members | Where-Object { $_.skipReason -eq 'sibling_missing' -and $_.role -eq 'qcPdf' })
Assert-Eq $missing.Count 1 'missing qc pdf logged with sibling_missing'
Assert-Eq $script:transitionCalls.Count 2 'only resolved siblings get transitions'

# 7. Prepend completion records all siblings (Ready for QC)
Reset-TestState
$prevReady = @{
    ($dgnGuid.ToLowerInvariant()) = 'QC Initiated'
    ($pdfGuid.ToLowerInvariant()) = 'QC Initiated'
    ($qcGuid.ToLowerInvariant()) = 'QC Initiated'
}
$ctx = @{
    job = @{ id = 'job-prepend-1' }
    sheetStateSync = @{
        updates = @(
            @{ documentGuid = $dgnGuid; fromState = 'QC Initiated'; toState = 'Ready for QC' }
            @{ documentGuid = $pdfGuid; fromState = 'QC Initiated'; toState = 'Ready for QC' }
            @{ documentGuid = $qcGuid; fromState = 'QC Initiated'; toState = 'Ready for QC' }
        )
    }
}
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $qcGuid `
    -TriggerDocumentName ($sheetStem + '-qc.pdf') -FolderPath $folder -SourceState 'QC Initiated' `
    -TargetState 'Ready for QC' -TransitionSource 'automation_prepend_completion' `
    -Members $members -PreviousStateByGuid $prevReady -JobId 'job-prepend-1' -JobType 'QC_PREPEND' `
    -Context $ctx -SuppressNotification | Out-Null
Assert-Eq $script:transitionCalls.Count 3 'prepend completion: all siblings get Ready for QC events'
Assert-Eq $script:notificationCalls 0 'prepend path suppresses sheet-group notification'

# 7b. PW already at target but sheet_index still shows prior state (trigger-doc bug)
Reset-TestState
$liveAligned = @{
    ($dgnGuid.ToLowerInvariant()) = 'QC Initiated'
    ($pdfGuid.ToLowerInvariant()) = 'QC Initiated'
    ($qcGuid.ToLowerInvariant()) = 'QC Initiated'
}
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $dgnGuid `
    -TriggerDocumentName ($sheetStem + '.dgn') -FolderPath $folder -SourceState 'In Production' `
    -TargetState 'QC Initiated' -TransitionSource 'user_audit' -Members $members `
    -PreviousStateByGuid $prevMap -StateByGuid $liveAligned -AuditEventId 5007 | Out-Null
Assert-Eq $script:transitionCalls.Count 3 'index-lag still records events for all siblings when PW already aligned'

# 8. Dashboard parity: DGN rows get same transition target as PDF/QC PDF
Reset-TestState
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $dgnGuid `
    -TriggerDocumentName ($sheetStem + '.dgn') -FolderPath $folder -TargetState 'QC Initiated' `
    -TransitionSource 'user_audit' -Members $members -PreviousStateByGuid $prevMap -AuditEventId 5008 | Out-Null
$dgnT = @($script:transitionCalls | Where-Object { $_.documentGuid -eq $dgnGuid })
$pdfT = @($script:transitionCalls | Where-Object { $_.documentGuid -eq $pdfGuid })
$qcT = @($script:transitionCalls | Where-Object { $_.documentGuid -eq $qcGuid })
Assert-Eq $dgnT[0].toValue 'QC Initiated' 'DGN dashboard row target state'
Assert-Eq $pdfT[0].toValue 'QC Initiated' 'PDF dashboard row target state'
Assert-Eq $qcT[0].toValue 'QC Initiated' 'QC PDF dashboard row target state'

Assert-Eq (Get-QCSheetGroupTransitionKey -SheetStem $sheetStem -DocumentGuid $pdfGuid -TargetState 'QC Initiated' `
    -TransitionSource 'user_audit' -AuditEventId 42) `
    ('sg|' + $sheetStem.ToLowerInvariant() + '|' + $pdfGuid.ToLowerInvariant() + '|qc initiated|audit:42|user_audit') `
    'stable sheet-group transition key'

Write-Host 'test_sheet_group_workflow_transition.ps1 passed'
