# Sheet-group workflow transition telemetry (no PW / SQL runtime).
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Database\Core.Database.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Workflow\QC.AuditTriggers.psm1') -Force

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
$script:mockSheetPackageId = [guid]::Parse('11111111-1111-1111-1111-111111111111')

function Resolve-SheetPackageIdForSheetGroup {
    param([hashtable]$Config, [string]$FolderPath = '', [string]$SheetStem = '', [string]$DocumentGuid = '', [string]$DocumentName = '')
    return $script:mockSheetPackageId
}

function Write-QCDocumentStateHistoryRow {
    param(
        [hashtable]$Config, [string]$DocumentGuid, [string]$EventType,
        [string]$DocumentName = '', [string]$FolderPath = '',
        [string]$OldValue = '', [string]$NewValue = '', [string]$FieldName = '',
        [Nullable[int]]$ChangedByUser = $null, [string]$ChangedByUsername = '',
        [Nullable[long]]$SourceAuditId = $null,
        [Nullable[guid]]$SheetPackageId = $null,
        [Nullable[guid]]$TransitionGroupId = $null
    )
    $script:historyCalls.Add(@{
        documentGuid = $DocumentGuid
        documentName = $DocumentName
        oldValue = $OldValue
        newValue = $NewValue
        sourceAuditId = $SourceAuditId
        sheetPackageId = $SheetPackageId
        transitionGroupId = $TransitionGroupId
    }) | Out-Null
    return [pscustomobject]@{ IsSuccess = $true; Data = @{ written = $true } }
}

function Ensure-QCTransitionEvent {
    param(
        [hashtable]$Config, [string]$DocumentGuid, [string]$TransitionType,
        [string]$DocumentName = '', [string]$FolderPath = '',
        [string]$FromValue = '', [string]$ToValue = '',
        [string]$JobId = '', [string]$JobType = '',
        [Nullable[long]]$TriggerAuditId = $null,
        [Nullable[guid]]$SheetPackageId = $null,
        [Nullable[guid]]$TransitionGroupId = $null
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
        sheetPackageId = $SheetPackageId
        transitionGroupId = $TransitionGroupId
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
        [Nullable[int]]$TransitionEventId = $null,
        [Nullable[guid]]$SheetPackageId = $null,
        [Nullable[guid]]$TransitionGroupId = $null
    )
    $script:workflowEvents.Add(@{
        documentId = $DocumentId
        eventType = $EventType
        previousPwState = $PreviousPwState
        targetPwState = $TargetPwState
        decisionCode = $DecisionCode
        transitionEventId = $TransitionEventId
        sheetPackageId = $SheetPackageId
        transitionGroupId = $TransitionGroupId
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
            qcPdfEventsOnly = $true
        }
    }
}

$sheetStem = '080J082001ab001'
$folder = 'Documents\X\CADD\Sheets'
$dgnGuid = 'guid-dgn'
$pdfGuid = 'guid-pdf'
$qcGuid = 'guid-prod'
$members = @(
    @{ documentGuid = $dgnGuid; documentName = ($sheetStem + '.dgn'); document = $null }
    @{ documentGuid = $pdfGuid; documentName = ($sheetStem + '.pdf'); document = $null }
    @{ documentGuid = $qcGuid; documentName = ($sheetStem + '-prod.pdf'); document = $null }
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

# 1. User changes DGN -> QC Initiated; only lane prod PDF gets telemetry when qcPdfEventsOnly
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
Assert-Eq $script:historyCalls.Count 1 'DGN trigger: history for lane prod PDF only'
Assert-Eq $script:transitionCalls.Count 1 'DGN trigger: transition for lane prod PDF only'
Assert-Eq $script:workflowEvents.Count 1 'DGN trigger: workflow mirror for lane prod PDF only'
Assert-Eq $script:transitionCalls[0].documentGuid $qcGuid 'DGN trigger: prod lane receives transition'
Assert-True (@($script:workflowEvents | Where-Object { $null -ne $_.transitionEventId -and $_.transitionEventId -gt 0 }).Count -eq 1) `
    'Workflow mirror row should link to transition_events.id'
Assert-Eq $script:notificationCalls 1 'DGN trigger: one notification per sheet group'
Assert-True ($null -ne $r1.transitionGroupId) 'transition_group_id returned'
Assert-Eq $r1.sheetPackageId $script:mockSheetPackageId 'sheet_package_id returned'
$groupIds = @($script:transitionCalls | ForEach-Object { $_.transitionGroupId.ToString() } | Select-Object -Unique)
Assert-Eq $groupIds.Count 1 'lane transition has transition_group_id'
$pkgIds = @($script:transitionCalls | ForEach-Object { $_.sheetPackageId.ToString() } | Select-Object -Unique)
Assert-Eq $pkgIds.Count 1 'lane transition has sheet_package_id'
Assert-Eq $pkgIds[0] $script:mockSheetPackageId.ToString() 'transition sheet_package_id matches package'

# 2. User changes Sheet PDF — same lane-only telemetry
Reset-TestState
$r2 = Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $pdfGuid `
    -TriggerDocumentName ($sheetStem + '.pdf') -FolderPath $folder -SourceState 'In Production' `
    -TargetState 'QC Initiated' -TransitionSource 'user_audit' -Members $members `
    -PreviousStateByGuid $prevMap -AuditEventId 5002
Assert-Eq $script:transitionCalls.Count 1 'PDF trigger: lane prod PDF receives transition'

# 3. User changes lane prod PDF
Reset-TestState
$r3 = Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $qcGuid `
    -TriggerDocumentName ($sheetStem + '-prod.pdf') -FolderPath $folder -SourceState 'In Production' `
    -TargetState 'QC Initiated' -TransitionSource 'user_audit' -Members $members `
    -PreviousStateByGuid $prevMap -AuditEventId 5003
Assert-Eq $script:transitionCalls.Count 1 'prod PDF trigger: lane document receives transition'
Assert-Eq $script:notificationCalls 1 'prod PDF trigger: one notification'

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
Assert-True ($script:transitionReuseCount -ge 1) 'duplicate audit id reuses transition row'

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
Assert-Eq $script:transitionCalls.Count 0 'no lane PDF resolved: no transitions when qcPdfEventsOnly'

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
    -TriggerDocumentName ($sheetStem + '-prod.pdf') -FolderPath $folder -SourceState 'QC Initiated' `
    -TargetState 'Ready for QC' -TransitionSource 'automation_prepend_completion' `
    -Members $members -PreviousStateByGuid $prevReady -JobId 'job-prepend-1' -JobType 'QC_PREPEND' `
    -Context $ctx -SuppressNotification | Out-Null
Assert-Eq $script:transitionCalls.Count 1 'prepend completion: lane prod PDF gets Ready for QC event'
Assert-Eq $script:notificationCalls 0 'prepend path suppresses sheet-group notification'

# 7c. Lane-independent initial prepend: sheet_index updated before telemetry must still record lane transition
Reset-TestState
$revGuid = [guid]::NewGuid().ToString()
$revMembers = @(
    @{ documentGuid = $dgnGuid; documentName = ($sheetStem + '.dgn'); document = $null }
    @{ documentGuid = $pdfGuid; documentName = ($sheetStem + '.pdf'); document = $null }
    @{ documentGuid = $revGuid; documentName = ($sheetStem + '-rev.pdf'); document = $null }
)
$laneSplitCtx = @{
    job = @{ id = 'job-prepend-rev-cycle2' }
    laneIndependentInitialPrepend = $true
    laneTargetState = 'Originated'
    referenceState = 'In Development'
    activeQcProcessType = 'review'
    qcProcessType = 'review'
    lanePostPrependSplit = @{
        lanePdfName = ($sheetStem + '-rev.pdf')
        laneStateVerified = $true
        updates = @(
            @{
                documentGuid = $revGuid
                documentName = ($sheetStem + '-rev.pdf')
                fromState = 'Redlines Received'
                toState = 'Originated'
                verified = $true
                isLaneAuthority = $true
            }
            @{
                documentGuid = $pdfGuid
                documentName = ($sheetStem + '.pdf')
                fromState = 'Initiate Origination'
                toState = 'In Development'
                verified = $true
                isStemReference = $true
            }
        )
    }
}
# Simulate sheet_index already updated to Originated before processor telemetry runs.
$indexLagPrev = @{
    ($revGuid.ToLowerInvariant()) = 'Originated'
    ($pdfGuid.ToLowerInvariant()) = 'In Development'
    ($dgnGuid.ToLowerInvariant()) = 'In Development'
}
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $pdfGuid `
    -TriggerDocumentName ($sheetStem + '.pdf') -FolderPath $folder -SourceState 'Initiate Origination' `
    -TargetState 'Originated' -TransitionSource 'automation_prepend_completion' `
    -Members $revMembers -PreviousStateByGuid $indexLagPrev -JobId 'job-prepend-rev-cycle2' -JobType 'QC_PREPEND' `
    -Context $laneSplitCtx -SuppressNotification | Out-Null
$revTransition = @($script:transitionCalls | Where-Object { $_.documentGuid -eq $revGuid })
Assert-Eq $revTransition.Count 1 'lane rev PDF gets Originated workflow event after cycle overwrite prepend'
Assert-Eq $revTransition[0].fromValue 'Redlines Received' 'lane event uses prepend from-state not post-write sheet_index'
Assert-Eq $revTransition[0].toValue 'Originated' 'lane event target is Originated'

# 7d. Review prepend with existing prod lane — prod must not get spurious transition telemetry
Reset-TestState
$revMembersWithProd = @(
    @{ documentGuid = $dgnGuid; documentName = ($sheetStem + '.dgn'); document = $null }
    @{ documentGuid = $pdfGuid; documentName = ($sheetStem + '.pdf'); document = $null }
    @{ documentGuid = $qcGuid; documentName = ($sheetStem + '-prod.pdf'); document = $null }
    @{ documentGuid = $revGuid; documentName = ($sheetStem + '-rev.pdf'); document = $null }
)
$indexLagPrevWithProd = @{
    ($revGuid.ToLowerInvariant()) = 'Redlines Received'
    ($qcGuid.ToLowerInvariant()) = 'Initiate Origination'
    ($pdfGuid.ToLowerInvariant()) = 'In Development'
    ($dgnGuid.ToLowerInvariant()) = 'In Development'
}
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $pdfGuid `
    -TriggerDocumentName ($sheetStem + '.pdf') -FolderPath $folder -SourceState 'Initiate Origination' `
    -TargetState 'Originated' -TransitionSource 'automation_prepend_completion' `
    -Members $revMembersWithProd -PreviousStateByGuid $indexLagPrevWithProd -JobId 'job-prepend-rev-with-prod' -JobType 'QC_PREPEND' `
    -Context $laneSplitCtx -SuppressNotification | Out-Null
$prodTransition = @($script:transitionCalls | Where-Object { $_.documentGuid -eq $qcGuid })
Assert-Eq $prodTransition.Count 0 'inactive prod lane skipped during review prepend telemetry'
$revWithProdTransition = @($script:transitionCalls | Where-Object { $_.documentGuid -eq $revGuid })
Assert-Eq $revWithProdTransition.Count 1 'active rev lane still records Originated transition'

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
Assert-Eq $script:transitionCalls.Count 1 'index-lag still records lane prod event when PW already aligned'

# 8. Lane prod row gets transition target when stem/DGN triggers
Reset-TestState
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $dgnGuid `
    -TriggerDocumentName ($sheetStem + '.dgn') -FolderPath $folder -TargetState 'QC Initiated' `
    -TransitionSource 'user_audit' -Members $members -PreviousStateByGuid $prevMap -AuditEventId 5008 | Out-Null
$qcT = @($script:transitionCalls | Where-Object { $_.documentGuid -eq $qcGuid })
Assert-Eq $qcT[0].toValue 'QC Initiated' 'lane prod dashboard row target state'

Assert-Eq (Get-QCSheetGroupTransitionKey -SheetStem $sheetStem -DocumentGuid $pdfGuid -TargetState 'QC Initiated' `
    -TransitionSource 'user_audit' -AuditEventId 42) `
    ('sg|' + $sheetStem.ToLowerInvariant() + '|' + $pdfGuid.ToLowerInvariant() + '|qc initiated|audit:42|user_audit') `
    'stable sheet-group transition key'

# 9. Member telemetry: finalState is logical target; preSyncLiveState captures PW before sync
Reset-TestState
$redlinesPrev = @{
    ($dgnGuid.ToLowerInvariant()) = 'Redlines Received'
    ($pdfGuid.ToLowerInvariant()) = 'Redlines Received'
    ($qcGuid.ToLowerInvariant()) = 'Redlines Received'
}
$preSync = @{
    ($dgnGuid.ToLowerInvariant()) = 'Redlines Received'
    ($pdfGuid.ToLowerInvariant()) = 'Redlines Received'
    ($qcGuid.ToLowerInvariant()) = 'Redlines Received'
}
$r9 = Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $qcGuid `
    -TriggerDocumentName ($sheetStem + '-prod.pdf') -FolderPath $folder -SourceState 'Redlines Received' `
    -TargetState 'Corrections Received' -TransitionSource 'user_audit' -Members $members `
    -PreviousStateByGuid $redlinesPrev -StateByGuid $preSync -AuditEventId 40598
$pdfMember = @($r9.members | Where-Object { $_.role -eq 'pdf' })[0]
Assert-Eq $pdfMember.preSyncLiveState 'Redlines Received' 'preSyncLiveState should reflect PW before sibling sync'
Assert-Eq $pdfMember.finalState 'Corrections Received' 'finalState should reflect logical target after sibling sync'
Assert-Eq $pdfMember.skipReason 'qc_pdf_events_only' 'stem PDF member skipped for workflow telemetry'
Assert-True ($null -eq $pdfMember.liveState) 'legacy liveState field should not be emitted'

# 10. Duplicate notification results should not count as sent in sheet-group telemetry
$script:notificationResultOverride = $null
function Invoke-QCWorkflowStateChangeNotification {
    param([hashtable]$Config, [hashtable]$Context, [string]$PreviousState, [string]$CurrentState, $Document)
    $script:notificationCalls++
    if ($script:notificationResultOverride) { return $script:notificationResultOverride }
    return [pscustomobject]@{ IsSuccess = $true; Code = 'QC_NOTIFICATION_SENT' }
}
Reset-TestState
$script:notificationResultOverride = [pscustomobject]@{
    IsSuccess = $true
    Code = 'QC_NOTIFICATION_SKIPPED_DUPLICATE'
    Data = @{ skipped = $true; dedupeKey = 'test-dedupe-key' }
}
$r10 = Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $qcGuid `
    -TriggerDocumentName ($sheetStem + '-prod.pdf') -FolderPath $folder -SourceState 'Redlines Received' `
    -TargetState 'Corrections Received' -TransitionSource 'user_audit' -Members $members `
    -PreviousStateByGuid $redlinesPrev -AuditEventId 40599
Assert-True $r10.notificationEvaluated 'duplicate skip should still mark notification evaluated'
Assert-True (-not $r10.notificationSent) 'duplicate skip should not mark notification sent'
Assert-Eq $r10.notificationResultCode 'QC_NOTIFICATION_SKIPPED_DUPLICATE' 'duplicate skip should preserve result code'
Assert-True (-not $r10.notificationEmitted) 'notificationEmitted should mirror notificationSent'

# 11. Lane prod PDF trigger should win over stale sibling when sending notification
$script:lastNotifyDocumentGuid = $null
$script:notificationResultOverride = $null
function Invoke-QCWorkflowStateChangeNotification {
    param([hashtable]$Config, [hashtable]$Context, [string]$PreviousState, [string]$CurrentState, $Document)
    $script:notificationCalls++
    if ($Document -and $Document.DocumentGUID) {
        $script:lastNotifyDocumentGuid = [string]$Document.DocumentGUID
    }
    if ($script:notificationResultOverride) { return $script:notificationResultOverride }
    return [pscustomobject]@{ IsSuccess = $true; Code = 'QC_NOTIFICATION_SENT' }
}
$oldQcGuid = '27d9a8ba-6aaa-4e51-a1d8-9759b20880bb'
$newQcGuid = '72bf6609-063c-40db-b067-94d1b064a5b5'
$mixedMembers = @(
    @{ documentGuid = $dgnGuid; documentName = ($sheetStem + '.dgn'); document = $null }
    @{ documentGuid = $pdfGuid; documentName = ($sheetStem + '.pdf'); document = $null }
    @{ documentGuid = $oldQcGuid; documentName = ($sheetStem + '-prod.pdf'); document = $null }
    @{ documentGuid = $newQcGuid; documentName = ($sheetStem + '-prod.pdf'); document = $null }
)
$stalePrevMap = @{
    ($dgnGuid.ToLowerInvariant()) = 'Ready for QC'
    ($pdfGuid.ToLowerInvariant()) = 'Ready for QC'
    ($oldQcGuid.ToLowerInvariant()) = 'Ready for QC'
    ($newQcGuid.ToLowerInvariant()) = 'Ready for QC'
}
Reset-TestState
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $newQcGuid `
    -TriggerDocumentName ($sheetStem + '-prod.pdf') -FolderPath $folder -SourceState 'Ready for QC' `
    -TargetState 'Redlines Received' -TransitionSource 'user_audit' -Members $mixedMembers `
    -PreviousStateByGuid $stalePrevMap -AuditEventId 41070 | Out-Null
Assert-Eq $script:lastNotifyDocumentGuid $newQcGuid 'Trigger lane prod member should be selected for notification'
$staleOnlyMembers = @(
    @{ documentGuid = $dgnGuid; documentName = ($sheetStem + '.dgn'); document = $null }
    @{ documentGuid = $pdfGuid; documentName = ($sheetStem + '.pdf'); document = $null }
    @{ documentGuid = $oldQcGuid; documentName = ($sheetStem + '-prod.pdf'); document = $null }
)
$staleOnlyPrevMap = @{
    ($dgnGuid.ToLowerInvariant()) = 'Ready for QC'
    ($pdfGuid.ToLowerInvariant()) = 'Ready for QC'
    ($oldQcGuid.ToLowerInvariant()) = 'Ready for QC'
}
Reset-TestState
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $newQcGuid `
    -TriggerDocumentName ($sheetStem + '-prod.pdf') -FolderPath $folder -SourceState 'Ready for QC' `
    -TargetState 'Redlines Received' -TransitionSource 'user_audit' -Members $staleOnlyMembers `
    -PreviousStateByGuid $staleOnlyPrevMap -AuditEventId 41071 | Out-Null
Assert-Eq $script:lastNotifyDocumentGuid $newQcGuid 'Missing trigger lane prod member should fall back to trigger identity'

# 12. qcPdfEventsOnly=false restores all-sibling telemetry (legacy dashboard parity)
$cfgAllSiblings = @{
    database = @{ enabled = $true; allowWritesInDryRun = $true }
    auditPoller = @{
        workflowTriggers = @{
            enabled = $true
            recordStateHistory = $true
            recordTransitions = $true
            qcPdfEventsOnly = $false
        }
    }
}
Reset-TestState
Invoke-QCSheetGroupWorkflowTransition -Config $cfgAllSiblings -TriggerDocumentGuid $dgnGuid `
    -TriggerDocumentName ($sheetStem + '.dgn') -FolderPath $folder -TargetState 'QC Initiated' `
    -TransitionSource 'user_audit' -Members $members -PreviousStateByGuid $prevMap -AuditEventId 5010 | Out-Null
Assert-Eq $script:transitionCalls.Count 3 'qcPdfEventsOnly=false: all siblings receive transitions'

Write-Host 'test_sheet_group_workflow_transition.ps1 passed'
