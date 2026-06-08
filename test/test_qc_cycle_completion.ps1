# QC cycle completion tracking (no PW / SQL runtime).
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

# Review type normalization
Assert-Eq (Get-QCReviewTypeBucket -ReviewType 'Production QC') 'production' 'Production QC -> production'
Assert-Eq (Get-QCReviewTypeBucket -ReviewType 'Production') 'production' 'Production -> production'
Assert-Eq (Get-QCReviewTypeBucket -ReviewType 'QC') 'production' 'QC -> production'
Assert-Eq (Get-QCReviewTypeBucket -ReviewType 'Peer Review') 'peer_review' 'Peer Review -> peer_review'
Assert-Eq (Get-QCReviewTypeBucket -ReviewType 'Peer') 'peer_review' 'Peer -> peer_review'
Assert-Eq (Get-QCReviewTypeBucket -ReviewType 'Independent Check') 'independent_check' 'Independent Check -> independent_check'
Assert-Eq (Get-QCReviewTypeBucket -ReviewType 'Independent Review') 'independent_check' 'Independent Review -> independent_check'
Assert-Eq (Get-QCReviewTypeBucket -ReviewType 'IC') 'independent_check' 'IC -> independent_check'
Assert-Eq (Get-QCReviewTypeBucket -ReviewType 'Unknown Type') $null 'Unknown review type returns null'

$script:completionCalls = [System.Collections.Generic.List[object]]::new()
$script:summaryCalls = [System.Collections.Generic.List[string]]::new()
$script:transitionCalls = [System.Collections.Generic.List[object]]::new()
$script:historyCalls = [System.Collections.Generic.List[object]]::new()
$script:workflowEvents = [System.Collections.Generic.List[object]]::new()
$script:cycleByGuid = @{}
$script:jsonLogs = [System.Collections.Generic.List[object]]::new()

function Write-QCDocumentStateHistoryRow {
    param(
        [hashtable]$Config, [string]$DocumentGuid, [string]$EventType,
        [string]$DocumentName = '', [string]$FolderPath = '',
        [string]$OldValue = '', [string]$NewValue = '', [string]$FieldName = '',
        [Nullable[int]]$ChangedByUser = $null, [string]$ChangedByUsername = '',
        [Nullable[long]]$SourceAuditId = $null
    )
    $script:historyCalls.Add(@{ documentGuid = $DocumentGuid; oldValue = $OldValue; newValue = $NewValue }) | Out-Null
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
    $script:transitionCalls.Add(@{
        documentGuid = $DocumentGuid; fromValue = $FromValue; toValue = $ToValue; triggerAuditId = $TriggerAuditId
    }) | Out-Null
    return [pscustomobject]@{
        IsSuccess = $true
        Data = @{ written = $true; reused = $false; transitionId = 2000 + $script:transitionCalls.Count }
    }
}

function Write-QCWorkflowEventRow {
    param(
        [hashtable]$Config, [string]$DocumentId = '', [string]$JobId = '',
        [string]$EventType = '', [string]$PreviousPwState = '', [string]$TargetPwState = '',
        [string]$DecisionCode = '', [string]$PayloadJson = '', [string]$QcReviewType = ''
    )
    $script:workflowEvents.Add(@{ documentId = $DocumentId; eventType = $EventType }) | Out-Null
    return [pscustomobject]@{ IsSuccess = $true; Data = @{ written = $true } }
}

function Ensure-QCCycleCompletion {
    param(
        [hashtable]$Config, [string]$DocumentGuid, [string]$QcCycleId, [string]$QcReviewType,
        [string]$DocumentName = '', [Nullable[int]]$QcCycleNumber = $null,
        [Nullable[long]]$TransitionEventId = $null, [Nullable[long]]$AuditEventId = $null,
        [string]$CompletedBy = '', [Nullable[datetime]]$CompletedAt = $null
    )
    $key = ($DocumentGuid + '|' + $QcCycleId + '|' + $QcReviewType)
    $existing = @($script:completionCalls | Where-Object {
        ($_.documentGuid + '|' + $_.qcCycleId + '|' + $_.qcReviewType) -eq $key
    })
    if ($existing.Count -gt 0) {
        return [pscustomobject]@{
            IsSuccess = $true
            Data = @{ inserted = $false; reused = $true; completionId = 1; completedAt = [datetime]::UtcNow }
        }
    }
    $script:completionCalls.Add(@{
        documentGuid = $DocumentGuid
        documentName = $DocumentName
        qcCycleId = $QcCycleId
        qcCycleNumber = $QcCycleNumber
        qcReviewType = $QcReviewType
        transitionEventId = $TransitionEventId
        auditEventId = $AuditEventId
        completedBy = $CompletedBy
    }) | Out-Null
    return [pscustomobject]@{
        IsSuccess = $true
        Data = @{ inserted = $true; reused = $false; completionId = $script:completionCalls.Count; completedAt = [datetime]::UtcNow }
    }
}

function Update-QCSheetCycleCompletionSummary {
    param([hashtable]$Config, [string]$DocumentGuid)
    $script:summaryCalls.Add([string]$DocumentGuid) | Out-Null
    return [pscustomobject]@{ IsSuccess = $true; Data = @{ written = $true } }
}

function Get-QCSheetIndexCycle {
    param([hashtable]$Config, [string]$DocumentGuid = '', [string]$FolderPath = '', [string]$SheetStem = '')
    if ($script:cycleByGuid.ContainsKey($DocumentGuid)) { return $script:cycleByGuid[$DocumentGuid] }
    return $null
}

function Invoke-QCWorkflowStateChangeNotification {
    param([hashtable]$Config, [hashtable]$Context, [string]$PreviousState, [string]$CurrentState, $Document)
    return [pscustomobject]@{ IsSuccess = $true; Code = 'QC_NOTIFICATION_SENT' }
}

function _QCAT-BuildNotificationDocument {
    param([hashtable]$Config, [string]$FolderPath, [string]$DocumentName, [string]$DocumentGuid, [hashtable]$Attributes)
    return [pscustomobject]@{ Name = $DocumentName; DocumentGUID = $DocumentGuid }
}

function Write-QCJsonLog {
    param([string]$Level, [string]$Code, [string]$Message, [hashtable]$Data)
    $script:jsonLogs.Add(@{ code = $Code; data = $Data }) | Out-Null
}

$cfg = @{
    database = @{ enabled = $true; allowWritesInDryRun = $true }
    auditPoller = @{
        workflowTriggers = @{
            enabled = $true
            recordStateHistory = $true
            recordTransitions = $true
            recordFromProcessor = $true
            notifyOnStateChange = $false
            qcPdfNotificationsOnly = $true
        }
    }
    qcWorkflow = @{
        states = @{ complete = 'QC Complete' }
        reviewTypes = @{
            productionQc = 'Production QC'
            peerReview = 'Peer Review'
            independentCheck = 'Independent Check'
        }
    }
}

$sheetStem = '081800063ea515'
$folder = 'documents\proj\cadd\sheets'
$dgnGuid = '11111111-1111-1111-1111-111111111111'
$sheetGuid = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
$qcGuid = '33333333-3333-3333-3333-333333333333'
$script:cycleByGuid[$sheetGuid] = @{ cycleId = 'cycle-2026-001'; cycleNumber = '3' }

$members = @(
    @{ documentGuid = $dgnGuid; documentName = ($sheetStem + '.dgn'); document = $null }
    @{ documentGuid = $sheetGuid; documentName = ($sheetStem + '.pdf'); document = [pscustomobject]@{ QC_Review_Type = 'Peer Review' } }
    @{ documentGuid = $qcGuid; documentName = ($sheetStem + '-qc.pdf'); document = $null }
)

function Reset-CompletionState {
    $script:completionCalls.Clear()
    $script:summaryCalls.Clear()
    $script:transitionCalls.Clear()
    $script:historyCalls.Clear()
    $script:workflowEvents.Clear()
    $script:jsonLogs.Clear()
}

$prevMap = @{
    ($dgnGuid.ToLowerInvariant()) = 'QC Finalizing'
    ($sheetGuid.ToLowerInvariant()) = 'QC Finalizing'
    ($qcGuid.ToLowerInvariant()) = 'QC Finalizing'
}

# Package identity: DGN trigger records one completion on sheet PDF guid
Reset-CompletionState
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $dgnGuid `
    -TriggerDocumentName ($sheetStem + '.dgn') -FolderPath $folder `
    -SourceState 'QC Finalizing' -TargetState 'QC Complete' -TransitionSource 'user_audit' `
    -Members $members -PreviousStateByGuid $prevMap -AuditEventId 9001 -ChangedByUsername 'reviewer@example.com'
Assert-Eq $script:completionCalls.Count 1 'DGN trigger should record one package completion'
Assert-Eq $script:completionCalls[0].documentGuid $sheetGuid 'DGN trigger should canonicalize to sheet PDF guid'
Assert-Eq $script:completionCalls[0].qcReviewType 'peer_review' 'Stored review type should be normalized'
Assert-Eq $script:historyCalls.Count 3 'Sibling history must remain per-member'
Assert-Eq $script:transitionCalls.Count 3 'Sibling transitions must remain per-member'

# Package identity: QC PDF trigger also canonicalizes to sheet PDF
Reset-CompletionState
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $qcGuid `
    -TriggerDocumentName ($sheetStem + '-qc.pdf') -FolderPath $folder `
    -SourceState 'QC Finalizing' -TargetState 'QC Complete' -TransitionSource 'user_audit' `
    -Members $members -PreviousStateByGuid $prevMap -AuditEventId 9002
Assert-Eq $script:completionCalls.Count 1 'QC PDF trigger should record one package completion'
Assert-Eq $script:completionCalls[0].documentGuid $sheetGuid 'QC PDF trigger should canonicalize to sheet PDF guid'

# Sibling retry / replay: repeated transitions still produce one row
Reset-CompletionState
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $sheetGuid `
    -TriggerDocumentName ($sheetStem + '.pdf') -FolderPath $folder `
    -SourceState 'QC Finalizing' -TargetState 'QC Complete' -TransitionSource 'user_audit' `
    -Members $members -PreviousStateByGuid $prevMap -AuditEventId 9010 | Out-Null
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $dgnGuid `
    -TriggerDocumentName ($sheetStem + '.dgn') -FolderPath $folder `
    -SourceState 'QC Finalizing' -TargetState 'QC Complete' -TransitionSource 'user_audit' `
    -Members $members -PreviousStateByGuid $prevMap -AuditEventId 9011 | Out-Null
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $qcGuid `
    -TriggerDocumentName ($sheetStem + '-qc.pdf') -FolderPath $folder `
    -SourceState 'QC Finalizing' -TargetState 'QC Complete' -TransitionSource 'user_audit' `
    -Members $members -PreviousStateByGuid $prevMap -AuditEventId 9012 | Out-Null
Assert-Eq $script:completionCalls.Count 1 'Sibling retries should not multiply completion rows'
Assert-True ($script:jsonLogs | Where-Object { $_.code -eq 'QC_CYCLE_COMPLETION_DUPLICATE' }) 'Sibling retries should log duplicates'

# Prepend completion excluded
$completionBeforePrepend = $script:completionCalls.Count
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $sheetGuid `
    -TriggerDocumentName ($sheetStem + '.pdf') -FolderPath $folder `
    -SourceState 'QC Finalizing' -TargetState 'QC Complete' -TransitionSource 'automation_prepend_completion' `
    -Members $members -PreviousStateByGuid $prevMap -JobId 'job-prepend-1'
Assert-Eq $script:completionCalls.Count $completionBeforePrepend 'Prepend completion must not record QC cycle completion'

# Unknown review type skipped
Reset-CompletionState
$unknownMembers = @(
    @{ documentGuid = $sheetGuid; documentName = ($sheetStem + '.pdf'); document = [pscustomobject]@{ QC_Review_Type = 'Mystery Review' } }
)
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $sheetGuid `
    -TriggerDocumentName ($sheetStem + '.pdf') -FolderPath $folder `
    -SourceState 'QC Finalizing' -TargetState 'QC Complete' -TransitionSource 'user_audit' `
    -Members $unknownMembers -AuditEventId 9020
Assert-Eq $script:completionCalls.Count 0 'Unknown review type must not insert completion'
Assert-True ($script:jsonLogs | Where-Object { $_.code -eq 'QC_CYCLE_COMPLETION_SKIPPED' }) 'Unknown review type should log skip'

# Audit trigger path canonicalizes via StaleCheckMembers and records normalized type
Reset-CompletionState
Invoke-QCAuditWorkflowStateChangeTriggers -Config $cfg -DocumentGuid $dgnGuid `
    -DocumentName ($sheetStem + '.dgn') -FolderPath $folder `
    -PreviousState 'QC Finalizing' -CurrentState 'QC Complete' `
    -PwAttributes @{ qc_review_type = 'Independent Check' } -AuditEventId 9030 `
    -ChangedByUsername 'checker@example.com' -StaleCheckMembers $members
Assert-Eq $script:completionCalls.Count 1 'Audit trigger path should record one canonical completion'
Assert-Eq $script:completionCalls[0].documentGuid $sheetGuid 'Audit trigger should canonicalize DGN to sheet PDF'
Assert-Eq $script:completionCalls[0].qcReviewType 'independent_check' 'Audit trigger should store normalized review type'

# Dual-path duplicate: sheet group then audit trigger for same cycle/review type
Reset-CompletionState
$script:cycleByGuid[$sheetGuid] = @{ cycleId = 'cycle-2026-001'; cycleNumber = '3' }
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $sheetGuid `
    -TriggerDocumentName ($sheetStem + '.pdf') -FolderPath $folder `
    -SourceState 'QC Finalizing' -TargetState 'QC Complete' -TransitionSource 'user_audit' `
    -Members $members -PreviousStateByGuid $prevMap -AuditEventId 9040 | Out-Null
$beforeDual = $script:completionCalls.Count
Invoke-QCAuditWorkflowStateChangeTriggers -Config $cfg -DocumentGuid $sheetGuid `
    -DocumentName ($sheetStem + '.pdf') -FolderPath $folder `
    -PreviousState 'QC Finalizing' -CurrentState 'QC Complete' `
    -PwAttributes @{ qc_review_type = 'Peer Review' } -AuditEventId 9041 -StaleCheckMembers $members
Assert-Eq $script:completionCalls.Count $beforeDual 'Dual entry points must not double-count same package cycle'

Write-Host 'ALL PASSED' -ForegroundColor Green
