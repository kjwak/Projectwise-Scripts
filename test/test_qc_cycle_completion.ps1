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

Assert-Eq (Get-QCReviewTypeBucket -ReviewType 'Production QC') 'production' 'Production QC -> production'
Assert-Eq (Get-QCReviewTypeBucket -ReviewType 'Peer Review') 'peer_review' 'Peer Review -> peer_review'
Assert-Eq (Get-QCReviewTypeBucket -ReviewType 'Independent Check') 'independent_check' 'Independent Check -> independent_check'
Assert-Eq (Get-QCReviewTypeBucket -ReviewType 'Unknown Type') $null 'Unknown review type returns null'

$packageId = [guid]'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
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
        [string]$DecisionCode = '', [string]$PayloadJson = '', [string]$QcReviewType = '',
        [Nullable[int]]$TransitionEventId = $null
    )
    $script:workflowEvents.Add(@{ documentId = $DocumentId; eventType = $EventType }) | Out-Null
    return [pscustomobject]@{ IsSuccess = $true; Data = @{ written = $true } }
}

function Resolve-QCCycleCompletionSheetPackageId {
    param(
        [hashtable]$Config,
        [string]$DocumentGuid = '',
        [Nullable[guid]]$SheetPackageId = $null
    )
    if ($null -ne $SheetPackageId -and $SheetPackageId -ne [guid]::Empty) { return $SheetPackageId }
    if ([string]$DocumentGuid -eq $dgnGuid) { return $packageId }
    return $null
}

function Ensure-QCCycleCompletion {
    param(
        [hashtable]$Config, [string]$DocumentGuid, [string]$QcCycleId, [string]$QcReviewType,
        [Nullable[guid]]$SheetPackageId = $null,
        [string]$DocumentName = '', [Nullable[int]]$QcCycleNumber = $null,
        [Nullable[long]]$TransitionEventId = $null, [Nullable[long]]$AuditEventId = $null,
        [string]$CompletedBy = '', [Nullable[datetime]]$CompletedAt = $null
    )
    $pkg = Resolve-QCCycleCompletionSheetPackageId -Config $Config -DocumentGuid $DocumentGuid -SheetPackageId $SheetPackageId
    if (-not $pkg) {
        return [pscustomobject]@{
            IsSuccess = $true
            Data = @{ inserted = $false; reused = $false; completionId = $null; sheetPackageId = $null; reason = 'sheet_package_not_found' }
        }
    }
    $key = ($pkg.ToString() + '|' + $QcCycleId + '|' + $QcReviewType)
    $existing = @($script:completionCalls | Where-Object {
        ($_.sheetPackageId.ToString() + '|' + $_.qcCycleId + '|' + $_.qcReviewType) -eq $key
    })
    if ($existing.Count -gt 0) {
        return [pscustomobject]@{
            IsSuccess = $true
            Data = @{ inserted = $false; reused = $true; completionId = 1; completedAt = [datetime]::UtcNow; sheetPackageId = $pkg }
        }
    }
    $script:completionCalls.Add(@{
        documentGuid = $DocumentGuid
        sheetPackageId = $pkg
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
        Data = @{ inserted = $true; reused = $false; completionId = $script:completionCalls.Count; completedAt = [datetime]::UtcNow; sheetPackageId = $pkg }
    }
}

function Update-QCSheetCycleCompletionSummary {
    param(
        [hashtable]$Config,
        [string]$DocumentGuid = '',
        [Nullable[guid]]$SheetPackageId = $null
    )
    $pkg = Resolve-QCCycleCompletionSheetPackageId -Config $Config -DocumentGuid $DocumentGuid -SheetPackageId $SheetPackageId
    if ($pkg) { $script:summaryCalls.Add($pkg.ToString()) | Out-Null }
    return [pscustomobject]@{ IsSuccess = $true; Data = @{ written = [bool]$pkg; sheetPackageId = $pkg } }
}

function Get-QCSheetIndexCycle {
    param([hashtable]$Config, [string]$DocumentGuid = '', [string]$FolderPath = '', [string]$SheetStem = '')
    if ($script:cycleByGuid.ContainsKey($DocumentGuid)) { return $script:cycleByGuid[$DocumentGuid] }
    if (-not [string]::IsNullOrWhiteSpace($SheetStem) -and -not [string]::IsNullOrWhiteSpace($FolderPath)) {
        $sheetKey = ($FolderPath + '|' + $SheetStem).ToLowerInvariant()
        if ($script:cycleByGuid.ContainsKey($sheetKey)) { return $script:cycleByGuid[$sheetKey] }
    }
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
$script:cycleByGuid[($folder + '|' + $sheetStem).ToLowerInvariant()] = @{ cycleId = 'cycle-2026-001'; cycleNumber = '3' }

$members = @(
    @{ documentGuid = $dgnGuid; documentName = ($sheetStem + '.dgn'); document = $null }
    @{ documentGuid = $sheetGuid; documentName = ($sheetStem + '.pdf'); document = [pscustomobject]@{ QC_Review_Type = 'Peer Review' } }
    @{ documentGuid = $qcGuid; documentName = ($sheetStem + '-qc.pdf'); document = $null }
)

$productionMembers = @(
    @{ documentGuid = $dgnGuid; documentName = ($sheetStem + '.dgn'); document = $null }
    @{ documentGuid = $sheetGuid; documentName = ($sheetStem + '.pdf'); document = [pscustomobject]@{ QC_Review_Type = 'Production QC' } }
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

# 1. DGN-triggered completion records against DGN GUID
Reset-CompletionState
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $dgnGuid `
    -TriggerDocumentName ($sheetStem + '.dgn') -FolderPath $folder `
    -SourceState 'QC Finalizing' -TargetState 'QC Complete' -TransitionSource 'user_audit' `
    -Members $members -PreviousStateByGuid $prevMap -AuditEventId 9001 -ChangedByUsername 'reviewer@example.com'
Assert-Eq $script:completionCalls.Count 1 'DGN trigger should record one package completion'
Assert-Eq $script:completionCalls[0].documentGuid $dgnGuid 'DGN trigger should keep audit document_guid'
Assert-Eq $script:completionCalls[0].sheetPackageId.ToString() $packageId.ToString() 'DGN trigger should resolve sheet_package_id'
Assert-Eq $script:completionCalls[0].qcReviewType 'peer_review' 'Stored review type should be normalized'
Assert-Eq $script:summaryCalls.Count 1 'Rollup should run after insert'
Assert-Eq $script:summaryCalls[0] $packageId.ToString() 'Rollup should target sheet_packages'
Assert-Eq $script:historyCalls.Count 3 'Sibling history must remain per-member'
Assert-Eq $script:transitionCalls.Count 3 'Sibling transitions must remain per-member'

# 2. Sheet PDF-triggered completion resolves to DGN GUID
Reset-CompletionState
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $sheetGuid `
    -TriggerDocumentName ($sheetStem + '.pdf') -FolderPath $folder `
    -SourceState 'QC Finalizing' -TargetState 'QC Complete' -TransitionSource 'user_audit' `
    -Members $members -PreviousStateByGuid $prevMap -AuditEventId 9003
Assert-Eq $script:completionCalls.Count 1 'Sheet PDF trigger should record one package completion'
Assert-Eq $script:completionCalls[0].documentGuid $dgnGuid 'Sheet PDF trigger should canonicalize to DGN GUID'
Assert-Eq $script:summaryCalls[0] $packageId.ToString() 'Sheet PDF trigger rollup should target sheet_packages'

# 3. QC PDF-triggered completion resolves to DGN GUID
Reset-CompletionState
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $qcGuid `
    -TriggerDocumentName ($sheetStem + '-qc.pdf') -FolderPath $folder `
    -SourceState 'QC Finalizing' -TargetState 'QC Complete' -TransitionSource 'user_audit' `
    -Members $members -PreviousStateByGuid $prevMap -AuditEventId 9002
Assert-Eq $script:completionCalls.Count 1 'QC PDF trigger should record one package completion'
Assert-Eq $script:completionCalls[0].documentGuid $dgnGuid 'QC PDF trigger should canonicalize to DGN GUID'

# 4. Sibling retries produce one completion row
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
Assert-Eq $script:completionCalls[0].documentGuid $dgnGuid 'Sibling retries should keep DGN identity'
Assert-True ($script:jsonLogs | Where-Object { $_.code -eq 'QC_CYCLE_COMPLETION_DUPLICATE' }) 'Sibling retries should log duplicates'

# 5. Rollup updates sheet_packages only once per package
Reset-CompletionState
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $dgnGuid `
    -TriggerDocumentName ($sheetStem + '.dgn') -FolderPath $folder `
    -SourceState 'QC Finalizing' -TargetState 'QC Complete' -TransitionSource 'user_audit' `
    -Members $members -PreviousStateByGuid $prevMap -AuditEventId 9015 | Out-Null
Assert-Eq $script:summaryCalls.Count 1 'Only one rollup should run'
Assert-Eq $script:summaryCalls[0] $packageId.ToString() 'Rollup should update sheet_packages row only'
Assert-Eq @($script:summaryCalls | Where-Object { $_ -eq $sheetGuid }).Count 0 'Sheet PDF guid should not receive rollup'
Assert-Eq @($script:summaryCalls | Where-Object { $_ -eq $qcGuid }).Count 0 'QC PDF guid should not receive rollup'

# 6. Finalizing prepend counts once
Reset-CompletionState
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $qcGuid `
    -TriggerDocumentName ($sheetStem + '-qc.pdf') -FolderPath $folder `
    -SourceState 'QC Finalizing' -TargetState 'QC Complete' -TransitionSource 'automation_prepend_completion' `
    -Members $productionMembers -PreviousStateByGuid $prevMap -JobId 'job-finalize-1' -JobType 'QC_PREPEND' -SuppressNotification | Out-Null
Assert-Eq $script:completionCalls.Count 1 'Finalizing prepend should record one QC cycle completion'
Assert-Eq $script:completionCalls[0].documentGuid $dgnGuid 'Finalizing prepend should use DGN GUID'
Assert-Eq $script:completionCalls[0].qcReviewType 'production' 'Finalizing prepend should store normalized production review type'
Assert-True ($script:jsonLogs | Where-Object { $_.code -eq 'QC_CYCLE_COMPLETION_RECORDED' }) 'Finalizing prepend should log QC_CYCLE_COMPLETION_RECORDED'

# 7. Initiation prepend does not count
Reset-CompletionState
$prevReady = @{
    ($dgnGuid.ToLowerInvariant()) = 'QC Initiated'
    ($sheetGuid.ToLowerInvariant()) = 'QC Initiated'
    ($qcGuid.ToLowerInvariant()) = 'QC Initiated'
}
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $qcGuid `
    -TriggerDocumentName ($sheetStem + '-qc.pdf') -FolderPath $folder `
    -SourceState 'QC Initiated' -TargetState 'Ready for QC' -TransitionSource 'automation_prepend_completion' `
    -Members $productionMembers -PreviousStateByGuid $prevReady -JobId 'job-prepend-intake' -JobType 'QC_PREPEND' -SuppressNotification | Out-Null
Assert-Eq $script:completionCalls.Count 0 'Prepend initiation must not record QC cycle completion'
Assert-Eq $script:transitionCalls.Count 3 'Prepend initiation should still record sibling transitions'

# Sheet PDF already QC Complete; DGN/QC PDF finalize -> one package completion on DGN
Reset-CompletionState
$prevLag = @{
    ($dgnGuid.ToLowerInvariant()) = 'QC Finalizing'
    ($sheetGuid.ToLowerInvariant()) = 'QC Complete'
    ($qcGuid.ToLowerInvariant()) = 'QC Finalizing'
}
$liveComplete = @{
    ($dgnGuid.ToLowerInvariant()) = 'QC Complete'
    ($sheetGuid.ToLowerInvariant()) = 'QC Complete'
    ($qcGuid.ToLowerInvariant()) = 'QC Complete'
}
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $qcGuid `
    -TriggerDocumentName ($sheetStem + '-qc.pdf') -FolderPath $folder `
    -SourceState 'QC Finalizing' -TargetState 'QC Complete' -TransitionSource 'automation_prepend_completion' `
    -Members $productionMembers -PreviousStateByGuid $prevLag -StateByGuid $liveComplete `
    -JobId 'job-finalize-lag' -JobType 'QC_PREPEND' -SuppressNotification | Out-Null
Assert-Eq $script:completionCalls.Count 1 'Lag sheet PDF should not block package completion when siblings finalize'
Assert-Eq $script:completionCalls[0].documentGuid $dgnGuid 'Lag finalize should canonicalize to DGN GUID'

# Watcher echo after QC Complete does not create duplicate completions
Reset-CompletionState
$prevEcho = @{
    ($dgnGuid.ToLowerInvariant()) = 'QC Complete'
    ($sheetGuid.ToLowerInvariant()) = 'QC Complete'
    ($qcGuid.ToLowerInvariant()) = 'QC Complete'
}
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $sheetGuid `
    -TriggerDocumentName ($sheetStem + '.pdf') -FolderPath $folder `
    -SourceState 'QC Complete' -TargetState 'QC Complete' -TransitionSource 'user_audit' `
    -Members $members -PreviousStateByGuid $prevEcho -StateByGuid $prevEcho -AuditEventId 9050 | Out-Null
Assert-Eq $script:completionCalls.Count 0 'Watcher echo QC Complete -> QC Complete must not record completion'

# Unknown review type skipped
Reset-CompletionState
$unknownMembers = @(
    @{ documentGuid = $dgnGuid; documentName = ($sheetStem + '.dgn'); document = $null }
    @{ documentGuid = $sheetGuid; documentName = ($sheetStem + '.pdf'); document = [pscustomobject]@{ QC_Review_Type = 'Mystery Review' } }
)
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $sheetGuid `
    -TriggerDocumentName ($sheetStem + '.pdf') -FolderPath $folder `
    -SourceState 'QC Finalizing' -TargetState 'QC Complete' -TransitionSource 'user_audit' `
    -Members $unknownMembers -AuditEventId 9020
Assert-Eq $script:completionCalls.Count 0 'Unknown review type must not insert completion'
Assert-True ($script:jsonLogs | Where-Object { $_.code -eq 'QC_CYCLE_COMPLETION_SKIPPED' }) 'Unknown review type should log skip'

# No DGN resolvable -> canonical_dgn_not_found
Reset-CompletionState
$pdfOnlyMembers = @(
    @{ documentGuid = $sheetGuid; documentName = ($sheetStem + '.pdf'); document = [pscustomobject]@{ QC_Review_Type = 'Production QC' } }
)
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $sheetGuid `
    -TriggerDocumentName ($sheetStem + '.pdf') -FolderPath $folder `
    -SourceState 'QC Finalizing' -TargetState 'QC Complete' -TransitionSource 'user_audit' `
    -Members $pdfOnlyMembers -AuditEventId 9021
Assert-Eq $script:completionCalls.Count 0 'Missing DGN should not insert completion'
$skipLog = @($script:jsonLogs | Where-Object { $_.code -eq 'QC_CYCLE_COMPLETION_SKIPPED' })
Assert-True ($skipLog.Count -gt 0) 'Missing DGN should log skip'
Assert-Eq $skipLog[0].data.reason 'canonical_dgn_not_found' 'Missing DGN skip reason should be canonical_dgn_not_found'

# DGN resolvable but sheet_package_id missing -> sheet_package_not_found
Reset-CompletionState
function Resolve-QCCycleCompletionSheetPackageId {
    param([hashtable]$Config, [string]$DocumentGuid = '', [Nullable[guid]]$SheetPackageId = $null)
    return $null
}
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $dgnGuid `
    -TriggerDocumentName ($sheetStem + '.dgn') -FolderPath $folder `
    -SourceState 'QC Finalizing' -TargetState 'QC Complete' -TransitionSource 'user_audit' `
    -Members $members -PreviousStateByGuid $prevMap -AuditEventId 9022
Assert-Eq $script:completionCalls.Count 0 'Missing sheet_package_id should not insert completion'
$pkgSkip = @($script:jsonLogs | Where-Object { $_.code -eq 'QC_CYCLE_COMPLETION_SKIPPED' -and $_.data.reason -eq 'sheet_package_not_found' })
Assert-True ($pkgSkip.Count -gt 0) 'Missing package should log sheet_package_not_found'
function Resolve-QCCycleCompletionSheetPackageId {
    param([hashtable]$Config, [string]$DocumentGuid = '', [Nullable[guid]]$SheetPackageId = $null)
    if ($null -ne $SheetPackageId -and $SheetPackageId -ne [guid]::Empty) { return $SheetPackageId }
    if ([string]$DocumentGuid -eq $dgnGuid) { return $packageId }
    return $null
}

# Audit trigger path canonicalizes via StaleCheckMembers to DGN
Reset-CompletionState
Invoke-QCAuditWorkflowStateChangeTriggers -Config $cfg -DocumentGuid $dgnGuid `
    -DocumentName ($sheetStem + '.dgn') -FolderPath $folder `
    -PreviousState 'QC Finalizing' -CurrentState 'QC Complete' `
    -PwAttributes @{ qc_review_type = 'Independent Check' } -AuditEventId 9030 `
    -ChangedByUsername 'checker@example.com' -StaleCheckMembers $members
Assert-Eq $script:completionCalls.Count 1 'Audit trigger path should record one canonical completion'
Assert-Eq $script:completionCalls[0].documentGuid $dgnGuid 'Audit trigger should use DGN GUID'
Assert-Eq $script:completionCalls[0].qcReviewType 'independent_check' 'Audit trigger should store normalized review type'

# Duplicate suppression via unique key
Reset-CompletionState
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $qcGuid `
    -TriggerDocumentName ($sheetStem + '-qc.pdf') -FolderPath $folder `
    -SourceState 'QC Finalizing' -TargetState 'QC Complete' -TransitionSource 'automation_prepend_completion' `
    -Members $members -PreviousStateByGuid $prevMap -JobId 'job-dup-1' -SuppressNotification | Out-Null
Assert-Eq $script:completionCalls.Count 1 'First finalize should insert one completion'
Assert-Eq $script:completionCalls[0].documentGuid $dgnGuid 'Duplicate key should remain on DGN GUID'
$beforeDup = $script:completionCalls.Count
$script:jsonLogs.Clear()
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $dgnGuid `
    -TriggerDocumentName ($sheetStem + '.dgn') -FolderPath $folder `
    -SourceState 'QC Finalizing' -TargetState 'QC Complete' -TransitionSource 'user_audit' `
    -Members $members -PreviousStateByGuid $prevMap -AuditEventId 9060 | Out-Null
Assert-Eq $script:completionCalls.Count $beforeDup 'Duplicate suppression must keep one completion row'
Assert-True ($script:jsonLogs | Where-Object { $_.code -eq 'QC_CYCLE_COMPLETION_DUPLICATE' }) 'Retry should log QC_CYCLE_COMPLETION_DUPLICATE'

# Processor finalize: stale sheet_index but sheetStateSync proves package finalized
Reset-CompletionState
$script:cycleByGuid.Clear()
$script:cycleByGuid[($folder + '|' + $sheetStem).ToLowerInvariant()] = @{ cycleId = 'qc_qcprepend_d0ca0819859b391d'; cycleNumber = '1' }
$processorCtx = @{
    sheetStateSync = @{
        updates = @(
            @{ documentGuid = $dgnGuid; documentName = ($sheetStem + '.dgn'); fromState = 'QC Finalizing'; toState = 'QC Complete'; applied = $true }
            @{ documentGuid = $sheetGuid; documentName = ($sheetStem + '.pdf'); fromState = 'QC Complete'; toState = 'QC Complete'; skipped = 'already_at_target' }
            @{ documentGuid = $qcGuid; documentName = ($sheetStem + '-qc.pdf'); fromState = 'QC Finalizing'; toState = 'QC Complete'; applied = $true }
        )
    }
    attributes = @{
        cycleId = 'qc_qcprepend_d0ca0819859b391d'
        cycleNumber = '1'
        qcReviewType = 'Production QC'
    }
}
$staleIndex = @{
    ($dgnGuid.ToLowerInvariant()) = 'QC Complete'
    ($sheetGuid.ToLowerInvariant()) = 'QC Complete'
    ($qcGuid.ToLowerInvariant()) = 'QC Complete'
}
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $qcGuid `
    -TriggerDocumentName ($sheetStem + '-qc.pdf') -FolderPath $folder `
    -SourceState 'QC Finalizing' -TargetState 'QC Complete' -TransitionSource 'automation_prepend_completion' `
    -Members $productionMembers -PreviousStateByGuid $staleIndex -Context $processorCtx `
    -JobId 'qc_qcprepend_d0ca0819859b391d' -JobType 'QC_PREPEND' -SuppressNotification | Out-Null
Assert-Eq $script:completionCalls.Count 1 'Processor finalize should record completion via sheetStateSync fallback'
Assert-Eq $script:completionCalls[0].documentGuid $dgnGuid 'Processor finalize should use DGN GUID'
Assert-Eq $script:completionCalls[0].qcCycleId 'qc_qcprepend_d0ca0819859b391d' 'Processor finalize should use context cycle id'
Assert-Eq $script:summaryCalls[0] $packageId.ToString() 'Processor finalize rollup should target sheet_packages'

# Empty sheet_index previous state must not throw or count as QC Complete (index baseline)
Reset-CompletionState
$emptyPrevMap = @{
    ($dgnGuid.ToLowerInvariant()) = ''
    ($sheetGuid.ToLowerInvariant()) = ''
    ($qcGuid.ToLowerInvariant()) = ''
}
Invoke-QCSheetGroupWorkflowTransition -Config $cfg -TriggerDocumentGuid $sheetGuid `
    -TriggerDocumentName ($sheetStem + '.pdf') -FolderPath $folder `
    -SourceState '' -TargetState 'QC Complete' -TransitionSource 'user_audit' `
    -Members $members -PreviousStateByGuid $emptyPrevMap -AuditEventId 9070 -SuppressNotification | Out-Null
Assert-Eq $script:completionCalls.Count 0 'Empty previous state must not record QC cycle completion'

Write-Host 'ALL PASSED' -ForegroundColor Green
