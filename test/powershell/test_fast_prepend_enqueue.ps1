# Fast DOCUMENT_STATE prepend enqueue: watcher order, deferred lane, worker preflight.
$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $PSScriptRoot '_Resolve-ModuleImplPath.ps1')

Import-Module "$repoRoot/modules/Core/Core.Results.psm1" -Force
Import-Module "$repoRoot/modules/Workflow/QC.ProcessType.psm1" -Force
Import-Module "$repoRoot/modules/Core/QC.WatcherOrchestration.psm1" -Force
Import-Module "$repoRoot/modules/Workflow/QC.Workflow.psm1" -Force
Import-Module "$repoRoot/modules/Processing/QC.Processors.psm1" -Force

# --- Flag parsing ---
Assert-True (-not (Test-QCFastAuditEnqueueEnabled -Config @{})) 'Missing flag is false'
Assert-True (-not (Test-QCFastAuditEnqueueEnabled -Config @{ qcPrepend = @{} })) 'Empty qcPrepend is false'
Assert-True (-not (Test-QCFastAuditEnqueueEnabled -Config @{ qcPrepend = @{ fastAuditEnqueue = $false } })) 'Explicit false'
Assert-True (Test-QCFastAuditEnqueueEnabled -Config @{ qcPrepend = @{ fastAuditEnqueue = $true } }) 'Explicit true'

# --- Watcher source call order ---
$watchPath = Join-Path $repoRoot 'scripts\service\Watch-QCTrigger.ps1'
$watchText = Get-Content -LiteralPath $watchPath -Raw -Encoding UTF8
Assert-True ($watchText -match 'Test-QCFastAuditEnqueueEnabled') 'Watcher reads qcPrepend.fastAuditEnqueue'
Assert-True ($watchText -match 'Invoke-QCFastAuditPrependEnqueue') 'Watcher calls fast prepend enqueue'
Assert-True ($watchText -match 'WATCH_FAST_AUDIT_SYNC_PASS') 'Watcher runs deferred sibling sync after enqueue'
Assert-True ($watchText -match 'skipPathDForPrepend') 'Watcher can skip Path D after fast enqueue'
Assert-True ($watchText -match 'WATCH_AUDIT_PATH_D_SKIPPED_FAST_ENQUEUE') 'Watcher logs Path D skip for fast enqueue'

$idxFlag = $watchText.IndexOf('Test-QCFastAuditEnqueueEnabled -Config $config')
$idxForeach = $watchText.IndexOf('foreach ($ac in $auditCandidates)')
$idxFastEnq = $watchText.IndexOf('Invoke-QCFastAuditPrependEnqueue -Config $config')
$idxSyncImmediate = $watchText.IndexOf('Sync-PWAssociatedSheetWorkflowState @syncPwParams')
$idxSecondPass = $watchText.IndexOf('WATCH_FAST_AUDIT_SYNC_PASS')
$idxPathD = $watchText.IndexOf('-not $skipPathDForPrepend')
Assert-True ($idxFlag -ge 0 -and $idxFlag -lt $idxForeach) 'Flag is read before the audit candidate loop'
Assert-True ($idxFastEnq -gt $idxForeach) 'Fast enqueue runs inside the candidate loop'
Assert-True ($idxFastEnq -lt $idxSecondPass) 'Fast enqueue happens before the deferred Sync-PW pass'
Assert-True ($idxSyncImmediate -gt $idxFastEnq) 'Flag-off Sync-PW stays in the DOCUMENT_STATE branch after the fast-path if'
Assert-True ($idxSyncImmediate -lt $idxSecondPass) 'In-loop Sync-PW (flag off) is still before the second pass'
Assert-True ($idxPathD -gt $idxFastEnq -and $idxPathD -lt $idxSecondPass) 'Path D skip gate is after enqueue and before deferred sync'
$afterForeachClose = $watchText.LastIndexOf('foreach ($syncPwParams in $fastAuditSyncPass)')
Assert-True ($afterForeachClose -gt $idxPathD) 'Second-pass Sync-PW iterates after Path D (and therefore after all fast enqueues)'

# Flag-off preserves current order: Sync-PW in the else of fastAuditEnqueue, still inside DOCUMENT_STATE.
$docStateBlock = $watchText.Substring($watchText.IndexOf('WATCH_AUDIT_DOCUMENT_STATE_CONTEXT'), $idxSecondPass - $watchText.IndexOf('WATCH_AUDIT_DOCUMENT_STATE_CONTEXT'))
Assert-True ($docStateBlock -match 'if \(\$fastAuditEnqueue\)') 'DOCUMENT_STATE branches on the fast-enqueue flag'
Assert-True ($docStateBlock -match 'Sync-PWAssociatedSheetWorkflowState @syncPwParams') 'Flag off still syncs siblings before the next sheet'

# --- Fast enqueue eligibility ---
InModuleScope -ModuleName QC.Processors {
    $script:addInitiatedCalls = 0
    $script:addFinalizingCalls = 0
    function Test-QCIsSheetPdfDocumentName { param([string]$DocumentName) return ($DocumentName -match '(?i)\.pdf$' -and $DocumentName -notmatch '(?i)-(prod|chk|rev)\.pdf$') }
    function Test-QCIsStatusSetOutputPdfName { param([string]$FileName) return ($FileName -match '(?i)_StatusSet\.pdf$') }
    function Test-QCIsAutomationPwActor { param($Config, $ChangedByUser, $ChangedByUsername) return ([string]$ChangedByUsername -eq 'svc_qc') }
    function Test-QCWorkflowStateIsQcInitiated { param([string]$StateName, [hashtable]$Config) return ($StateName -eq 'Initiate Origination') }
    function Test-QCWorkflowStateIsQcFinalizing { param([string]$StateName, [hashtable]$Config) return ($StateName -eq 'Initiate Verification') }
    function Add-QCPrependJobForQcInitiatedStateChange {
        param($Config, $TriggerDocumentGuid, $TriggerDocumentName, $FolderPath, $CurrentStateName, $DryRun, $ChangedByUser, $ChangedByUsername, $LastAuditEventAt, $AuditEventId)
        $script:addInitiatedCalls++
        return New-QCSuccessResult -Code 'QUEUE_ENQUEUED' -Message 'queued' -Data @{ jobId = 'j1' }
    }
    function Add-QCPrependJobForQcFinalizingStateChange {
        param($Config, $TriggerDocumentGuid, $TriggerDocumentName, $FolderPath, $CurrentStateName, $DryRun, $ChangedByUser, $ChangedByUsername, $LastAuditEventAt, $AuditEventId)
        $script:addFinalizingCalls++
        return New-QCSuccessResult -Code 'QUEUE_ENQUEUED' -Message 'queued' -Data @{ jobId = 'j2' }
    }

    $cfg = @{ qcPrepend = @{ fastAuditEnqueue = $true } }
    $base = @{
        Config = $cfg
        TriggerDocumentGuid = '11111111-1111-1111-1111-111111111111'
        FolderPath = 'Documents\proj\CADD\Sheets\N_Seg'
        EnableQcPrepend = $true
    }

    $notSheets = Invoke-QCFastAuditPrependEnqueue @base -TriggerDocumentName 'sheet.pdf' -LiveStateName 'Initiate Origination' -IsSheetsFolder:$false
    if ($notSheets.reason -ne 'not_sheets_folder') { throw "ASSERT FAILED: Non-Sheets folder does not enqueue (got $($notSheets.reason))" }
    if ($notSheets.attempted) { throw 'ASSERT FAILED: Non-Sheets is not an enqueue attempt' }

    $lanePdf = Invoke-QCFastAuditPrependEnqueue @base -TriggerDocumentName 'sheet-rev.pdf' -LiveStateName 'Initiate Origination' -IsSheetsFolder:$true
    if ($lanePdf.reason -ne 'not_stem_sheet_pdf') { throw "ASSERT FAILED: Lane PDF does not enqueue from this path (got $($lanePdf.reason))" }

    $auto = Invoke-QCFastAuditPrependEnqueue @base -TriggerDocumentName 'sheet.pdf' -LiveStateName 'Initiate Origination' -IsSheetsFolder:$true -ChangedByUsername 'svc_qc'
    if ($auto.reason -ne 'automation_actor') { throw "ASSERT FAILED: Automation actor does not enqueue (got $($auto.reason))" }

    $wrongState = Invoke-QCFastAuditPrependEnqueue @base -TriggerDocumentName 'sheet.pdf' -LiveStateName 'In Development' -IsSheetsFolder:$true
    if ($wrongState.reason -ne 'state_not_prepend') { throw "ASSERT FAILED: Non-prepend live state does not enqueue (got $($wrongState.reason))" }
    if ($wrongState.skipPathD) { throw 'ASSERT FAILED: Path D remains available when state is not a prepend trigger' }

    $script:addInitiatedCalls = 0
    $okInit = Invoke-QCFastAuditPrependEnqueue @base -TriggerDocumentName '0818000063ia501.pdf' -LiveStateName 'Initiate Origination' -IsSheetsFolder:$true
    if (-not $okInit.attempted) { throw 'ASSERT FAILED: Initiated state attempts enqueue' }
    if (-not $okInit.enqueued) { throw 'ASSERT FAILED: Initiated enqueue reports success' }
    if (-not $okInit.skipPathD) { throw 'ASSERT FAILED: Successful initiated enqueue skips Path D' }
    if ($script:addInitiatedCalls -ne 1) { throw "ASSERT FAILED: Initiated helper called once (got $($script:addInitiatedCalls))" }

    $script:addFinalizingCalls = 0
    $okFin = Invoke-QCFastAuditPrependEnqueue @base -TriggerDocumentName '0818000063ia502.pdf' -LiveStateName 'Initiate Verification' -IsSheetsFolder:$true
    if (-not $okFin.enqueued) { throw 'ASSERT FAILED: Finalizing enqueue reports success' }
    if ($script:addFinalizingCalls -ne 1) { throw "ASSERT FAILED: Finalizing helper called once (got $($script:addFinalizingCalls))" }

    $script:addInitiatedCalls = 0
    1..5 | ForEach-Object {
        $r = Invoke-QCFastAuditPrependEnqueue @base -TriggerDocumentName ('0818000063ia50' + $_ + '.pdf') -LiveStateName 'Initiate Origination' -IsSheetsFolder:$true
        if (-not $r.enqueued) { throw "ASSERT FAILED: Burst sheet $_ should enqueue" }
    }
    if ($script:addInitiatedCalls -ne 5) { throw "ASSERT FAILED: Five initiated sheets enqueue without sibling sync (got $($script:addInitiatedCalls))" }
}

Import-Module "$repoRoot/modules/Processing/QC.Processors.psm1" -Force

# --- Add-QCPrependJob*: fast path does not require live process type ---
InModuleScope -ModuleName QC.Processors {
    $script:queueAdds = 0
    function Get-QCWorkflowSettings { param($Config) return @{ states = @{ qcInitiated = 'Initiate Origination'; qcFinalizing = 'Initiate Verification' } } }
    function Get-QCWorkflowStateName {
        param($Settings, $StateKey)
        if ($StateKey -eq 'qcInitiated') { return 'Initiate Origination' }
        if ($StateKey -eq 'qcFinalizing') { return 'Initiate Verification' }
        return $StateKey
    }
    function Test-QCIsAutomationPwActor { param($Config, $ChangedByUser, $ChangedByUsername) return $false }
    function Test-QCPrependBlockedByMissingEmailAttributes { param($Config, $FolderPath, $SheetPdfName, $DocumentGuid) return @{ blocked = $false } }
    function _QCP-EnsureQueueModulesLoaded { }
    function Test-QCPrependEnqueueBlockedForSheet { param($Config, $FolderPath, $SheetPdfName, $PrependTrigger, $QcProcessType, $StateTransitionKey) return @{ blocked = $false } }
    function Test-QCDuplicateJob { param($DedupeKey, $Config) return New-QCSuccessResult -Code 'OK' -Data @{ isDuplicate = $false } }
    function New-QCJobObject {
        param($Candidate, $Rule, $Config)
        return New-QCSuccessResult -Code 'OK' -Data @{
            job = @{
                id = 'qc_qcprepend_fast1'
                type = 'QC_PREPEND'
                dedupeKey = 'dk1'
                sourceFolder = [string]$Candidate.sourceFolder
                sourceName = [string]$Candidate.fileName
                sourcePath = [string]$Candidate.path
                metadata = @{}
            }
        }
    }
    function Add-QCQueueJob {
        param($Job, $Config)
        $script:queueAdds++
        return New-QCSuccessResult -Code 'QUEUE_ENQUEUED' -Message 'queued' -Data @{ job = $Job }
    }
    function _QCP-TryResolvePrependLaneContext {
        param($Job, $Config)
        return New-QCFailureResult -Code 'QC_PROCESS_TYPE_UNKNOWN' -Message 'no process type in test' -Data @{}
    }
    function _QCP-ResolveProcessTypeFromSheetIndex { param($Config, $FolderPath, $SourceDocumentName) return '' }
    function _QCP-ResolveIntendedPrependProcessType { param($Config, $FolderPath, $SheetPdfName, $SheetPdfGuid) return '' }
    function Write-QCJsonLog { param($Level, $Code, $Message, $Data) }
    function Get-QCTimestamp { return '2026-08-24T22:00:00.0000000Z' }

    $cfgOff = @{ qcPrepend = @{ fastAuditEnqueue = $false } }
    $cfgOn = @{ qcPrepend = @{ fastAuditEnqueue = $true } }
    $common = @{
        TriggerDocumentGuid = '22222222-2222-2222-2222-222222222222'
        TriggerDocumentName = '0818000063ia501.pdf'
        FolderPath = 'Documents\proj\CADD\Sheets\N_Seg'
        ChangedByUsername = 'jflint'
    }

    $script:queueAdds = 0
    $off = Add-QCPrependJobForQcInitiatedStateChange @common -Config $cfgOff -CurrentStateName 'Initiate Origination'
    Assert-True ($null -ne $off) 'Flag-off initiated enqueue returns a result'
    Assert-Eq ([string]$off.Code) 'QC_PREPEND_SKIPPED_NO_PROCESS_TYPE' 'Flag off still skips when process type is missing'
    Assert-Eq $script:queueAdds 0 'Flag off does not queue without process type'

    $script:queueAdds = 0
    $on = Add-QCPrependJobForQcInitiatedStateChange @common -Config $cfgOn -CurrentStateName 'Initiate Origination'
    Assert-True $on.IsSuccess 'Fast path enqueue succeeds without live process type'
    Assert-Eq ([string]$on.Code) 'QUEUE_ENQUEUED' 'Fast path uses the queue'
    Assert-Eq $script:queueAdds 1 'Fast path queued despite missing process type'

    $script:queueAdds = 0
    $finOff = Add-QCPrependJobForQcFinalizingStateChange @common -Config $cfgOff -CurrentStateName 'Initiate Verification'
    Assert-Eq ([string]$finOff.Code) 'QC_PREPEND_SKIPPED_NO_PROCESS_TYPE' 'Flag-off finalizing still requires process type'
    $finOn = Add-QCPrependJobForQcFinalizingStateChange @common -Config $cfgOn -CurrentStateName 'Initiate Verification'
    Assert-Eq ([string]$finOn.Code) 'QUEUE_ENQUEUED' 'Fast path finalizing enqueues without process type'
    Assert-Eq $script:queueAdds 1 'Fast path finalizing queued once'
}

Import-Module "$repoRoot/modules/Processing/QC.Processors.psm1" -Force

# --- Worker preflight: skip = success, no overlay; writeback resume skips preflight ---
InModuleScope -ModuleName QC.Processors {
    $script:spawnCalls = 0
    $script:writebackCalls = 0
    $script:preflightCalls = 0
    function _QCP-StartAndWaitLaunchedProcess {
        param(
            [string]$FilePath, [string]$ArgumentList, [string]$StdoutPath, [string]$StderrPath,
            [hashtable]$Job = $null, [hashtable]$Config = $null,
            [int]$TimeoutSeconds = 0, [int]$PollMilliseconds = 500, [int]$HeartbeatSeconds = 15
        )
        $script:spawnCalls++
        return @{
            process = $null; processId = 7; exited = $true; timedOut = $false; exitCode = 0
            elapsedMs = 5; stdout = 'ok'; stderr = ''; startFailed = $false
        }
    }
    function _QCP-InvokePrependPostSuccessWriteback {
        param(
            [hashtable]$Job, [hashtable]$Config, [object]$SuccessResult,
            [string]$IncomingFolder, [string]$IncomingDocName, [string]$DatasourceName,
            [hashtable]$ProjectWiseCfg, [switch]$ClearTriggerTag
        )
        $script:writebackCalls++
        return New-QCSuccessResult -Code 'QC_PREPEND_OK' -Message 'writeback ok' -Data @{
            resumedFromCheckpoint = [bool]$SuccessResult.Data.resumedFromCheckpoint
        }
    }
    function Get-PWDocumentWorkflowStateName { param($FolderPath, $DocumentName, $DocumentGuid) return 'In Development' }
    function Test-PWSheetPdfHasMatchingPair { param($FolderPath, $DocumentName) return $true }
    function Test-QCWorkflowStateIsQcInitiated { param([string]$StateName, [hashtable]$Config) return $false }
    function Test-QCWorkflowStateIsQcFinalizing { param([string]$StateName, [hashtable]$Config) return $false }

    $job = @{
        id = 'qc_qcprepend_preflight_skip'
        type = 'QC_PREPEND'
        sourceFolder = 'documents\proj\cadd\sheets\n_seg'
        sourceName = '0818000063ia501.pdf'
        sourcePath = 'documents\proj\cadd\sheets\n_seg\0818000063ia501.pdf'
        metadata = @{ qcProcessType = 'review'; prependTrigger = 'initialQcPdf' }
    }
    $cfg = @{
        dryRun = $false
        qcPrepend = @{ mode = 'projectWise' }
        projectWise = @{ datasourceName = 'typsa-us-pw.bentley.com:typsa-us-pw-03' }
    }

    $skipState = Invoke-QCPrependProcessor -Job $job -Config $cfg
    Assert-True $skipState.IsSuccess 'Non-initiated live state is a successful skip'
    Assert-Eq ([string]$skipState.Code) 'QC_PREPEND_SKIPPED_NOT_QC_INITIATED' 'Skip code is NOT_QC_INITIATED'
    Assert-Eq $script:spawnCalls 0 'Overlay child must not start on state skip'

    function Get-PWDocumentWorkflowStateName { param($FolderPath, $DocumentName, $DocumentGuid) return 'Initiate Origination' }
    function Test-QCWorkflowStateIsQcInitiated { param([string]$StateName, [hashtable]$Config) return $true }
    function Test-PWSheetPdfHasMatchingPair { param($FolderPath, $DocumentName) return $false }
    $skipPair = Invoke-QCPrependProcessor -Job $job -Config $cfg
    Assert-True $skipPair.IsSuccess 'Missing DGN pair is a successful skip'
    Assert-Eq ([string]$skipPair.Code) 'QC_PREPEND_SKIPPED_NO_DGN_PAIR' 'Skip code is NO_DGN_PAIR'
    Assert-Eq $script:spawnCalls 0 'Overlay child must not start on DGN skip'

    function Test-PWSheetPdfHasMatchingPair { param($FolderPath, $DocumentName) return $true }
    function Test-QCWorkflowStateIsQcFinalizing { param([string]$StateName, [hashtable]$Config) return $true }
    function Get-PWDocumentWorkflowStateName { param($FolderPath, $DocumentName, $DocumentGuid) return 'Initiate Verification' }
    function _QCP-TryResolvePrependLaneContext {
        param($Job, $Config)
        return New-QCFailureResult -Code 'QC_PROCESS_TYPE_UNKNOWN' -Message 'no lane' -Data @{}
    }
    $jobFinal = @{
        id = 'qc_qcprepend_preflight_lane'
        type = 'QC_PREPEND'
        sourceFolder = 'documents\proj\cadd\sheets\n_seg'
        sourceName = '0818000063ia501.pdf'
        sourcePath = 'documents\proj\cadd\sheets\n_seg\0818000063ia501.pdf'
        metadata = @{ prependTrigger = 'finalQcComplete' }
    }
    $skipLane = Invoke-QCPrependProcessor -Job $jobFinal -Config $cfg
    Assert-True $skipLane.IsSuccess 'Missing process type at preflight is a successful skip'
    Assert-Eq ([string]$skipLane.Code) 'QC_PREPEND_SKIPPED_NO_PROCESS_TYPE' 'Skip code is NO_PROCESS_TYPE'
    Assert-Eq $script:spawnCalls 0 'Overlay child must not start when lane is unresolved'

    $script:preflightCalls = 0
    $script:spawnCalls = 0
    $script:writebackCalls = 0
    function Invoke-QCPrependPreflight {
        param($Job, $Config, $IncomingFolder, $IncomingDocName)
        $script:preflightCalls++
        throw 'preflight must not run on writeback-only resume'
    }
    $resumeJob = @{
        id = 'qc_qcprepend_preflight_resume'
        type = 'QC_PREPEND'
        checkpoint = 'prepend_complete'
        sourceFolder = 'documents\proj\cadd\sheets\n_seg'
        sourceName = '0818000063ia501.pdf'
        sourcePath = 'documents\proj\cadd\sheets\n_seg\0818000063ia501.pdf'
        metadata = @{ qcProcessType = 'review'; prependTrigger = 'initialQcPdf' }
    }
    $resume = Invoke-QCPrependProcessor -Job $resumeJob -Config $cfg
    Assert-True $resume.IsSuccess 'Writeback-only resume succeeds'
    Assert-Eq $script:preflightCalls 0 'Writeback resume must not re-run preflight'
    Assert-Eq $script:spawnCalls 0 'Writeback resume must not spawn overlay'
    Assert-Eq $script:writebackCalls 1 'Writeback resume still runs writeback'
}

Write-Host 'All fast prepend enqueue tests passed.' -ForegroundColor Green
