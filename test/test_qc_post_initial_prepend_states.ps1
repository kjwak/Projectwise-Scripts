# Post-Initiate Origination prepend: lane PDF -> Originated; stem PDF -> In Development (verified); DGN untouched.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.ProcessType.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.Workflow.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\PW.Discovery.psm1') -Force

function Assert-True($cond, $msg) { if (-not $cond) { throw "ASSERT FAILED: $msg" } }
function Assert-Eq($a, $b, $msg) { if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" } }
function Assert-False($cond, $msg) { if ($cond) { throw "ASSERT FAILED: $msg" } }

Assert-True (Test-QCResetProcessTypeAfterLanePrepend -Config @{}) 'ResetProcessTypeAfterLanePrepend defaults true'
Assert-True (Test-QCProcessTypeResetsAfterPrepend -ProcessType 'production' -Config @{}) 'production lane resets to Production after prepend by default'
Assert-True (Test-QCProcessTypeResetsAfterPrepend -ProcessType 'review' -Config @{}) 'review lane resets to Production after prepend by default'

# Lane QC PDFs are never modified by associated review/process type sync (legacy path only)
InModuleScope -ModuleName PW.Discovery {
    function Write-QCJsonLog { param($Level, $Code, $Message, $Data) }
    function Get-PWQcProcessTypeAttributeName { param($Config) return 'QC_Process_Type' }
    function Test-PWQcReviewTypeAttributesEnabled { param($Config, $FolderPath) return $true }
    function _PWD-GetSheetIndexQcReviewType { param($Config, $DocumentGuid) return '' }
    function _PWD-NormalizeSheetIndexValue { param($Value) return ([string]$Value).Trim().ToLowerInvariant() }
    function _PWD-GetPwAttributeValue { param($PwAttributes, $ColumnName) return '' }
    function _PWD-ResolvePwDocumentInFolder { param($DocByGuid, $FolderPath, $DocumentName, $DocumentGuid) return [pscustomobject]@{ Name = $DocumentName } }
    function _PWD-InvokeUpdatePWDocumentAttributes { param($Document, $Attributes, $Config) $script:attrWrites += ,@($Attributes); return $true }
    function Get-PWAssociatedSheetSyncMembers {
        param($Config, $FolderPath, $DocumentName, $DocumentGuid, $TriggerSource)
        return @(
            @{ documentGuid = '1'; documentName = 'CA001.pdf' }
            @{ documentGuid = '2'; documentName = 'CA001.dgn' }
            @{ documentGuid = '3'; documentName = 'CA001-rev.pdf' }
        )
    }
    function Get-PWDocumentAttributesByColumns { param($FolderPath, $DocumentName, $ColumnsToReturn) return @{ found = $true; attributes = @{} } }
    function Get-PWQcPdfLaneFromDocumentName { param($DocumentName) if ($DocumentName -match '(?i)-rev\.pdf$') { return 'review' } return $null }
    function Invoke-QCAuditWorkflowAttributeChangeTriggers { param($Config, $DocumentGuid, $DocumentName, $FolderPath, $FieldChanges) }
    function Write-QCSheetIndex { param($Config, $DocumentGuid, $DocumentName, $FolderPath, $WatchRoot, $Extension, $SourceType, $QcReviewType, $LastAuditEventAt, $SetOwnershipFromProjectWise) }

    $legacyCfg = @{ QCProcess = @{ EnableLegacyReviewTypeAttributeSync = $true; EnableLegacySiblingStateSync = $true } }
    $script:attrWrites = @()
    Sync-PWAssociatedSheetReviewTypeAttributes -Config $legacyCfg -DocumentGuid '1' -DocumentName 'CA001.pdf' `
        -FolderPath 'Drawings\X' -CanonicalReviewType 'production'
    Assert-Eq $script:attrWrites.Count 1 'lane QC PDF and DGN excluded from legacy process type sync'
    Assert-False ($script:attrWrites[0].ContainsKey('QC_Review_Type')) 'legacy sync writes QC_Process_Type only'
}

# Sync-PWPostInitialPrependLaneStates uses ExpectedLanePdfName and rejects empty process type
InModuleScope -ModuleName PW.Discovery {
    function Write-QCJsonLog { param($Level, $Code, $Message, $Data) }
    function Get-PWSheetStemFromDocumentName { param($DocumentName) return 'CA001' }
    function Format-QCWorkflowStateName { param($StateName, $Config) return $StateName }
    function Format-QCProcessTypeAttributeValue { param($ProcessType) switch ($ProcessType) { 'review' { return 'Review' } 'check' { return 'Check' } default { return 'Production' } } }
    function Get-PWQcProcessTypeAttributeName { param($Config) return 'QC_Process_Type' }
    function Test-PWQcReviewTypeAttributesEnabled { param($Config, $FolderPath) return $true }
    function Get-PWDocumentAttributesByColumns { param($FolderPath, $DocumentName, $ColumnsToReturn) return @{ found = $true; attributes = @{} } }
    function _PWD-InvokeUpdatePWDocumentAttributes { param($Document, $Attributes, $Config) $script:laneAttrWrite = $Attributes; return $true }
    function _PWD-ResolvePwDocumentInFolder {
        param($DocByGuid, $FolderPath, $DocumentName, $DocumentGuid)
        return [pscustomobject]@{ DocumentGUID = 'lane-guid'; Name = $DocumentName }
    }
    function Get-PWDocumentWorkflowStateName { param($FolderPath, $DocumentName, $DocumentGuid) return 'Initiate Origination' }
    function _PWD-InvokeSetPwDocumentState { param($Document, $StateName, $GuardContext) }
    function Update-QCSheetIndexPwStateName { param($Config, $DocumentGuid, $PwStateName) return $true }
    function Get-PWQcPdfLaneFromDocumentName { param($DocumentName) if ($DocumentName -match '(?i)-rev\.pdf$') { return 'review' } return $null }
    function Update-SheetPackageQcPdfLaneState { param($Config, $DocumentGuid, $CurrentPwState, $QcProcessType) return $true }
    function _PWD-SyncReferenceSheetProcessTypeAttributes {
        param($Config, $DocumentGuid, $DocumentName, $FolderPath, $CanonicalProcessType, $WatchRoot, $LastAuditEventAt, $DryRun, $ControlDocumentOnly)
        return @{ reset = $true; toProcessType = $CanonicalProcessType; documentName = $DocumentName }
    }
    function Get-QCWorkflowSettings { param($Config) return @{ expectedWorkflowName = 'TYPSA QC' } }
    function Test-QCWorkflowStateWritebackUseForce { param($Config) return $true }
    function Set-PWDocumentWorkflowStateVerified {
        param($Config, $FolderPath, $DocumentName, $DocumentGuid, $TargetState, $QcProcessType, $DryRun, $WorkflowName, $IsLaneAuthority)
        return @{
            documentGuid = 'guid-' + $DocumentName
            documentName = $DocumentName
            fromState = 'Initiate Origination'
            targetState = $TargetState
            readBackState = $TargetState
            applied = $true
            verified = $true
            usedForce = $true
            isLaneAuthority = [bool]$IsLaneAuthority
            commandShape = "Set-PWDocumentState -InputDocuments @(`$cleanDoc) -State '$TargetState' -Force"
        }
    }

    $split = Sync-PWPostInitialPrependLaneStates -Config @{} -FolderPath 'Drawings\X' `
        -DocumentName 'CA001.pdf' -DocumentGuid 'stem-guid' -QcProcessType 'review' `
        -ExpectedLanePdfName 'CA001-rev.pdf' -LaneTargetState 'Originated' -ReferenceState 'In Development'
    Assert-Eq $split.lanePdfName 'CA001-rev.pdf' 'uses ExpectedLanePdfName'
    Assert-Eq $split.laneTargetState 'Originated' 'lane target preserved'
    Assert-Eq $split.referenceState 'In Development' 'reference target preserved'
    Assert-True $split.writeStemPdfReferenceState 'prepend writes stem PDF back to reference state'
    Assert-Eq $split.stemPdfName 'CA001.pdf' 'stem PDF name resolved'
    Assert-True ($null -ne $split.processTypeReset) 'post-prepend process type reset attempted'
    Assert-True $split.laneProcessTypeEnsure.ensured 'sets lane qc_process_type when unset'
    Assert-Eq $split.laneProcessTypeEnsure.toValue 'Review' 'lane qc_process_type matches triggering review lane'

    $script:laneAttrWrite = $null
    function Get-PWDocumentAttributesByColumns {
        param($FolderPath, $DocumentName, $ColumnsToReturn)
        return @{ found = $true; attributes = @{ 'QC_Process_Type' = 'Production' } }
    }
    $wrong = Sync-PWPostInitialPrependLaneStates -Config @{} -FolderPath 'Drawings\X' `
        -DocumentName 'CA001.pdf' -DocumentGuid 'stem-guid' -QcProcessType 'check' `
        -ExpectedLanePdfName 'CA001-chk.pdf' -LaneTargetState 'Originated' -ReferenceState 'In Development'
    Assert-True $wrong.laneProcessTypeEnsure.ensured 'corrects wrong lane qc_process_type'
    Assert-True $wrong.laneProcessTypeEnsure.corrected 'marks lane qc_process_type correction'
    Assert-Eq $wrong.laneProcessTypeEnsure.toValue 'Check' 'corrects wrong lane qc_process_type to Check'

    $empty = Sync-PWPostInitialPrependLaneStates -Config @{} -FolderPath 'Drawings\X' `
        -DocumentName 'CA001.pdf' -QcProcessType '' -LaneTargetState 'Originated' -ReferenceState 'In Development'
    Assert-True $empty.skipped 'empty process type skips split'
    Assert-Eq $empty.skipReason 'empty_qc_process_type' 'empty process type reason'
}
# Set-PWQCWorkflowState: initialQcPdf always splits lane states even when legacy sibling sync is enabled
$script:stateWrites = @()
function Get-PWWorkflows { [CmdletBinding()] param() [pscustomobject]@{ Name = 'TYPSA QC' } }
function Get-PWWorkflowStateLinks { [CmdletBinding()] param() @() }
function Set-PWDocumentState {
    [CmdletBinding()] param([object[]]$InputDocuments, [string]$StateName, [switch]$ReturnBoolean)
    $script:stateWrites += $StateName
    return $true
}

$cfg = @{
    qcWorkflow = @{
        enabled = $true
        strictMode = $false
        dryRunWriteback = $false
        mode = 'StateAndAttributes'
        autoSetState = $true
        expectedWorkflowName = 'TYPSA QC'
        states = @{
            production = 'In Development'
            readyForQc = 'Originated'
            qcInitiated = 'Initiate Origination'
        }
        stateAfterPrependByTrigger = @{ initialQcPdf = 'Originated' }
    }
    QCProcess = @{
        EnableLegacySiblingStateSync = $true
        ProcessTypes = @{ production = @{ SyncWithSiblingSheets = $true } }
    }
}
$settings = Get-QCWorkflowSettings -Config $cfg
$doc = [pscustomobject]@{ WorkflowName = 'TYPSA QC'; StateName = 'Initiate Origination'; DocumentGUID = 'stem-guid'; Name = 'CA001.pdf'; FolderPath = 'Drawings\X' }
$ctx = @{
    config = $cfg
    prependTrigger = 'initialQcPdf'
    qcProcessType = 'production'
    expectedLanePdfName = 'CA001-prod.pdf'
    skipSiblingStateSync = $false
    job = @{ id = 'j1'; sourceFolder = 'Drawings\X'; sourceName = 'CA001.pdf' }
    document = $doc
}
$result = Set-PWQCWorkflowState -Settings $settings -Context $ctx -StateName 'Originated' -DryRun:$false
Assert-True $result.IsSuccess 'initialQcPdf state write should succeed'
Assert-True ($null -ne $result.Data.lanePostPrependSplit) 'lane split should run for initialQcPdf even when sibling sync would be enabled'
Assert-Eq $result.Data.lanePostPrependSplit.lanePdfName 'CA001-prod.pdf' 'lane split uses expected lane PDF name'
Assert-True (-not $result.Data.sheetStateSync) 'sibling sync must not run after lane-independent initial prepend'
Assert-Eq $script:stateWrites.Count 0 'stem primary write skipped; lane PDF is workflow authority'
Assert-True $result.Data.skippedPrimaryReferenceWrite 'marks skipped primary reference write'
Assert-True $ctx.laneIndependentInitialPrepend 'context marks lane-independent prepend'

# finalQcComplete: lane split without stem reference state must not throw on empty referenceState
$finalCfg = @{
    qcWorkflow = @{
        enabled = $true
        strictMode = $false
        dryRunWriteback = $false
        mode = 'StateAndAttributes'
        autoSetState = $true
        expectedWorkflowName = 'TYPSA QC'
        states = @{
            production = 'In Development'
            qcFinalizing = 'Initiate Verification'
            readyForQc = 'Ready for Verification'
        }
        stateAfterPrependByTrigger = @{ finalQcComplete = 'Ready for Verification' }
    }
}
$finalSettings = Get-QCWorkflowSettings -Config $finalCfg
$finalDoc = [pscustomobject]@{
    WorkflowName = 'TYPSA QC'
    StateName = 'Initiate Verification'
    DocumentGUID = 'stem-guid'
    Name = '080J082001ab001.pdf'
    FolderPath = 'documents\caltrans\cafwy2200-i-15_elpse\cadd\sheets\seg_1'
}
$finalCtx = @{
    config = $finalCfg
    prependTrigger = 'finalQcComplete'
    qcProcessType = 'review'
    expectedLanePdfName = '080J082001ab001-rev.pdf'
    job = @{ id = 'j-final'; sourceFolder = $finalDoc.FolderPath; sourceName = '080J082001ab001.pdf' }
    document = $finalDoc
}
$finalResult = Set-PWQCWorkflowState -Settings $finalSettings -Context $finalCtx -StateName 'Ready for Verification' -DryRun:$false
Assert-True $finalResult.IsSuccess 'finalQcComplete state write should succeed without empty referenceState throw'
Assert-False ($finalResult.Data.lanePostPrependSplitError) 'finalQcComplete lane split should not fail on empty referenceState'
Assert-True ($null -ne $finalResult.Data.lanePostPrependSplit) 'finalQcComplete lane split should run'
Assert-True $finalCtx.laneIndependentInitialPrepend 'finalQcComplete marks lane-independent prepend'
Assert-Eq $finalCtx.laneTargetState 'Ready for Verification' 'finalQcComplete lane target is post-prepend state'
Assert-False $finalCtx.writeStemPdfReferenceState 'finalQcComplete does not write stem reference state'
Assert-False ($finalCtx.ContainsKey('referenceState') -and -not [string]::IsNullOrWhiteSpace($finalCtx.referenceState)) 'finalQcComplete leaves referenceState unset'

# DGN is never written by automation (workflow state or QC_Process_Type)
InModuleScope -ModuleName PW.Discovery {
    function Write-QCJsonLog { param($Level, $Code, $Message, $Data) }
    Assert-True (_PWD-TestAutomationDgnProtectedDocument -DocumentName 'CA001.dgn') 'dgn extension is protected'
    Assert-False (_PWD-TestAutomationDgnProtectedDocument -DocumentName 'CA001.pdf') 'stem pdf is not protected'
    $blocked = _PWD-InvokeSetPwDocumentState -Document $null -StateName 'Initiate Origination' -GuardContext @{
        documentName = 'CA001.dgn'
        folderPath = 'Drawings\X'
        callSite = 'test'
    }
    Assert-True $blocked.skipped 'dgn state write is skipped'
    Assert-Eq $blocked.skipReason 'dgn_automation_write_blocked' 'dgn state skip reason'
    $names = Get-PWAssociatedSheetSyncDocumentNames -SheetStem 'CA001'
    Assert-False ($names -contains 'CA001.dgn') 'sync document names exclude dgn'
}

Write-Host 'test_qc_post_initial_prepend_states: OK' -ForegroundColor Green
