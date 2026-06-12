# Post-Initiate Origination prepend: lane PDF -> Originated, stem/DGN -> In Development;
# stem/DGN qc_process_type -> Production; lane QC PDFs keep lane type and are never overwritten.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.ProcessType.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.Workflow.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\PW.Discovery.psm1') -Force

function Assert-True($cond, $msg) { if (-not $cond) { throw "ASSERT FAILED: $msg" } }
function Assert-Eq($a, $b, $msg) { if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" } }
function Assert-False($cond, $msg) { if ($cond) { throw "ASSERT FAILED: $msg" } }

# initialQcPdf resets stem/DGN qc_process_type to Production (lane PDFs excluded)
Assert-False (Test-QCProcessTypeResetsAfterPrepend -ProcessType 'production' -Config @{}) 'production lane default ResetToProductionAfterPrepend is false'
Assert-True (Test-QCProcessTypeResetsAfterPrepend -ProcessType 'review' -Config @{}) 'review lane resets by default'

# Lane QC PDFs are never modified by associated review/process type sync
InModuleScope -ModuleName PW.Discovery {
    function Write-QCJsonLog { param($Level, $Code, $Message, $Data) }
    function Get-PWQcProcessTypeAttributeName { param($Config) return 'QC_Process_Type' }
    function Get-PWQcReviewTypeAttributeName { param($Config) return 'QC_Review_Type' }
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

    $script:attrWrites = @()
    Sync-PWAssociatedSheetReviewTypeAttributes -Config @{} -DocumentGuid '1' -DocumentName 'CA001.pdf' `
        -FolderPath 'Drawings\X' -CanonicalReviewType 'production'
    Assert-Eq $script:attrWrites.Count 2 'lane QC PDF excluded from process type sync'
}

# Sync-PWPostInitialPrependLaneStates uses ExpectedLanePdfName and rejects empty process type
InModuleScope -ModuleName PW.Discovery {
    function Write-QCJsonLog { param($Level, $Code, $Message, $Data) }
    function Get-PWSheetStemFromDocumentName { param($DocumentName) return 'CA001' }
    function Format-QCWorkflowStateName { param($StateName, $Config) return $StateName }
    function Format-QCProcessTypeAttributeValue { param($ProcessType) if ($ProcessType -eq 'review') { return 'Review' } return 'Production' }
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
    $script:reviewTypeSync = $null
    function Sync-PWAssociatedSheetReviewTypeAttributes {
        param($Config, $DocumentGuid, $DocumentName, $FolderPath, $CanonicalReviewType, $DryRun)
        $script:reviewTypeSync = @{
            documentName = $DocumentName
            canonicalReviewType = $CanonicalReviewType
        }
    }

    $split = Sync-PWPostInitialPrependLaneStates -Config @{} -FolderPath 'Drawings\X' `
        -DocumentName 'CA001.pdf' -DocumentGuid 'stem-guid' -QcProcessType 'review' `
        -ExpectedLanePdfName 'CA001-rev.pdf' -LaneTargetState 'Originated' -ReferenceState 'In Development'
    Assert-Eq $split.lanePdfName 'CA001-rev.pdf' 'uses ExpectedLanePdfName'
    Assert-Eq $split.laneTargetState 'Originated' 'lane target preserved'
    Assert-Eq $split.referenceState 'In Development' 'reference target preserved'
    Assert-Eq $script:reviewTypeSync.canonicalReviewType 'production' 'resets stem/DGN qc_process_type to production'
    Assert-True $split.laneProcessTypeEnsure.ensured 'sets lane qc_process_type when unset'
    Assert-Eq $split.laneProcessTypeEnsure.toValue 'Review' 'lane qc_process_type matches triggering review lane'

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
Assert-Eq $script:stateWrites[0] 'In Development' 'stem primary write should target In Development'
Assert-True $ctx.laneIndependentInitialPrepend 'context marks lane-independent prepend'

Write-Host 'test_qc_post_initial_prepend_states: OK' -ForegroundColor Green
