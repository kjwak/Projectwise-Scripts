# v1.19 lane workflow writeback: lane PDF is authority; reference docs untouched by default.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.ProcessType.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.Workflow.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.Notifications.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\PW.Discovery.psm1') -Force

function Assert-True($cond, $msg) { if (-not $cond) { throw "ASSERT FAILED: $msg" } }
function Assert-Eq($a, $b, $msg) { if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" } }
function Assert-False($cond, $msg) { if ($cond) { throw "ASSERT FAILED: $msg" } }

Assert-False (Test-QCResetProcessTypeAfterLanePrepend -Config @{}) 'ResetProcessTypeAfterLanePrepend defaults false'
Assert-False (Test-QCProcessTypeResetsAfterPrepend -ProcessType 'review' -Config @{}) 'review lane does not reset by default'
Assert-False (Test-QCProcessTypeResetsAfterPrepend -ProcessType 'check' -Config @{}) 'check lane does not reset by default'

InModuleScope -ModuleName PW.Discovery {
    function Write-QCJsonLog { param($Level, $Code, $Message, $Data) $script:jsonLogs += ,@{ Code = $Code; Data = $Data } }
    function Get-PWSheetStemFromDocumentName { param($DocumentName) return '080J082001ab001' }
    function Format-QCWorkflowStateName { param($StateName, $Config) return $StateName }
    function Format-QCProcessTypeAttributeValue { param($ProcessType) switch ($ProcessType) { 'review' { 'Review' } 'check' { 'Check' } default { 'Production' } } }
    function Get-PWQcProcessTypeAttributeName { param($Config) return 'QC_Process_Type' }
    function Test-PWQcReviewTypeAttributesEnabled { param($Config, $FolderPath) return $true }
    function Test-QCResetProcessTypeAfterLanePrepend { param($Config) return $false }
    function Get-QCWorkflowSettings { param($Config) return @{ expectedWorkflowName = 'TYPSA QC' } }
    function Get-PWDocumentAttributesByColumns { param($FolderPath, $DocumentName, $ColumnsToReturn) return @{ found = $true; attributes = @{} } }
    function _PWD-InvokeUpdatePWDocumentAttributes { param($Document, $Attributes, $Config) return $true }
    function Update-QCSheetIndexPwStateName { param($Config, $DocumentGuid, $PwStateName) return $true }
    function Get-PWQcPdfLaneFromDocumentName { param($DocumentName) if ($DocumentName -match '(?i)-rev\.pdf$') { return 'review' } elseif ($DocumentName -match '(?i)-chk\.pdf$') { return 'check' } return $null }
    function Update-SheetPackageQcPdfLaneState { param($Config, $DocumentGuid, $CurrentPwState, $QcProcessType) return $true }
    function _PWD-SyncReferenceSheetProcessTypeAttributes { param($Config, $DocumentGuid, $DocumentName, $FolderPath, $CanonicalProcessType, $WatchRoot, $LastAuditEventAt, $DryRun, $ControlDocumentOnly) $script:processResetCalled = $true; return @{ updates = @() } }
    function _PWD-ResolvePwDocumentInFolder { param($DocByGuid, $FolderPath, $DocumentName, $DocumentGuid) return [pscustomobject]@{ DocumentGUID = ('guid-' + $DocumentName); Name = $DocumentName } }
    function Test-QCWorkflowStateWritebackUseForce { param($Config) return $true }
    function Set-PWDocumentWorkflowStateVerified {
        param($Config, $FolderPath, $DocumentName, $DocumentGuid, $TargetState, $QcProcessType, $DryRun, $WorkflowName)
        $script:stateTargets += [string]$DocumentName
        return @{
            documentGuid = 'guid-' + $DocumentName
            documentName = $DocumentName
            fromState = 'In Development'
            targetState = $TargetState
            readBackState = $TargetState
            applied = $true
            verified = $true
            usedForce = $true
            commandShape = "Set-PWDocumentState -InputDocuments @(`$cleanDoc) -State '$TargetState' -Force"
        }
    }

    foreach ($lane in @('review', 'check')) {
        $script:stateTargets = @()
        $script:processResetCalled = $false
        $expectedPdf = if ($lane -eq 'review') { '080J082001ab001-rev.pdf' } else { '080J082001ab001-chk.pdf' }
        $split = Sync-PWPostInitialPrependLaneStates -Config @{} -FolderPath 'Drawings\X' `
            -DocumentName '080J082001ab001.pdf' -DocumentGuid 'stem-guid' -QcProcessType $lane `
            -ExpectedLanePdfName $expectedPdf -LaneTargetState 'Originated' -ReferenceState 'In Development'
        Assert-Eq $split.lanePdfName $expectedPdf "$lane prepend targets expected lane PDF"
        Assert-False $split.writeReferenceStates "$lane prepend does not write reference states"
        Assert-False $script:processResetCalled "$lane prepend does not reset source/DGN process type by default"
        $uniqueTargets = @($script:stateTargets | Select-Object -Unique)
        Assert-Eq $uniqueTargets.Count 1 "$lane prepend writes lane PDF state only"
        Assert-Eq $uniqueTargets[0] $expectedPdf "$lane prepend targets expected lane PDF"
    }
}

InModuleScope -ModuleName PW.Discovery {
    function Write-QCJsonLog { param($Level, $Code, $Message, $Data) }
    function Get-PWSheetStemFromDocumentName { param($DocumentName) return 'CA001' }
    function Format-QCWorkflowStateName { param($StateName, $Config) return $StateName }
    function Get-QCWorkflowSettings { param($Config) return @{ expectedWorkflowName = 'TYPSA QC' } }
    function Test-QCResetProcessTypeAfterLanePrepend { param($Config) return $false }
    function Update-QCSheetIndexPwStateName { param($Config, $DocumentGuid, $PwStateName) return $true }
    function Get-PWQcPdfLaneFromDocumentName { param($DocumentName) return 'review' }
    function Update-SheetPackageQcPdfLaneState { param($Config, $DocumentGuid, $CurrentPwState, $QcProcessType) return $true }
    function Get-PWQcProcessTypeAttributeName { param($Config) return 'QC_Process_Type' }
    function Test-PWQcReviewTypeAttributesEnabled { param($Config, $FolderPath) return $true }
    function Get-PWDocumentAttributesByColumns { param($FolderPath, $DocumentName, $ColumnsToReturn) return @{ found = $true; attributes = @{} } }
    function _PWD-InvokeUpdatePWDocumentAttributes { param($Document, $Attributes, $Config) return $true }
    function Format-QCProcessTypeAttributeValue { param($ProcessType) return 'Review' }
    function _PWD-SyncReferenceSheetProcessTypeAttributes { param($Config, $DocumentGuid, $DocumentName, $FolderPath, $CanonicalProcessType, $WatchRoot, $LastAuditEventAt, $DryRun, $ControlDocumentOnly) }
    function _PWD-ResolvePwDocumentInFolder { param($DocByGuid, $FolderPath, $DocumentName, $DocumentGuid) return [pscustomobject]@{ DocumentGUID = 'lane-guid'; Name = $DocumentName } }
    function Test-QCWorkflowStateWritebackUseForce { param($Config) return $true }
    function Set-PWDocumentWorkflowStateVerified {
        param($Config, $FolderPath, $DocumentName, $DocumentGuid, $TargetState, $QcProcessType, $DryRun, $WorkflowName)
        return @{
            documentGuid = 'lane-guid'
            documentName = $DocumentName
            fromState = 'In Development'
            targetState = $TargetState
            readBackState = 'In Development'
            applied = $true
            verified = $false
            usedForce = $true
            commandShape = "Set-PWDocumentState -InputDocuments @(`$cleanDoc) -State '$TargetState' -Force"
        }
    }

    $split = Sync-PWPostInitialPrependLaneStates -Config @{} -FolderPath 'Drawings\X' `
        -DocumentName 'CA001.pdf' -QcProcessType 'review' -ExpectedLanePdfName 'CA001-rev.pdf' `
        -LaneTargetState 'Originated' -ReferenceState 'In Development'
    Assert-False $split.laneStateVerified 'unverified lane write marks laneStateVerified false when read-back mismatches'
    Assert-False $split.allVerified 'allVerified false when lane state unverified'
}

InModuleScope -ModuleName QC.Notifications {
    $target = _QCN-ResolveQcPdfNotificationTarget -Document $null -Config @{} `
        -Job @{ metadata = @{ qcProcessType = 'review'; expectedLanePdfName = '080J082001ab001-rev.pdf' } } `
        -DocumentName '080J082001ab001.pdf' -DocumentGuid 'stem' -DocumentPath 'Drawings\X\080J082001ab001.pdf'
    if ($target.documentName -ne '080J082001ab001-rev.pdf') {
        throw "ASSERT FAILED: notification uses expected lane PDF name (got '$($target.documentName)')"
    }
    $chk = _QCN-NormalizeQcPdfDocumentName -DocumentName '080J082001ab001.pdf' -SheetStem '080J082001ab001' -ProcessType 'check' -Config @{}
    if ($chk -ne '080J082001ab001-chk.pdf') { throw "ASSERT FAILED: check notification name (got '$chk')" }
    $legacy = _QCN-NormalizeQcPdfDocumentName -DocumentName '080J082001ab001-qc.pdf' -SheetStem '080J082001ab001' -ProcessType 'review' -Config @{}
    if ($legacy -ne '080J082001ab001-rev.pdf') { throw "ASSERT FAILED: legacy -qc.pdf normalized to review lane (got '$legacy')" }
}

Write-Host 'test_qc_lane_workflow_writeback: OK' -ForegroundColor Green
