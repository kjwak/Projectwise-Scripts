# Verified lane workflow state writeback: clean doc reload, -Force, read-back verification.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Workflow\QC.Workflow.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\ProjectWise\PW.Discovery.psm1') -Force

function Assert-True($cond, $msg) { if (-not $cond) { throw "ASSERT FAILED: $msg" } }
function Assert-Eq($a, $b, $msg) { if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" } }
function Assert-False($cond, $msg) { if ($cond) { throw "ASSERT FAILED: $msg" } }

Assert-True (Test-QCWorkflowStateWritebackUseForce -Config @{}) 'useForce defaults true'
Assert-True (Test-QCWorkflowStateWritebackUseVerified -Config @{}) 'useVerified defaults true'
Assert-False (Test-QCWorkflowStateWritebackUseVerified -Config @{ qcWorkflow = @{ stateWriteback = @{ useVerified = $false } } }) 'useVerified false only when explicitly configured'

InModuleScope -ModuleName PW.Discovery {
    function Write-QCJsonLog { param($Level, $Code, $Message, $Data) $script:jsonLogs += ,@{ Code = $Code; Data = $Data } }
    function Format-QCWorkflowStateName { param($StateName, $Config) return $StateName }
    function Test-QCWorkflowStateWritebackUseForce { param($Config) return $true }
    function Get-PWDocumentWorkflowStateName { param($FolderPath, $DocumentName, $DocumentGuid) return $script:workflowStateByName[$DocumentName] }
    function _PWD-NormalizeSheetIndexValue { param($v) return ([string]$v).Trim().ToLowerInvariant() }

    $script:cleanReloadCount = 0
    $script:stateWriteParams = @()
    $script:workflowStateByName = @{
        'lane-chk.pdf' = 'In Development'
    }
    $script:jsonLogs = @()

    function _PWD-GetCleanPwDocumentForStateChange {
        param($FolderPath, $DocumentName, $DocumentGuid)
        $script:cleanReloadCount++
        return [pscustomobject]@{
            DocumentGUID = 'lane-guid'
            Name = $DocumentName
            WorkflowState = $script:workflowStateByName[$DocumentName]
        }
    }

    function Set-PWDocumentState {
        param($InputDocuments, $State, [switch]$Force, [switch]$ReturnBoolean)
        $script:stateWriteParams += @{
            state = [string]$State
            force = [bool]$Force
            docName = [string]$InputDocuments[0].Name
        }
        if ($Force) {
            $script:workflowStateByName[$InputDocuments[0].Name] = [string]$State
        }
        if ($ReturnBoolean) { return $true }
    }

    $result = Set-PWDocumentWorkflowStateVerified -Config @{} -FolderPath 'Drawings\X' `
        -DocumentName 'lane-chk.pdf' -TargetState 'Originated' -QcProcessType 'check'
    Assert-Eq $script:cleanReloadCount 2 'reloads clean document before write and before read-back'
    Assert-Eq $script:stateWriteParams.Count 1 'single Set-PWDocumentState call'
    Assert-True $script:stateWriteParams[0].force 'lane write uses -Force'
    Assert-Eq $script:stateWriteParams[0].state 'Originated' 'writes target state'
    Assert-True $result.verified 'force write with matching read-back is verified'
    Assert-Eq $result.readBackState 'Originated' 'read-back matches target'
    Assert-True $result.usedForce 'result reports usedForce true'
    Assert-True ($result.commandShape -match '-Force$') 'commandShape includes -Force'

    $script:cleanReloadCount = 0
    $script:stateWriteParams = @()
    $script:workflowStateByName['lane-chk.pdf'] = 'In Development'
    function Test-QCWorkflowStateWritebackUseForce { param($Config) return $false }
    $noForce = Set-PWDocumentWorkflowStateVerified -Config @{} -FolderPath 'Drawings\X' `
        -DocumentName 'lane-chk.pdf' -TargetState 'Originated' -QcProcessType 'check'
    Assert-False $noForce.verified 'non-force write with read-back mismatch is unverified'
    Assert-Eq $noForce.readBackState 'In Development' 'read-back unchanged without force'
    Assert-False $script:stateWriteParams[0].force 'explicit useForce false skips -Force switch'

    $unverifiedCodes = @($script:jsonLogs | Where-Object { $_.Code -eq 'QC_LANE_STATE_WRITE_UNVERIFIED' })
    Assert-True ($unverifiedCodes.Count -ge 1) 'unverified write emits QC_LANE_STATE_WRITE_UNVERIFIED'
}

InModuleScope -ModuleName PW.Discovery {
    function Write-QCJsonLog { param($Level, $Code, $Message, $Data) }
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
    function Get-PWQcPdfLaneFromDocumentName { param($DocumentName) if ($DocumentName -match '(?i)-rev\.pdf$') { return 'review' } elseif ($DocumentName -match '(?i)-chk\.pdf$') { return 'check' } elseif ($DocumentName -match '(?i)-prod\.pdf$') { return 'production' } return $null }
    function Update-SheetPackageQcPdfLaneState { param($Config, $DocumentGuid, $CurrentPwState, $QcProcessType) return $true }
    function _PWD-SyncReferenceSheetProcessTypeAttributes { param($Config, $DocumentGuid, $DocumentName, $FolderPath, $CanonicalProcessType, $WatchRoot, $LastAuditEventAt, $DryRun, $ControlDocumentOnly) $script:processResetCalled = $true; return @{ updates = @() } }
    function _PWD-ResolvePwDocumentInFolder { param($DocByGuid, $FolderPath, $DocumentName, $DocumentGuid) return [pscustomobject]@{ DocumentGUID = ('guid-' + $DocumentName); Name = $DocumentName } }

    $script:stateTargets = @()
    function Set-PWDocumentWorkflowStateVerified {
        param($Config, $FolderPath, $DocumentName, $DocumentGuid, $TargetState, $QcProcessType, $DryRun, $WorkflowName, $IsLaneAuthority)
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
            isLaneAuthority = [bool]$IsLaneAuthority
            commandShape = "Set-PWDocumentState -InputDocuments @(`$cleanDoc) -State '$TargetState' -Force"
        }
    }

    foreach ($lane in @(
        @{ lane = 'production'; pdf = '080J082001ab001-prod.pdf' }
        @{ lane = 'check'; pdf = '080J082001ab001-chk.pdf' }
        @{ lane = 'review'; pdf = '080J082001ab001-rev.pdf' }
    )) {
        $script:stateTargets = @()
        $script:processResetCalled = $false
        $split = Sync-PWPostInitialPrependLaneStates -Config @{} -FolderPath 'Drawings\X' `
            -DocumentName '080J082001ab001.pdf' -DocumentGuid 'stem-guid' -QcProcessType $lane.lane `
            -ExpectedLanePdfName $lane.pdf -LaneTargetState 'Originated' -ReferenceState 'In Development'
        $uniqueTargets = @($script:stateTargets | Select-Object -Unique)
        Assert-Eq $uniqueTargets.Count 2 "$($lane.lane) prepend writes lane and stem PDF states"
        Assert-Eq $uniqueTargets[0] $lane.pdf "$($lane.lane) prepend targets expected lane PDF"
        Assert-True ($uniqueTargets -contains '080J082001ab001.pdf') "$($lane.lane) prepend returns stem PDF to reference state"
        Assert-True $split.writeStemPdfReferenceState "$($lane.lane) prepend enables stem reference writeback"
    }
}

Write-Host 'test_qc_lane_state_write_verified: OK' -ForegroundColor Green
