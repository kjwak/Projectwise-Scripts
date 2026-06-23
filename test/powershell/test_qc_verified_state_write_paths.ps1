# Verified workflow state write routing: default production paths require read-back verification.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Workflow\QC.Workflow.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\ProjectWise\PW.Discovery.psm1') -Force

function Assert-True($cond, $msg) { if (-not $cond) { throw "ASSERT FAILED: $msg" } }
function Assert-Eq($a, $b, $msg) { if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" } }
function Assert-False($cond, $msg) { if ($cond) { throw "ASSERT FAILED: $msg" } }

Assert-True (Test-QCWorkflowStateWritebackUseVerified -Config @{}) 'useVerified defaults true'

InModuleScope -ModuleName PW.Discovery {
    function Write-QCJsonLog { param($Level, $Code, $Message, $Data) $script:jsonLogs += ,@{ Code = $Code; Data = $Data } }
    function Format-QCWorkflowStateName { param($StateName, $Config) return $StateName }
    function Test-QCWorkflowStateWritebackUseForce { param($Config) return $true }
    function Test-QCWorkflowStateWritebackUseVerified { param($Config) return $true }
    function _PWD-NormalizeSheetIndexValue { param($v) return ([string]$v).Trim().ToLowerInvariant() }
    function Get-PWDocumentWorkflowStateName { param($FolderPath, $DocumentName, $DocumentGuid) return $script:workflowStateByName[$DocumentName] }

    $script:jsonLogs = @()
    $script:workflowStateByName = @{ 'member.pdf' = 'In Development' }
    $script:stateWriteParams = @()

    function _PWD-GetCleanPwDocumentForStateChange {
        param($FolderPath, $DocumentName, $DocumentGuid)
        return [pscustomobject]@{
            DocumentGUID = 'member-guid'
            Name = $DocumentName
            WorkflowState = $script:workflowStateByName[$DocumentName]
        }
    }

    function Set-PWDocumentState {
        param($InputDocuments, $State, [switch]$Force, [switch]$ReturnBoolean)
        $script:stateWriteParams += @{ state = [string]$State; force = [bool]$Force }
        if ($Force) { $script:workflowStateByName[$InputDocuments[0].Name] = [string]$State }
        if ($ReturnBoolean) { return $true }
    }

    $result = _PWD-InvokeSetPwDocumentState -Document $null -StateName 'Originated' -GuardContext @{
        callSite = 'test.verified.default'
        config = @{}
        folderPath = 'Drawings\X'
        documentName = 'member.pdf'
        documentGuid = 'member-guid'
        writeScope = 'workflow'
    }
    Assert-True $result.verified 'default wrapper uses verified helper'
    Assert-True $script:stateWriteParams[0].force 'verified routing uses -Force'
    Assert-Eq $script:stateWriteParams[0].state 'Originated' 'verified routing writes target state'

    $legacyLogs = @($script:jsonLogs | Where-Object { $_.Code -eq 'QC_LEGACY_UNVERIFIED_STATE_WRITE' })
    Assert-Eq $legacyLogs.Count 0 'default path does not emit legacy telemetry'

    $script:jsonLogs = @()
    $script:stateWriteParams = @()
    $script:workflowStateByName['member.pdf'] = 'In Development'
    function Test-QCWorkflowStateWritebackUseVerified { param($Config) return $false }

    $legacy = _PWD-InvokeSetPwDocumentState -Document $null -StateName 'Originated' -GuardContext @{
        callSite = 'test.legacy.config'
        config = @{ qcWorkflow = @{ stateWriteback = @{ useVerified = $false } } }
        folderPath = 'Drawings\X'
        documentName = 'member.pdf'
        documentGuid = 'member-guid'
        writeScope = 'workflow'
    }
    Assert-False $legacy.verified 'legacy config returns verified=false'
    Assert-True $legacy.legacyUnverified 'legacy config marks legacyUnverified'
    $legacyLogs = @($script:jsonLogs | Where-Object { $_.Code -eq 'QC_LEGACY_UNVERIFIED_STATE_WRITE' })
    Assert-True ($legacyLogs.Count -ge 1) 'legacy path emits QC_LEGACY_UNVERIFIED_STATE_WRITE'

    $script:workflowStateByName['member.pdf'] = 'In Development'
    function Test-QCWorkflowStateWritebackUseVerified { param($Config) return $true }
    function Set-PWDocumentState {
        param($InputDocuments, $State, [switch]$Force, [switch]$ReturnBoolean)
        if (-not $Force) { return $true }
        if ($ReturnBoolean) { return $true }
    }
    $threw = $false
    try {
        [void](_PWD-InvokeSetPwDocumentState -Document $null -StateName 'Originated' -GuardContext @{
            callSite = 'test.unverified.throw'
            config = @{}
            folderPath = 'Drawings\X'
            documentName = 'member.pdf'
            documentGuid = 'member-guid'
            writeScope = 'workflow'
        })
    } catch {
        $threw = $true
    }
    Assert-True $threw 'verified default throws when read-back mismatches without -Force'
}

InModuleScope -ModuleName QC.Workflow {
    function Get-PWWorkflows { param() [pscustomobject]@{ Name = 'TYPSA QC' } }
    function Get-PWWorkflowStateLinks { param() [pscustomobject]@{ FromStateName = 'In Development'; ToStateName = 'Originated' } }
    function Set-PWDocumentWorkflowStateVerified {
        param($Config, $FolderPath, $DocumentName, $DocumentGuid, $TargetState, $QcProcessType, $DryRun, $WorkflowName, $IsLaneAuthority, $WriteScope)
        return @{
            applied = $true
            verified = $true
            readBackState = $TargetState
            targetState = $TargetState
        }
    }

    $cfg = @{
        qcWorkflow = @{
            enabled = $true
            strictMode = $false
            mode = 'StateAndAttributes'
            autoSetState = $true
            expectedWorkflowName = 'TYPSA QC'
            stateWriteback = @{ useForce = $true; useVerified = $true }
            states = @{ production = 'In Development'; readyForQc = 'Originated' }
        }
    }
    $settings = Get-QCWorkflowSettings -Config $cfg
    $ctx = @{
        config = $cfg
        job = @{ sourceFolder = 'Drawings\X'; sourceName = 'sheet.pdf' }
        document = [pscustomobject]@{ WorkflowName = 'TYPSA QC'; StateName = 'In Development'; Name = 'sheet.pdf'; DocumentGUID = 'g1' }
    }
    $result = Set-PWQCWorkflowState -Settings $settings -Context $ctx -StateName 'Originated' -DryRun:$false
    Assert-True $result.IsSuccess 'Set-PWQCWorkflowState uses verified helper when useVerified=true'
    Assert-True $result.Data.stateWriteVerified 'primary write records verified=true'
}

InModuleScope -ModuleName PW.Discovery {
    function Write-QCJsonLog { param($Level, $Code, $Message, $Data) }
    function Format-QCWorkflowStateName { param($StateName, $Config) return $StateName }
    function Test-QCWorkflowStateWritebackUseForce { param($Config) return $true }
    function Test-QCWorkflowStateWritebackUseVerified { param($Config) return $true }
    function Get-PWAssociatedSheetSyncMembers {
        param($Config, $FolderPath, $DocumentName, $DocumentGuid, $TriggerSource)
        return @(@{ documentGuid = '1'; documentName = 'CA001.pdf'; document = $null })
    }
    function Get-PWDocumentWorkflowStateMapByGuid { param($DocumentGuids) return @{ '1' = 'In Development' } }
    function _PWD-GetWorkflowStateFromDocumentRow { param($DocRow) return '' }
    function _PWD-TestStateNameNotEmpty { param($StateName) return $true }
    function Update-QCSheetIndexPwStateName { param($Config, $DocumentGuid, $PwStateName) return $true }
    function Test-QCDatabaseEnabled { param($Config) return $false }

    function _PWD-InvokeSetPwDocumentState {
        param($Document, $StateName, $GuardContext)
        throw 'Workflow state write unverified for member.pdf'
    }

    $threw = $false
    $result = $null
    try {
        $result = Sync-PWAssociatedSheetMembersToWorkflowState -Config @{ QCProcess = @{ EnableLegacySiblingStateSync = $true } } `
            -FolderPath 'Drawings\X' -DocumentName 'CA001.pdf' -DocumentGuid '1' -TargetStateName 'Originated'
    } catch {
        $threw = $true
    }
    Assert-True ($threw -or ($result -and $result.updates -and $result.updates[0].error)) 'legacy sibling sync does not treat unverified write as success'
    if (-not $threw -and $result) {
        Assert-False $result.updates[0].verified 'unverified sibling update remains unverified'
        Assert-False $result.updates[0].applied 'unverified sibling update is not marked applied'
    }
}

Write-Host 'test_qc_verified_state_write_paths: OK' -ForegroundColor Green
