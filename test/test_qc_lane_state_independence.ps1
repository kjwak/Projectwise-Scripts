# Lane-independent workflow state: sibling sync gated off by default.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.ProcessType.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Database.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path $repoRoot 'modules\PW.Discovery.psm1') -Force

function Assert-True($cond, $msg) { if (-not $cond) { throw "ASSERT FAILED: $msg" } }
function Assert-Eq($a, $b, $msg) { if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" } }
function Assert-False($cond, $msg) { if ($cond) { throw "ASSERT FAILED: $msg" } }

# Legacy sibling sync disabled by default
Assert-False (Test-QCLegacySiblingStateSyncEnabled -Config @{}) 'legacy sibling sync default off'
Assert-False (Test-QCProcessTypeSyncsWithSiblingSheets -ProcessType 'production' -Config @{}) 'production does not sync siblings by default'
Assert-False (Test-QCProcessTypeSyncsWithSiblingSheets -ProcessType 'check' -Config @{}) 'check does not sync siblings'
Assert-False (Test-QCProcessTypeSyncsWithSiblingSheets -ProcessType 'review' -Config @{}) 'review does not sync siblings'

$legacyCfg = @{ QCProcess = @{ EnableLegacySiblingStateSync = $true; ProcessTypes = @{ production = @{ SyncWithSiblingSheets = $true } } } }
Assert-True (Test-QCLegacySiblingStateSyncEnabled -Config $legacyCfg) 'legacy flag enables gate'
Assert-True (Test-QCProcessTypeSyncsWithSiblingSheets -ProcessType 'production' -Config $legacyCfg) 'production syncs when legacy enabled'
Assert-False (Test-QCProcessTypeSyncsWithSiblingSheets -ProcessType 'production' -Config @{ QCProcess = @{ EnableLegacySiblingStateSync = $true; ProcessTypes = @{ production = @{ SyncWithSiblingSheets = $false } } } }) 'lane SyncWithSiblingSheets=false blocks even when legacy flag on'

# Post-prepend sibling sync returns early when legacy disabled
InModuleScope -ModuleName PW.Discovery {
    function _PWD-TestLegacySiblingStateSyncEnabled { param($Config) return $false }
    function Write-QCJsonLog { param($Level, $Code, $Message, $Data) }
    function Update-SheetPackageQcPdfLaneState { param($Config, $DocumentGuid, $CurrentPwState, $QcProcessType) return $true }

    $result = Sync-PWAssociatedSheetMembersToWorkflowState -Config @{} -FolderPath 'Documents\X' `
        -DocumentName 'sheet-prod.pdf' -DocumentGuid '11111111-1111-1111-1111-111111111111' `
        -TargetStateName 'Ready for QC' -DryRun:$false
    Assert-True $result.siblingSyncDisabled 'sibling sync skipped flag set'
    Assert-Eq $result.memberCount 0 'no members updated when legacy sync disabled'
}

# Lane-independent audit members: only trigger document
$allMembers = @(
    @{ documentGuid = '11111111-1111-1111-1111-111111111111'; documentName = 'sheet-prod.pdf' }
    @{ documentGuid = '22222222-2222-2222-2222-222222222222'; documentName = 'sheet-chk.pdf' }
    @{ documentGuid = '33333333-3333-3333-3333-333333333333'; documentName = 'sheet.pdf' }
    @{ documentGuid = '44444444-4444-4444-4444-444444444444'; documentName = 'sheet.dgn' }
)
$laneOnly = @(_PWD-GetLaneIndependentAuditMembers -AllMembers $allMembers `
    -TriggerDocumentGuid '22222222-2222-2222-2222-222222222222' -TriggerDocumentName 'sheet-chk.pdf')
Assert-Eq $laneOnly.Count 1 'only trigger member returned'
Assert-Eq ([string]$laneOnly[0].documentName) 'sheet-chk.pdf' 'trigger is check lane PDF'

Assert-True (_PWD-TestAutomationStemPdfStateWriteBlocked -DocumentName 'sheet.pdf' -TargetState 'Redlines Received') `
    'stem PDF blocks lane workflow states'
Assert-False (_PWD-TestAutomationStemPdfStateWriteBlocked -DocumentName 'sheet.pdf' -TargetState 'In Development' -WriteScope 'stem') `
    'post-prepend stem writeback allowed'
Assert-False (_PWD-TestAutomationStemPdfStateWriteBlocked -DocumentName 'sheet-chk.pdf' -TargetState 'Redlines Received') `
    'lane PDF is not stem-protected'

InModuleScope -ModuleName PW.Discovery {
    function Get-PWAssociatedSheetMembers {
        param($Config, $FolderPath, $DocumentName, $DocumentGuid)
        return @(
            @{ documentGuid = '22222222-2222-2222-2222-222222222222'; documentName = 'sheet-chk.pdf' }
            @{ documentGuid = '33333333-3333-3333-3333-333333333333'; documentName = 'sheet.pdf' }
        )
    }
    function _PWD-TestLegacySiblingStateSyncEnabled { param($Config) return $true }
    function _PWD-LogLaneStateIndependentTelemetry { param($AllMembers, $FolderPath, $TriggerSource, $TriggerDocumentName, $TriggerDocumentGuid) }

    $syncMembers = @(Get-PWAssociatedSheetSyncMembers -Config @{} -FolderPath 'Documents\X' `
        -DocumentName 'sheet-chk.pdf' -DocumentGuid '22222222-2222-2222-2222-222222222222')
    Assert-Eq $syncMembers.Count 0 'lane trigger never participates in legacy sibling sync'
}

InModuleScope -ModuleName PW.Discovery {
    $script:stateWrites = @()
    function Get-PWDocumentWorkflowStateName { param($FolderPath, $DocumentName, $DocumentGuid) return 'In Development' }
    function _PWD-GetSheetIndexPwStateName { param($Config, $DocumentGuid) return 'In Development' }
    function _PWD-GetSheetIndexStateSnapshot { param($Config, $DocumentGuid) return @{ pwStateName = 'In Development'; lastAuditEventAt = $null } }
    function _PWD-TestShouldBlockStaleRestartOverwrite { return $false }
    function Test-QCDatabaseEnabled { return $false }
    function _PWD-WriteDocumentStateLiveVerificationLog { }
    function Test-QCDocumentStateAuditEventIsStale { return @{ isStale = $false } }
    function Invoke-QCWorkflowStateEmailAttributeGate { return @{ blocked = $false } }
    function _PWD-InvokeSetPwDocumentState {
        param($Document, $StateName, $GuardContext)
        $script:stateWrites += [pscustomobject]@{
            documentName = [string]$GuardContext.documentName
            stateName = [string]$StateName
        }
        return @{ applied = $true; verified = $true; readBackState = $StateName }
    }
    function Get-PWAssociatedSheetMembers {
        return @(
            @{ documentGuid = '22222222-2222-2222-2222-222222222222'; documentName = 'sheet-chk.pdf'; document = $null }
            @{ documentGuid = '33333333-3333-3333-3333-333333333333'; documentName = 'sheet.pdf'; document = $null }
        )
    }
    function _PWD-TestLegacySiblingStateSyncEnabled { return $false }
    function _PWD-LogLaneStateIndependentTelemetry { }
    function Get-PWDocumentWorkflowStateMapByGuid { return @{} }
    function Invoke-QCSheetGroupWorkflowTransition { param($Config, $TriggerDocumentGuid, $TriggerDocumentName, $FolderPath, $SourceState, $TargetState, $TransitionSource, $Members, $StateByGuid, $PreviousStateByGuid, $AuditEventId, $ChangedByUser, $ChangedByUsername, $LastAuditEventAt, $DryRun, $AuditActionName) return @{ members = @() } }
    function _PWD-EnqueuePrependJobsFromAssociatedQcPdfState { }
    function Test-QCWorkflowStateIsQcInitiated { return $false }

    $script:stateWrites = @()
    Sync-PWAssociatedSheetWorkflowState -Config @{} -DocumentGuid '22222222-2222-2222-2222-222222222222' `
        -DocumentName 'sheet-chk.pdf' -FolderPath 'Documents\X' -LastAuditEventAt '2026-06-15 11:14:21'
    $stemWrites = @($script:stateWrites | Where-Object { $_.documentName -eq 'sheet.pdf' })
    Assert-Eq $stemWrites.Count 0 'chk DOCUMENT_STATE must not write stem PDF state'
}

# Resolve-SheetPackageFromDocument: lane PDFs are qc_pdf role with shared stem
InModuleScope -ModuleName Core.Database {
    $prod = Resolve-SheetPackageFromDocument -DocumentName 'CA001-prod.pdf' -FolderPath 'Documents\X'
    $chk = Resolve-SheetPackageFromDocument -DocumentName 'CA001-chk.pdf' -FolderPath 'Documents\X'
    $rev = Resolve-SheetPackageFromDocument -DocumentName 'CA001-rev.pdf' -FolderPath 'Documents\X'
    $clean = Resolve-SheetPackageFromDocument -DocumentName 'CA001.pdf' -FolderPath 'Documents\X'
    Assert-Eq $prod.documentRole 'qc_pdf' 'prod is qc_pdf'
    Assert-Eq $chk.documentRole 'qc_pdf' 'chk is qc_pdf'
    Assert-Eq $rev.documentRole 'qc_pdf' 'rev is qc_pdf'
    Assert-Eq $clean.documentRole 'sheet_pdf' 'clean pdf is sheet_pdf'
    Assert-Eq $prod.sheetStem 'CA001' 'prod stem'
    Assert-Eq $chk.sheetStem 'CA001' 'chk stem'
    Assert-Eq $rev.sheetStem 'CA001' 'rev stem'
}

Write-Host 'test_qc_lane_state_independence: OK' -ForegroundColor Green
