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
