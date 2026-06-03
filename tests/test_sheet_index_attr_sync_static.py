"""Static checks for DOCUMENT_ATTR sheet_index sync from ProjectWise."""
import json
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
DISCOVERY = REPO_ROOT / "modules" / "PW.Discovery.psm1"
DATABASE = REPO_ROOT / "modules" / "Core.Database.psm1"
WATCHER = REPO_ROOT / "scripts" / "Watch-QCTrigger.ps1"
TELEMETRY_DOC = REPO_ROOT / "docs" / "database-telemetry.md"


def test_discovery_exposes_sheet_index_sync_helpers():
    text = DISCOVERY.read_text(encoding="utf-8")
    assert "function Get-PWSheetIndexSyncColumnNames" in text
    assert "function ConvertTo-SheetIndexFieldValues" in text
    assert "function Get-PWDocumentAttributesByColumns" in text
    assert "forceAttrSync   = $isDocumentAttr" in text
    assert "QC_Review_Type" in text or "reviewType" in text


def test_sync_always_refreshes_on_document_attr():
    text = DISCOVERY.read_text(encoding="utf-8")
    assert "$isDocumentAttr" in text and "DOCUMENT_ATTR" in text
    assert "Get-PWSheetIndexSyncColumnNames -Config $Config" in text
    assert "ConvertTo-SheetIndexFieldValues -Config $Config" in text
    assert "-not $isDocumentAttr -and -not $emailsDiffer" in text


def test_database_schema_includes_qc_attribute_columns():
    text = DATABASE.read_text(encoding="utf-8")
    assert "checker_email" in text
    assert "qc_review_type" in text
    assert "qc_assigned_to" in text
    assert "_QDB-GetSchemaV1dot5" in text
    assert "targetVersion = '1.5.0'" in text


def test_write_sheet_index_merges_checker_and_qc_fields():
    text = DATABASE.read_text(encoding="utf-8")
    assert "[string]$CheckerEmail" in text
    assert "checker_email = CASE WHEN @setOwnership = 1 THEN COALESCE(@checkerEmail, tgt.checker_email)" in text
    assert "qc_review_type = CASE WHEN @setOwnership = 1 THEN COALESCE(@qcReviewType, tgt.qc_review_type)" in text


def test_document_attr_propagates_qc_review_type_to_associated_sheet():
    text = DISCOVERY.read_text(encoding="utf-8")
    assert "Sync-PWAssociatedSheetReviewTypeAttributes -Config $Config" in text
    assert "$reviewTypeDiffer" in text
    assert "WATCH_SHEET_REVIEW_TYPE_SYNC" in text
    assert "-not $isDocumentAttr" in text
    assert "CustomAttributes" in text
    assert "_PWD-ResolveSheetIndexQcReviewType" in text
    assert "CanonicalReviewType $pwReviewType" in text
    assert "_PWD-NormalizePwEnvironmentForQcReviewType" in text
    assert "pwWritesEnabled" in text


def test_watcher_audit_loop_calls_sync_on_document_attr():
    text = WATCHER.read_text(encoding="utf-8")
    assert "Sync-PWSheetIndexOwnership -Config $config" in text
    assert "'DOCUMENT_ATTR', 'DOCUMENT_STATE'" in text


def test_reconciliation_scan_uses_full_attribute_batch_builder():
    text = DISCOVERY.read_text(encoding="utf-8")
    assert "function Build-PWSheetIndexRowsForPairedSheets" in text
    assert "Get-PWSheetIndexSyncColumnNames -Config $Config" in text

    watch = WATCHER.read_text(encoding="utf-8")
    assert "Build-PWSheetIndexRowsForPairedSheets -Config $config" in watch
    assert "reconciliationCycle = [bool]$isReconciliationCycle" in watch

    db = DATABASE.read_text(encoding="utf-8")
    assert "checker_email NVARCHAR(200)" in db
    assert "sheetParamsPerRow = 12" in db


def test_database_schema_upgrade_applies_additive_patches_before_version_check():
    text = DATABASE.read_text(encoding="utf-8")
    assert "function _QDB-InvokeSchemaSqlBatches" in text
    assert "_QDB-GetSchemaV1dot5Additive" in text
    assert "DB_SCHEMA_UPGRADED" in text
    assert "IF NOT EXISTS (SELECT 1 FROM schema_version WHERE version = @version)" in text


def test_watcher_initializes_database_schema_on_startup():
    text = WATCHER.read_text(encoding="utf-8")
    assert "Initialize-QCDatabaseSchema -Config $config" in text
    assert "WATCH_DB_SCHEMA_READY" in text


def test_document_state_propagates_to_associated_sheet_files():
    discovery = DISCOVERY.read_text(encoding="utf-8")
    assert "function Sync-PWAssociatedSheetWorkflowState" in discovery
    assert "function Sync-PWAssociatedSheetReviewTypeAttributes" in discovery
    assert "function Get-PWAssociatedSheetMembers" in discovery
    assert "function Get-PWSheetStemFromDocumentName" in discovery
    assert "_PWD-InvokeSetPwDocumentState" in discovery

    watch = WATCHER.read_text(encoding="utf-8")
    assert "Sync-PWAssociatedSheetWorkflowState -Config $config" in watch
    assert "WATCH_SHEET_STATE_SYNC" in discovery


def test_qc_prepend_defaults_review_type_on_associated_sheet():
    text = DISCOVERY.read_text(encoding="utf-8")
    assert "function Get-PWQcDefaultReviewType" in text
    assert "function Ensure-PWQcReviewTypeOnAssociatedSheet" in text
    assert "function Test-PWQcReviewTypeAttributesEnabled" in text
    assert "function Get-PWQcReviewTypeEnabledEnvironments" in text
    assert "Get-PWAssociatedSheetMembers" in text
    assert "Get-PWQcDefaultReviewType -Config $Config" in text
    assert "Test-PWQcReviewTypeAttributesEnabled -Config $Config -FolderPath $FolderPath" in text

    legacy = (REPO_ROOT / "legacy" / "prepend_qc.ps1").read_text(encoding="utf-8")
    assert "Invoke-PrependQcReviewTypeDefaultIfNeeded" in legacy
    assert "Ensure-PWQcReviewTypeOnAssociatedSheet" in legacy

    appsettings = json.loads((REPO_ROOT / "appsettings.json").read_text(encoding="utf-8-sig"))
    enabled = appsettings["projectWise"]["qcReviewTypeAttributes"]["enabledEnvironments"]
    assert "Caltrans" in enabled


def test_qc_prepend_audit_checks_description_on_paired_pdf_actions():
    discovery = DISCOVERY.read_text(encoding="utf-8")
    assert "function Test-PWSheetPdfHasMatchingPair" in discovery

    watch = WATCHER.read_text(encoding="utf-8")
    assert "Get-QCPrependAuditActions -Config $config" in watch
    assert "Test-PWSheetPdfHasMatchingPair -FolderPath $fp" in watch
    assert "qcPrependAuditActions -notcontains $actionName" in watch
    assert "WATCH_AUDIT_STATUSSET_SKIP_CURRENT" in watch
    assert "Audit PDF skipped (action not configured for QC_Archivist description check)" in watch

    orch = (REPO_ROOT / "modules" / "QC.WatcherOrchestration.psm1").read_text(encoding="utf-8")
    assert "function Get-QCPrependAuditActions" in orch
    assert "DOCUMENT_MODIFY" in orch
