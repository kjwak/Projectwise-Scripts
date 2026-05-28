from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
DB = (REPO / "modules" / "Core.Database.psm1").read_text(encoding="utf-8")
POLLER = (REPO / "modules" / "PW.AuditPoller.psm1").read_text(encoding="utf-8")


def test_write_audit_event_rows_exported() -> None:
    assert "function Write-QCAuditEventRows" in DB
    assert "Write-QCAuditEventRows" in DB.split("Export-ModuleMember", 1)[1]


def test_audit_poller_delegates_to_database_writer() -> None:
    scan = POLLER.split("function Invoke-AuditTrailScan", 1)[1].split("function Get-AuditPollCycleCounter", 1)[0]
    assert "Write-QCAuditEventRows -Config $Config -Rows $dbRows" in scan
    assert "INSERT INTO audit_events" not in scan
    assert "o_action IN" not in scan
    assert "Ingest every fetched row into audit_events" in scan


def test_audit_poller_reads_sql_rows_case_insensitively() -> None:
    assert "OrdinalIgnoreCase" in POLLER
    assert "_AuditPoller-GetSqlResultRows" in POLLER
    assert "AUDIT_EVENTS_INGEST" in POLLER


def test_audit_events_insert_uses_dedupe_guard() -> None:
    assert "UX_audit_events_natural_key" in DB or "NOT EXISTS" in DB
    body = DB.split("function Write-QCAuditEventRows", 1)[1].split("function _QDB-SafeWrite", 1)[0]
    assert "INSERT INTO audit_events" in body
    assert "NOT EXISTS" in body
    assert "pw_objguid IS NOT NULL" in body
