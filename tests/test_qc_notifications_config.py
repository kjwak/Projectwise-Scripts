import json
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
APPSETTINGS = REPO_ROOT / "appsettings.json"
NOTIFICATIONS = REPO_ROOT / "modules" / "QC.Notifications.psm1"
GRAPH = REPO_ROOT / "modules" / "QC.NotificationGraph.psm1"
DOC = REPO_ROOT / "docs" / "qc-notifications.md"


def test_appsettings_notifications_defaults_disabled_mock():
    config = json.loads(APPSETTINGS.read_text(encoding="utf-8-sig"))
    notifications = config["notifications"]

    assert notifications["enabled"] is False
    assert notifications["provider"] == "Mock"
    assert notifications["dryRun"] is True
    assert notifications["dedupe"]["enabled"] is True
    assert notifications["dedupe"]["keyFields"] == [
        "documentGuid",
        "eventType",
        "currentState",
    ]


def test_appsettings_notification_events_include_lifecycle_states():
    events = json.loads(APPSETTINGS.read_text(encoding="utf-8-sig"))["notifications"]["events"]

    assert events["Ready for QC"]["eventType"] == "READY_FOR_QC"
    assert events["Review In Progress"]["eventType"] == "REVIEW_IN_PROGRESS"
    assert events["Redlines Issued"]["eventType"] == "REDLINES_ISSUED"
    assert events["Corrections In Progress"]["eventType"] == "CORRECTIONS_IN_PROGRESS"
    assert events["Verification In Progress"]["eventType"] == "VERIFICATION_IN_PROGRESS"
    assert events["Error Needs Attention"]["eventType"] == "QC_ERROR"


def test_notifications_module_exports_state_change_entrypoint():
    text = NOTIFICATIONS.read_text(encoding="utf-8")

    for fn in [
        "New-QCNotificationEvent",
        "Resolve-QCNotificationRecipients",
        "Send-QCNotification",
        "Invoke-QCNotificationForStateChange",
        "Write-QCNotificationResult",
        "Get-QCNotificationDedupeKey",
    ]:
        assert f"function {fn}" in text


def test_graph_module_documents_not_configured_and_todo():
    text = GRAPH.read_text(encoding="utf-8")

    assert "QC_NOTIFICATION_GRAPH_NOT_CONFIGURED" in text
    assert "TODO:" in text
    assert "sendMail" in text


def test_notifications_doc_covers_qc_pdf_authority_and_graph():
    text = DOC.read_text(encoding="utf-8")

    assert "QC PDF" in text
    assert "not synchronized" in text.lower() or "not synchron" in text.lower()
    assert "Microsoft Graph" in text
    assert "tenantId" in text
    assert "certificateThumbprint" in text
