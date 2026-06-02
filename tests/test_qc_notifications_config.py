import json
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
APPSETTINGS = REPO_ROOT / "appsettings.json"
NOTIFICATIONS = REPO_ROOT / "modules" / "QC.Notifications.psm1"
TEMPLATES = REPO_ROOT / "modules" / "QC.NotificationTemplates.psm1"
GRAPH = REPO_ROOT / "modules" / "QC.NotificationGraph.psm1"
EMAIL_TEMPLATE = REPO_ROOT / "email" / "templates" / "qc_notification.html"
DOC = REPO_ROOT / "docs" / "qc-notifications.md"


def test_appsettings_notifications_production_delivery_defaults():
    config = json.loads(APPSETTINGS.read_text(encoding="utf-8-sig"))
    notifications = config["notifications"]
    rendition = config["qcRendition"]

    assert notifications["dryRun"] is False
    assert notifications["email"]["bodyFormat"] == "Html"
    assert rendition["deferReadyForQcNotification"] is False
    assert notifications["dedupe"]["enabled"] is True
    assert notifications["dedupe"]["keyFields"] == [
        "documentGuid",
        "eventType",
        "currentState",
    ]


def test_appsettings_notification_events_include_lifecycle_states():
    events = json.loads(APPSETTINGS.read_text(encoding="utf-8-sig"))["notifications"]["events"]

    assert events["QC Received"]["eventType"] == "QC_RECEIVED"
    assert events["QC Received"]["enabled"] is True
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
        "Resolve-QCNotificationQcPdfUrl",
        "Send-QCNotification",
        "Invoke-QCNotificationForStateChange",
        "Write-QCNotificationResult",
        "Get-QCNotificationDedupeKey",
    ]:
        assert f"function {fn}" in text


def test_templates_module_exports_html_renderer():
    text = TEMPLATES.read_text(encoding="utf-8")
    for fn in [
        "ConvertTo-QCEmailHtml",
        "New-QCNotificationEmailTemplateData",
        "Expand-QCNotificationTemplate",
    ]:
        assert f"function {fn}" in text


def test_graph_module_supports_html_and_client_secret_sendmail():
    text = GRAPH.read_text(encoding="utf-8")

    assert "QC_NOTIFICATION_GRAPH_NOT_CONFIGURED" in text
    assert "clientSecret" in text
    assert "sendMail" in text
    assert "New-QCGraphEmailMessage" in text
    assert "contentId" in text
    assert "typsa-logo" in text


def test_appsettings_notifications_email_defaults_html():
    email = json.loads(APPSETTINGS.read_text(encoding="utf-8-sig"))["notifications"]["email"]
    assert email["bodyFormat"] == "Html"
    assert "qc_notification.html" in email["templatePath"]
    assert email["logoPath"].endswith("typsalogo.png.webp")


def test_html_email_template_exists():
    assert EMAIL_TEMPLATE.is_file()
    content = EMAIL_TEMPLATE.read_text(encoding="utf-8")
    assert "cid:typsa-logo" in content
    assert "{QCPdfUrl}" in content
    assert "#bf1425" in content


def test_notifications_doc_covers_qc_pdf_authority_and_graph():
    text = DOC.read_text(encoding="utf-8")

    assert "QC PDF" in text
    assert "not synchronized" in text.lower() or "not synchron" in text.lower()
    assert "Microsoft Graph" in text
    assert "tenantId" in text
    assert "clientSecret" in text
    assert "ConvertTo-QCEmailHtml" in text
    assert "Resolve-QCNotificationQcPdfUrl" in text
