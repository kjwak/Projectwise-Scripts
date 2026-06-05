#!/usr/bin/env python3
"""Static validation for QC workflow notification dedupe/logging semantics."""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
notifications = (ROOT / "modules" / "QC.Notifications.psm1").read_text(encoding="utf-8")
workflow = (ROOT / "modules" / "QC.Workflow.psm1").read_text(encoding="utf-8")

required_snippets = {
    "logical sheet default key includes recipient set": "keyFields = @('notificationType', 'sheetStem', 'targetState', 'cycleId', 'recipientKey')",
    "recipient key helper exists": "function _QCN-BuildRecipientKey",
    "transition id is opt-in only": "Backward-compatible opt-in only. Not part of the default key",
    "send attempt log code exists": "QC_NOTIFICATION_SEND_ATTEMPT",
    "sent log code exists": "QC_NOTIFICATION_SENT",
    "deduped log code exists": "QC_NOTIFICATION_DEDUPED",
    "automation echo skip log code exists": "QC_NOTIFICATION_SKIPPED_AUTOMATION_ECHO",
    "lifecycle logger includes auditEventId": "auditEventId = _QCN-GetNotificationAuditEventId -Event $Event",
    "invoke path computes recipient key": "$recipientKey = _QCN-BuildRecipientKey -To @($resolved.to) -Cc @($resolved.cc)",
    "invoke path stores cycle id": "$event['cycleId'] = [string]$event.stateTransitionKey",
}
missing = [name for name, snippet in required_snippets.items() if snippet not in notifications]

# The prior defect was caused by transitionId short-circuiting all configured key fields.
forbidden = "if ($Event.ContainsKey('transitionId') -and $null -ne $Event.transitionId)"
get_start = notifications.index("function Get-QCNotificationDedupeKey")
test_start = notifications.index("function Test-QCNotificationDedupe", get_start)
get_body = notifications[get_start:test_start]
if forbidden in get_body and "return ('transition:{0}'" in get_body:
    missing.append("Get-QCNotificationDedupeKey must not return transition:<id> before logical key fields")

if "Get-QCNotificationDedupeKey -Event $eventForDedupe" not in workflow:
    missing.append("workflow notification enqueue path still calls shared dedupe helper")

if missing:
    print("Notification dedupe validation failed:")
    for item in missing:
        print(f" - {item}")
    raise SystemExit(1)
print("notification dedupe validation passed")
