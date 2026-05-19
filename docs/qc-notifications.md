# QC workflow email notifications

## Purpose

Send actionable email notifications when **QC PDF** workflow states change in ProjectWise (for example `In Production` → `QC Received`). The framework is configuration-driven, supports **Mock/dry-run** delivery today, and is structured for **Microsoft Graph** production sending once IT provides Entra app and mailbox details.

## QC PDF is the state authority

- The **`*-qc.pdf`** document (history/overlay QC PDF in ProjectWise) is the authoritative object for QC workflow state.
- **DGN**, clean sheet PDF, and QC PDF states are **not synchronized** by this automation. Notifications fire from observed state on the QC PDF only.
- Do not infer QC routing state from non-QC documents.

## Providers

| Provider | Status | Notes |
| --- | --- | --- |
| **Mock** | Implemented | Writes JSON payloads under `notifications/mock/` and logs structured events. Default while Graph credentials are unavailable. |
| **MicrosoftGraph** | Stub | Validates config; returns `not configured` or `not implemented` until certificate-based sendMail is added. |

Enable notifications in `appsettings.json`:

```json
"notifications": {
  "enabled": true,
  "provider": "Mock",
  "dryRun": true
}
```

Set `enabled: false` to suppress all sends (workers and tests remain safe).

## Recipients (ProjectWise attributes)

Recipients are resolved from the **QC PDF document** attribute bags (see `docs/pw-environment-email-attributes.md`). Default column names:

- `EM_Reviewer_Email` → reviewers
- `EM_Designer_Email` → designers
- Optional `CcEmails`

Override field names under `notifications.attributes` without code changes.

## State → event mapping

Configured under `notifications.events` keyed by **current workflow state name** (must match ProjectWise exactly):

| State | Event type | Typical audience |
| --- | --- | --- |
| QC Received | `QC_RECEIVED` | To reviewers, CC designers |
| Corrections In Progress | `CORRECTIONS_IN_PROGRESS` | To designers, CC reviewers |
| Backcheck In Progress | `BACKCHECK_IN_PROGRESS` | To reviewers, CC designers |
| Error Needs Attention | `QC_ERROR` | To reviewers + designers |

## Deduplication

When `notifications.dedupe.enabled` is true, the same notification is not sent twice for the same key (default: `documentGuid|eventType|currentState`). Keys are stored in `notifications/dedupe/sent-keys.jsonl`.

## Integration points

1. **`Invoke-QCNotificationForStateChange`** — Call when a QC PDF transition is detected (previous state ≠ current state).
2. **`QC.Workflow.psm1`** — After a successful `Set-PWDocumentState` write (`changed = true`), invokes notifications when `Context.config` is present.
3. **`QC.Processors.psm1`** — Passes `config`, `job`, and optional `document` into workflow context after successful prepend writeback.

Future watchers can call `Invoke-QCNotificationForStateChange` directly when polling QC PDF workflow state.

## Microsoft Graph (future production)

Do **not** store secrets in the repository. IT will provide:

| Setting | `appsettings.json` path |
| --- | --- |
| Entra tenant ID | `notifications.graph.tenantId` |
| Application (client) ID | `notifications.graph.clientId` |
| Sender mailbox (UPN) | `notifications.graph.senderMailbox` |
| Certificate thumbprint **or** path to `.pfx` | `notifications.graph.certificateThumbprint` / `certificatePath` |

Implementation plan (see `modules/QC.NotificationGraph.psm1`):

1. Certificate-based client credentials token against Entra.
2. `POST /v1.0/users/{senderMailbox}/sendMail` with plain-text body from templates.
3. Honor `dryRun` by building the payload without calling Graph.

Switch provider:

```json
"provider": "MicrosoftGraph",
"dryRun": false
```

## Modules

| Module | Role |
| --- | --- |
| `QC.Notifications.psm1` | Settings, events, dedupe, orchestration |
| `QC.NotificationTemplates.psm1` | Subject/body templates |
| `QC.NotificationMock.psm1` | Mock file + result envelope |
| `QC.NotificationGraph.psm1` | Graph stub + config validation |

## Public functions

- `New-QCNotificationEvent`
- `Resolve-QCNotificationRecipients`
- `Send-QCNotification`
- `Invoke-QCNotificationForStateChange`
- `Write-QCNotificationResult`
- `Get-QCNotificationDedupeKey`

## Result object

Successful Mock send returns a hashtable similar to:

```json
{
  "success": true,
  "provider": "Mock",
  "dryRun": true,
  "eventType": "QC_RECEIVED",
  "documentName": "D-101-qc.pdf",
  "to": ["reviewer@company.com"],
  "cc": ["designer@company.com"],
  "message": "Mock notification written.",
  "timestampUtc": "2026-05-19T12:00:00.0000000Z",
  "filePath": "notifications/mock/..."
}
```

## Tests

- PowerShell: `test/test_qc_notifications.ps1`
- Python config guards: `tests/test_qc_notifications_config.py`
