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
| **MicrosoftGraph** | Implemented | Client-secret app auth + `sendMail`. Certificate auth reserved for future use. |

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

- `EM_Reviewer_Email` → reviewers (except **Independent Check**; see below)
- `EM_Designer_Email` → designers
- `EM_Checker_Email` → checker; also used as the **reviewers** audience when `QC_Review_Type` is **Independent Check**
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

When `notifications.dedupe.enabled` is true, the same notification is not sent twice for the same key.

- Default `notifications.dedupe.keyFields` use **sheet stem + logical transition** (`logicalTransitionAnchor`), not per-file `documentGuid`, so a replacement `*-qc.pdf` GUID does not fork dedupe.
- `notifications.dedupe.sheetPackageKeyFields` (`sheetStem`, `currentState`, `cycleId`) register a coarser **sheet-package** key used to suppress `audit_trigger` echo notifications (for example `DOCUMENT_CIN` after sheet-group `DOCUMENT_STATE` already enqueued).
- If `transitionId` is supplied (from `transition_events`), the dedupe key is `transition:{id}` — **one email per transition row**.
- Otherwise the key uses `stateTransitionKey` (e.g. `audit:38268`), then `sheetStem`, `eventType`, `previousState`, `currentState` so a **later QC cycle** to the same target state can email again.
- `transition_events.notification_sent` is set to `1` only after a **successful** send. `notification_sent = 0` means no successful email yet (not “deduped forever”).
- Pending transitions (`notification_sent = 0`) are not blocked by older `sent-keys.jsonl` entries that used sheet-only keys.

Keys are stored in `notifications/dedupe/sent-keys.jsonl` and `notification_log`.

## Missing email attributes on state change

When `notifications.rollbackWhenEmailAttributesMissing` is **true** (default), a human `DOCUMENT_STATE` change to a state with an **enabled** `notifications.events` entry is validated before sibling sync runs. If required role emails (`EM_Designer_Email`, `EM_Reviewer_Email`, or `EM_Checker_Email` for Independent Check) cannot be resolved on the sheet:

1. **DGN, sheet PDF, and `*-qc.pdf`** are reverted to their prior workflow states (`sheet_index.pw_state_name`).
2. The user who changed state (`pw_users.pw_user_email` for audit `o_userno`) receives a plain email listing the missing attribute column names.
3. Normal workflow notifications, prepend enqueue, and sibling alignment for that transition are skipped.

Automation accounts (`auditPoller.workflowTriggers.automationPwUsernames`) are not gated. States without a configured notification event (for example **QC Initiated**) are not gated.

## Integration points

1. **`Invoke-QCNotificationForStateChange`** — Call when a QC PDF transition is detected (previous state ≠ current state).
2. **`QC.Workflow.psm1`** — After a successful primary `Set-PWDocumentState`, syncs **DGN / sheet PDF / `*-qc.pdf`** via `Sync-PWAssociatedSheetMembersToWorkflowState`, then invokes **QC Received** notification when `Context.config` is present.
3. **`QC.Processors.psm1`** — Passes `config`, `job`, and optional `document` into workflow context after successful prepend writeback.
4. **`QC.AuditTriggers.psm1`** — When `audit.workflowTriggers.notifyOnStateChange` is true, sends on QC PDF state changes in ProjectWise. Default `ignoreStateChangeFromAutomation: false` includes automation account transitions (dedupe avoids double-send with prepend).

### When emails fire (production)

| Phase | Trigger |
| --- | --- |
| After prepend | All sheet members → **QC Received**; `QC.Workflow` calls `Invoke-QCNotificationForStateChange` for **QC Received** when `qcRendition.deferReadyForQcNotification` is **false** (current default). |
| After rendition (optional defer) | Set `deferReadyForQcNotification: true` to hold **QC Received** email until prepend **and** rendition complete (`Invoke-QCReadyForQcNotificationIfReady`). |
| Rest of review cycle | Each configured state (Review In Progress, Redlines Issued, Corrections In Progress, Verification In Progress, Error Needs Attention) via audit on user-driven transitions, or workflow writeback when automation changes state. |
| QC PDF already at new state in PW | `Sync-PWAssociatedSheetWorkflowState` compares `sheet_index.pw_state_name` to the canonical state and records/notifies via `_PWD-InvokeStaleSheetIndexAuditStateTriggers` when PW moved first (common when you change state on the QC PDF). |

**Gotcha:** Email is only sent for `*-qc.pdf` (`qcPdfNotificationsOnly`). Sibling sync updates to DGN/sheet PDF do not email. If PW already shows the new state on the QC PDF before sync runs, the member loop used to skip notification entirely; the sheet-index comparison above fixes that path.

Future watchers can call `Invoke-QCNotificationForStateChange` directly when polling QC PDF workflow state.

## Microsoft Graph (production)

Keep **non-secret** notification settings in `appsettings.json` (`enabled`, `provider`, `events`, etc.). Put Entra credentials in **`appsettings.secrets.json`** (gitignored):

```powershell
Copy-Item .\appsettings.secrets.json.example .\appsettings.secrets.json
```

| Setting | Path |
| --- | --- |
| Entra tenant ID | `notifications.graph.tenantId` |
| Application (client) ID | `notifications.graph.clientId` |
| Client secret **Value** (not Secret ID) | `notifications.graph.clientSecret` |
| Sender mailbox (UPN) | `notifications.graph.senderMailbox` |

`Read-QCAppSettings` merges `appsettings.secrets.json` after `appsettings.json` and `appsettings.local.json`.

The app registration needs **application** permission `Mail.Send` with admin consent. HTML notifications and conversation threading prefer `Mail.ReadWrite` (create draft / reply APIs). When only `Mail.Send` is granted, the provider automatically falls back to `sendMail` (emails still deliver; thread message IDs may be unavailable).

In `appsettings.json`, set `"provider": "MicrosoftGraph"` and `"dryRun": false` for live sends (use `"dryRun": true` only while validating payloads).

Smoke test: `.\scripts\Test-QCNotificationGraph.ps1 -To you@company.com -Live`

## HTML email templates

TYPSA-branded HTML notifications use `email/templates/qc_notification.html` when `notifications.email.bodyFormat` is `"Html"` (production default). Use `"Text"` for plain-body only.

| Setting | Path | Description |
| --- | --- | --- |
| Body format | `notifications.email.bodyFormat` | `Text` or `Html` |
| Template | `notifications.email.templatePath` | HTML file (repo-relative) |
| Logo | `notifications.email.logoPath` | Inline CID attachment (`cid:typsa-logo`) |
| Environment badge | `notifications.email.environment` | e.g. `Production`, `Test` |
| pwlink base | `notifications.email.pwLinkBaseUrl` | Used when building `{QCPdfUrl}` from GUID |
| pwlink app | `notifications.email.pwLinkApp` | `pwe`, `web`, or `webview` |

Preview without sending:

```powershell
.\scripts\Test-QCEmailTemplate.ps1
```

Output: `output/test_qc_email.html`, `output/test_qc_graph_payload.json`. Send live test: `.\scripts\Test-QCEmailTemplate.ps1 -To you@company.com -Send -Live`

### Submitted By (email template)

The HTML field **Submitted By** (`{SubmittedBy}`) is a display label (email, display name, or PW username), not the raw integer stored in SQL.

It is resolved from the same actor as `transition_events.changed_by_user` / `changed_by_username`:

1. When `stateTransitionKey` is `transition:{id}`, actor fields are read from that `transition_events` row.
2. When the key is `audit:{id}`, actor fields come from `audit_events.pw_userno` (and `pw_users` when synced).
3. Otherwise callers pass `ChangedByUser` / `ChangedByUsername`, or the QC job metadata keys `changedByUser` / `changedByUsername`.

`Resolve-QCNotificationSubmittedBy` turns that PW user number into the human-readable label shown in email. It does not use designer/reviewer attribute emails for Submitted By.

### QC PDF link (`QCPdfUrl`)

`Resolve-QCNotificationQcPdfUrl` resolves the **Open QC PDF** button URL (first match wins):

1. `Event.qcPdfUrl` (explicit)
2. PW attribute `notifications.attributes.qcPdfUrlField`
3. `Get-PWDocumentsByGUIDs` → `ProjectWiseWebLink` on the document object
4. `notifications.email.qcPdfUrlTemplate` with `{documentGuid}`, `{documentName}`, etc.
5. Built pwlink URL from `pwLinkBaseUrl` + `documentGuid` + encoded datasource

Discovery: `tools/discovery/Test-PWDocumentProperties.ps1` prints `ProjectWiseWebLink`, `DocumentURN`, and `DocumentGUID`.

## Modules

| Module | Role |
| --- | --- |
| `QC.Notifications.psm1` | Settings, events, dedupe, orchestration, `Resolve-QCNotificationQcPdfUrl` |
| `QC.NotificationTemplates.psm1` | Subject/body + `ConvertTo-QCEmailHtml` |
| `QC.NotificationMock.psm1` | Mock file + result envelope |
| `QC.NotificationGraph.psm1` | Graph send + `New-QCGraphEmailMessage` (HTML + inline logo) |

## Public functions

- `New-QCNotificationEvent`
- `Resolve-QCNotificationRecipients`
- `Resolve-QCNotificationQcPdfUrl`
- `Resolve-QCNotificationStateChangeActor`
- `ConvertTo-QCEmailHtml`
- `New-QCNotificationEmailTemplateData`
- `New-QCGraphEmailMessage`
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

- PowerShell: `test/test_qc_notifications.ps1`, `test/test_qc_email_templates.ps1`
- Python config guards: `tests/test_qc_notifications_config.py`
