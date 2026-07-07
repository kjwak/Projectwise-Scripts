# QC workflow email notifications

## Purpose

Send actionable email notifications when **lane QC PDF** workflow states change in ProjectWise (for example `In Development` → `Originated`). The framework is configuration-driven, supports **Mock/dry-run** delivery, and uses **Microsoft Graph** for production sending.

## Lane QC PDFs are the notification authority (current)

- **Current model:** Each sheet stem may have up to three lane QC PDFs: `*-prod.pdf`, `*-rev.pdf`, and `*-chk.pdf`, keyed by `QC_Process_Type` (`Production`, `Review`, `Check`).
- Notifications fire from observed state on the **lane QC PDF** that changed. Dedupe keys include `qcProcessType` so parallel lanes do not suppress each other.
- **DGN** and clean sheet PDF states are **not synchronized** by default; lane-independent behavior is canonical (`QCProcess.EnableLegacySiblingStateSync: false`). Legacy sibling sync remains available behind that flag.

### Legacy (`*-qc.pdf`)

Older deployments used a single `*-qc.pdf` per sheet. Code still supports legacy naming for compatibility (normalization helpers, standalone legacy prepend, older `sheet_index` rows). **Do not** configure new projects around `*-qc.pdf` as the primary lane document.

## Providers

| Provider | Status | Notes |
| --- | --- | --- |
| **Mock** | Implemented | Writes JSON payloads under `notifications/mock/` and logs structured events. |
| **MicrosoftGraph** | Implemented | Client-secret app auth + `sendMail`. |

Enable notifications in `appsettings.json`:

```json
"notifications": {
  "enabled": true,
  "provider": "MicrosoftGraph",
  "dryRun": false
}
```

Set `enabled: false` to suppress all sends (workers and tests remain safe).

## Recipients (ProjectWise attributes)

Recipients are resolved from the **lane QC PDF document** attribute bags (see `docs/workflow/pw-environment-email-attributes.md`). Default column names:

- `EM_Reviewer_Email` → reviewers (Production and Review lanes)
- `EM_Designer_Email` → designers
- `EM_Checker_Email` → checker; also used as the **reviewers** audience when `QC_Process_Type` is **Check**
- Optional `CcEmails`

Override field names under `notifications.attributes` without code changes.

## State → event mapping (TYPSA)

Configured under `notifications.events` keyed by **current workflow state name** (must match ProjectWise exactly):

| State (TYPSA) | Event type | Enabled (committed) | Typical audience |
| --- | --- | --- | --- |
| `Originated` | `READY_FOR_QC` | **yes** | To reviewers, CC designers |
| `Redlines Received` | `REDLINES_RECEIVED` | **yes** | To designers |
| `Ready for Verification` | `READY_FOR_VERIFICATION` | **yes** | To reviewers |
| `Verified` | `QC_COMPLETE` | **yes** | To designers, reviewers, checkers |
| `Error Needs Attention` | `QC_ERROR` | **yes** | To reviewers + designers |

### Legacy event keys (disabled in committed config)

These keys remain in `appsettings.json` for unmigrated environments but are **`enabled: false`** in TYPSA production config:

| Legacy state key | Notes |
| --- | --- |
| `QC Received` | Pre-TYPSA intake-complete state; superseded by `Originated` |
| `Ready for QC` | Pre-TYPSA; superseded by `Originated` |
| `QC Complete` | Pre-TYPSA; superseded by `Verified` |

## Deduplication

When `notifications.dedupe.enabled` is true, the same notification is not sent twice for the same key.

- Default `notifications.dedupe.keyFields` include **`qcProcessType`** plus sheet stem and logical transition anchor, so parallel lanes and replacement GUIDs do not fork dedupe incorrectly.
- `notifications.dedupe.sheetPackageKeyFields` (`sheetStem`, `currentState`, `cycleId`) register a coarser **sheet-package** key used to suppress `audit_trigger` echo notifications.
- If `transitionId` is supplied (from `transition_events`), the dedupe key is `transition:{id}` — **one email per transition row**.
- `transition_events.notification_sent` is set to `1` only after a **successful** send.

Keys are stored in `notifications/dedupe/sent-keys.jsonl` and `notification_log`.

## Missing email attributes

**QC processing always runs** when workflow states change or prepend is triggered. Missing `EM_Designer_Email`, `EM_Reviewer_Email`, or `EM_Checker_Email` values do **not** block prepend, rendition, workflow writeback, or state sync.

When a configured notification fires but the resolved **To** audience for that event is empty (for example `Redlines Received` with no designer email), the send is skipped with `QC_NOTIFICATION_SKIPPED_NO_RECIPIENTS` at information level. The worker job succeeds and QC processing is unaffected.

Recipient resolution order: ProjectWise `EM_*` attributes on DGN and sheet PDF (via `Get-PWQcPrependRoleFieldsFromSourcePdf`), then the lane QC PDF (`*-prod.pdf` / `*-rev.pdf` / `*-chk.pdf`), then `sheet_index`, then `sheet_packages`.

### Optional rollback gate (legacy / opt-in)

When `notifications.rollbackWhenEmailAttributesMissing` is **true**, a human `DOCUMENT_STATE` change to a state with an **enabled** `notifications.events` entry is validated before sibling sync runs. If required role emails cannot be resolved on the sheet:

1. **DGN, sheet PDF, and lane QC PDFs** may be reverted to their prior workflow states (`sheet_index.pw_state_name`).
2. The user who changed state receives a plain email listing missing attribute column names.
3. Normal workflow notifications and sibling alignment for that transition are skipped.

Committed production config sets this to **`false`**. Prepend is never blocked by missing emails even when rollback is enabled.

Automation accounts (`auditPoller.workflowTriggers.automationPwUsernames`) are not gated. States without a configured notification event (for example **Initiate Origination**) are not gated.

## Integration points

1. **`Invoke-QCNotificationForStateChange`** — Call when a lane QC PDF transition is detected (previous state ≠ current state).
2. **`QC.Workflow.psm1`** — After successful `Set-PWDocumentState`, may sync associated members when legacy sibling sync is enabled; invokes notifications when `Context.config` is present.
3. **`QC.Processors.psm1`** — Passes `config`, `job`, and optional `document` into workflow context after successful prepend writeback.
4. **`QC.AuditTriggers.psm1`** — When `audit.workflowTriggers.notifyOnStateChange` is true, sends on lane QC PDF state changes in ProjectWise.

### When emails fire (production)

| Phase | Trigger |
| --- | --- |
| After prepend | Workflow writeback sets lane-appropriate states; **`Originated`** notification when `notifications.events['Originated'].enabled` is true. |
| After rendition (optional defer) | Set `qcRendition.deferReadyForQcNotification: true` to hold intake email until prepend **and** rendition complete (`Invoke-QCReadyForQcNotificationIfReady`). |
| Rest of review cycle | Each enabled TYPSA state (`Redlines Received`, `Ready for Verification`, `Verified`, `Error Needs Attention`) via audit on user-driven transitions, or workflow writeback when automation changes state. |
| Lane PDF already at new state in PW | `Sync-PWAssociatedSheetWorkflowState` compares `sheet_index.pw_state_name` to the canonical state and records/notifies via `_PWD-InvokeStaleSheetIndexAuditStateTriggers` when PW moved first. |

**Lane routing:** Email is sent for lane QC PDFs (`*-prod.pdf`, `*-rev.pdf`, `*-chk.pdf`). Legacy `*-qc.pdf` documents may still notify when present and matched by trigger rules. Sibling sync updates to DGN/sheet PDF do not email unless legacy sibling sync is enabled.

## Microsoft Graph (production)

Keep **non-secret** notification settings in `appsettings.json`. Put Entra credentials in **`appsettings.secrets.json`** (gitignored):

```powershell
Copy-Item .\appsettings.secrets.json.example .\appsettings.secrets.json
```

| Setting | Path |
| --- | --- |
| Entra tenant ID | `notifications.graph.tenantId` |
| Application (client) ID | `notifications.graph.clientId` |
| Client secret **Value** | `notifications.graph.clientSecret` |
| Sender mailbox (UPN) | `notifications.graph.senderMailbox` |

`Read-QCAppSettings` merges `appsettings.secrets.json` after `appsettings.json` and `appsettings.local.json`.

Smoke test: `.\scripts\Test-QCNotificationGraph.ps1 -To you@company.com -Live`

## HTML email templates

TYPSA-branded HTML notifications use `email/templates/qc_notification.html` when `notifications.email.bodyFormat` is `"Html"` (production default).

Preview without sending:

```powershell
.\scripts\Test-QCEmailTemplate.ps1
```

### Submitted By (email template)

The HTML field **Submitted By** (`{SubmittedBy}`) is resolved from the same actor as `transition_events.changed_by_user` / `changed_by_username` (audit row, transition row, or job metadata).

### QC PDF link (`QCPdfUrl`)

`Resolve-QCNotificationQcPdfUrl` resolves the **Open QC PDF** button URL for the lane document that triggered the notification.

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
- `Invoke-QCNotificationForStateChange`
- `Get-QCNotificationDedupeKey`

## Tests

- PowerShell: `test/powershell/test_qc_notifications.ps1`, `test/powershell/test_qc_email_templates.ps1`
- Python config guards: `test/python/test_qc_notifications_config.py`
