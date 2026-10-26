# Database Telemetry Layer

## Status: Operational (May 2026)

SQL Server serves as a **telemetry and reporting layer** alongside the existing JSON queue. The JSON queue remains the primary execution source; the database provides historical tracking, dashboards, and the `sheet_index` for project status.

---

## Connection

- **Server**: `localhost\SQLEXPRESS` (SQL Server 2025 Express Edition)
- **Database**: `QC_Pipeline`
- **Authentication**: Windows Authentication (`Trusted_Connection=True`)
- **Module**: `modules/Database/Core.Database.psm1`

### Configuration (`appsettings.json`)

```json
{
  "database": {
    "enabled": true,
    "provider": "SqlServer",
    "connectionString": "Server=localhost\\SQLEXPRESS;Database=QC_Pipeline;Trusted_Connection=True;",
    "commandTimeout": 60
  }
}
```

Set `enabled: false` to disable all database writes. The pipeline continues to function normally without the database.

---

## Execution source of truth (invariants)

| Layer | Role | Canonical for |
|-------|------|----------------|
| **JSON queue** (`modules/Queue/QC.Queue.Json.psm1`) | Job scheduling and worker dispatch | `pending/` → `running/` → `succeeded/` / `failed/` |
| **SQL** (`QC_Pipeline`) | Telemetry, audit ingest, package index, mirrors | Reporting, `sheet_index`, `audit_events`, `sheet_packages` |

**Rules:**

1. Workers dequeue and transition jobs **only** from the JSON queue.
2. `processing_jobs` mirrors queue outcomes for dashboards — it is **not** a second scheduler.
3. `sheet_packages` / `sheet_package_qc_pdfs` are canonical for lane/document relationships, not for job execution state.
4. When `database.enabled` is true, the audit capture watermark is **DB-primary** (`watcher_state`) with `queue/_watcher/audit-capture-watermark.txt` as mirror/fallback. See `docs/architecture/hybrid-polling.md`.

---

## Fire-and-Forget Pattern

All telemetry writes use the `_QDB-SafeWrite` internal helper, which silently no-ops when the database is disabled or unreachable. **Pipeline execution never fails because telemetry fails.** This is critical: the JSON queue is the execution backbone, and database availability must not block job processing.

---

## Schema Management

Schema is managed by `Initialize-QCDatabaseSchema` in `Core.Database.psm1`. It is idempotent — all `CREATE TABLE` statements use `IF OBJECT_ID(...) IS NULL` guards. Schema version is tracked in the `schema_version` table. Additive patches (new columns, indexes) run on every init call, including watcher startup, so existing databases upgrade without manual SQL.

`Watch-QCTrigger.ps1` calls `Initialize-QCDatabaseSchema` when `database.enabled` is true.

| Version | Description |
|---------|-------------|
| 1.0.0 | Core tables: `audit_events`, `document_activity`, `document_state_history`, `transition_events`, `poll_runs`, `processing_jobs`, `notification_log` |
| 1.1.0 | Added `sheet_index` table and `v_sheet_status` view |
| 1.2.0 | Comment sync: `qc_comment_runs`, `qc_comments`, `qc_comment_status_history`, `qc_workflow_events` |
| 1.3.0 | Audit natural-key index (`UX_audit_events_natural_key`) |
| 1.4.0 | `pw_users` lookup + `v_audit_events_with_user` view |
| 1.5.0 | `sheet_index` QC attribute columns (`checker_email`, `qc_review_type`, `qc_assigned_to`) |
| 1.14.0 | `qc_workflow_events.transition_event_id` FK to `transition_events.id` |
| 1.21.0 | Notification email threading tables |
| 1.22.0 | `processing_jobs.worker_machine_name` / `worker_pid` (which host claimed the job) |

### Schema 1.2.0 — comment sync (see `docs/workflow/qc-comment-status-sync.md`)

| Table | Purpose |
|-------|---------|
| `qc_comment_runs` | One processor execution per `QC_COMMENT_STATUS_SYNC` job; includes `processor_version` |
| `qc_comments` | Annotation snapshot per run |
| `qc_comment_status_history` | Status transitions per annotation (analytics, recurring issues) |
| `qc_workflow_events` | Workflow/audit events separate from comment snapshots (replay, debugging) |

**Population:** `Write-QCWorkflowEventRow` from `Invoke-QCAuditWorkflowStateChangeTriggers` and processor workflow telemetry on every successful `transition_events` write; comment sync continues to use the same writer via `Write-QCWorkflowEvent`.

Dry-run: set `database.allowWritesInDryRun: false` (default) so `dryRun: true` does not mutate production data.

---

## Tables

### `schema_version`
Tracks applied schema migrations.

| Column | Type | Description |
|--------|------|-------------|
| `version` | NVARCHAR(50) | Semantic version string |
| `applied_at` | DATETIME2(3) | When the migration was applied |
| `description` | NVARCHAR(500) | Human-readable description |

### `audit_events`
Raw ProjectWise audit trail records captured from `dms_audt` during audit polling.

| Column | Type | Description |
|--------|------|-------------|
| `pw_acttime` | NVARCHAR(50) | Timestamp from PW audit trail |
| `pw_action` / `pw_action_name` | INT / NVARCHAR | Action code and human-readable name |
| `pw_objguid` | NVARCHAR(50) | Document GUID |
| `pw_parentguid` | NVARCHAR(50) | Parent folder GUID |
| `pw_itemname` / `pw_itemdesc` | NVARCHAR | Document name and description |
| `resolved_folder` | NVARCHAR(1000) | Resolved folder path |
| `candidate_type` | NVARCHAR(50) | Classified trigger type |
| `processed` | BIT | Whether the event was processed |
| `poll_run_id` | INT | FK to `poll_runs` |
| `pw_userno` | INT | ProjectWise user number (`o_userno` from `dms_audt`) |

Join to human-readable identity via `pw_users` or view `v_audit_events_with_user`.

### `pw_users`
Maps ProjectWise `pw_userno` to login and email (resolved via `Select-PWSQL` on `dms_user`, then `Get-PWUsersByMatch` as fallback).

| Column | Type | Description |
|--------|------|-------------|
| `pw_userno` | INT | Primary key; matches `audit_events.pw_userno` |
| `pw_username` | NVARCHAR(128) | ProjectWise login name |
| `pw_user_email` | NVARCHAR(320) | Email from PW user record |
| `display_name` | NVARCHAR(256) | PW description / display name |
| `first_seen_at` | DATETIMEOFFSET(3) | First time this user was stored |
| `last_synced_at` | DATETIMEOFFSET(3) | Last successful PW lookup |

**Population:** The audit poller calls `Sync-PWUserDirectory` for user numbers in each ingest batch (up to 25 per scan). Backfill historical users:

```powershell
.\scripts\Sync-PWUserDirectory.ps1 -FromAuditEventsOnly
```

### `document_activity`
Enriched per-document summary, upserted on each event. Keyed by `document_guid` (unique constraint).

### `document_state_history`
Time-series of state and attribute changes per document. Used by `v_qc_cycle_aging` for QC duration tracking. `folder_path` is stored in canonical `documents\...` form via `Normalize-QCDocumentsFolderPath` on insert.

**Population:** `Write-QCDocumentStateHistoryRow` from `QC.AuditTriggers.psm1` when audit `DOCUMENT_STATE` or `DOCUMENT_ATTR` events produce real diffs (see `auditPoller.workflowTriggers`).

### `transition_events`
Business-level events (QC stage changes, attribute changes). Links to notifications via `notification_log.transition_id`. `folder_path` uses the same canonical normalization as `document_state_history`.

**Population:** `Write-QCTransitionEvent` from audit workflow triggers and processor telemetry; `notification_sent` updated after a successful `Invoke-QCNotificationForStateChange` on lane QC PDFs. When PW already shows the new state before sync runs, `_PWD-InvokeStaleSheetIndexAuditStateTriggers` compares `sheet_index.pw_state_name` to the canonical state and records the transition anyway.

### `poll_runs`
Operational health of the audit poller. One row per watcher tick.

| Column | Type | Description |
|--------|------|-------------|
| `watermark_before` / `watermark_after` | NVARCHAR(50) | High-water mark before and after the poll |
| `events_fetched` / `events_relevant` | INT | Raw vs. filtered event counts |
| `candidates_created` / `jobs_enqueued` | INT | Candidate and job counts |
| `duration_ms` | INT | Poll cycle duration |
| `is_reconciliation` | BIT | Whether this was a full reconciliation scan |
| `error_message` | NVARCHAR(2000) | Error details if the poll failed |

### `processing_jobs`
Mirrors JSON queue job outcomes for dashboards. Written by `Write-QCJobTelemetry` in `Run-QCProcessor.ps1`.

Queue type `QC_COMMENT_STATUS_SYNC` is stored as **`QC_STATE`** in `job_type` (mapping via `Get-QCProcessingJobType`). Automation state writes (audit sheet sync, QC_PREPEND workflow state) use **`Write-QCStateChangeJobTelemetry`**, which always upserts `job_type = QC_STATE` (often with `job_id` like `{queueJobId}|state` or `qc-state-{guid}-{from}-{to}`). Other queue types are stored unchanged unless `database.processingJobTypeMap` overrides them.

| Column | Type | Description |
|--------|------|-------------|
| `job_id` | NVARCHAR(200) | Unique job ID (unique constraint) |
| `job_type` | NVARCHAR(50) | `STATUS_SET_GEN`, `QC_PREPEND`, etc. |
| `status` | NVARCHAR(20) | `pending`, `succeeded`, `failed`, `dead` |
| `duration_ms` | INT | Job execution time in milliseconds |
| `source_path` / `source_folder` | NVARCHAR | Document and folder paths (canonical: lowercase `documents\...`; see `Normalize-QCDocumentsFolderPath` in `Core.Paths.psm1`) |
| `worker_machine_name` | NVARCHAR(128) | Host that claimed the job (`COMPUTERNAME`). Stamped on claim and kept on succeeded/failed. Schema 1.22.0. |
| `worker_pid` | INT | Worker process ID on that host. Schema 1.22.0. |
| `error_code` / `error_message` | NVARCHAR | Error details on failure |
| `result_data` | NVARCHAR(MAX) | JSON result payload |

### `notification_log`
Tracks sent notifications with dedupe keys. `folder_path` is normalized on insert.

**Folder path canonical form:** lowercase `documents\<project>\...` (see `Normalize-QCDocumentsFolderPath` in `Core.Paths.psm1`). Repair existing rows: `scripts/sql/normalize-documents-folder-paths.sql` or `scripts/Repair-QCDocumentsFolderPaths.ps1 -Table all`.

### `sheet_index`
Indexes all documents in watched `CADD\Sheets` folders. Supports project status reporting, QC PDF pairing, and ownership tracking.

| Column | Type | Description |
|--------|------|-------------|
| `document_guid` | NVARCHAR(40) | PW document GUID (unique constraint) |
| `document_name` | NVARCHAR(500) | Filename |
| `folder_path` | NVARCHAR(1000) | Full PW folder path |
| `project_name` | NVARCHAR(200) | Extracted project name |
| `watch_root` | NVARCHAR(500) | Matched watch root from config |
| `extension` | NVARCHAR(20) | File extension (`.pdf`, `.dgn`) |
| `qc_pdf_guid` / `qc_pdf_name` | NVARCHAR | Linked QC PDF document |
| `source_type` | NVARCHAR(10) | `pdf` or `dgn` |
| `designer_email` | NVARCHAR(200) | `EM_Designer_Email`, else `QC_Designer_Email` |
| `reviewer_email` | NVARCHAR(200) | `EM_Reviewer_Email`, else `QC_Reviewer_Email` |
| `checker_email` | NVARCHAR(200) | `QC_Checker_Email` |
| `qc_review_type` | NVARCHAR(100) | `QC_Review_Type` |
| `qc_assigned_to` | NVARCHAR(200) | `QC_Assigned_To` |
| `pw_state_name` | NVARCHAR(100) | Current workflow state |
| `qc_stage` / `qc_status` | NVARCHAR | Legacy stage column; `qc_status` mirrors `QC_Status` |

---

## Views

| View | Purpose |
|------|---------|
| `v_qc_cycle_aging` | Documents in QC with duration since **Originated** state (reads `document_state_history.folder_path`) |
| `v_folder_activity` | Folder event counts and last activity, 7-day window (reads `audit_events.resolved_folder`) |

Fix path inconsistencies in the **base tables**; views reflect those columns automatically.
| `v_poller_health` | Last 100 poll runs with status and watermark |
| `v_job_summary` | Job counts by type and status with average duration |
| `v_sheet_status` | Project status overview of all indexed sheets |

---

## Public API (`Core.Database.psm1`)

### Connection & Query

| Function | Purpose |
|----------|---------|
| `Test-QCDatabaseEnabled` | Check if database is enabled in config |
| `Get-QCDatabaseConnection` | Open a `SqlConnection` (caller disposes) |
| `Invoke-QCDatabaseQuery` | Parameterized SELECT → `DataTable` |
| `Invoke-QCDatabaseNonQuery` | Parameterized INSERT/UPDATE/DELETE |
| `Invoke-QCDatabaseScalar` | Scalar query returning a single value |
| `Invoke-QCDatabaseBatch` | Execute multi-statement SQL with GO separators |
| `Initialize-QCDatabaseSchema` | Idempotent schema creation and migration |

### Telemetry Writers (Fire-and-Forget)

| Function | Called From | Purpose |
|----------|------------|---------|
| `Write-QCJobTelemetry` | `Run-QCProcessor.ps1` | Record job outcome with duration |
| `Write-QCPollRunTelemetry` | `Watch-QCTrigger.ps1` | Record poll cycle metrics |
| `Write-QCNotificationTelemetry` | `QC.Notifications.psm1` | Record sent notifications |
| `Write-QCDocumentStateHistoryRow` | `QC.AuditTriggers.psm1` | State/attribute change time-series |
| `Write-QCTransitionEvent` | `QC.AuditTriggers.psm1` | Business-level transition row |
| `Write-QCWorkflowEventRow` | `QC.AuditTriggers.psm1`, `QC.CommentSync.Database.psm1` | Mirror row in `qc_workflow_events` |
| `Update-QCTransitionEventNotification` | `QC.AuditTriggers.psm1` | Mark transition after email sent |
| `Write-QCStateChangeJobTelemetry` | `QC.AuditTriggers`, `PW.Discovery` | `processing_jobs` row with `job_type = QC_STATE` for automation state applies |
| `New-QCStateChangeJobId` | `Core.Database.psm1` | Stable `job_id` for state-only telemetry |
| `Invoke-QCProcessorWorkflowStateTelemetry` | `QC.Workflow.psm1`, `QC.CommentStatusProcessor.psm1` | Processor state writes (no duplicate email) |
| `Invoke-QCProcessorWorkflowAttributeTelemetry` | `QC.Workflow.psm1` | Processor attribute writeback rows |
| `Write-QCSheetIndex` | `Watch-QCTrigger.ps1` | Upsert sheet document to index |
| `Sync-PWAssociatedSheetWorkflowState` | `Watch-QCTrigger.ps1` (via audit) | On `DOCUMENT_STATE`, set associated DGN, sheet PDF, and lane QC PDFs to the same workflow state when legacy sibling sync is enabled |
| `Sync-PWSheetIndexOwnership` | `Watch-QCTrigger.ps1` (via audit) | On `DOCUMENT_ATTR`, re-read EM_* and QC_* from PW into `sheet_index` (audit has no old/new values) |
| `Sync-PWAssociatedSheetReviewTypeAttributes` | `PW.Discovery.psm1` (from `Sync-PWSheetIndexOwnership`) | On `QC_Process_Type` / `QC_Review_Type` change, align DGN / sheet PDF / lane QC PDFs in PW and `sheet_index` |
| `Update-QCSheetQcPdf` | `Watch-QCTrigger.ps1`, `Run-QCProcessor.ps1` | Link QC PDF to source document |

---

## Power BI Integration

Connect Power BI to `QC_Pipeline` via SQL Server connector:

1. **Get Data → SQL Server** → `localhost\SQLEXPRESS` / `QC_Pipeline`
2. Import tables and views
3. Create relationships: `processing_jobs.job_id` ↔ related tables, `sheet_index.document_guid` ↔ `audit_events.pw_objguid`

---

## Retention / cleanup scripts

Execution state lives in the JSON queue and operational tables (`sheet_index`, `processing_jobs`); these scripts do not touch those.

**Audit events** — periodic retention via `Invoke-QCDatabaseRetention.ps1` and `database.retention` in appsettings (default: processed rows older than 90 days).

**Workflow events** — no automatic retention. Use `Remove-QCWorkflowEvents.ps1` with `-FolderPathLike` to delete bad telemetry for specific folder path fragments (matches `sheet_index.folder_path` and `qc_comment_runs.pw_path`).

| Script | Purpose |
|--------|---------|
| `scripts/Remove-QCAuditEvents.ps1` | Delete aged `audit_events` by `captured_at` (default: processed rows only) |
| `scripts/Remove-QCWorkflowEvents.ps1` | Delete `qc_workflow_events` (and optional comment tables) by folder path / document / job |
| `scripts/Invoke-QCDatabaseRetention.ps1` | Scheduled `audit_events` cleanup only |

Preview by default; pass `-ConfirmDeletes` to apply.

```powershell
# Scheduled audit retention
.\scripts\Invoke-QCDatabaseRetention.ps1 -ConfirmDeletes

# One-off bad workflow telemetry under specific projects
.\scripts\Remove-QCWorkflowEvents.ps1 -FolderPathLike 'AZFWY2302-018', 'SomeOtherProject'
.\scripts\Remove-QCWorkflowEvents.ps1 -ConfirmDeletes -FolderPathLike 'AZFWY2302-018'
```

---

## Validation Scripts

| Script | Purpose |
|--------|---------|
| `tools/discovery/Test-SQLServerConnectivity.ps1` | Verify connection, schema init, CRUD operations |
| `tools/discovery/Test-SheetIndexAndAuditPoller.ps1` | End-to-end test of sheet index population and audit scanning |

---

## Automation Events (schema 1.20.0)

`automation_events` is the **primary diagnostic source** for worker/watcher logging and MCP debugging. Structured JSONL files (`Run-QCProcessor_*.jsonl`, `Watch-QCTrigger_*.jsonl`) remain as a short-retention emergency backup only.

### Write path

All `Write-QCJsonLog` calls route through `Write-QCAutomationEvent` (`modules/Core/Core.Telemetry.psm1`):

1. JSONL file line is written first (crash-recovery backup when `QC_JSON_LOG_DIR` is set).
2. Filtered events are inserted into `automation_events` (never throws; failures are returned as `QCResult`).

Call `Set-QCAutomationTelemetryContext` once at process start (`Run-QCProcessor`, `Watch-QCTrigger`) to bind `process_name`, `run_id`, and JSONL retention settings.

### Configuration (`telemetry.automationEvents`)

```json
"telemetry": {
  "automationEvents": {
    "enabled": true,
    "jsonLogRetentionDays": 7,
    "jsonLogMaxFileSizeMb": 50,
    "excludeCodes": ["WATCH_TICK_START", "WATCH_TICK_SLEEP", "WORKER_NO_JOB"]
  }
}
```

### Persisted vs filtered event codes

| Rule | Behavior |
|------|----------|
| All `Warning` / `Error` | Always persisted |
| `WATCH_TICK_START`, `WATCH_TICK_SLEEP`, `WORKER_NO_JOB` | Excluded at Information level |
| `WORKER_STAGE` | Excluded when message/stage matches queue-polling noise; kept when `job_id`, `document_guid`, or `sheet_package_id` is present |
| Everything else | Persisted |

Full original payload is stored in `data_json`; common fields are indexed (`job_id`, `document_guid`, `sheet_package_id`, `audit_event_id`, `folder_path`).

### MCP debug views

| View | Purpose |
|------|---------|
| `v_mcp_automation_events_recent` | Last 14 days of events |
| `v_mcp_process_health` | Per-process 24h error/warning counts |
| `v_mcp_job_timeline` | Events with `job_id` |
| `v_mcp_document_debug_events` | Events with `document_guid` |
| `v_mcp_package_debug_events` | Events with `sheet_package_id` |
| `v_mcp_audit_scan_history` | `WATCH_AUDIT_*` / `AUDIT_*` codes |
| `v_mcp_recent_errors` | Warning/error events (7 days) |

MCP tools (`get_recent_errors`, `get_process_health`, `get_audit_scan_history`, `get_job_timeline`, `get_document_debug_events`, `get_package_debug_events`) query these views first. JSONL fallback is used only when `automation_events` is unavailable or `force_jsonl_fallback` is set.

### Historical backfill

```powershell
.\scripts\Import-QCJsonlLogsToAutomationEvents.ps1 -LogDirectory C:\path\to\queue\_logs
```

Re-runs are idempotent via `dedupe_key` (SHA-256 of ts/process/code/message/job/run).

---

## Future: Queue Migration

The database is positioned for a phased migration where it could eventually replace the JSON queue for job management. This is planned for after the telemetry layer has proven stable. The current priority is building confidence in SQL Server reliability before moving execution-critical functionality to it.

---

## Timestamp Convention

All database timestamps use `DATETIMEOFFSET(3)` with `SYSDATETIMEOFFSET()` defaults. Application code standardizes display to Mountain Time (MST/MDT) using `Get-QCTimestamp` from `Core.Runtime.psm1`. Raw UTC values are stored for accuracy.
