# Modules directory file guide

Modules own the reusable pipeline logic. Scripts should call these functions instead of duplicating credential parsing, ProjectWise compatibility handling, queue transitions, status-set generation, trigger evaluation, or processor dispatch.

**Inventory sync:** every module implementation under `modules/` must have a row with its relative path. Run `test/powershell/test_module_inventory.ps1` and `test/powershell/test_module_folder_shims.ps1` after adding or moving modules.

## Folder layout (Phase 4E)

| Folder | Role |
| --- | --- |
| `Core/` | Results, runtime, paths, config, logging, hashing, telemetry, watcher orchestration |
| `Database/` | SQL connectivity and telemetry writers |
| `ProjectWise/` | PW connection, discovery, audit poller, users |
| `Workflow/` | Workflow writeback, audit triggers, process type |
| `Queue/` | JSON queue, job factory, worker policy, filters, triggers |
| `Processing/` | Processors, status set, rendition, comment sync, PDF export |
| `Notifications/` | Email notifications, Graph, templates, threads, watcher alerts |
| `Reporting/` | Reporting scan aggregation |
| `Diagnostics/` | MCP debug module |
| `Ops/` | Logon ops console control plane |

Active scripts and tests import the organized implementation paths below. Some callers retain compatibility logic for older checkouts.

**Phase 4G prototype:** `modules/Core/QC.Core.psd1` and `modules/Queue/QC.Queue.psd1` bundle nested `.psm1` files for test/documentation only — not used by production entrypoints.

## Refactoring assessment completed

- ProjectWise command-line diagnostics now reuse `PW.Connection.psm1` for key/value credential parsing, ProjectWise module loading, connection cleanup, folder-view compatibility, child splitting, and document enumeration.
- This reduces duplicate script-local credential and `Get-PWFolderView` compatibility logic and makes future ProjectWise diagnostics safer to implement as thin wrappers.
- Future cleanup should continue moving long-lived orchestration/UI behavior from scripts into modules: dashboard rendering can become a dashboard module, and watcher ProjectWise expansion can be consolidated behind `PW.Discovery.psm1`.

## File purposes

### Core

| File | Purpose |
| --- | --- |
| `Core/Core.Config.psm1` | General appsettings loading, setting access, and compatibility validation helpers. |
| `Database/Core.Database.psm1` | SQL Server connectivity, schema management, and telemetry writers for QC pipeline data. |
| `Core/Core.Hashing.psm1` | Shared SHA-256 helpers for file and text hashing. |
| `Core/Core.Logging.psm1` | Structured application/audit logging primitives. |
| `Core/Core.Paths.psm1` | Path normalization, root-containment checks, and path splitting for local and ProjectWise-style paths. |
| `Core/Core.Results.psm1` | Standard QC result constructors (`IsSuccess`, `Code`, `Message`, `Data`) used across modules. |
| `Core/Core.Runtime.psm1` | Runtime helpers for deep hashtable conversion, appsettings loading, merged config construction, and JSON log output. |
| `Core/Core.Telemetry.psm1` | Durable `automation_events` telemetry and JSONL log routing helpers. |

### QC pipeline

| File | Purpose |
| --- | --- |
| `Workflow/QC.AuditTriggers.psm1` | Audit-trail-driven trigger helpers and workflow event hooks. |
| `Processing/QC.CommentExtract.psm1` | Invokes `tools/overlay/qc_pdf_comments.py`; normalized PDF comment annotations. |
| `Processing/QC.CommentStatusDecision.psm1` | Pure `Resolve-QCCommentWorkflowState` decision logic (no side effects). |
| `Processing/QC.CommentStatusProcessor.psm1` | Orchestrator for `QC_COMMENT_STATUS_SYNC` jobs. |
| `Processing/QC.CommentSync.Database.psm1` | `qc_comment_*` and related workflow event database writers. |
| `Processing/QC.CommentSync.Job.psm1` | Comment-sync job metadata and PW document resolution. |
| `Processing/QC.CommentSync.Notifications.psm1` | State-based notification routing for comment sync. |
| `Processing/QC.CommentSync.State.psm1` | Thin `Set-PWQCWorkflowState` wrapper for comment sync. |
| `Diagnostics/QC.DebugMcp.psm1` | Read-only QC workflow diagnostics for MCP and interactive debugging. |
| `Queue/QC.Filters.psm1` | Whitelist/blacklist filter evaluation for candidate paths before jobs are created. |
| `Queue/QC.HostThrottle.psm1` | Remote worker throttle policy based on configured host activity. |
| `Queue/QC.JobFactory.psm1` | Queue job ID/dedupe generation, payload construction, required-field validation, and dedupe-key generation. |
| `Notifications/QC.NotificationGraph.psm1` | Microsoft Graph email message construction for notifications. |
| `Notifications/QC.NotificationMock.psm1` | Mock notification transport for testing. |
| `Notifications/QC.NotificationTemplates.psm1` | HTML email template rendering (`ConvertTo-QCEmailHtml`). |
| `Notifications/QC.NotificationThreads.psm1` | Durable notification email threading by sheet package and review type. |
| `Notifications/QC.Notifications.psm1` | Configurable QC workflow email notifications (Mock + Graph). |
| `Processing/QC.PdfExport.psm1` | `Export-QCPdfToStaging` — ProjectWise PDF download to staging. |
| `Workflow/QC.ProcessType.psm1` | `QC_Process_Type` normalization, lane PDF resolution, and stamp configuration. |
| `Processing/QC.PrependOverlayPolicy.psm1` | Decides whether the incoming PDF requires a prepend overlay. |
| `Processing/QC.Processors.psm1` | Processor dispatch and implementations for QC prepend, status-set, comment sync, reporting, and related job types. |
| `Queue/QC.Queue.Json.psm1` | JSON-backed queue storage, locking, state transitions, stale-job recovery, dedupe checks, stats, recent-job reporting, and recon-slot succeeded/`_logs` retention. |
| `Processing/QC.Rendition.psm1` | ProjectWise PDF rendition request helpers. |
| `Reporting/QC.Reporting.psm1` | Read-only QC reporting aggregation and JSON snapshot generation for `QC_REPORTING_SCAN` jobs. |
| `Processing/QC.ReviewStamp.psm1` | QC review stamp processor integration. |
| `Processing/QC.StatusSet.psm1` | Native status-set implementation: manifest/cache, PW export/writeback, PDF merging, workspace reconciliation. |
| `Processing/QC.StatusSetPublication.psm1` | Dated ProjectWise archive copy of configured Status Set PDFs after the Monday 06:00 reconciliation slot. |
| `Queue/QC.Triggers.psm1` | Trigger rule ordering, matching, and job-type classification. |
| `Notifications/QC.WatcherAlerts.psm1` | Operational email alerts when the watcher loses ProjectWise connectivity. |
| `Core/QC.StatusSetBatching.psm1` | Status Set batch scheduling and queue gating. |
| `Core/QC.WatcherOrchestration.psm1` | Watcher tick orchestration shared by trigger scan entrypoints. |
| `Queue/QC.Worker.psm1` | Shared worker retry/state-transition policy for moving locked jobs between queue states. |
| `Workflow/QC.Workflow.psm1` | Optional QC workflow/state/attribute writeback framework for ProjectWise. |

### Ops

| File | Purpose |
| --- | --- |
| `Ops/QC.OpsConsole.psm1` | Logon-session pipeline status, scheduled-task on/off, queue inspect, status-set batch countdown, ops-request full scan. Does not start the dashboard. |

### ProjectWise

| File | Purpose |
| --- | --- |
| `ProjectWise/PW.AuditPoller.psm1` | ProjectWise audit trail polling, ingestion, and sheet-index telemetry. |
| `ProjectWise/PW.Connection.psm1` | ProjectWise session boundary: module loading, credentials, connect/disconnect, folder helpers, CLI diagnostics. |
| `ProjectWise/PW.Discovery.psm1` | Read-only PW/local discovery: watch paths, trigger candidates, metadata, folder/document listing, sheets expansion. |
| `ProjectWise/PW.Users.psm1` | ProjectWise user lookup helpers for notifications and workflow. |

### Documentation

| File | Purpose |
| --- | --- |
| `README.md` | Architectural notes for module responsibilities, exports, and contracts. |
| `FILES.md` | This file; quick purpose index (must list every `.psm1`). |
