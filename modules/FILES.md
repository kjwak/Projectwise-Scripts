# Modules directory file guide

Modules own the reusable pipeline logic. Scripts should call these functions instead of duplicating credential parsing, ProjectWise compatibility handling, queue transitions, status-set generation, trigger evaluation, or processor dispatch.

## Refactoring assessment completed

- ProjectWise command-line diagnostics now reuse `PW.Connection.psm1` for key/value credential parsing, ProjectWise module loading, connection cleanup, folder-view compatibility, child splitting, and document enumeration.
- This reduces duplicate script-local credential and `Get-PWFolderView` compatibility logic and makes future ProjectWise diagnostics safer to implement as thin wrappers.
- Future cleanup should continue moving long-lived orchestration/UI behavior from scripts into modules: dashboard rendering can become a dashboard module, and watcher ProjectWise expansion can be consolidated behind `PW.Discovery.psm1`/`Orchestrator.Pipeline.psm1` ports.

## File purposes

| File | Purpose |
| --- | --- |
| `Core.Config.psm1` | General appsettings loading, setting access, and compatibility validation helpers. |
| `Core.Hashing.psm1` | Shared SHA-256 helpers for file and text hashing. |
| `Core.Logging.psm1` | Structured application/audit logging primitives. |
| `Core.Metrics.psm1` | Queue and processing metrics initialization, updates, reset, and flush helpers. |
| `Core.Paths.psm1` | Path normalization, root-containment checks, and path splitting for local and ProjectWise-style paths. |
| `Core.Results.psm1` | Standard QC result constructors used by modules to return consistent success/failure objects. |
| `Core.Runtime.psm1` | Runtime helpers for deep hashtable conversion, appsettings loading, merged config construction, and JSON log output. |
| `Orchestrator.Pipeline.psm1` | Testable orchestration functions that compose discovery, trigger evaluation, enqueueing, worker selection, processing, and queue transitions. |
| `PW.Connection.psm1` | ProjectWise session boundary: module loading, credential loading, connect/disconnect, discovery-cmdlet checks, child-folder helpers, and CLI diagnostics for browsing/listing folders. |
| `PW.Discovery.psm1` | Read-only ProjectWise/local discovery: watch path resolution, trigger candidate construction, metadata extraction, folder/document listing, and sheets-folder expansion. |
| `QC.Filters.psm1` | Whitelist/blacklist filter evaluation for candidate paths before jobs are created. |
| `QC.JobFactory.psm1` | Queue job ID/dedupe generation, payload construction, required-field validation, and dedupe-key generation. |
| `QC.Notifications.psm1` | Notification dispatch placeholder for workflow events. |
| `QC.Processors.psm1` | Processor dispatch and implementations for QC prepend and status-set generation jobs, including readiness checks and external tool invocation. |
| `QC.Queue.Json.psm1` | JSON-backed queue storage, locking, state transitions, stale-job recovery, dedupe checks, stats, and recent-job reporting. |
| `QC.StatusSet.psm1` | Native status-set implementation: manifest/cache management, ProjectWise export/writeback, PDF merging, workspace reconciliation, and related helpers. |
| `QC.Triggers.psm1` | Trigger rule ordering, matching, and job-type classification. |
| `QC.Worker.psm1` | Shared worker retry/state-transition policy for moving locked jobs between queue states. |
| `README.md` | Existing architectural notes for module responsibilities and contracts. |
| `FILES.md` | This file; quick purpose index and refactoring assessment for the modules directory. |
