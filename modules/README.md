# `modules/` — Module documentation

Production QC uses the **TYPSA three-lane model** (`Production` / `Review` / `Check` via `QC_Process_Type`, lane PDFs `*-prod/-rev/-chk.pdf`). See [`docs/workflow/qc-workflow-framework.md`](../docs/workflow/qc-workflow-framework.md).

## Folder layout (Phase 4E+)

Active module implementations live in responsibility-based subfolders (`Core/`, `Database/`, `ProjectWise/`, `Workflow/`, `Queue/`, `Processing/`, `Notifications/`, `Reporting/`, `Diagnostics/`).

**Phase 4F:** Production scripts, tests, and loaders import **folder implementation paths** (for example `modules/Core/Core.Results.psm1`, `modules/Queue/QC.Queue.Json.psm1`).

Flat paths such as `modules/QC.Queue.Json.psm1` remain as **silent compatibility shims** that forward `Import-Module` to the subfolder implementation. Do not remove until Phase 4H.

**Phase 4G prototype (test-only):** `modules/Core/QC.Core.psd1` and `modules/Queue/QC.Queue.psd1` bundle nested implementations via `NestedModules`. Production scripts still import folder `.psm1` paths directly.

Implementations live under `modules/<Folder>/<Name>.psm1`; see [`FILES.md`](FILES.md) for the full inventory.

All modules primarily communicate via a shared result envelope from `Core.Results.psm1`:

- Result shape: **`IsSuccess`**, **`Code`**, **`Message`**, **`Data`**

## Core modules

### `Core.Results.psm1`
- **Purpose**: standardized result constructors.
- **Exports**: `New-QCResult`, `New-QCSuccessResult`, `New-QCFailureResult`.

### `Core.Paths.psm1`
- **Purpose**: normalize paths (local + ProjectWise-ish `pw:\...`) so filters/triggers match consistently.
- **Exports**: `Normalize-QCPath`, `Normalize-QCDocumentsFolderPath`, `Normalize-QCPaths`, `Test-PathUnderRoot`, `Split-QCPathParts`.
- **Notable behavior**: converts `pw:\...\Documents\...` into canonical `Documents\...`, collapses slashes, trims trailing `\`, lowercases.

### `Core.Config.psm1`
- **Purpose**: config validation helpers and compatibility wrappers (`Read-AppConfig` delegates to `Read-QCAppSettings`).
- **Declared exports**: `Read-AppConfig`, `Test-AppSettings`, `Test-AppSettingsCompatibility`, `Get-QCProjectNameFromFolderPath`.

### `Core.Logging.psm1`
- **Purpose**: intended structured logging/audit façade.
- **Declared exports**: `Write-QCLog`, `Write-QCAudit`.

### `Core.Runtime.psm1`
- **Purpose**: script runtime helpers shared by entrypoints.
- **Exports**: `ConvertTo-HashtableDeep`, `Read-QCAppSettings`, `Get-QCAppSettingsConfig`, `Write-QCJsonLog`.

### `Core.Hashing.psm1`
- **Purpose**: reusable SHA-256 helpers.
- **Exports**: `Get-Sha256FileHex`, `Get-Sha256TextHex`.

## QC pipeline modules

### `QC.Filters.psm1`
- **Purpose**: allow/deny decisions via whitelist/blacklist before triggers/job creation.
- **Export**: `Test-QCPathAllowed(CandidatePath, Config)`.
- **Uses**: `Core.Paths` normalization and root containment.

### `QC.Triggers.psm1`
- **Purpose**: rule ordering + trigger evaluation (turn candidates into job types).
- **Exports**: `Get-OrderedTriggerRules`, `Test-TriggerRule`, `Test-QCTriggerCandidate`, `Resolve-QCTriggerMatch`.
- **Supports**: extension checks, filename/path regex, and “description contains” (PW metadata).

### `QC.JobFactory.psm1`
- **Purpose**: build validated job payloads, compute deterministic job IDs and dedupe keys.
- **Exports**: `New-QCJobId`, `New-QCJobObject`, `Test-QCJobRequiredFields`, `Get-QCDedupeKey`.
- **Notable behavior**:
  - `STATUS_SET_GEN` IDs/dedupe incorporate `folderStateHash` when present (folder content change → new identity).
  - `QC_PREPEND` dedupe can incorporate file hash when available.

### Sheet package model (production)

Production package grouping is **SQL-backed** via `Core.Database.psm1` (`sheet_packages`, `sheet_documents`, `sheet_package_qc_pdfs`). See [`docs/architecture/qc-package-model.md`](../docs/architecture/qc-package-model.md).

An earlier in-memory `QC.Package*` module cluster was archived in Phase 3 under [`archive/package-model-v1/`](../archive/package-model-v1/) and is not wired into the watcher, processor, or notification paths.

### `QC.Queue.Json.psm1`
- **Purpose**: JSON-backed durable queue + locks + recovery.
- **Layout** (under `queue.rootDir`): `pending/ running/ succeeded/ failed/ locks/`.
- **Key exports**:
  - enqueue/selection: `Add-QCQueueJob`, `Get-NextQCJob`
  - lifecycle: `Lock-QCJob`, `Unlock-QCJob`, `Move-QCJob`, `Get-QCJobById`, `Update-QCJob`
  - observability: `Get-QCQueueStats`, `Get-QCRecentJobs`
  - recovery: `Recover-QCStaleJobs` (requeue/fail stale running jobs, clean orphan locks)
- **Robustness**: atomic writes, PID-validated lock stealing, retryable moves (AV handle contention).

### `QC.Processors.psm1`
- **Purpose**: job readiness validation and job-type dispatch.
- **Exports**: `Test-QCJobReady`, `Invoke-QCProcessorByType`, `Invoke-QCPrependProcessor`, `Invoke-StatusSetProcessor`.
- **QC_PREPEND modes**:
  - `projectWise` / `pw` / `legacyPw` (alias): runs `scripts\processing\Invoke-QCPrependPw.ps1` (PW export + overlay + history) and can clear `QC_Archivist` tag on success.
  - local: updates history with `qpdf`; optional overlay via `dist\qc_overlay_prepend\qc_overlay_prepend.exe`.
- **STATUS_SET_GEN modes**:
  - `native`: imports `QC.StatusSet.psm1` and runs `Invoke-StatusSetNativeJob`
  - `legacy`: runs `legacy\combine_status_set.ps1`
  - `stub`: explicit no-op (only if configured)

### `QC.Reporting.psm1`
- **Purpose**: read-only QC reporting aggregation over existing `CADD/Sheets` folders using QC document attributes first and workflow states second.
- **Key exports**: `Get-QCReportingSettings`, `Get-QCReportingDocuments`, `ConvertTo-QCReportingDocument`, `New-QCReportingSnapshot`, `Write-QCReportingSnapshot`, `Invoke-QCReportingScan`, `New-QCReportingScanJob`.
- **Output**: JSON snapshots for `QC_REPORTING_SCAN` under `metrics/qc/<timestamp>/<project>.json`; database ingestion is intentionally deferred.

### `QC.StatusSet.psm1`
- **Purpose**: native Status Set generation and reconcile/writeback.
- **Core flow**: pair PDF sheets with matching CAD docs by base name → order → build `_StatusSet.pdf` → optional PW write-back.
- **Key exports**:
  - state: `Get-StatusSetLocalFolderState`, `Get-StatusSetPWFolderState`
  - workspace/manifest: `Get-StatusSetWorkspaceDirectory`, `Get-StatusSetManifestPath`, `Read-StatusSetManifestFile`, `Write-StatusSetManifestFile`, `New-StatusSetManifestObject`
  - decisions: `Test-StatusSetRebuildNeeded`, `Test-StatusSetWatcherShouldEnqueue`
  - operations: `Merge-StatusSetPdfWithQpdf`, `Export-StatusSetPdfToFolder`, `Invoke-StatusSetNativeJob`, `Sync-StatusSetWorkspaceToPw`, `Invoke-StatusSetReconcile`
- **Operational constraints handled**: AV-friendly throttling for file ops and PW export; manifest-driven sheet-cache reuse; staged/atomic `_StatusSet.pdf` replacement with `_history`; retention-based staging cleanup instead of per-job scratch deletion.

### Comment-status sync (`QC_COMMENT_STATUS_SYNC`)

- **`QC.PdfExport.psm1`**: `Export-QCPdfToStaging` — PW PDF download to staging.
- **`QC.CommentExtract.psm1`**: invokes `tools/overlay/qc_pdf_comments.py`; normalized annotations only.
- **`QC.CommentStatusDecision.psm1`**: pure `Resolve-QCCommentWorkflowState` (no side effects).
- **`QC.CommentSync.Database.psm1`**: `qc_comment_*` + `qc_workflow_events` writers.
- **`QC.CommentSync.State.psm1`**: thin `Set-PWQCWorkflowState` wrapper.
- **`QC.CommentSync.Notifications.psm1`**: state-based notification routing.
- **`QC.CommentStatusProcessor.psm1`**: thin orchestrator (`Invoke-QCCommentStatusSyncProcessor`).
- **`QC.CommentSync.Job.psm1`**: job metadata + `Get-QCCommentSyncPwDocument`.
- See `docs/workflow/qc-comment-status-sync.md`.

### `QC.Notifications.psm1`
- **Purpose**: configurable QC workflow email notifications (Mock + Microsoft Graph).
- **Exports**: `Invoke-QCNotificationForStateChange`, `Send-QCNotification`, `Resolve-QCNotificationRecipients`, `Resolve-QCNotificationQcPdfUrl`, and related helpers.
- **Related**: `QC.NotificationTemplates.psm1` (`ConvertTo-QCEmailHtml`), `QC.NotificationMock.psm1`, `QC.NotificationGraph.psm1` (`New-QCGraphEmailMessage`). See `docs/workflow/qc-notifications.md`.

### `QC.Worker.psm1`
- **Purpose**: reusable worker retry/transition policy.
- **Exports**: `Move-QCJobWithLockRetries`.

## ProjectWise port modules

### `PW.Connection.psm1`
- **Purpose**: ProjectWise connection wrapper + helper utilities used by watcher/status-set.
- **Key exports**: `Get-PWCredentialFromFile`, `Connect-PW`, `Disconnect-PW`, `Get-PWImmediateChildFolders`.
- **Requires**: pwps_dab cmdlets (run in ProjectWise PowerShell).

### `PW.Discovery.psm1`
- **Purpose**: discovery interface for PW candidates plus shared ProjectWise metadata/folder traversal helpers used by the watcher.
- **Key exports**: `Resolve-WatchPaths`, `Get-PWTriggerCandidates`, `Get-PWCandidateMetadata`, `Get-PWDocName`, `Get-PWDocDescription`, `Get-PWDocLastModifiedUtc`, `ConvertTo-PWCmdletFolderPath`, `ConvertTo-PWCanonicalDocumentsFolderPath`, `Get-PWDocumentsInFolder`, `Find-PWSheetsFoldersUnderRoot`.
