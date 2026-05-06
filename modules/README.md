# `modules/` — Module documentation

All modules primarily communicate via a shared result envelope from `Core.Results.psm1`:

- Result shape: **`IsSuccess`**, **`Code`**, **`Message`**, **`Data`**

## Core modules

### `Core.Results.psm1`
- **Purpose**: standardized result constructors.
- **Exports**: `New-QCResult`, `New-QCSuccessResult`, `New-QCFailureResult`.

### `Core.Paths.psm1`
- **Purpose**: normalize paths (local + ProjectWise-ish `pw:\...`) so filters/triggers match consistently.
- **Exports**: `Normalize-QCPath`, `Normalize-QCPaths`, `Test-PathUnderRoot`, `Split-QCPathParts`.
- **Notable behavior**: converts `pw:\...\Documents\...` into canonical `Documents\...`, collapses slashes, trims trailing `\`, lowercases.

### `Core.Config.psm1`
- **Purpose**: intended config loader/validator façade.
- **Declared exports**: `Read-AppConfig`, `Test-AppSettings`, `Get-AppSetting`, `Test-AppSettingsCompatibility`.

### `Core.Runtime.psm1`
- **Purpose**: script runtime helpers shared by entrypoints.
- **Exports**: `ConvertTo-HashtableDeep`, `Read-QCAppSettings`, `Get-QCAppSettingsConfig`, `Write-QCJsonLog`.

### `Core.Hashing.psm1`
- **Purpose**: reusable SHA-256 helpers.
- **Exports**: `Get-Sha256FileHex`, `Get-Sha256TextHex`.

### `Core.Logging.psm1`
- **Purpose**: intended structured logging/audit façade.
- **Declared exports**: `Write-QCLog`, `Write-QCAudit`.

### `Core.Metrics.psm1`
- **Purpose**: intended counters/timing metrics layer.
- **Declared exports**: `Initialize-QCMetrics`, `Get-QCMetricsSnapshot`, `Update-QCMetricsOn*`, `Reset-QCMetricsDaily`, `Flush-QCMetrics`.

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
  - `legacyPw`: runs `legacy\prepend_qc.ps1` (PW export + overlay + history) and can clear `QC_Archivist` tag on success.
  - local: updates history with `qpdf`; optional overlay via `dist\qc_overlay_prepend\qc_overlay_prepend.exe`.
- **STATUS_SET_GEN modes**:
  - `native`: imports `QC.StatusSet.psm1` and runs `Invoke-StatusSetNativeJob`
  - `legacy`: runs `legacy\combine_status_set.ps1`
  - `stub`: explicit no-op (only if configured)

### `QC.StatusSet.psm1`
- **Purpose**: native Status Set generation and reconcile/writeback.
- **Core flow**: pair PDF sheets with matching CAD docs by base name → order → build `_StatusSet.pdf` → optional PW write-back.
- **Key exports**:
  - state: `Get-StatusSetLocalFolderState`, `Get-StatusSetPWFolderState`
  - workspace/manifest: `Get-StatusSetWorkspaceDirectory`, `Get-StatusSetManifestPath`, `Read-StatusSetManifestFile`, `Write-StatusSetManifestFile`, `New-StatusSetManifestObject`
  - decisions: `Test-StatusSetRebuildNeeded`, `Test-StatusSetWatcherShouldEnqueue`
  - operations: `Merge-StatusSetPdfWithQpdf`, `Export-StatusSetPdfToFolder`, `Invoke-StatusSetNativeJob`, `Sync-StatusSetWorkspaceToPw`, `Invoke-StatusSetReconcile`
- **Operational constraints handled**: AV-friendly throttling for file ops and PW export; manifest-driven sheet-cache reuse; staged/atomic `_StatusSet.pdf` replacement with `_history`; retention-based staging cleanup instead of per-job scratch deletion.

### `QC.Notifications.psm1`
- **Purpose**: placeholder notification routing for job events.
- **Declared export**: `Send-QCNotification(Event, Job, Config)`.

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

## Orchestration

### `Orchestrator.Pipeline.psm1`
- **Purpose**: composable pipeline/worker ticks that call “port” functions and enforce QCResult shape.
- **Exports**: `Invoke-QCPipelineTick`, `Invoke-QCWorkerTick`.

