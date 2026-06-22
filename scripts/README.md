# `scripts/` — Operational entrypoints

This folder contains the runnable PowerShell entrypoints for the QC pipeline.

## Primary entrypoints

### `Start-QCPipelineDashboard.ps1`
- **Purpose**: unified pipeline runner + live terminal dashboard.
- **Runs**: `Watch-QCTrigger.ps1` (enqueue) + `Run-QCProcessor.ps1` (dequeue/process) concurrently.
- **Key behaviors**:
  - Production dashboard view by default (`-DashboardView Production`) that shows only critical health, queue, watcher, worker, and warning/error information.
  - Detailed operations view remains available with `-DashboardView Detailed` for recent job tables, scan paths, and processor activity.
  - Singleton guard via `queueRoot\_dashboard.lock` (prevents multiple dashboards spawning too many processes).
  - Persistent child logs under `queueRoot\_logs\` (stdout/stderr) for post-mortem when AV kills processes.
  - Periodic `Recover-QCStaleJobs` to requeue orphaned `running\` jobs.

### `Watch-QCTrigger.ps1`
- **Purpose**: one-shot watcher tick (detect → filter → trigger → job → dedupe → enqueue).
- **Sources**:
  - ProjectWise watch list: `projectWise.watchList` in `appsettings.json`
  - Local folders: `watchFolders` in `appsettings.json`
- **Outputs**: JSON log lines to stdout (ingested by the dashboard).
- **Notes**:
  - Can reconcile local status sets back to PW on startup via `-ReconcileStatusSetsFirst`.
  - Uses `Test-StatusSetWatcherShouldEnqueue` to skip `STATUS_SET_GEN` enqueue when already current.

### `Run-QCProcessor.ps1`
- **Purpose**: worker loop (select → lock → dispatch → move state).
- **Modes**:
  - one-shot: default (`-MaxJobs 1`)
  - long-running: use `-MaxJobs`, `-LeaseSeconds`, and/or `-IdleSleepMs`
- **Dispatch**: `modules/QC.Processors.psm1` (`QC_PREPEND`, `STATUS_SET_GEN`).
- **Observability**: logs JSON lines like `WORKER_SELECTED`, `WORKER_SUCCEEDED`, `WORKER_FAILED`.

### `run_prepend_qc.ps1`
- **Purpose**: launcher wrapper.
- **Default**: runs the dashboard unless `-NoDashboard` is provided.
- **Legacy**: supports `-Legacy` to run the older monolithic workflow.

## Diagnostics and maintenance

Diagnostic and discovery scripts live under `scripts/diagnostics/`. Compatibility wrappers at the former `scripts/*.ps1` paths forward to the new locations (Phase 4C).

### `diagnostics/Show-QCStatus.ps1` (wrapper: `Show-QCStatus.ps1`)
- **Purpose**: read-only queue snapshot (counts + recent jobs).

### `diagnostics/Show-QCQueueDiag.ps1` (wrapper: `Show-QCQueueDiag.ps1`)
- **Purpose**: deeper queue health view (running ages, lock owner PID liveness, orphan analysis).

### `diagnostics/Test-PWConnection.ps1` (wrapper: `Test-PWConnection.ps1`)
- **Purpose**: connect/disconnect smoke test using `appsettings.json` credentials.

See `scripts/diagnostics/` for ProjectWise browse/list/smoke probes, metrics scripts (`Get-PWFolderStateCounts`, `Scan-PWProjectMetrics`), and `Test-PWEmailAttributes*` / `Test-QC*` smoke scripts.

### `Stop-QCPipeline.ps1`
- **Purpose**: kill dashboard/watcher/worker PowerShell processes (useful for cleaning stale runs).

### `Reset-QCFolderWorkflow.ps1`
- **Purpose**: reset PW workflow states + clear folder-scoped QC telemetry for a clean prepend cycle.
- **Lane PDF recycle**: after manually deleting `*-prod/-chk/-rev.pdf` in PW, run with `-ConfirmReset` to delete lane
  `sheet_index` ghosts, `sheet_documents` `qc_pdf` rows, and `sheet_package_qc_pdfs` (default). Use `-KeepLanePdfRegistry`
  for legacy UPDATE-only behavior.

### `Publish-QCPipelineCode.ps1`
- **Purpose**: copy `modules/` and key scripts to the production worker root; optional `-ConfirmRestart`.

### Other helper scripts

This repo also includes queue/status-set helpers such as `Requeue-QCJobs.ps1`, `Purge-QCPendingByFilters.ps1`,
and status-set wrappers (`Combine-StatusSet.ps1`, `Run-CombineStatusSet.ps1`, `Reconcile-QCStatusSets.ps1`).

