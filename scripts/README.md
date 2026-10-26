# `scripts/` — Operational entrypoints

This folder contains the runnable PowerShell entrypoints for the QC pipeline.

## Primary entrypoints

### `Start-QCPipelineDashboard.ps1`
- **Purpose**: unified pipeline runner + live terminal dashboard.
- **Runs**: `Watch-QCTrigger.ps1` (enqueue) + `Run-QCProcessor.ps1` (dequeue/process) concurrently.
- **Host**: QC server only. Do not run this on the modelling PC.
- **Key behaviors**:
  - Production dashboard view by default (`-DashboardView Production`) that shows only critical health, queue, watcher, worker, and warning/error information.
  - Detailed operations view remains available with `-DashboardView Detailed` for recent job tables, scan paths, and processor activity.
  - Singleton guard via `queueRoot\_dashboard.lock` (prevents multiple dashboards spawning too many processes).
  - Persistent child logs under `queueRoot\_logs\` (stdout/stderr) for post-mortem when AV kills processes. Files older than `queue.retention.logsHours` (default 24) are deleted at the start of each reconciliation slot.
  - Periodic `Recover-QCStaleJobs` to requeue orphaned `running\` jobs.
- **Boot task (QC server only):** `scripts/deployment/Register-QCPipelineDashboardTask.ps1` registers `QC-PipelineDashboard` at startup, whether the user is logged on or not (Session 0, no TUI). On `PXBENTLEY01` it defaults to `C:\Users\jflint\Documents\github\Prepend PDF QC`. It also grants the run-as account a Full Control ACE so the logon ops GUI can toggle Pipeline on/off without elevation. For a task that already exists, re-run elevated with `-GrantOperatorAccess` (no-ops only when a Full Control ACE is already present).
- **Logon console:** `scripts/deployment/Register-QCPipelineDashboardLogonConsole.ps1` (Startup shortcut → `Start-QCOpsConsole.ps1` WinForms; `-NoGui` → `Watch-QCPipelineDashboardConsole.ps1` text tail). Does not start a second dashboard. Do not register this on the modelling PC.

### `Start-QCOpsConsole.ps1`
- **Purpose**: logon-session WinForms console (STA) on the QC server.
- **Does not** start `Start-QCPipelineDashboard.ps1`. Close anytime; Session 0 keeps running.
- **On/off**: Task Scheduler COM Enable/Disable of `QC-PipelineDashboard` (honors the task DACL), then `Stop-QCPipeline.ps1` (GUI process names are excluded).
- **Full scan while live**: writes `queue/_watcher/ops-request.json`; the live watcher honors it.
- **Host**: `PXBENTLEY01` only (`-Force` to override).

### `Watch-QCPipelineDashboardConsole.ps1`
- **Purpose**: read-only log tail for the QC server boot task.
- **Does not** start watcher, workers, or a second dashboard. Close anytime.

### `Start-QCRemoteWorkerHost.ps1`
- **Purpose**: processor-only supervisor for a remote worker host.
- **Runs**: one or more `Run-QCProcessor.ps1` children. Does **not** start the watcher or dashboard.
- **Job types**: honors `workers.enabledJobTypes` (empty/omitted = all types).
- **UNC**: refuses `\\server\share` queue roots unless `-AllowUncQueue` or `workers.remoteHost.allowUncQueue`.
- **Stop**: `Stop-QCPipeline.ps1` (also matches this supervisor).
- **Boot task**: `scripts/deployment/Register-QCRemoteWorkerHostTask.ps1 -AllowUncQueue` (runs at startup when user is logged off).
- **Logon console**: `scripts/deployment/Register-QCRemoteWorkerHostLogonConsole.ps1` (Startup shortcut → read-only log tail).
- **Console**: in-flight work comes from `queueRoot\running` for this host (`machineName`). JSONL is stage/success/failure only. Heartbeat shows `idle` or `busy=N <document>`.

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
- **Dispatch**: `modules/Processing/QC.Processors.psm1` (`QC_PREPEND`, `STATUS_SET_GEN`).
- **Job types**: `workers.enabledJobTypes` allow-list (empty/omitted = all). UNC queue roots require `-AllowUncQueue`.
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
- **Purpose**: deeper queue health view (running ages, lock owner host/PID liveness, orphan analysis, bounded Warning/Error log scan).
- **Locations**: uses `Resolve-QCDebugLocations.ps1` (env / `config/qc-debug-locations.local.json` / default `\\192.168.22.90\QC_Queue`); does not default to production `queue.rootDir`. Pass `-NoLogs` to skip log scan; `-UseAppSettingsQueueRoot` for old local-queue debugging.

### `diagnostics/Test-PWConnection.ps1` (wrapper: `Test-PWConnection.ps1`)
- **Purpose**: connect/disconnect smoke test using `appsettings.json` credentials.

See `scripts/diagnostics/` for ProjectWise browse/list/smoke probes, metrics scripts (`Get-PWFolderStateCounts`, `Scan-PWProjectMetrics`), and `Test-PWEmailAttributes*` / `Test-QC*` smoke scripts.

Maintenance and operator recovery scripts live under `scripts/maintenance/`. Compatibility wrappers at the former `scripts/*.ps1` paths forward to the new locations (Phase 4D).

### `maintenance/Reset-QCFolderWorkflow.ps1` (wrapper: `Reset-QCFolderWorkflow.ps1`)
- **Purpose**: reset PW workflow states + clear folder-scoped QC telemetry for a clean prepend cycle.
- **Lane PDF recycle**: after manually deleting `*-prod/-chk/-rev.pdf` in PW, run with `-ConfirmReset` to delete lane
  `sheet_index` ghosts, `sheet_documents` `qc_pdf` rows, and `sheet_package_qc_pdfs` (default). Use `-KeepLanePdfRegistry`
  for legacy UPDATE-only behavior.

### `Stop-QCPipeline.ps1`
- **Purpose**: kill dashboard/watcher/worker PowerShell processes (useful for cleaning stale runs). Does **not** match `Start-QCOpsConsole.ps1`, `Watch-QCPipelineDashboardConsole.ps1`, or `Invoke-QCOpsPwCompare.ps1`.

See `scripts/maintenance/` for queue purge/requeue/repair, database retention/removal, sheet-index sync/reconcile, and related operator tools (`Purge-QCPendingByFilters`, `Requeue-QCJobs`, `Invoke-QCDatabaseRetention`, etc.).

## Processing and deployment helpers

Status-set processing scripts live under `scripts/processing/`. Deployment helpers live under `scripts/deployment/`. Compatibility wrappers at the former `scripts/*.ps1` paths forward to the new locations.

### `processing/Combine-StatusSet.ps1` (wrapper: `Combine-StatusSet.ps1`)
- **Purpose**: manual `_StatusSet.pdf` build/refresh via `QC.StatusSet.psm1`.

### `processing/Run-CombineStatusSet.ps1` (wrapper: `Run-CombineStatusSet.ps1`)
- **Purpose**: launcher that reads `statusSet.localRoot` from `appsettings.json` and forwards to `Combine-StatusSet.ps1`.

### `deployment/Promote-DevToMain.ps1` (wrapper: `Promote-DevToMain.ps1`)
- **Purpose**: promote `dev` to `main` and push (git workflow helper).

### `deployment/Register-QCPipelineDashboardTask.ps1`
- **Purpose**: QC server only. Registers `QC-PipelineDashboard` to run `Start-QCPipelineDashboard.ps1` at startup whether the user is logged on or not. Defaults to the `Prepend PDF QC` clone on `PXBENTLEY01`. Grants the run-as account task ACL for unelevated Enable/Disable. Existing task: `-GrantOperatorAccess`.

### `deployment/Set-QCScheduledTaskOperatorAcl.ps1`
- **Purpose**: Elevated helper that adds a Full Control ACE on one scheduled task for a Windows account. Called by the dashboard registrar; can be run standalone. No-ops only when that account already has FA/GA on the DACL (owner SID or Read is not enough).

### `deployment/Register-QCPipelineDashboardLogonConsole.ps1`
- **Purpose**: QC server only. Startup shortcut that opens `Start-QCOpsConsole.ps1` (`-STA`). Does not start a second dashboard. Text tail: `Start-QCOpsConsole.ps1 -NoGui`.

### `deployment/Register-QCRemoteWorkerHostTask.ps1`
- **Purpose**: modelling PC only. Registers `QC-RemoteWorkerHost` to run `Start-QCRemoteWorkerHost.ps1` at startup with `-AllowUncQueue`.

### `deployment/Sync-OverlayReviewStamp.ps1` (wrapper: `Sync-OverlayReviewStamp.ps1`)
- **Purpose**: sync `tools/overlay/qc_review_stamp.py` into the PyInstaller onedir bundle.

### `Publish-QCPipelineCode.ps1`
- **Purpose**: copy `modules/` and key scripts to the production worker root; optional `-ConfirmRestart`.

`Import-QCScriptModules.ps1` is dot-sourced by maintenance scripts under `scripts/maintenance/`. `Restore-QCModuleExports.ps1` is dot-sourced by production entrypoints for PS 5.1 export restore.

