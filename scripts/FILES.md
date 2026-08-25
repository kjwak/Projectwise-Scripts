# Scripts directory file guide

The PowerShell files in this directory are intended to stay thin: they parse command-line parameters, load the relevant modules from `../modules`, and delegate reusable behavior to module functions. If a script starts to need substantial branching, queue mutation, ProjectWise compatibility logic, or PDF/status-set behavior, that logic should move into a module first and the script should remain a launcher.

## Refactoring assessment completed

- `PW-BrowseFolder.ps1` was simplified so credential loading, ProjectWise module loading, connection cleanup, folder-view compatibility, and output shaping are provided by `modules/ProjectWise/PW.Connection.psm1`.
- `PW-ListDocsInFolder.ps1` was simplified the same way; it now delegates connection/session handling and document enumeration to module functions.
- The remaining larger entrypoints (`Watch-QCTrigger.ps1`, `Run-QCProcessor.ps1`, and `Start-QCPipelineDashboard.ps1`) already depend heavily on modules but still contain orchestration/UI code. The next simplification target should be moving dashboard rendering and watcher ProjectWise expansion into module functions while keeping the scripts as parameter-only launchers.

## Diagnostics folder (Phase 4C)

Read-only probes, smoke tests, and queue/PW inspection scripts are implemented under `diagnostics/`. The same filenames remain at the `scripts/` root as silent compatibility wrappers that forward arguments and exit codes.

| Location | Examples |
| --- | --- |
| `diagnostics/` | `Show-QCStatus`, `Show-QCQueueDiag`, `Get-PWFolderStateCounts`, `Scan-PWProjectMetrics`, `PW-BrowseFolder`, `Test-PWConnection`, `Test-PWEmailAttributes*`, `Test-QC*` |

`tools/discovery/` is unchanged in Phase 4C (deferred).

## Maintenance folder (Phase 4D)

Queue, database, sheet-index, and ProjectWise operator recovery scripts are implemented under `maintenance/`. The same filenames remain at the `scripts/` root as silent compatibility wrappers.

| Location | Examples |
| --- | --- |
| `maintenance/` | `Reset-QCFolderWorkflow`, `Purge-QCPendingByFilters`, `Requeue-QCJobs`, `Repair-*`, `Remove-*`, `Invoke-QCDatabaseRetention`, `Sync-*`, `Refresh-*`, `Reconcile-*` |

`Import-QCScriptModules.ps1` stays at `scripts/` (dot-sourced helper). `Stop-QCPipeline.ps1` stays at `scripts/` (pipeline stop helper).

## Processing folder

Status-set and related processing helpers are implemented under `processing/`. The same filenames remain at the `scripts/` root as silent compatibility wrappers.

| Location | Examples |
| --- | --- |
| `processing/` | `Combine-StatusSet`, `Run-CombineStatusSet`, `Invoke-QCPrependPw` |

## Deployment folder

Low-risk deployment helpers are implemented under `deployment/`. The same filenames remain at the `scripts/` root as silent compatibility wrappers.

| Location | Examples |
| --- | --- |
| `deployment/` | `Promote-DevToMain`, `Sync-OverlayReviewStamp`, `Register-QCPipelineDashboardTask`, `Register-QCPipelineDashboardLogonConsole`, `Register-QCRemoteWorkerHostTask` |

`Publish-QCPipelineCode.ps1` remains at `scripts/` root (high-risk publish entrypoint; not moved in Phase 4 processing/deployment move).

## File purposes

| File | Purpose |
| --- | --- |
| `Combine-StatusSet.ps1` | Compatibility wrapper → `processing/Combine-StatusSet.ps1` (manual `_StatusSet.pdf` build via `QC.StatusSet.psm1`). |
| `PW-BrowseFolder.ps1` | Compatibility wrapper → `diagnostics/PW-BrowseFolder.ps1` (read-only ProjectWise folder browser). |
| `PW-ListDocsInFolder.ps1` | Compatibility wrapper → `diagnostics/PW-ListDocsInFolder.ps1`. |
| `PW-SmokeTest.ps1` | Compatibility wrapper → `diagnostics/PW-SmokeTest.ps1`. |
| `PW-TestWatchRoots.ps1` | Compatibility wrapper → `diagnostics/PW-TestWatchRoots.ps1`. |
| `Purge-QCPendingByFilters.ps1` | Compatibility wrapper → `maintenance/Purge-QCPendingByFilters.ps1`. |
| `Remove-QCAuditEvents.ps1` | Compatibility wrapper → `maintenance/Remove-QCAuditEvents.ps1`. |
| `Remove-QCWorkflowEvents.ps1` | Compatibility wrapper → `maintenance/Remove-QCWorkflowEvents.ps1`. |
| `Invoke-QCDatabaseRetention.ps1` | Compatibility wrapper → `maintenance/Invoke-QCDatabaseRetention.ps1`. |
| `Reconcile-QCStatusSets.ps1` | Compatibility wrapper → `maintenance/Reconcile-QCStatusSets.ps1`. |
| `Requeue-QCJobs.ps1` | Compatibility wrapper → `maintenance/Requeue-QCJobs.ps1`. |
| `Promote-DevToMain.ps1` | Compatibility wrapper → `deployment/Promote-DevToMain.ps1` (git dev→main promote helper). |
| `Register-QCPipelineDashboardTask.ps1` | QC-server boot task for `Start-QCPipelineDashboard.ps1` (Session 0; lives in `deployment/` only). |
| `Register-QCPipelineDashboardLogonConsole.ps1` | QC-server Startup shortcut → `Start-QCOpsConsole.ps1` (WinForms; `-NoGui` text tail). |
| `Start-QCOpsConsole.ps1` | Logon ops GUI; observes/controls `QC-PipelineDashboard`; never starts a second dashboard. |
| `Invoke-QCOpsPwCompare.ps1` | MTA child helper for SQL vs ProjectWise compare (GUI cannot host pwps_dab). |
| `Watch-QCPipelineDashboardConsole.ps1` | Read-only tail of watcher/worker logs; does not start the dashboard. |
| `Register-QCRemoteWorkerHostTask.ps1` | Modelling-PC boot task for `Start-QCRemoteWorkerHost.ps1` (lives in `deployment/` only). |
| `Run-CombineStatusSet.ps1` | Compatibility wrapper → `processing/Run-CombineStatusSet.ps1` (launcher reading `statusSet.localRoot` from appsettings). |
| `Run-QCProcessor.ps1` | Worker entrypoint that claims pending queue jobs, dispatches processors, renews locks, records logs, and transitions job states. |
| `Show-QCQueueDiag.ps1` | Compatibility wrapper → `diagnostics/Show-QCQueueDiag.ps1`. |
| `Show-QCStatus.ps1` | Compatibility wrapper → `diagnostics/Show-QCStatus.ps1`. |
| `Start-QCPipelineDashboard.ps1` | Unified dashboard entrypoint that runs watcher and worker processes while rendering live terminal status. |
| `Stop-QCPipeline.ps1` | Safety/operations helper that finds and stops PowerShell processes running this repo's watcher, worker, or dashboard scripts. Does not match the logon ops GUI. |
| `Sync-OverlayReviewStamp.ps1` | Compatibility wrapper → `deployment/Sync-OverlayReviewStamp.ps1` (sync overlay Python into PyInstaller bundle). |
| `Test-PWConnection.ps1` | Compatibility wrapper → `diagnostics/Test-PWConnection.ps1`. |
| `Watch-QCTrigger.ps1` | One-shot trigger scan that discovers local/ProjectWise candidates, evaluates rules and filters, and enqueues queue jobs. |
| `README.md` | Existing overview and usage notes for script entrypoints. |
| `FILES.md` | This file; quick purpose index and refactoring assessment for the scripts directory. |
| `Import-QCScriptModules.ps1` | Dot-sourced by DB cleanup scripts to load Core.* modules into global scope. |
| `run_prepend_qc.ps1` | Compatibility launcher for the modular watcher/worker pipeline, with optional legacy monolith mode. |
