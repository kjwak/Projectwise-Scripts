# Scripts directory file guide

The PowerShell files in this directory are intended to stay thin: they parse command-line parameters, load the relevant modules from `../modules`, and delegate reusable behavior to module functions. If a script starts to need substantial branching, queue mutation, ProjectWise compatibility logic, or PDF/status-set behavior, that logic should move into a module first and the script should remain a launcher.

## Refactoring assessment completed

- `PW-BrowseFolder.ps1` was simplified so credential loading, ProjectWise module loading, connection cleanup, folder-view compatibility, and output shaping are provided by `modules/PW.Connection.psm1`.
- `PW-ListDocsInFolder.ps1` was simplified the same way; it now delegates connection/session handling and document enumeration to module functions.
- The remaining larger entrypoints (`Watch-QCTrigger.ps1`, `Run-QCProcessor.ps1`, and `Start-QCPipelineDashboard.ps1`) already depend heavily on modules but still contain orchestration/UI code. The next simplification target should be moving dashboard rendering and watcher ProjectWise expansion into module functions while keeping the scripts as parameter-only launchers.

## File purposes

| File | Purpose |
| --- | --- |
| `Combine-StatusSet.ps1` | Manual wrapper that builds or refreshes `_StatusSet.pdf` for a specified sheets folder by calling `Invoke-StatusSetNativeJob` in `QC.StatusSet.psm1`. |
| `PW-BrowseFolder.ps1` | Read-only ProjectWise diagnostic that prints immediate child folders and documents for a folder path through `Show-PWFolderBrowser`. |
| `PW-ListDocsInFolder.ps1` | Read-only ProjectWise diagnostic that prints document/PDF counts and a sample of PDFs through `Show-PWFolderDocumentList`. |
| `PW-SmokeTest.ps1` | End-to-end local environment check for ProjectWise connection, document search, history lookup, and optional export assumptions. |
| `PW-TestWatchRoots.ps1` | Diagnostic that expands configured ProjectWise watch-list roots and shows discovered sheets folders/candidates. |
| `Purge-QCPendingByFilters.ps1` | Queue maintenance utility that re-evaluates pending jobs against current filters and moves disallowed jobs to failed. |
| `Remove-QCAuditEvents.ps1` | Database maintenance: preview/delete aged `audit_events` rows (batched, processed-only by default). |
| `Remove-QCWorkflowEvents.ps1` | Database maintenance: preview/delete `qc_workflow_events` by folder path fragment(s) (not scheduled retention). |
| `Invoke-QCDatabaseRetention.ps1` | Scheduled `audit_events` retention only (`database.retention` in appsettings). |
| `Reconcile-QCStatusSets.ps1` | Startup/catch-up utility that scans local status-set manifests and reconciles generated `_StatusSet.pdf` files back to ProjectWise. |
| `Requeue-QCJobs.ps1` | Queue maintenance utility that moves matching succeeded or failed jobs back to pending for reprocessing. |
| `Run-CombineStatusSet.ps1` | Convenience launcher that reads `statusSet.localRoot` from `appsettings.json` and forwards to `Combine-StatusSet.ps1`. |
| `Run-QCProcessor.ps1` | Worker entrypoint that claims pending queue jobs, dispatches processors, renews locks, records logs, and transitions job states. |
| `Show-QCQueueDiag.ps1` | Queue diagnostic snapshot showing state counts, running-job age, lock ownership, and watcher-active sentinel status. |
| `Show-QCStatus.ps1` | Human-readable queue status summary with state counts and recent jobs. |
| `Start-QCPipelineDashboard.ps1` | Unified dashboard entrypoint that runs watcher and worker processes while rendering live terminal status. |
| `Stop-QCPipeline.ps1` | Safety/operations helper that finds and stops PowerShell processes running this repo's watcher, worker, or dashboard scripts. |
| `Test-PWConnection.ps1` | Read-only smoke test for loading app settings, reading ProjectWise credentials, connecting, verifying discovery cmdlets, and disconnecting. |
| `Watch-QCTrigger.ps1` | One-shot trigger scan that discovers local/ProjectWise candidates, evaluates rules and filters, and enqueues queue jobs. |
| `README.md` | Existing overview and usage notes for script entrypoints. |
| `FILES.md` | This file; quick purpose index and refactoring assessment for the scripts directory. |
| `run_prepend_qc.ps1` | Compatibility launcher for the modular watcher/worker pipeline, with optional legacy monolith mode. |
