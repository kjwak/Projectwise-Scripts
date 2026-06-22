# Phase 4D — Maintenance script move summary

**Branch:** `phase-4/maintenance-scripts`  
**Base branch:** `phase-4/integration`  
**Status:** Implementation complete — **not committed** (awaiting review)

## Goal

Path-organization only: move maintenance and operator recovery scripts into `scripts/maintenance/` while preserving operator-facing paths via silent wrappers. No changes to maintenance behavior, queue logic, database logic, ProjectWise behavior, or production service entry points.

## Scripts moved (16)

All implementations now live under [`scripts/maintenance/`](../scripts/maintenance/):

| Script | Category |
|--------|----------|
| `Reset-QCFolderWorkflow.ps1` | PW workflow + folder telemetry reset |
| `Purge-QCPendingByFilters.ps1` | Queue filter purge |
| `Requeue-QCJobs.ps1` | Queue requeue |
| `Repair-QCQueueDuplicates.ps1` | Queue duplicate repair |
| `Repair-QCDocumentsFolderPaths.ps1` | Folder path normalization |
| `Invoke-QCDatabaseRetention.ps1` | Scheduled audit retention |
| `Remove-QCAuditEvents.ps1` | Audit event deletion |
| `Remove-QCWorkflowEvents.ps1` | Workflow event deletion |
| `Remove-LegacyQcPdfDatabaseRows.ps1` | Legacy QC PDF row cleanup |
| `Remove-InvalidSheetIndexRows.ps1` | Sheet index row cleanup |
| `Import-QCJsonlLogsToAutomationEvents.ps1` | Log import to automation events |
| `Sync-QCFolderSheetIndex.ps1` | Sheet index sync from PW |
| `Refresh-SheetIndexStates.ps1` | Sheet index state refresh |
| `Reconcile-QCSheetOwnership.ps1` | Sheet ownership reconcile |
| `Reconcile-QCStatusSets.ps1` | Status set reconcile to PW |
| `Sync-PWUserDirectory.ps1` | PW user directory sync |

### Path adjustments in moved scripts

Scripts that resolved repo root or modules from `$PSScriptRoot` were updated for the extra folder depth:

- `$repoRoot = Split-Path -Parent $PSScriptRoot` → double parent
- `$repoRoot = Split-Path -Parent $scriptDir` → double parent on `$scriptDir`
- `Join-Path $PSScriptRoot '..\modules\...'` → `..\..\modules\...`
- `Join-Path $PSScriptRoot 'Import-QCScriptModules.ps1'` → `..\Import-QCScriptModules.ps1` (helper stays in `scripts/`)
- `Invoke-QCDatabaseRetention.ps1` still invokes `Remove-QCAuditEvents.ps1` via same-folder path within `maintenance/`

## Wrappers retained (16)

Silent compatibility wrappers remain at former [`scripts/*.ps1`](../scripts/) paths:

```powershell
$target = Join-Path $PSScriptRoot 'maintenance\<ScriptName>.ps1'
& $target @args
exit $LASTEXITCODE
```

`Publish-QCPipelineCode.ps1` still copies `scripts\Reset-QCFolderWorkflow.ps1` — the wrapper preserves that publish path without code changes.

## Scripts intentionally deferred

| Path | Reason |
|------|--------|
| `scripts/Import-QCScriptModules.ps1` | Dot-sourced helper, not an operator entry point; stays in `scripts/` |
| `scripts/Stop-QCPipeline.ps1` | Pipeline stop helper; not in maintenance move list |
| `scripts/Combine-StatusSet.ps1`, `Run-CombineStatusSet.ps1` | Processing scripts (Phase 4 scope exclusion) |
| Service, deployment, diagnostics, discovery, prepend | Forbidden per Phase 4D scope |
| `tools/discovery/*` | Unchanged |

## References updated

| File | Change |
|------|--------|
| [`scripts/README.md`](../scripts/README.md) | Maintenance folder section |
| [`scripts/FILES.md`](../scripts/FILES.md) | Maintenance folder table; wrapper notes |
| [`docs/phase-4-module-script-organization-plan.md`](phase-4-module-script-organization-plan.md) | Phase 4D marked complete; Appendix E added |
| [`test/test_maintenance_script_wrappers.ps1`](../test/test_maintenance_script_wrappers.ps1) | **New** — wrapper/target checks + parse only |

Operator docs citing `scripts/Reset-*`, `scripts/Purge-*`, etc. remain valid via wrappers.

## Validation results

### Baseline (on `phase-4/integration`, before moves)

| Command | Result |
|---------|--------|
| `python -m pytest tests/ -q` | **PASS** — 86 passed, 3 skipped |
| `./test/run_focus_tests.ps1` | **PASS** — 21/21 |
| `./test/test_module_inventory.ps1` | **PASS** — 41 modules |
| `./test/test_watcher_module_bootstrap.ps1` | **PASS** |
| `./test/test_diagnostic_script_wrappers.ps1` | **PASS** — 23/23 |

### Post-implementation (on `phase-4/maintenance-scripts`)

| Command | Result |
|---------|--------|
| `python -m pytest tests/ -q` | **PASS** — 86 passed, 3 skipped |
| `./test/run_focus_tests.ps1` | **PASS** — 21/21 |
| `./test/test_module_inventory.ps1` | **PASS** — 41 modules |
| `./test/test_watcher_module_bootstrap.ps1` | **PASS** |
| `./test/test_diagnostic_script_wrappers.ps1` | **PASS** — 23/23 |
| `./test/test_maintenance_script_wrappers.ps1` | **PASS** — 16/16 |

Live maintenance operations (PW reset, SQL deletes, queue mutation) were not executed in this session.

## Rollback plan

1. `git checkout phase-4/integration` (or revert the branch).
2. Move all 16 scripts from `scripts/maintenance/` back to `scripts/`.
3. Revert path-resolution edits inside those scripts.
4. Delete compatibility wrappers at `scripts/*.ps1` for moved names.
5. Remove `scripts/maintenance/` if empty.
6. Revert doc/test updates and this summary.
7. Re-run baseline validation commands.

## Production impact

**None expected.** No modules, `appsettings.json`, processor routing, prepend behavior, or service entry points were changed. `Publish-QCPipelineCode.ps1` continues to target wrapper paths that forward correctly.
