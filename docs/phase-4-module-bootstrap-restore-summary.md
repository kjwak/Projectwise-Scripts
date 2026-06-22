# Phase 4 — Module bootstrap restore summary

**Branch:** `phase-4/module-bootstrap-restore`  
**Base branch:** `phase-4/integration`  
**Status:** Complete — **not committed** (awaiting review)

## Goal

Shared PowerShell 5.1 module export/session restore helper for production/high-traffic entrypoints. Stabilization only — no script/module moves, no manifest adoption, no behavior changes.

## Root cause

On Windows PowerShell 5.1, when module A is imported in a script and module B internally performs `Import-Module A -Force`, **session-scoped exports from A can disappear**. Entrypoints that import foundation modules, then feature modules that nested-reload them, then call foundation commands without a final restore are at risk.

**Dashboard failure (trigger):**

```text
The term 'Get-QCAppSettingsConfig' is not recognized...
scripts/Start-QCPipelineDashboard.ps1 line 1550
```

## Confirmed missing-command chains (pre-restore)

| Chain | Missing commands |
|-------|------------------|
| Dashboard: Config → Queue.Json → WatcherOrchestration → WatcherAlerts | `Get-QCAppSettingsConfig`, `Get-QCTimestamp` |
| Test-QCWatcherSessionAlert: Runtime → NotificationGraph → WatcherAlerts | `Write-QCJsonLog` |

## Helper added

**File:** `scripts/Restore-QCModuleExports.ps1`

Dot-sourced by entrypoints (not a `.psm1` module — avoids new production module during stabilization).

| Function | Purpose |
|----------|---------|
| `Resolve-QCModulePath` | Folder implementation paths only (maps bare names; never flat shims) |
| `Import-QCModuleGlobal` | `Import-Module -Force -Global` for session export restore |
| `Restore-QCFoundationModuleExports` | Standard foundation re-import order |
| `Test-QCRequiredCommands` | Throws listing all missing command names |
| `Import-QCModuleBootstrapSet` | Feature imports → foundation restore → command validation |

**Foundation restore order:**

```text
Core.Results → Core.Paths → Core.Runtime → Core.Config → Core.Logging → Core.Hashing → Core.Database
```

## Entrypoints updated (P0/P1)

| Script | Change |
|--------|--------|
| `scripts/Start-QCPipelineDashboard.ps1` | `_Import-DashboardModules` uses shared helper |
| `scripts/Run-QCProcessor.ps1` | Replaced narrow Runtime/Database re-import checks with full bootstrap set |
| `scripts/run_prepend_qc.ps1` | `-NoDashboard` path uses shared bootstrap before queue startup |

### Required commands validated

**Dashboard:** `Get-QCAppSettingsConfig`, `Get-QCTimestamp`, `Recover-QCStaleJobs`, `Clear-QCWatcherActive`

**Processor:** `Read-QCAppSettings`, `Get-NextQCJob`, `Lock-QCJob`, `Move-QCJob`, `Move-QCJobWithLockRetries`, `Invoke-QCPrependProcessor`, `Invoke-QCProcessorByType`, `Write-QCJsonLog`, `Test-QCDatabaseEnabled`

**run_prepend_qc (-NoDashboard):** `Get-QCAppSettingsConfig`, `Get-QCTimestamp`, `Get-NextQCJob`, `Get-QCQueueStats`, `Invoke-QCQueueStartupCheck`

## Intentionally deferred (P2/P3)

- `scripts/Combine-StatusSet.ps1` / `Run-CombineStatusSet.ps1` — StatusSet nested imports touch PW/Processing; defer to avoid scope creep
- All diagnostic scripts with hand-rolled imports (`Show-QCStatus.ps1`, `Get-PWFolderStateCounts.ps1`, etc.)
- All maintenance scripts not already using `Import-QCScriptModules.ps1`

## Tests added/updated

| Test | Purpose |
|------|---------|
| `test/test_module_bootstrap_restore.ps1` | **New** — helper path resolution, error messages, post-restore guarantees |
| `test/test_entrypoint_imports.ps1` | **Updated** — dashboard/processor/prepend chains via shared helper |

Not added to `run_focus_tests.ps1` (run separately).

## Validation results

| Command | Result |
|---------|--------|
| `python -m pytest tests/ -q` | **PASS** — 86 passed, 3 skipped |
| `./test/run_focus_tests.ps1` | **PASS** |
| `./test/test_module_inventory.ps1` | **PASS** |
| `./test/test_watcher_module_bootstrap.ps1` | **PASS** |
| `./test/test_watch_foundation_restore.ps1` | **PASS** |
| `./test/test_diagnostic_script_wrappers.ps1` | **PASS** |
| `./test/test_maintenance_script_wrappers.ps1` | **PASS** |
| `./test/test_module_folder_shims.ps1` | **PASS** |
| `./test/test_entrypoint_imports.ps1` | **PASS** |
| `./test/test_psd1_manifest_prototype.ps1` | **PASS** |
| `./test/test_module_bootstrap_restore.ps1` | **PASS** |

## Manual dashboard validation

No safe load-only or `-Help` mode exists for `Start-QCPipelineDashboard.ps1` (starts watcher/worker loop). Manual dashboard smoke not run.

## Known risks

- Foundation restore order may need tuning if new cross-folder nested imports appear.
- Dot-sourcing helper in multiple entrypoints duplicates function definitions (safe but not DRY with watcher’s inline restore).
- Diagnostics/maintenance scripts remain vulnerable until P2/P3 migration.
- `.psd1` manifest production adoption must include the same post-import restore validation.

## Explicit notes

- **No manifests adopted** in production entrypoints.
- **No script/module moves** performed.
- **Flat shims retained.**
- **No queue/PW/SQL/notification behavior changes.**

## Rollback plan

1. Revert changes to `Start-QCPipelineDashboard.ps1`, `Run-QCProcessor.ps1`, `run_prepend_qc.ps1`.
2. Delete `scripts/Restore-QCModuleExports.ps1`.
3. Revert `test/test_entrypoint_imports.ps1`; delete `test/test_module_bootstrap_restore.ps1`.
4. Revert doc and `.cursor/rules/module-imports.mdc` updates.
5. Re-run baseline validation.

Do not remove Phase 4E/4F folder layout or flat shims.
