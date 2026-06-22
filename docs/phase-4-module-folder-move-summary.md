# Phase 4E — Module folder move summary

**Branch:** `phase-4/module-folders`  
**Base branch:** `phase-4/integration`  
**Status:** Implementation complete — **not committed** (awaiting review)

## Goal

Physical module organization only: move 41 active `.psm1` implementations into responsibility-based subfolders while preserving flat `modules/*.psm1` paths as silent compatibility shims. No production script import path changes in this phase (deferred to Phase 4F).

## Modules moved (41)

| Folder | Modules |
|--------|---------|
| `modules/Core/` | `Core.Results`, `Core.Runtime`, `Core.Paths`, `Core.Config`, `Core.Logging`, `Core.Hashing`, `Core.Telemetry`, `QC.WatcherOrchestration` |
| `modules/Database/` | `Core.Database` |
| `modules/ProjectWise/` | `PW.Connection`, `PW.Discovery`, `PW.AuditPoller`, `PW.Users` |
| `modules/Workflow/` | `QC.Workflow`, `QC.AuditTriggers`, `QC.ProcessType` |
| `modules/Queue/` | `QC.Queue.Json`, `QC.JobFactory`, `QC.Worker`, `QC.Filters`, `QC.Triggers` |
| `modules/Processing/` | `QC.Processors`, `QC.StatusSet`, `QC.Rendition`, `QC.ReviewStamp`, `QC.PdfExport`, `QC.CommentExtract`, `QC.CommentStatusDecision`, `QC.CommentStatusProcessor`, `QC.CommentSync.*` (4) |
| `modules/Notifications/` | `QC.Notifications`, `QC.NotificationGraph`, `QC.NotificationMock`, `QC.NotificationTemplates`, `QC.NotificationThreads`, `QC.WatcherAlerts` |
| `modules/Reporting/` | `QC.Reporting` |
| `modules/Diagnostics/` | `QC.DebugMcp` |

## Flat shims retained (41)

Each flat `modules/<Name>.psm1`:

```powershell
$ErrorActionPreference = 'Stop'
$target = Join-Path $PSScriptRoot '<Folder>\<Name>.psm1'
if (-not (Test-Path -LiteralPath $target)) {
    throw "Module implementation not found: $target"
}
Import-Module $target -Force -Global
```

Callers using `Import-Module "$repoRoot\modules\QC.Queue.Json.psm1"` continue to work unchanged.

**Note:** Importing via a flat shim briefly registers both the shim script module and the folder implementation under the same module name. Public exports resolve correctly. Tests that use `InModuleScope` or invoke private module functions should target the implementation module via [`test/_Resolve-ModuleImplPath.ps1`](../test/_Resolve-ModuleImplPath.ps1) helpers (`Get-QCModuleImplementation`, `Remove-QCModuleFlatShims`).

## Path-resolution changes (implementations only)

| Pattern | Adjustment |
|---------|------------|
| `Join-Path $PSScriptRoot 'Sibling.psm1'` | `Join-Path (Split-Path -Parent $PSScriptRoot) 'Sibling.psm1'` (flat shim path) |
| `Split-Path -Parent $PSScriptRoot` (repo root) | `Split-Path -Parent (Split-Path -Parent $PSScriptRoot)` |
| `$orchRoot` / `$modulesRoot` for `modules/` root | Single `Split-Path -Parent $PSScriptRoot` (not double) |
| `_QCN-GetRepoRoot`, `_QCR-GetRepoRoot`, `_QCWA-GetRepoRoot`, `_QCNT-ResolveOutputRoot`, `_QCNG-ResolveRepoPath` | Return repo root via double parent from subfolder |

No module function bodies refactored beyond path resolution.

## References updated

| File | Change |
|------|--------|
| [`modules/FILES.md`](../modules/FILES.md) | Folder layout table; shim note |
| [`modules/README.md`](../modules/README.md) | Phase 4E folder structure section |
| [`docs/phase-4-module-script-organization-plan.md`](phase-4-module-script-organization-plan.md) | Phase 4E complete; Appendix F |
| [`test/test_module_folder_shims.ps1`](../test/test_module_folder_shims.ps1) | **New** — 41 shim/layout checks + 11 flat-path import probes |
| [`tests/module_impl.py`](../tests/module_impl.py) | **New** — static pytest helper follows shims to implementation source |
| [`test/_Resolve-ModuleImplPath.ps1`](../test/_Resolve-ModuleImplPath.ps1) | **New** — resolve shim → implementation; helpers for tests needing impl module scope |
| Focus tests (`test_lock_steal_dead_pid`, `test_qc_workflow`, static source readers) | Use impl path helpers where flat shims affect module identity or source reads |

Production scripts, service entry points, diagnostics/maintenance wrappers, `appsettings.json`, SQL, MCP config: **unchanged**.

## Validation results

### Baseline (on `phase-4/integration`, before moves)

| Command | Result |
|---------|--------|
| `python -m pytest tests/ -q` | **PASS** — 86 passed, 3 skipped |
| `./test/run_focus_tests.ps1` | **PASS** — 21/21 |
| `./test/test_module_inventory.ps1` | **PASS** — 41 modules |
| `./test/test_watcher_module_bootstrap.ps1` | **PASS** |
| `./test/test_diagnostic_script_wrappers.ps1` | **PASS** — 23/23 |
| `./test/test_maintenance_script_wrappers.ps1` | **PASS** — 16/16 |

### Post-implementation (on `phase-4/module-folders`)

| Command | Result |
|---------|--------|
| `python -m pytest tests/ -q` | **PASS** — 86 passed, 3 skipped |
| `./test/run_focus_tests.ps1` | **PASS** — 21/21 |
| `./test/test_module_inventory.ps1` | **PASS** — 41 modules |
| `./test/test_watcher_module_bootstrap.ps1` | **PASS** |
| `./test/test_diagnostic_script_wrappers.ps1` | **PASS** — 23/23 |
| `./test/test_maintenance_script_wrappers.ps1` | **PASS** — 16/16 |
| `./test/test_module_folder_shims.ps1` | **PASS** — 41 shims + 11 import probes |
| Flat-path import smoke | **PASS** — `Core.Results`, `QC.Queue.Json`, `QC.JobFactory` → `New-QCJobObject` |

## Known risks

- Deep import chains (`PW.Discovery`, `QC.Processors`, `QC.Notifications`) depend on flat shims resolving correctly when implementations import siblings.
- Flat shim imports register a transient duplicate module name alongside the folder implementation; harmless for normal `Import-Module` + public API use, but `Get-Module <Name>` may return multiple entries until flat shim instances are removed from the session.
- `QC.WatcherOrchestration` uses `$orchRoot = Split-Path -Parent $PSScriptRoot` (modules root) — must not be double-parented to repo root.
- Phase 4F must update production `Import-Module` paths and `Publish-QCPipelineCode.ps1` copy lists before shim removal in 4H.

## Rollback plan

1. Revert branch or move implementations back to flat `modules/`.
2. Delete subfolder copies and flat shims.
3. Revert path-resolution edits inside implementation files.
4. Revert docs/tests (`module_impl.py`, static pytest updates, shim test).
5. Re-run baseline validation.

## Production impact

**None expected** for runtime behavior when importing via flat paths. Physical layout changes only; Phase 4F handles import path migration.
