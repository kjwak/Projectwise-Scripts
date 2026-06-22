# Phase 4G — PSD1 manifest prototype summary

**Branch:** `phase-4/psd1-manifest-prototype`  
**Base branch:** `phase-4/integration`  
**Status:** Prototype complete — **not committed** (awaiting review)

## Goal

Prototype `.psd1` module manifests for low-risk groups only (`QC.Core`, `QC.Queue`). Not a full Phase 4G migration. Production entrypoints unchanged.

## Manifests created

| Manifest | Location | Nested modules |
|----------|----------|----------------|
| `QC.Core.psd1` | `modules/Core/QC.Core.psd1` | `Core.Results`, `Core.Runtime`, `Core.Paths`, `Core.Config`, `Core.Logging`, `Core.Hashing`, `Core.Telemetry`, `QC.WatcherOrchestration` |
| `QC.Queue.psd1` | `modules/Queue/QC.Queue.psd1` | `QC.Filters`, `QC.Triggers`, `QC.JobFactory`, `QC.Queue.Json`, `QC.Worker` |

**Location decision:** Manifests live **inside category folders** beside their `.psm1` implementations. Root-level `modules/QC.Core.psd1` would duplicate naming with flat shims and blur the folder model; category-local manifests keep `NestedModules` paths short and colocated with implementations.

## Export strategy (prototype)

- `FunctionsToExport = '*'` on both manifests (technical debt — leaks private nested helpers such as `_CC-ToHashtableDeep` through the aggregate module surface).
- No `.psm1` file bodies changed.
- Flat `modules/*.psm1` shims retained.
- No production script, `Import-QCScriptModules.ps1`, or service entrypoint uses manifests yet.

## Commands verified (via `test/test_psd1_manifest_prototype.ps1`)

**QC.Core:** `New-QCResult`, `Get-QCTimestamp`, `Normalize-QCPath`, `Read-QCAppSettings`, `Write-QCJsonLog`, `Get-Sha256TextHex`, `Write-QCAutomationEvent`, `Get-QCWatcherMode`

**QC.Queue:** `Get-NextQCJob`, `New-QCJobObject`, `Move-QCJobWithLockRetries`, `Test-QCPathAllowed`, `Resolve-QCTriggerMatch`

## Design question answers

| # | Question | Prototype result |
|---|----------|------------------|
| 1 | QC.Core imports cleanly on PS 5.1? | **Yes** — `Test-ModuleManifest` passes; `Import-Module` succeeds |
| 2 | QC.Queue imports after QC.Core? | **Yes** — queue commands available after core+queue import |
| 3 | Expected public commands exposed? | **Yes** — via `FunctionsToExport = '*'` re-export from nested modules |
| 4 | Duplicate / export clobber? | **Partial risk** — nested modules still load individually; `Core.Database` pulled by `Core.Telemetry`; `QC.ProcessType` pulled by `QC.Triggers`; same PS 5.1 clobber patterns as direct imports |
| 5 | Coexist with folder paths + flat shims? | **Yes** — flat shim and folder `.psm1` imports work after manifest import in same session |
| 6 | Future production use? | **Defer decision** — manifests viable for packaging/docs; production should not switch until export narrowing, dependency boundaries, and load-order restore strategy are defined |
| 7 | Folder vs root placement? | **Prefer category folder** for this repo layout |

## PowerShell version coverage

| Runtime | Result |
|---------|--------|
| Windows PowerShell 5.1 | **PASS** — manifest import smoke + prototype test |
| PowerShell 7 (`pwsh`) | **Not tested** — `pwsh` not available in validation environment |

## Validation results

| Command | Result |
|---------|--------|
| `python -m pytest tests/ -q` | **PASS** — 86 passed, 3 skipped |
| `./test/run_focus_tests.ps1` | **PASS** |
| `./test/test_module_inventory.ps1` | **PASS** — 41 modules |
| `./test/test_watcher_module_bootstrap.ps1` | **PASS** |
| `./test/test_watch_foundation_restore.ps1` | **PASS** |
| `./test/test_diagnostic_script_wrappers.ps1` | **PASS** — 23 scripts |
| `./test/test_maintenance_script_wrappers.ps1` | **PASS** — 16 scripts |
| `./test/test_module_folder_shims.ps1` | **PASS** — 41 modules, 11 probes |
| `./test/test_entrypoint_imports.ps1` | **PASS** |
| `./test/test_psd1_manifest_prototype.ps1` | **PASS** |
| PS 5.1 direct smoke (`New-QCResult`, `Get-QCTimestamp`) | **PASS** |
| PS 5.1 direct smoke (`Get-NextQCJob`, `New-QCJobObject`) | **PASS** |

## Known risks

- `FunctionsToExport = '*'` exports private nested functions from aggregate module names (`QC.Core`, `QC.Queue`).
- `QC.Core` nested import chain loads `Core.Database` (telemetry optional SQL path) without live SQL use.
- `QC.Queue` nested `QC.Triggers` imports `Workflow/QC.ProcessType.psm1` — broader than queue-only surface.
- Manifest + direct nested-module imports in one session can register duplicate module identities (same as flat-shim era).
- Wildcard exports incompatible with strict public API contracts until narrowed.

## Recommendation for full Phase 4G

1. Keep manifests **test/documentation only** until export lists are explicit per nested module.
2. Add `RequiredModules` / ordered `NestedModules` documentation for cross-folder deps (or split `QC.Triggers` ProcessType import).
3. Do **not** adopt manifests in `Watch-QCTrigger.ps1` / `Run-QCProcessor.ps1` until load-order restore (`_Watch-RestoreFoundationModules`) is revalidated against manifest import.
4. Expand manifests only after Core+Queue production soak; defer PW/Workflow/Processing/Notifications manifests.
5. Consider root-level aggregate manifests only if PSModulePath publishing requires single-folder discovery.

## Rollback plan

1. Delete `modules/Core/QC.Core.psd1` and `modules/Queue/QC.Queue.psd1`.
2. Delete `test/test_psd1_manifest_prototype.ps1`.
3. Revert doc updates from this branch.
4. Re-run baseline validation.

## Explicit notes

- **No full Phase 4G complete** — prototype only.
- **Flat shims remain.**
- **No production behavior changes.**
