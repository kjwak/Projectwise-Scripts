# Phase 4F — Import path update summary

**Branch:** `phase-4/import-updates`  
**Base branch:** `phase-4/integration`  
**Status:** Implementation complete — **not committed** (awaiting review)

## Goal

Update active `Import-Module` references from flat `modules/<Name>.psm1` shim paths to Phase 4E folder implementation paths. Flat shims remain for compatibility. **No `.psd1` manifests** introduced.

## Import paths updated

| Area | Examples |
|------|----------|
| Service entrypoints | `Watch-QCTrigger.ps1`, `Run-QCProcessor.ps1`, `Start-QCPipelineDashboard.ps1` |
| Shared loader | `Import-QCScriptModules.ps1` (+ `_QCResolve-ModuleRelativePath` for `-AdditionalModules`) |
| Diagnostics / maintenance | All `scripts/diagnostics/*.ps1`, `scripts/maintenance/*.ps1` under repo |
| Other scripts | `Combine-StatusSet.ps1`, `run_prepend_qc.ps1`, etc. |
| Tests | `test/*.ps1` (except intentional flat-shim probes in `test_module_folder_shims.ps1`) |
| Static pytest | `tests/*.py` via existing `module_impl` helper paths |
| Module implementations | Cross-folder `Import-Module` in `modules/*/*.psm1` (flat parent shim → folder path) |
| Cursor rule | `.cursor/rules/module-imports.mdc` |
| Docs | `modules/README.md`, `modules/FILES.md`, organization plan Appendix G |

### Mapping (unchanged from Phase 4E)

`modules/<Name>.psm1` shim → `modules/<Folder>/<Name>.psm1` implementation (see Phase 4E plan table).

## Intentionally left on flat shim paths

| File | Reason |
|------|--------|
| `scripts/processing/Invoke-QCPrependPw.ps1` | Forbidden in Phase 4F scope; prepend path isolation |
| `modules/*.psm1` (41 flat shims) | Compatibility layer — must remain |
| `test/test_module_folder_shims.ps1` | Explicitly tests flat shim forwarding |
| `archive/*` | Historical / out of scope |

## Files changed (approximate)

~190+ files touched by systematic path update (scripts, tests, module implementations, docs). No changes to `appsettings.json`, SQL, MCP config, prepend behavior, or diagnostic/maintenance **behavior** (import paths only).

## New / updated tests

| File | Purpose |
|------|---------|
| `test/test_entrypoint_imports.ps1` | **New** — watcher/processor/dashboard/MCP folder imports + flat shim compatibility |
| `test/test_watch_foundation_restore.ps1` | Load/restore order arrays use folder-relative paths |
| `scripts/Import-QCScriptModules.ps1` | Folder load order + bare-name resolver for `-AdditionalModules` |

## Validation results

| Command | Result |
|---------|--------|
| `python -m pytest tests/ -q` | **PASS** — 86 passed, 3 skipped |
| `./test/run_focus_tests.ps1` | **PASS** — 21/21 |
| `./test/test_module_inventory.ps1` | **PASS** — 41 modules |
| `./test/test_watcher_module_bootstrap.ps1` | **PASS** |
| `./test/test_diagnostic_script_wrappers.ps1` | **PASS** — 23/23 |
| `./test/test_maintenance_script_wrappers.ps1` | **PASS** — 16/16 |
| `./test/test_module_folder_shims.ps1` | **PASS** — 41 shims + 11 import probes |
| `./test/test_entrypoint_imports.ps1` | **PASS** — folder paths + flat shim smoke |
| `./test/test_watch_foundation_restore.ps1` | **PASS** |

## Known risks

- PowerShell 5.1 export clobber still requires foundation restore after nested `-Force` imports (unchanged behavior).
- `Invoke-QCPrependPw.ps1` still imports via flat paths — mixed import style until a future prepend import update.
- `tools/pw-qc-mcp/pw_qc_worker.ps1` is gitignored; MCP worker should be updated manually if it still references flat paths.

## Rollback plan

1. Revert branch or restore flat `modules/<Name>.psm1` import strings in scripts/tests/modules.
2. Keep Phase 4E folder implementations and flat shims.
3. Revert `test/test_entrypoint_imports.ps1` and doc updates.
4. Re-run baseline validation.

## Explicit notes

- **No `.psd1` manifests** were added.
- **Flat shims remain** at `modules/*.psm1` (41 files).
- **No public export changes** intentional.
