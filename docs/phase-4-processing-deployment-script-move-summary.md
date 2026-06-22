# Phase 4 — Processing/deployment script move summary

**Branch:** `phase-4/processing-deployment-scripts`  
**Base branch:** `phase-4/integration`  
**Status:** Complete — **not committed** (awaiting review)

## Goal

Move remaining low/medium-risk processing and deployment helper scripts into `scripts/processing/` and `scripts/deployment/` with silent compatibility wrappers at former paths. Path organization only — no behavior changes.

## Scripts moved

| Old path | New path |
|----------|----------|
| `scripts/Combine-StatusSet.ps1` | `scripts/processing/Combine-StatusSet.ps1` |
| `scripts/Run-CombineStatusSet.ps1` | `scripts/processing/Run-CombineStatusSet.ps1` |
| `scripts/Promote-DevToMain.ps1` | `scripts/deployment/Promote-DevToMain.ps1` |
| `scripts/Sync-OverlayReviewStamp.ps1` | `scripts/deployment/Sync-OverlayReviewStamp.ps1` |

## Wrappers retained

Silent wrappers at all four former `scripts/*.ps1` paths forward `@args` and `exit $LASTEXITCODE` to the subfolder implementation.

## Path adjustments in moved scripts

| Script | Change |
|--------|--------|
| `processing/Combine-StatusSet.ps1` | `$repoRoot` → two levels up from `processing/` |
| `processing/Run-CombineStatusSet.ps1` | `$repoRoot` and default `AppSettingsPath` → two levels up |
| `deployment/Promote-DevToMain.ps1` | `$repoRoot` → two levels up from `deployment/` |
| `deployment/Sync-OverlayReviewStamp.ps1` | `$repoRoot` → two levels up |

`Run-CombineStatusSet.ps1` still invokes `Combine-StatusSet.ps1` in the same `processing/` folder (unchanged relative target).

## Restore-QCModuleExports.ps1

**Not used** by moved scripts. `Combine-StatusSet.ps1` imports only `Core.Results` and `QC.StatusSet` and calls `Invoke-StatusSetNativeJob` at module scope — no foundation commands at script session level. Helper unchanged.

## Intentionally deferred / not moved

| Script | Reason |
|--------|--------|
| `scripts/Publish-QCPipelineCode.ps1` | High-risk publish entrypoint; copy list unchanged |
| Service entrypoints | `Watch-QCTrigger`, `Run-QCProcessor`, `Start-QCPipelineDashboard`, `run_prepend_qc`, etc. |
| `scripts/processing/Invoke-QCPrependPw.ps1` | Already in processing (Phase 4B) |
| Diagnostics/maintenance | Completed in Phase 4C/4D |

## References updated

- `scripts/README.md`
- `scripts/FILES.md`
- `docs/phase-4-module-script-organization-plan.md` (Appendix J)

No changes to `Publish-QCPipelineCode.ps1` (does not copy moved scripts).

## Tests added

| Test | Purpose |
|------|---------|
| `test/test_processing_deployment_script_wrappers.ps1` | Wrapper existence, target path, parse syntax (no script execution) |

Not added to `run_focus_tests.ps1`.

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
| `./test/test_processing_deployment_script_wrappers.ps1` | **PASS** — 4 scripts |

## Known risks

- External docs/bookmarks pointing at old paths still work via wrappers.
- Direct invocation of `scripts/processing/*.ps1` bypasses wrappers (intentional).
- `Combine-StatusSet` remains vulnerable to PS 5.1 export clobber if extended to call Runtime commands at script level (deferred to bootstrap-restore P2).

## Explicit notes

- **Service scripts not moved.**
- **`Publish-QCPipelineCode.ps1` not moved** and publish behavior unchanged.
- **Manifests not changed**; no production manifest adoption.
- **Flat module shims retained.**

## Rollback plan

1. `git mv` implementations back to `scripts/` root.
2. Delete wrappers at old paths (or restore pre-move content).
3. Revert path adjustments in moved scripts.
4. Delete `test/test_processing_deployment_script_wrappers.ps1`.
5. Revert doc updates.
6. Re-run baseline validation.
