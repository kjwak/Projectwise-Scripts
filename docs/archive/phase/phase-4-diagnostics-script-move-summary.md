# Phase 4C — Diagnostics script move summary

**Branch:** `phase-4/diagnostics-scripts`  
**Base branch:** `phase-4/integration`  
**Status:** Implementation complete — **not committed** (awaiting review)

## Goal

Path-organization only: move low-risk diagnostic/discovery scripts into `scripts/diagnostics/` while preserving operator-facing paths via silent wrappers. No production behavior changes.

## Scripts moved (23)

All implementations now live under [`scripts/diagnostics/`](../scripts/diagnostics/):

| Script | Notes |
|--------|--------|
| `Get-PWFolderStateCounts.ps1` | PW folder state counts (read-only) |
| `Scan-PWProjectMetrics.ps1` | Project metrics scan |
| `Show-QCStatus.ps1` | Queue status summary |
| `Show-QCQueueDiag.ps1` | Queue health diagnostic |
| `PW-BrowseFolder.ps1` | PW folder browser |
| `PW-ListDocsInFolder.ps1` | PW document list |
| `PW-TestWatchRoots.ps1` | Watch-list root expansion |
| `PW-SmokeTest.ps1` | PW end-to-end smoke (writes) |
| `Test-PWConnection.ps1` | Connect/disconnect smoke |
| `Test-PWDocumentStateChange.ps1` | Workflow state change probe |
| `Test-PWEmailAttributes.ps1` | Email attribute probe |
| `Test-PWEmailAttributes-AttributesBag.ps1` | Variant probes (9 scripts) |
| `Test-PWEmailAttributes-Caltrans.ps1` | |
| `Test-PWEmailAttributes-DeepProbe.ps1` | |
| `Test-PWEmailAttributes-DumpBag.ps1` | |
| `Test-PWEmailAttributes-EnvCount.ps1` | |
| `Test-PWEmailAttributes-Extract.ps1` | |
| `Test-PWEmailAttributes-FolderEnv.ps1` | |
| `Test-PWEmailAttributes-InspectOne.ps1` | |
| `Test-PWEmailAttributes-ScanPdfs.ps1` | |
| `Test-QCEmailTemplate.ps1` | Email template smoke |
| `Test-QCNotificationGraph.ps1` | Graph notification smoke |
| `Test-QCWatcherSessionAlert.ps1` | Watcher session alert preview |

Moved scripts that resolve `appsettings.json` / `modules/` via `$PSScriptRoot` were updated to use one additional `..` segment (from `scripts/` → `scripts/diagnostics/`).

## Wrappers retained (23)

Silent compatibility wrappers remain at former [`scripts/*.ps1`](../scripts/) paths:

```powershell
$target = Join-Path $PSScriptRoot 'diagnostics\<ScriptName>.ps1'
& $target @args
exit $LASTEXITCODE
```

Operators, docs, and command history can continue using `.\scripts\Show-QCStatus.ps1` (etc.) unchanged.

## Scripts intentionally deferred

| Path | Reason |
|------|--------|
| `tools/discovery/*.ps1` (18 scripts) | Docs (`database-telemetry.md`, `qc-workflow-framework.md`, `appsettings-reference.md`, etc.), cross-references between discovery scripts, and operator runbooks still pin `tools/discovery/`. Moving would require broader doc + MCP path audit. |
| All forbidden production/service scripts | Per Phase 4C scope (`Watch-QCTrigger`, `Run-QCProcessor`, prepend, maintenance, deployment, etc.) |
| `legacy/prepend_qc.ps1`, `scripts/processing/Invoke-QCPrependPw.ps1` | Prepend path promotion (Phase 4B) — not touched |
| Modules, `appsettings.json`, SQL, MCP config | Out of scope |

## References updated

| File | Change |
|------|--------|
| [`scripts/README.md`](../scripts/README.md) | Diagnostics section points to `scripts/diagnostics/` + wrappers |
| [`scripts/FILES.md`](../scripts/FILES.md) | Diagnostics folder table; wrapper notes for moved scripts |
| [`docs/archive/phase/phase-4-module-script-organization-plan.md`](phase-4-module-script-organization-plan.md) | Phase 4C marked complete; Appendix D added |
| [`test/test_diagnostic_script_wrappers.ps1`](../test/test_diagnostic_script_wrappers.ps1) | **New** — wrapper/target existence and path resolution (no live PW) |

Docs that cite `scripts/Test-*` or `scripts/Show-*` paths remain valid via wrappers. No doc path churn required for operator commands.

## Validation results

### Baseline (on `phase-4/integration`, before moves)

| Command | Result |
|---------|--------|
| `python -m pytest tests/ -q` | **PASS** — 86 passed, 3 skipped |
| `./test/run_focus_tests.ps1` | **PASS** — 21/21 |
| `./test/test_module_inventory.ps1` | **PASS** — 41 modules |
| `./test/test_watcher_module_bootstrap.ps1` | **PASS** |

### Post-implementation (on `phase-4/diagnostics-scripts`)

| Command | Result |
|---------|--------|
| `python -m pytest tests/ -q` | **PASS** — 86 passed, 3 skipped |
| `./test/test_diagnostic_script_wrappers.ps1` | **PASS** — 23/23 wrapper checks |
| `./test/run_focus_tests.ps1` | **PASS** — 21/21 |
| `./test/test_module_inventory.ps1` | **PASS** — 41 modules |
| `./test/test_watcher_module_bootstrap.ps1` | **PASS** |
| Manual wrapper resolution | **PASS** — `scripts/Show-QCQueueDiag.ps1` resolves `scripts/diagnostics/Show-QCQueueDiag.ps1` |

Live ProjectWise diagnostics were not executed in this session (require credentials and may perform writes).

## Rollback plan

1. `git checkout phase-4/integration` (or revert the branch).
2. Move all 23 scripts from `scripts/diagnostics/` back to `scripts/`.
3. Revert path-resolution edits inside those scripts (one fewer `..` to repo root).
4. Delete compatibility wrappers at `scripts/*.ps1` for moved names.
5. Remove `scripts/diagnostics/` if empty.
6. Revert doc/test updates (`scripts/README.md`, `scripts/FILES.md`, plan Appendix D, `test/test_diagnostic_script_wrappers.ps1`, this summary).
7. Re-run baseline validation commands.

## Production impact

**None expected.** No modules, `appsettings.json`, processor routing, prepend behavior, or service entry points were changed. Diagnostic scripts are operator/dev tools only.
