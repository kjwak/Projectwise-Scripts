# Phase 4 — Diagnostic/maintenance bootstrap summary

**Branch:** `phase-4/diagnostic-maintenance-bootstrap`  
**Base branch:** `phase-4/integration`  
**Status:** Complete — awaiting review

## Goal

Harden PS 5.1 module import/bootstrap for diagnostic, maintenance, and `Combine-StatusSet` scripts using the shared `Restore-QCModuleExports.ps1` helper. No behavior, path, or service changes.

## Pattern

Scripts dot-source the helper from their subfolder:

```powershell
. (Join-Path $PSScriptRoot '..\Restore-QCModuleExports.ps1') -RepoRoot $repoRoot
Import-QCModuleBootstrapSet -FeatureModules @(...) -RequiredCommands @(...) -Context '<ScriptName> bootstrap'
```

`processing/Combine-StatusSet.ps1` uses `Join-Path (Split-Path $PSScriptRoot -Parent) 'Restore-QCModuleExports.ps1'`.

### Paths post-restore

When a script calls `Normalize-QCDocumentsFolderPath` at session scope, re-import Paths **after** bootstrap (foundation restore ends with `Core.Database`, which clobbers Paths exports):

```powershell
Import-QCModuleGlobal -RelativePath 'Core\Core.Paths.psm1'
Test-QCRequiredCommands -Names @('Normalize-QCDocumentsFolderPath') -Context '... paths post-restore'
```

Applied to: `Reset-QCFolderWorkflow.ps1`, `Repair-QCDocumentsFolderPaths.ps1`, `Sync-QCFolderSheetIndex.ps1`.

## Scripts updated

| Area | Count | Notes |
|------|-------|--------|
| `scripts/diagnostics/` | 11 | QC-module probes only |
| `scripts/maintenance/` | 16 | Replaced hand-rolled imports and `Import-QCScriptModules.ps1` |
| `scripts/processing/` | 1 | `Combine-StatusSet.ps1` |

### PW-only exempt (unchanged)

12 scripts import only `pwps` / `pwps_dab`: `Test-PWEmailAttributes*` (10), `PW-SmokeTest`, `Test-PWDocumentStateChange`.

## Tests added

| Test | Purpose |
|------|---------|
| `test/test_diagnostic_maintenance_bootstrap.ps1` | Static bootstrap usage checks + clobber-prone import chains |

## Validation results

| Command | Result |
|---------|--------|
| `./test/test_diagnostic_maintenance_bootstrap.ps1` | **PASS** |
| `./test/test_entrypoint_imports.ps1` | **PASS** |
| `./test/test_module_bootstrap_restore.ps1` | **PASS** |

## Explicit notes

- **No service script moves**
- **`Publish-QCPipelineCode.ps1` unchanged** (see `test/test_notification_template_paths.ps1` for documented `email/` publish gap)
- **No shim removal**
- **No production `.psd1` adoption**
- **`QC.NotificationTemplates.psm1` repo-root fix** — `_QCNT-GetRepoRoot` updated for Phase 4E `modules/Notifications/` layout (fixes `QC_EMAIL_TEMPLATE_NOT_FOUND` on worker)
- **`Import-QCScriptModules.ps1` retained** (unused by maintenance after this branch; still documented for reference)

## Rollback plan

1. Revert bootstrap blocks in affected scripts; restore prior `Import-Module` / `Import-QCScriptModules.ps1` calls.
2. Delete `test/test_diagnostic_maintenance_bootstrap.ps1`.
3. Revert doc and `.cursor/rules/module-imports.mdc` updates.
