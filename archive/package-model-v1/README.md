# Package model v1 (archived)

**Archived:** Phase 3 (`phase-3/package-model-archive`)  
**Status:** Not wired into production. Retained for reference and optional manual testing.

## Why this was archived

Phase 2 investigation ([`docs/phase-2-package-model-decision.md`](../../docs/phase-2-package-model-decision.md)) found that this in-memory `QC.Package*` cluster had **zero production callers**. Production package grouping already runs through SQL tables in [`modules/Core.Database.psm1`](../../modules/Core.Database.psm1).

The archived design diverged from production in important ways:

| Dimension | Archived v1 | Production (SQL) |
|-----------|-------------|------------------|
| Package key | `qcpkg_*` hash | `sheet_package_id` (`UNIQUEIDENTIFIER`) |
| QC PDF naming | `_QC`, `-QC` suffixes | `-prod`, `-rev`, `-chk` lanes |
| State model | Configurable pre-TYPSA labels in tests | TYPSA per-document / per-lane rows |
| Cache | Fictional `dbo.QCPackageCache` stub | `sheet_packages` + `sheet_package_qc_pdfs` |

## Production-canonical implementation

See [`docs/qc-package-model.md`](../../docs/qc-package-model.md) for the current SQL-backed model:

- **Tables:** `sheet_packages`, `sheet_documents`, `sheet_package_qc_pdfs`, `sheet_index` (dual-write)
- **Module:** `Core.Database.psm1` — `Ensure-SheetPackage`, `Resolve-SheetPackageFromDocument`, `Get-SheetPackageIdForDocument`, lane registry sync, etc.
- **Consumers:** watcher, audit triggers, notifications, discovery, MCP debug, reporting (`v_sheet_package_status`)

## Contents of this archive

### `modules/`

| File | Role |
|------|------|
| `QC.PackageResolver.psm1` | `Resolve-QCPackage`, canonical document selection |
| `QC.AttributePolicy.psm1` | User vs automation attribute ownership |
| `QC.StatePolicy.psm1` | Package-level state precedence |
| `QC.PackageSync.psm1` | `Sync-QCPackageAttributes` orchestration |
| `QC.Package.Database.psm1` | Stub `QCPackageCache` writer (table never existed) |

### `test/`

| File | Role |
|------|------|
| `test_qc_package_model.ps1` | Sole automated test for the v1 cluster |

## Run archived test manually

From repository root:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File archive/package-model-v1/test/test_qc_package_model.ps1
```

The test imports archived modules from this folder and `QC.JobFactory.psm1` from `modules/` (package dedupe branch only).

Archived modules resolve `Core.*` and `PW.*` dependencies from `modules/` via `$script:_QCPkgV1RepoModules`. Nested cluster imports use `-Global` so exported commands remain visible after reload.

**Note:** On Windows PowerShell 5.1, `test_qc_package_model.ps1` may fail with `InvokeMethodOnNull` inside `Resolve-QCPackage` — the same behavior reproduces with the pre-archive module copy and is not caused by the SQL production path. The archived test is optional reference coverage only.

## Restore to `modules/` (if needed)

Targeted restore from git history or a prior commit — do **not** restore the entire `modules/` folder:

```powershell
git checkout <commit-or-branch> -- `
  modules/QC.PackageResolver.psm1 `
  modules/QC.AttributePolicy.psm1 `
  modules/QC.StatePolicy.psm1 `
  modules/QC.PackageSync.psm1 `
  modules/QC.Package.Database.psm1 `
  test/test_qc_package_model.ps1
```

Then re-add the five module rows to `modules/FILES.md` and remove or relocate this archive folder as appropriate.

## Deferred follow-up

`QC.JobFactory.psm1` still contains an unwired `metadata.package` dedupe branch (`New-QCPackageJobDedupeKey`). Phase 3 did not modify job factory behavior. Product decision pending: wire to `sheet_package_id` at job creation or remove the dead branch.
