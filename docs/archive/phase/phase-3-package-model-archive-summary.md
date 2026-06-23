# Phase 3 — Package Model Archive Summary

**Branch:** `phase-3/package-model-archive`  
**Base:** `dev` @ `5916117` (Update documentation in QC package model and legacy README for improved clarity)  
**Date:** 2026-06-22

---

## Branch setup

| Item | Value |
|------|-------|
| Branch existed before work | No — created from updated `dev` |
| Branch point | `dev` @ `5916117` |
| Working tree at start | Clean |

### Baseline validation (on `dev`, before changes)

| Suite | Result |
|-------|--------|
| `python -m pytest tests/ -q` | **86 passed, 3 skipped** |
| `test/run_focus_tests.ps1` | **ALL PASSED** |
| `test/test_module_inventory.ps1` | **PASS** (46 modules) |
| `test/test_watcher_module_bootstrap.ps1` | **PASS** |

---

## Files moved

| From | To |
|------|-----|
| `modules/QC.PackageResolver.psm1` | `archive/package-model-v1/modules/QC.PackageResolver.psm1` |
| `modules/QC.AttributePolicy.psm1` | `archive/package-model-v1/modules/QC.AttributePolicy.psm1` |
| `modules/QC.StatePolicy.psm1` | `archive/package-model-v1/modules/QC.StatePolicy.psm1` |
| `modules/QC.PackageSync.psm1` | `archive/package-model-v1/modules/QC.PackageSync.psm1` |
| `modules/QC.Package.Database.psm1` | `archive/package-model-v1/modules/QC.Package.Database.psm1` |
| `test/test_qc_package_model.ps1` | `archive/package-model-v1/test/test_qc_package_model.ps1` |

## Files created

| File | Purpose |
|------|---------|
| `archive/package-model-v1/README.md` | Archive rationale, SQL canonical pointers, restore steps |
| `docs/archive/phase/phase-3-package-model-archive-summary.md` | This document |

## Files edited

| File | Change |
|------|--------|
| `modules/FILES.md` | Removed 5 archived module rows (46 → 41 modules) |
| `modules/README.md` | Added "Sheet package model (production)" section + archive pointer |
| `docs/architecture/qc-package-model.md` | Rewritten for SQL-backed production model |
| `docs/archive/phase/phase-2-package-model-decision.md` | Added "Implemented in Phase 3" section |
| `archive/package-model-v1/test/test_qc_package_model.ps1` | Patched import paths for archive layout (`-Global` imports) |
| `archive/package-model-v1/modules/*.psm1` | Added `$script:_QCPkgV1RepoModules` for `Core.*`/`PW.*`; nested cluster imports use `-Global`; `QC.PackageResolver` adds `[hashtable]` guard on `$CachedPackage` |

## Files not modified (explicit)

- `modules/Database/Core.Database.psm1`, SQL schema/migrations, Power BI views
- `modules/Queue/QC.JobFactory.psm1` (`New-QCPackageJobDedupeKey`, `metadata.package` branch)
- Watcher, processor, notifications, ProjectWise, MCP, Graph, prepend (`qcPrepend.mode`, `legacy/prepend_qc.ps1`)
- `modules/Reporting/QC.Reporting.psm1` SQL reporting helpers

---

## Fresh caller search results

| Symbol / file | Classification | Phase 3 action |
|---------------|----------------|----------------|
| `QC.PackageResolver.psm1` | test-only, docs, internal | Archived |
| `QC.AttributePolicy.psm1` | test-only, docs, internal | Archived |
| `QC.StatePolicy.psm1` | test-only, docs, internal | Archived |
| `QC.PackageSync.psm1` | test-only, docs, internal | Archived |
| `QC.Package.Database.psm1` | test-only, docs | Archived |
| `test/test_qc_package_model.ps1` | test-only | Archived |
| `Resolve-QCPackage`, `Write-QCPackageCache`, etc. | internal / archived test | Moved with cluster |
| `New-QCPackage` | not found in repo | N/A |
| `New-QCPackageJobDedupeKey`, `metadata.package` in `QC.JobFactory.psm1` | production module, dead branch | **Unchanged** (deferred) |
| `New-QCPackageReportingMetrics`, `Get-QCPackageReportingRows` in `QC.Reporting.psm1` | production (SQL views) | **Unchanged** |
| `docs/architecture/qc-package-model.md` | documentation | Rewritten for SQL |
| `docs/engineering/dead-code-and-deprecation-audit.md` | documentation | Unchanged |

**Production import chain:** `test/test_watcher_module_bootstrap.ps1` does not import any archived module. No production script or module imported `QC.Package*` before archive.

---

## Why production behavior is unchanged

1. **Zero production callers** — Phase 2 confirmed and Phase 3 re-verified: only the archived test imported the in-memory cluster.
2. **SQL path untouched** — `Core.Database.psm1`, `sheet_packages`, `sheet_documents`, `sheet_package_qc_pdfs`, and all production consumers remain in place.
3. **Watcher bootstrap unchanged** — import chain still loads Core + QC pipeline modules without `QC.Package*`.
4. **Job factory unchanged** — queue dedupe keys and `metadata.package` branch behavior identical (branch still unwired).
5. **No schema or view changes** — legacy `qc_pdf_name` / `qc_review_type` columns retained.

---

## SQL package model tests retained

All SQL-backed tests under `test/` remain in place, including:

- `test_sheet_package_resolution.ps1`
- `test_sheet_package_incomplete.ps1`, `test_sheet_package_dual_write.ps1`, `test_sheet_package_backfill.ps1`
- `test_sheet_package_qc_pdfs_schema.ps1`, `test_sheet_package_phase4.ps1`
- `test_qc_lane_state_independence.ps1`, `test_qc_lane_pdf_guid_sync.ps1`
- `test_qc_notification_guid_resolve.ps1`
- Additional lane/trigger/completion tests listed in `docs/archive/phase/phase-2-package-model-decision.md`

`test/test_qc_package_model.ps1` was **not** in `test/run_focus_tests.ps1`; no runner update required.

---

## Post-change validation

| Suite | Result |
|-------|--------|
| `python -m pytest tests/ -q` | **86 passed, 3 skipped** |
| `test/run_focus_tests.ps1` | **ALL PASSED** |
| `test/test_module_inventory.ps1` | **PASS** (41 modules) |
| `test/test_watcher_module_bootstrap.ps1` | **PASS** |
| `test/test_sheet_package_resolution.ps1` | **FAIL** — pre-existing: QC PDF stem `-qc` assertion (unrelated to archive) |
| `test/test_sheet_package_qc_pdfs_schema.ps1` | **FAIL** — pre-existing: schema target version 1.19.0 (environment/SQL fixture) |
| `test/test_qc_lane_state_independence.ps1` | **PASS** |
| `test/test_qc_notification_guid_resolve.ps1` | **PASS** |
| `archive/package-model-v1/test/test_qc_package_model.ps1` (optional) | **FAIL** — pre-existing `Resolve-QCPackage` / `_QCPR-IsBlank` issue on Windows PS 5.1 (reproduces with dev module copy; not in `run_focus_tests.ps1`) |

---

## Rollback plan (targeted)

Do **not** use `git checkout dev -- modules/` — that would restore the entire modules folder.

**Option A — revert the Phase 3 commit(s):**

```powershell
git revert <phase-3-commit-sha>
```

**Option B — targeted file restore from `dev`:**

```powershell
git checkout dev -- `
  modules/QC.PackageResolver.psm1 `
  modules/QC.AttributePolicy.psm1 `
  modules/QC.StatePolicy.psm1 `
  modules/QC.PackageSync.psm1 `
  modules/QC.Package.Database.psm1 `
  test/test_qc_package_model.ps1 `
  modules/FILES.md `
  modules/README.md `
  docs/architecture/qc-package-model.md `
  docs/archive/phase/phase-2-package-model-decision.md
```

Then remove `archive/package-model-v1/` and `docs/archive/phase/phase-3-package-model-archive-summary.md` if no longer wanted.

Re-run baseline validation after rollback.

---

## Remaining follow-up

**QC.JobFactory package dedupe branch:** `Get-QCDedupeKey` can use `job.metadata.package` via `New-QCPackageJobDedupeKey`, but no production code populates that metadata today. Product decision pending:

- Wire dedupe to `sheet_package_id` at job creation, **or**
- Remove the dead `metadata.package` branch and `New-QCPackageJobDedupeKey`

This was intentionally deferred in Phase 3 to avoid changing production queue identity behavior.
