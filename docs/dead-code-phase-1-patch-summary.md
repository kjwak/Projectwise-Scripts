# Phase 1 Dead Code Cleanup — Patch Summary

**Date:** 2026-06-22  
**Basis:** `docs/dead-code-and-deprecation-audit.md` §9 (Phase 1 patch candidates only)  
**Branch:** `cleanup/dead-code-phase-1` (create from current working tree if not already checked out)

---

## Files changed

| File | Change |
|------|--------|
| `modules/Orchestrator.Pipeline.psm1` | **Deleted** |
| `modules/Core.Metrics.psm1` | **Deleted** |
| `void/prepend_watcher.ps1` | **Deleted** |
| `void/test_description_update.ps1` | **Deleted** |
| `void/dump_pw_powershell_inventory.ps1` | **Deleted** |
| `modules/Core.Config.psm1` | Removed unused `Get-AppSetting` function |
| `modules/QC.StatusSet.psm1` | Merged duplicate `Export-ModuleMember` into one block (export set preserved) |
| `modules/QC.Notifications.psm1` | Updated 3 comments: `*-qc.pdf` → lane QC PDF wording |
| `modules/README.md` | Removed `Core.Metrics`, `Orchestrator.Pipeline`; fixed `Get-AppSetting` export list |
| `modules/FILES.md` | Removed `Core.Metrics`, `Orchestrator.Pipeline`; added 5 missing module entries; updated cleanup note |
| `test/run_focus_tests.ps1` | Removed duplicate `test_audit_watch_match.ps1` entry |
| `tests/test_qc_workflow_config_defaults.py` | TYPSA state/review-type assertions |
| `tests/test_qc_notifications_config.py` | `Originated`/`Verified` event assertions; `qcProcessType` dedupe field |
| `tests/test_sheet_index_attr_sync_static.py` | Schema version `1.21.0`; watcher audit-skip message |
| `tests/test_audit_events_db_static.py` | Audit poller ingest comment assertion |
| `tools/discovery/Test-PWRenditionRequest.ps1` | Comment no longer references deleted `void/` script |

---

## Items removed

| Item | Why safe |
|------|----------|
| `Orchestrator.Pipeline.psm1` | Fresh repo search: no `Import-Module`, dot-sourcing, or `Publish-QCPipelineCode.ps1` copy target. Only self-references and documentation. |
| `Core.Metrics.psm1` | Fresh repo search: never imported; only `modules/README.md` / docs referenced it. |
| `void/` (3 scripts) | Abandoned experiments; no production imports. One discovery script comment updated. |
| `Get-AppSetting` | Fresh repo search: defined only in `Core.Config.psm1`; zero callers. `Export-ModuleMember -Function *` still exports remaining functions. |

---

## Export consolidation (`QC.StatusSet.psm1`)

Removed the early `Export-ModuleMember` at line ~1495 and extended the final block to include the four legacy-named helpers that were only exported in the first block:

- `Get-StatusSetManifestPathLegacy`
- `Get-StatusSetCacheDirLegacy`
- `Read-StatusSetManifestLegacy`
- `Write-StatusSetManifestLegacy`

Final export set = union of both prior blocks (21 functions). No runtime behavior change.

---

## Tests run

### Before Phase 1 (baseline on `dev`, audit session)

| Suite | Result |
|-------|--------|
| `pytest tests/` | **77 passed, 9 failed, 3 skipped** |
| `test/run_focus_tests.ps1` | **18 passed, 3 failed** (duplicate `test_audit_watch_match` counted twice) |
| `test/test_module_inventory.ps1` | **Not run in baseline** |
| `test/test_watcher_module_bootstrap.ps1` | **Passed** (via focus suite) |

**Baseline pytest failures (all stale static expectations):**

- `test_qc_workflow_config_defaults.py` (4)
- `test_qc_notifications_config.py` (2)
- `test_sheet_index_attr_sync_static.py` (2)
- `test_audit_events_db_static.py` (1)

### After Phase 1 (initial patch)

| Suite | Result |
|-------|--------|
| `pytest tests/` | **86 passed, 0 failed, 3 skipped** |
| `test/run_focus_tests.ps1` | **18 passed, 3 failed** |
| `test/test_module_inventory.ps1` | **Failed** — `FILES.md` missing 5 on-disk modules (pre-existing drift) |
| `test/test_watcher_module_bootstrap.ps1` | **Passed** |

### After `FILES.md` inventory sync (pre-commit)

| Suite | Result |
|-------|--------|
| `pytest tests/` | **86 passed, 0 failed, 3 skipped** |
| `test/run_focus_tests.ps1` | **18 passed, 3 failed** |
| `test/test_module_inventory.ps1` | **Passed** (46 modules) |
| `test/test_watcher_module_bootstrap.ps1` | **Passed** |

**`modules/FILES.md` entries added:**

- `Core.Telemetry.psm1`
- `QC.DebugMcp.psm1`
- `QC.NotificationThreads.psm1`
- `QC.ProcessType.psm1`
- `QC.WatcherAlerts.psm1`

---

## Tests still failing (after Phase 1 + FILES.md sync)

| Test | Failure | Notes |
|------|---------|-------|
| `test/test_queue_json.ps1` | Stale job recovery assertion | Pre-existing; environmental/timing (`ASSERT FAILED: Recovery should requeue stale job-b`) |
| `test/test_qc_workflow.ps1` | `initialQcPdf -> Ready for QC` vs expected `Originated` | Pre-existing module-default vs test-config mismatch; **not modified in Phase 1** |
| `test/test_audit_poll_window.ps1` | Watermark overlap assertion | Pre-existing timing/watermark sensitivity |

---

## Tests not run

| Suite | Reason |
|-------|--------|
| Full `test/*.ps1` inventory (~100 scripts) | Out of Phase 1 scope; focus suite used as regression gate |
| Integration tests requiring SQL Server / ProjectWise / credentials | Not available in local audit environment |

---

## Rollback

Single revert of the Phase 1 commit restores:

- `modules/Orchestrator.Pipeline.psm1`
- `modules/Core.Metrics.psm1`
- `void/` scripts
- `Get-AppSetting`
- Test and doc edits

No database, config, or workflow behavior was changed.

```powershell
git revert <phase-1-commit-sha>
# or
git checkout dev -- modules/Orchestrator.Pipeline.psm1 modules/Core.Metrics.psm1 void/ ...
```

---

## Items deliberately skipped (not safe for Phase 1)

| Item | Reason |
|------|--------|
| `legacy/` scripts | Production `qcPrepend.mode: legacyPw` |
| `qcPrepend.mode` / native prepend path | Forbidden |
| Disabled notification event keys (`QC Received`, etc.) | Forbidden — code may still enumerate |
| `qc_pdf_name` / `qc_review_type` DB columns | Forbidden — still referenced |
| `QC.Package*` modules | Phase 2 — product decision required |
| `docs/qc-workflow-framework.md` full rewrite | Beyond comment-only scope; large doc drift |
| `Read-AppConfig` → `Read-QCAppSettings` consolidation | Phase 2 |
| `test/test_qc_workflow.ps1` updates | Runtime module defaults still emit `Ready for QC` for empty config merge — fixing requires production code change (forbidden) |

---

## Verification commands

```powershell
python -m pytest tests/ -q
.\test\run_focus_tests.ps1
.\test\test_module_inventory.ps1
.\test\test_watcher_module_bootstrap.ps1
```
