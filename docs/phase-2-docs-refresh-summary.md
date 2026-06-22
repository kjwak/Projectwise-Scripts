# Phase 2 Branch 1 — TYPSA Workflow Docs Refresh Summary

**Branch:** `phase-2/docs-typsa-workflow-refresh`  
**Date:** 2026-06-22  
**Scope:** Documentation only (no production code, tests, or config changes).

---

## Docs changed

| File | Change summary |
|------|----------------|
| `docs/qc-workflow-framework.md` | Full rewrite to TYPSA three-lane model, lifecycle states, `QC_Process_Type`, legacy section |
| `docs/appsettings-reference.md` | TYPSA defaults for `qcWorkflow`, `auditPoller`, `qcRendition`, `qcPrepend`; lane PDF wording |
| `docs/qc-notifications.md` | Lane QC PDF authority, TYPSA event keys, legacy disabled events, Graph integration |
| `docs/qc-reporting.md` | TYPSA reporting buckets, `sheet_package_qc_pdfs`, legacy label note |
| `docs/hybrid-polling.md` | `Initiate Origination` prepend trigger, lane PDF sync, process-type attr propagation |
| `docs/database-telemetry.md` | Lane QC PDF references, `Originated` in aging view |
| `docs/audit-trail-architecture.md` | “As built” notifications row updated for lane PDFs |
| `docs/qc-comment-status-sync.md` | Lane PDF targets, legacy `*-qc.pdf` trigger note |
| `docs/pw-environment-email-attributes.md` | TYPSA state examples, legacy prepend sync section |
| `docs/qc-package-model.md` | Banner: SQL `sheet_packages` is canonical; in-memory modules unwired |
| `legacy/README.md` | `legacyPw` production default, lane naming, prepend relevance |
| `modules/README.md` | TYPSA three-lane pointer to workflow framework doc |
| `docs/phase-2-docs-refresh-summary.md` | This deliverable |

---

## Old terminology — removed or marked legacy

| Old term | Treatment |
|----------|-----------|
| `QC Initiated` | Replaced with **Initiate Origination** (current); old name noted as legacy in workflow doc |
| `Ready for QC` | Replaced with **Originated** for TYPSA; retained as disabled notification key / legacy config key |
| `QC Complete` | Replaced with **Verified**; retained as disabled notification key |
| `QC Received` | Replaced with **Originated**; retained as disabled notification key |
| `In Production` | Replaced with **In Development** |
| `Production QC` / `Peer Review` / `Independent Check` | Replaced with **Production** / **Review** / **Check**; legacy labels documented |
| Single authoritative `*-qc.pdf` | Marked **legacy bridge**; current authority is lane PDFs per `QC_Process_Type` |
| `Independent Check` (routing) | Replaced with **Check** process type |

---

## Current canonical workflow summary

| Dimension | Current (committed `appsettings.json`) |
|-----------|----------------------------------------|
| **Workflow** | TYPSA QC |
| **States** | `In Development` → `Initiate Origination` → `Originated` → `Redlines Received` → `Initiate Verification` → `Ready for Verification` → `Verified` |
| **Process types** | `Production`, `Review`, `Check` via `QC_Process_Type` |
| **Lane PDFs** | `*-prod.pdf`, `*-rev.pdf`, `*-chk.pdf` |
| **Prepend** | `qcPrepend.mode: legacyPw` → `legacy/prepend_qc.ps1` |
| **Sibling sync** | Lane-independent by default (`EnableLegacySiblingStateSync: false`) |
| **Notifications (enabled)** | `Originated`, `Redlines Received`, `Ready for Verification`, `Verified`, `Error Needs Attention` |
| **Package model (production)** | SQL `sheet_packages` / `sheet_package_qc_pdfs` via `Core.Database.psm1` |

### Legacy compatibility (still supported in code)

- `*-qc.pdf` naming bridge in legacy prepend and normalization helpers
- Disabled notification keys: `QC Received`, `Ready for QC`, `QC Complete`
- `reviewTypes` config keys mapping to same three lane values
- `legacy/prepend_qc.ps1` remains production-relevant while `qcPrepend.mode: legacyPw`

---

## Docs intentionally left historical/stale

| Doc | Reason |
|-----|--------|
| `docs/dead-code-and-deprecation-audit.md` | Intentional drift inventory; not rewritten |
| `docs/dead-code-phase-1-patch-summary.md` | Phase 1 changelog; historical |
| `docs/status-set-av-refactor.md` | Historical refactor notes |
| `docs/watcher-architecture-refactor-summary.md` | Historical refactor notes |
| `docs/scripts-organization-review.md` | Historical review |
| `docs/audit-trail-architecture.md` (body) | Design history; only “as built” table row updated |
| `docs/qc-package-model.md` (body) | Full SQL rewrite deferred to Branch 3 |

---

## Validation results

| Suite | Result |
|-------|--------|
| `python -m pytest tests/ -q` | **86 passed, 3 skipped** |
| `test/run_focus_tests.ps1` | **ALL PASSED** |
| `test/test_module_inventory.ps1` | **PASS** (46 modules) |
| `test/test_watcher_module_bootstrap.ps1` | **PASS** |

**Note:** One pytest failure on first pass (`test_notifications_doc_covers_qc_pdf_authority_and_graph`) was resolved by restoring the phrase “not synchronized” in `docs/qc-notifications.md` (static doc guard; no test file edits).

---

## Rollback

```powershell
git revert <branch-1-commit-sha>
```

No runtime behavior changed.
