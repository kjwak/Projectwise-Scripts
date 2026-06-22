# Dead Code and Deprecation Audit

**Date:** 2026-06-22  
**Scope:** Full-repository audit (read-only). No production code was modified during this pass.  
**Canonical design baseline:** Three-lane QC (`Production` / `Review` / `Check` via `QC_Process_Type`), lane PDFs `*-prod.pdf` / `*-rev.pdf` / `*-chk.pdf`, TYPSA workflow states (`In Development` → `Initiate Origination` → `Originated` → `Redlines Received` → `Initiate Verification` → `Ready for Verification` → `Verified`).

---

## 1. Executive summary

### Overall level of accumulated dead or duplicate code

The repository carries **moderate-to-high** accumulated compatibility and parallel-implementation debt from a multi-year migration:

- **Single-lane `*-qc.pdf` → multi-lane `*-prod/-chk/-rev.pdf`**
- **Old PW state vocabulary → TYPSA state names**
- **Monolithic scripts → modular queue pipeline**
- **In-memory package model (`QC.Package*`) → SQL `sheet_packages` model**

Much of the old behavior remains as **fallback branches, default literals, disabled config entries, tests, and documentation** even where production config has moved forward.

### Highest-value cleanup opportunities

| Priority | Area | Impact |
|----------|------|--------|
| 1 | **Documentation + static Python tests** still describing `QC Initiated`, `Ready for QC`, `*-qc.pdf` | Removes false signals; unblocks CI confidence |
| 2 | **`QC.Notifications.psm1` default event keys** and `_QCN-NormalizeQcPdfDocumentName` bridge | Reduces dual-path notification routing |
| 3 | **`Get-QCWorkflowSettings` hardcoded fallbacks** (`Ready for QC`, `Production QC`) in `QC.Workflow.psm1` | Aligns runtime defaults with `appsettings.json` |
| 4 | **Unwired `QC.Package*` module cluster** vs active `Core.Database` `sheet_packages` | Eliminates a competing package abstraction |
| 5 | **`Orchestrator.Pipeline.psm1`, `Core.Metrics.psm1`, `void/`** | Low-risk deletions |
| 6 | **Config loader duplication** (`Core.Config` vs `Core.Runtime`) | Simplifies bootstrap |

### Highest-risk areas

| Risk | Why |
|------|-----|
| **`qcPrepend.mode: "legacyPw"`** (committed default) | Production `QC_PREPEND` still invokes `legacy\prepend_qc.ps1`. Removing `legacy/` or flipping mode without parity validation breaks prepend. |
| **`legacy\prepend_qc.ps1` + `PW.Discovery` sync helpers** | Active callers include `Sync-PWQcPdfEmailAttributesFromSourcePdf`, `Sync-PWQcPdfReviewTypeFromSourcePdf`, and legacy sibling-sync paths gated by flags. |
| **Database columns `qc_pdf_name`, `qc_review_type`** | Still read/written by `sheet_index`, `sheet_packages`, notifications, and backfill SQL. Lane table `sheet_package_qc_pdfs` is canonical for new logic but old columns are not yet retired. |
| **Disabled notification events** (`QC Received`, `Ready for QC`, `QC Complete`) | Code still registers default handlers for these keys; removal requires verifying no environment re-enables them. |
| **ProjectWise Rules Engine / scheduled scripts** | External entry points (`Publish-QCPipelineCode.ps1` deploys scripts) cannot be classified unused from repo references alone. |

### Competing implementations

**Yes.** The repository currently has multiple implementations for several responsibilities:

| Responsibility | Canonical (active) | Competing / legacy |
|----------------|-------------------|-------------------|
| QC prepend execution | `legacy\prepend_qc.ps1` via `qcPrepend.mode=legacyPw` | Native path in `QC.Processors.psm1` (lines 1571+) |
| Package grouping | SQL `sheet_packages` / `sheet_package_qc_pdfs` via `Core.Database.psm1` | In-memory `QC.PackageResolver` / `QC.PackageSync` / `QC.Package.Database` (test-only) |
| Config loading | `Read-QCAppSettings` (`Core.Runtime.psm1`) — all pipeline scripts | `Read-AppConfig` (`Core.Config.psm1`) — 3 diagnostic scripts |
| Watcher orchestration | `Watch-QCTrigger.ps1` + `QC.WatcherOrchestration.psm1` | `Orchestrator.Pipeline.psm1` (unreferenced) |
| Status set build | `statusSet.mode=native` → `QC.StatusSet.psm1` | `legacy\combine_status_set.ps1` when `mode=legacy` |
| Overlay exe resolution | `QC.ReviewStamp.psm1` | `legacy\Resolve-OverlayExe.ps1` (still dot-sourced) |

---

## 2. Active architecture

### Production entry points

```
scripts/Start-QCPipelineDashboard.ps1
  ├── scripts/Watch-QCTrigger.ps1     (audit poll → triggers → enqueue)
  └── scripts/Run-QCProcessor.ps1   (dequeue → dispatch processor)
```

Root shims (`Start-QCPipelineDashboard.ps1`, `Watch-QCTrigger.ps1`, etc.) forward to `scripts/`.

### Canonical execution paths by responsibility

| Responsibility | Canonical implementation | Entry / trigger |
|----------------|-------------------------|-----------------|
| **Audit polling** | `PW.AuditPoller.psm1` → `Invoke-AuditTrailScan` | `Watch-QCTrigger.ps1` via `QC.WatcherOrchestration.psm1` |
| **Trigger evaluation** | `QC.AuditTriggers.psm1`, `QC.Triggers.psm1` | Watcher after audit ingest |
| **Job creation** | `QC.JobFactory.psm1` | Watcher enqueue path |
| **Queue** | `QC.Queue.Json.psm1` | Worker + watcher |
| **QC prepend** | `QC.Processors.psm1` → `Invoke-QCPrependProcessor` → **`legacy\prepend_qc.ps1`** (current default) | `QC_PREPEND` jobs |
| **Status set** | `QC.StatusSet.psm1` → `Invoke-StatusSetNativeJob` | `STATUS_SET_GEN` jobs |
| **Comment sync** | `QC.CommentStatusProcessor.psm1` | `QC_COMMENT_STATUS_SYNC` jobs |
| **Notifications** | `QC.Notifications.psm1` → `QC_NOTIFICATION` jobs / inline enqueue | Workflow transitions, comment sync |
| **Graph delivery** | `QC.NotificationGraph.psm1` | Notification processor |
| **Workflow writeback** | `QC.Workflow.psm1` → `Invoke-QCWorkflowWriteback` | Post-prepend, audit state changes |
| **Lane / process type** | `QC.ProcessType.psm1` | Prepend, stamps, triggers, notifications |
| **Post-prepend lane split** | `PW.Discovery.psm1` → `Sync-PWPostInitialPrependLaneStates` | `Set-PWQCWorkflowState` for `initialQcPdf` |
| **Database telemetry** | `Core.Database.psm1` | Watcher, worker, processors |
| **MCP diagnostics** | `QC.DebugMcp.psm1` + `tools/pw-qc-mcp/server.py` (gitignored) | Cursor MCP registration |
| **Reporting** | `QC.Reporting.psm1` → `QC_REPORTING_SCAN` | Enqueued scan jobs |
| **Rendition** | `QC.Rendition.psm1` | `QC_RENDITION` jobs (disabled in committed config) |
| **Maintenance** | `scripts/Invoke-QCDatabaseRetention.ps1`, etc. | Task Scheduler / manual ops |

### Watcher modes (from `QC.WatcherOrchestration.psm1`)

Committed `appsettings.json` uses hybrid audit-driven polling with reconciliation hooks. Modes: `audit_only`, `reconciliation`, `recovery`, `hybrid`.

### Sibling state sync status

- `QCProcess.EnableLegacySiblingStateSync: false` (committed)
- `Test-QCLegacySiblingStateSyncEnabled` defaults **false** (`QC.ProcessType.psm1`)
- Lane-independent behavior is canonical; legacy sync code remains behind flags

---

## 3. Confirmed dead code

Items with **no in-repo callers** and **no known external invocation**. Risk: **Low** (still confirm no external dot-sourcing).

### 3.1 `modules/Orchestrator.Pipeline.psm1`

| Field | Value |
|-------|-------|
| **Lines** | 1–222 |
| **What it does** | Defines `Invoke-QCPipelineTick` and `Invoke-QCWorkerTick` as composable pipeline ticks over `Resolve-WatchPaths` / `Get-PWTriggerCandidates`. |
| **Why obsolete** | Never imported. Production uses `Watch-QCTrigger.ps1` and `Run-QCProcessor.psm1` directly. |
| **Replacement** | Active watcher/worker scripts + `QC.WatcherOrchestration.psm1`. |
| **Callers** | None (grep: only self-references). |
| **Risk** | Low |
| **Action** | **Safe to delete** after confirming no external automation references the module path. |
| **Tests** | None dedicated; no updates required. |
| **External** | Confirm `Publish-QCPipelineCode.ps1` deployment targets do not import it. |

### 3.2 `modules/Core.Metrics.psm1`

| Field | Value |
|-------|-------|
| **Lines** | 1–189 |
| **What it does** | In-memory metrics counters (`Initialize-QCMetrics`, `Get-QCMetricsSnapshot`, etc.). |
| **Why obsolete** | Never imported anywhere in the repository. |
| **Replacement** | `Core.Telemetry.psm1` + SQL telemetry tables. |
| **Callers** | None. |
| **Risk** | Low |
| **Action** | **Safe to delete**. |
| **Tests** | None. |

### 3.3 `void/` directory

| File | What it does | Why obsolete |
|------|-------------|--------------|
| `void/prepend_watcher.ps1` | Early 32-bit PW + qpdf prepend experiment | Superseded by modular pipeline |
| `void/test_description_update.ps1` | PW description update probe | One-off; referenced only in a discovery script comment |
| `void/dump_pw_powershell_inventory.ps1` | Generates PW cmdlet dumps | Output already committed as `pwps_dump_20260209_125642/` |

| Field | Value |
|-------|-------|
| **Risk** | Low |
| **Action** | **Safe to delete** `void/` (or move dump generator to `tools/discovery/` if still wanted). |
| **Tests** | None |

### 3.4 `Get-AppSetting` (`Core.Config.psm1`)

| Field | Value |
|-------|-------|
| **Lines** | ~110+ |
| **What it does** | Dot-path config value accessor. |
| **Why obsolete** | Defined and exported but **never called** outside `Core.Config.psm1`. |
| **Replacement** | Direct hashtable access after `Read-QCAppSettings` / `Read-AppConfig`. |
| **Risk** | Low |
| **Action** | **Safe to delete** function (keep module if `Read-AppConfig` retained). |

### 3.5 Duplicate focus-test entry

| Field | Value |
|-------|-------|
| **File** | `test/run_focus_tests.ps1` lines 19–21 |
| **What** | `test_audit_watch_match.ps1` listed twice |
| **Action** | **Safe to delete** duplicate line |

---

## 4. Superseded implementations

Old implementations **still present** alongside replacements.

### 4.1 QC prepend: legacy script vs native processor

| | Legacy (active default) | Native (superseded-in-config but implemented) |
|---|------------------------|-----------------------------------------------|
| **Path** | `legacy/prepend_qc.ps1` | `QC.Processors.psm1` lines 1571+ |
| **Config** | `qcPrepend.mode: "legacyPw"` | Set `mode` to non-`legacyPw` (e.g. `local`) |
| **Callers** | `Invoke-QCPrependProcessor` when `mode -eq 'legacypw'` | Same function, else branch |
| **Risk of removal** | **Critical** — committed production default |
| **Action** | **Keep** until native path reaches parity and config is flipped |
| **Tests** | `test/test_qc_prepend_stamp_integration.ps1` uses legacy fixtures; add native-path integration tests before migration |

### 4.2 Status set: legacy combine script

| Field | Value |
|-------|-------|
| **Legacy** | `legacy/combine_status_set.ps1` |
| **Canonical** | `QC.StatusSet.psm1` / `Invoke-StatusSetNativeJob` |
| **Config** | `statusSet.mode: "native"` (committed) |
| **Risk** | Medium if any environment sets `mode: "legacy"` |
| **Action** | **Likely obsolete** for committed config; confirm all deployed environments use `native` |

**Note:** `QC.StatusSet.psm1` functions named `*Legacy` (`Get-StatusSetManifestPathLegacy`, etc.) are **misnamed** — they are called by the **native** job path (lines 1119+), not only by `combine_status_set.ps1`. Rename/consolidate in Phase 2, do not delete.

### 4.3 Monolithic watcher: `prepend_qc_on_trigger.ps1`

| Field | Value |
|-------|-------|
| **Path** | `legacy/prepend_qc_on_trigger.ps1` |
| **Replacement** | `scripts/Watch-QCTrigger.ps1` + queue |
| **Still invoked** | `scripts/run_prepend_qc.ps1 -Legacy` |
| **Action** | **Keep** until `-Legacy` path is retired |

### 4.4 In-memory package model vs SQL package model

| | `QC.Package*` modules | `Core.Database` `sheet_packages` |
|---|----------------------|----------------------------------|
| **Modules** | `QC.PackageResolver`, `QC.PackageSync`, `QC.Package.Database`, `QC.AttributePolicy`, `QC.StatePolicy` | `Ensure-SheetPackage`, `Resolve-SheetPackageFromDocument`, `sheet_package_qc_pdfs` |
| **Production use** | **None** — only `test/test_qc_package_model.ps1`, `test/test_sheet_package_phase4.ps1` | Watcher, notifications, MCP, lane resolution |
| **Docs** | `docs/qc-package-model.md` describes in-memory model as current | `docs/database-telemetry.md` describes SQL model |
| **Action** | **Likely obsolete** as runtime code; **consolidate** docs to SQL model or wire modules to DB |
| **Risk** | Medium — docs mislead maintainers |

### 4.5 Config loaders: `Core.Config` vs `Core.Runtime`

| Function | Module | Callers |
|----------|--------|---------|
| `Read-QCAppSettings` | `Core.Runtime.psm1` | All pipeline scripts, most tests |
| `Read-AppConfig` | `Core.Config.psm1` | `Scan-PWProjectMetrics.ps1`, `Get-PWFolderStateCounts.ps1`, `Test-BluebeamCommentExtractAndDb.ps1` |

Both duplicate `_ToHashtableDeep` logic. **Consolidate** on `Read-QCAppSettings`.

### 4.6 Overlay exe resolver

| Field | Value |
|-------|-------|
| **Legacy** | `legacy/Resolve-OverlayExe.ps1` |
| **Canonical** | `QC.ReviewStamp.psm1` (line 114 still dot-sources legacy helper) |
| **Also used by** | `legacy/prepend_qc.ps1`, `test/run_f0548dv206_qc_two_step.ps1` |
| **Action** | **Consolidate** into `QC.ReviewStamp.psm1` after prepend migration |

### 4.7 Workflow state vocabulary fallbacks

| Location | Superseded literals | Canonical (appsettings) |
|----------|--------------------|-----------------------|
| `QC.Workflow.psm1` `Get-QCWorkflowSettings` defaults | `Ready for QC`, `Production QC` | `Originated`, `Production` |
| `QC.WatcherOrchestration.psm1` line ~143 | `'QC Initiated'` fallback | `Initiate Origination` |
| `QC.Notifications.psm1` default event keys | `QC Received`, `Ready for QC`, `QC Complete` | `Originated`, `Verified` |
| `PW.Discovery.psm1` `_PWD-GetConfiguredWorkflowStateNames` | `In Production`, `QC Initiated`, etc. | Configured TYPSA names |
| `QC.Reporting.psm1` metrics buckets | `Corrections Received`, `Verification In Progress` | TYPSA states |

**Action:** **Simplify** — remove fallbacks once all environments use TYPSA names (Phase 3).

### 4.8 Single-lane `*-qc.pdf` handling

| Field | Value |
|-------|-------|
| **Superseded** | `*-qc.pdf` as authoritative QC document |
| **Canonical** | `*-prod.pdf`, `*-chk.pdf`, `*-rev.pdf` |
| **Remaining** | `Test-QCLegacyQcPdfDocumentName`, `_QCN-NormalizeQcPdfDocumentName`, `Sync-PWQcPdfReviewTypeFromSourcePdf`, `sheet_index.qc_pdf_name`, cleanup script `Remove-LegacyQcPdfDatabaseRows.ps1` |
| **Action** | **Keep** bridge functions until DB/PW data fully migrated; then remove |

---

## 5. Duplicate logic

### 5.1 Hashtable deep-conversion (4+ copies)

| Location | Function |
|----------|----------|
| `Core.Runtime.psm1` | `ConvertTo-HashtableDeep` (canonical) |
| `Core.Config.psm1` | `_CC-ToHashtableDeep` |
| `scripts/Scan-PWProjectMetrics.ps1` | `_ToHashtable` |
| `scripts/Get-PWFolderStateCounts.ps1` | `_ToHashtable` |
| `QC.WatcherAlerts.psm1` | `_QCWA-ToHashtable` |
| `QC.PackageResolver.psm1` | `_QCPR-ToHashtable` |

**Recommendation:** **Consolidate** on `ConvertTo-HashtableDeep`.

### 5.2 Logging fallbacks

Repeated pattern across `QC.AttributePolicy`, `QC.StatePolicy`, `QC.PackageResolver`, `QC.PackageSync`, `QC.Package.Database`:

```powershell
if (Get-Command Write-QCJsonLog ...) { ... }
elseif (Get-Command Write-QCLog ...) { ... }
```

`Core.Logging.psm1` is imported only as optional fallback; primary path is `Write-QCJsonLog` everywhere.

**Recommendation:** **Simplify** to `Write-QCJsonLog` only, or make `Core.Logging` a thin alias.

### 5.3 QC PDF name normalization

| Function | Module | Role |
|----------|--------|------|
| `Get-QCLaneQcPdfExpectedName` | `QC.ProcessType.psm1` | Canonical lane naming |
| `_QCN-NormalizeQcPdfDocumentName` | `QC.Notifications.psm1` | Bridge stem/legacy → lane PDF |
| `Get-PWQcPdfLaneFromDocumentName` | `QC.ProcessType.psm1` | Suffix → lane |

**Recommendation:** **Consolidate** notification bridge into `QC.ProcessType.psm1`.

### 5.4 Associated sheet member lists

| Function | Purpose |
|----------|---------|
| `Get-PWAssociatedSheetDocumentNames` | Canonical lane PDFs + stem + DGN |
| `Get-PWAssociatedSheetSyncDocumentNames` | Legacy sibling sync (sheet + `-prod` only) |

**Recommendation:** **Keep** both while legacy flag exists; document clearly.

### 5.5 `QC.StatusSet.psm1` duplicate `Export-ModuleMember`

Lines 1495 and 3239 each export a overlapping function set. PowerShell uses the **last** export block; the first block partially shadows exports.

**Recommendation:** **Consolidate** to a single `Export-ModuleMember` (Phase 1 safe if exports are merged correctly).

### 5.6 Watcher script-local helpers

`Watch-QCTrigger.ps1` still contains ~20 `_Watch-*` / cache helpers that `docs/scripts-organization-review.md` flagged for module extraction (Priority 2). Partially addressed by `PW.AuditPoller.psm1`.

---

## 6. Obsolete configuration and schema

### 6.1 `appsettings.json` — disabled or legacy entries

| Key | Status | Evidence |
|-----|--------|----------|
| `qcPrepend.mode: "legacyPw"` | **Active** (not obsolete) | Production prepend path |
| `qcRendition.enabled: false` | Feature **dormant** | Full `QC.Rendition.psm1` stack exists but disabled |
| `qcRendition.deferReadyForQcNotification: false` | Inactive while rendition disabled | |
| `QCProcess.EnableLegacySiblingStateSync: false` | Compatibility flag (off) | Code paths remain |
| `QCProcess.EnableLegacyReviewTypeAttributeSync: false` | Compatibility flag (off) | |
| `notifications.events["QC Received"]` | **Disabled** | Replaced by `Originated` |
| `notifications.events["Ready for QC"]` | **Disabled** | Replaced by `Originated` |
| `notifications.events["QC Complete"]` | **Disabled** | Replaced by `Verified` |
| `qcPrepend.reviewStamps.peerReview` / `independentCheck` | Old review-type naming | Maps to Review/Check stamps; rename in Phase 3 |
| `qcWorkflow.reviewTypes` | Maps old review program names → `QC_Process_Type` values | Still used for `QC_Review_Type` read paths |

### 6.2 Deprecated config keys (warned, not removed)

From `Get-QCWorkflowDeprecationWarnings` (`QC.Workflow.psm1`):

- `qcWorkflow.stageMap` — ignored
- `qcWorkflow.correctionsInProgressStateName` — maps to `states.qcFinalizing`
- `stateAfterPrependByTrigger.reviewerRedlineUpdate` / `designerCorrectionComplete`
- State keys: `reviewInProgress`, `correctionsInProgress`, `verificationInProgress`, `redlinesIssued`, `correctionsReceived`

### 6.3 Schema fields — legacy but still referenced

| Column / object | Status | References |
|-----------------|--------|------------|
| `sheet_index.qc_pdf_name` | Legacy single-lane | Sync scripts, notifications fallback, backfill SQL |
| `sheet_index.qc_review_type` | Superseded by `qc_process_type` | Sync, MCP, backfill SQL |
| `sheet_packages.qc_pdf_name` / `qc_pdf_guid` | Legacy single-lane columns | Views, backfill; lane table added alongside |
| `sheet_package_qc_pdfs` | **Canonical** lane registry | Notifications, prepend resolution |
| `qc_workflow_events.qc_review_type` | Historical telemetry | Reporting |
| `peer_review_completed_count` / `independent_check_completed_count` on `sheet_packages` | Old review-type naming | `QC.Reporting.psm1` |

**Do not drop columns in Phase 1.** Run `scripts/Remove-LegacyQcPdfDatabaseRows.ps1` and backfill scripts only after operational sign-off.

### 6.4 Migrations

Historical migrations in `Core.Database.psm1` `Initialize-QCDatabaseSchema` are **idempotent ALTER** blocks. **Keep** — do not delete migration history.

### 6.5 MCP configuration

`.cursor/mcp.json` registers `pw-qc-debug` → `tools/pw-qc-mcp/server.py` (gitignored). `QC.DebugMcp.psm1` is active and must be retained.

### 6.6 Environment variables

No standalone env-var dispatch layer found; configuration is file-driven via `Read-QCAppSettings` merge chain (`appsettings.json` → `appsettings.local.json` → `appsettings.secrets.json`).

---

## 7. Obsolete tests and documentation

### 7.1 Test suite results (this audit run)

#### Python (`pytest tests/`)

| Result | Count |
|--------|-------|
| Passed | 77 |
| Failed | 9 |
| Skipped | 3 |

**Failed tests (obsolete expectations vs current `appsettings.json`):**

| Test | Why it fails |
|------|--------------|
| `tests/test_qc_workflow_config_defaults.py` (4 tests) | Expects `QC Initiated`, `Ready for QC`, `QC Complete`, `Independent Check`, `Production QC` — committed config uses TYPSA names |
| `tests/test_qc_notifications_config.py` (2 tests) | Expects old lifecycle event keys / delivery defaults |
| `tests/test_sheet_index_attr_sync_static.py` (2 tests) | Static assertions out of date with schema/audit behavior |
| `tests/test_audit_events_db_static.py::test_audit_poller_delegates_to_database_writer` | Static wiring assertion drift |

**Classification:** These failures are **documentation-of-drift**, not necessarily production bugs. Update or remove in Phase 1.

#### PowerShell (`test/run_focus_tests.ps1`)

| Result | Count |
|--------|-------|
| Passed | 18 unique tests |
| Failed | 3 |

| Failed test | Notes |
|-------------|-------|
| `test/test_queue_json.ps1` | Requires investigation — may be environmental path/permission |
| `test/test_qc_workflow.ps1` | May conflict with updated default state names |
| `test/test_audit_poll_window.ps1` | May be timing/watermark sensitive |

**Not run:** Full ~100 script suite (time constraints). Many tests use `*-qc.pdf` fixtures intentionally to test legacy bridge behavior.

### 7.2 Tests encoding obsolete single-lane model

| Test | Obsolete aspect | Still valuable? |
|------|----------------|-----------------|
| `test/test_sheet_group_workflow_transition.ps1` | `*-qc.pdf`, `QC Initiated`, `Ready for QC`, `In Production` | Partially — tests notification dedupe; update fixtures to lane PDFs |
| `test/test_qc_cycle_completion.ps1` | `*-qc.pdf`, `QC Finalizing`, `QC Complete` | Update to TYPSA states |
| `test/test_qc_comment_trigger.ps1` | `-qc.pdf` extension rule | Verify rule still exists or update |
| `test/test_qc_notifications.ps1` | Heavy `*-qc.pdf` usage | Update fixtures |

### 7.3 Tests for unwired package modules

| Test | Modules tested | Production wired? |
|------|---------------|-------------------|
| `test/test_qc_package_model.ps1` | `QC.PackageResolver`, `QC.PackageSync`, `QC.StatePolicy`, `QC.AttributePolicy` | **No** |
| `test/test_sheet_package_phase4.ps1` | SQL views + `QC.Reporting` | Partially (SQL yes, in-memory package no) |

### 7.4 Obsolete documentation

| Document | Issue |
|----------|-------|
| `docs/qc-workflow-framework.md` | States `In Production`, `QC Initiated`, `QC Received`, `QC Complete`, `stageMap` era |
| `docs/qc-notifications.md` | `*-qc.pdf` as authoritative |
| `docs/hybrid-polling.md` | Old state names, single-lane sibling sync |
| `docs/appsettings-reference.md` | Defaults list `In Production`, `Ready for QC` |
| `docs/database-telemetry.md` | `*-qc.pdf` references |
| `docs/qc-package-model.md` | Describes in-memory `Resolve-QCPackage` as current; production uses SQL |
| `modules/README.md` | Lists `Orchestrator.Pipeline` as active module |
| `legacy/README.md` | Correctly notes legacy role but understates that `legacyPw` is production default |

---

## 8. Removal plan

### Phase 1: Low-risk deletions

**Targets:** Unused modules, `void/`, duplicate test list entry, unused exports, obsolete static tests, stale comments.

| Item | Validation | Rollback |
|------|------------|----------|
| Delete `Orchestrator.Pipeline.psm1` | `test/test_module_inventory.ps1`; grep deploy scripts | Restore from git |
| Delete `Core.Metrics.psm1` | Module inventory test | Restore from git |
| Delete `void/` | None | Restore from git |
| Remove duplicate `test_audit_watch_match.ps1` entry | Re-run `run_focus_tests.ps1` | Restore line |
| Remove `Get-AppSetting` | Grep confirms zero callers | Restore function |
| Update/delete 9 failing Python static tests | `pytest tests/` green | Revert test commit |
| Fix `QC.StatusSet.psm1` duplicate export | Status set focus tests | Restore single block |

**Validation steps:**
1. `pytest tests/`
2. `.\test\run_focus_tests.ps1`
3. `.\test\test_module_inventory.ps1`
4. `.\test\test_watcher_module_bootstrap.ps1`

**Rollback:** Single git revert per change batch.

### Phase 2: Consolidation

| Item | Validation |
|------|------------|
| Merge `Read-AppConfig` callers onto `Read-QCAppSettings` | Metrics/scan scripts smoke test |
| Remove `_ToHashtable` copies in scripts | Script dry-runs |
| Consolidate `_QCN-NormalizeQcPdfDocumentName` into `QC.ProcessType.psm1` | Notification tests |
| Rename `QC.StatusSet` `*Legacy` helpers | Status set test suite |
| Resolve `QC.Package*` vs SQL model (delete modules or wire to DB) | Package + notification tests |
| Extract remaining `Watch-QCTrigger.ps1` helpers to modules | Watcher bootstrap test |

**Rollback:** Keep deprecated wrappers for one release with `Write-Warning` shims.

### Phase 3: Compatibility behavior removal

**Prerequisites:** All PW projects on TYPSA states; lane PDFs deployed; `qcPrepend.mode` flipped to native.

| Item | Validation |
|------|------------|
| Remove `EnableLegacySiblingStateSync` code paths | `test_qc_lane_state_independence.ps1` |
| Remove `*-qc.pdf` bridges after DB cleanup | `Remove-LegacyQcPdfDatabaseRows.ps1` dry-run empty |
| Remove disabled notification event definitions + default old keys | Notification routing tests |
| Remove `legacy/prepend_qc.ps1` after native parity | E2E prepend on PW test folder |
| Retire `run_prepend_qc.ps1 -Legacy` | Manual sign-off |

**Rollback:** Feature flags in `appsettings.json` retained until Phase 3 complete.

### Phase 4: Schema / migration cleanup

| Item | Validation |
|------|------------|
| Deprecate `sheet_index.qc_pdf_name` reads (read from `sheet_package_qc_pdfs` only) | MCP compare tools |
| Add migration to stop writing `qc_review_type` where `qc_process_type` suffices | Power BI dashboards |
| Drop views using old columns | BI regression |
| Remove `peer_review_*` column names (rename) | Reporting scan |

**Rollback:** Schema changes via additive migrations only; column drops require DB backup + BI coordination.

---

## 9. Phase 1 patch candidates

Only changes with **demonstrable safety** from in-repo evidence.

| # | File | Item | Classification | Action |
|---|------|------|----------------|--------|
| 1 | `modules/Orchestrator.Pipeline.psm1` | Entire module | Safe to delete | Delete |
| 2 | `modules/Core.Metrics.psm1` | Entire module | Safe to delete | Delete |
| 3 | `void/*` | 3 scripts | Safe to delete | Delete directory |
| 4 | `modules/Core/Core.Config.psm1` | `Get-AppSetting` | Safe to delete | Remove function + export |
| 5 | `test/run_focus_tests.ps1` | Duplicate line 21 | Safe to delete | Remove duplicate |
| 6 | `tests/test_qc_workflow_config_defaults.py` | 4 tests with wrong state names | Obsolete tests | Update assertions to TYPSA names |
| 7 | `tests/test_qc_notifications_config.py` | 2 stale tests | Obsolete tests | Update to `Originated`/`Verified` |
| 8 | `tests/test_sheet_index_attr_sync_static.py` | 2 stale tests | Obsolete tests | Update static expectations |
| 9 | `tests/test_audit_events_db_static.py` | 1 stale test | Obsolete tests | Fix delegate assertion |
| 10 | `modules/Processing/QC.StatusSet.psm1` | Duplicate `Export-ModuleMember` at 1495 | Simplify | Merge into single export at 3239 |
| 11 | `modules/README.md` | Lists `Orchestrator.Pipeline` | Obsolete docs | Remove from inventory |
| 12 | `QC.Notifications.psm1` ~175, ~1704 | Comments referencing `*-qc.pdf` as authoritative | Obsolete comments | Update comment text only |

### Explicitly NOT Phase 1

| Item | Reason |
|------|--------|
| `legacy/prepend_qc.ps1` | `qcPrepend.mode: legacyPw` production default |
| `legacy/combine_status_set.ps1` | Fallback for `statusSet.mode: legacy` |
| `QC.Package*` modules | Documented model; needs product decision |
| Database columns `qc_pdf_name`, `qc_review_type` | Still read in production paths |
| `notifications.events` disabled entries | Code may still enumerate keys |
| MCP tools / `QC.DebugMcp.psm1` | Active diagnostic surface |
| `scripts/Test-PWEmailAttributes*.ps1` (10 files) | Operational discovery probes |
| `pwps_dump_20260209_125642/` | Referenced by cursor rules; archive later |
| Migrations in `Core.Database.psm1` | Never delete historical DDL |
| `Sync-PWQcPdfReviewTypeFromSourcePdf` | Called from `legacy/prepend_qc.ps1` |

---

## Appendix A: Test gaps preventing confident deletion

| Gap | Blocks |
|-----|--------|
| No automated parity test: `legacyPw` vs native `QC_PREPEND` | Removing `legacy/prepend_qc.ps1` |
| No production-path test for `QC.Package*` modules | Deleting package module cluster |
| No CI job running full PowerShell `test/*.ps1` suite | Broad dead-code confidence |
| Python static tests drifted from `appsettings.json` | Config cleanup without test update |
| No inventory of ProjectWise Rules Engine callbacks | Classifying any exported function as unused |
| `qcRendition.enabled: false` — no enabled-config integration test | Removing rendition module |

## Appendix B: External dependencies to confirm before removal

1. **Deployed `appsettings.json` on production workers** — Is `qcPrepend.mode` still `legacyPw` everywhere?
2. **Task Scheduler jobs** — Which `scripts/*.ps1` are registered?
3. **ProjectWise Rules Engine** — Any rules calling `legacy/*.ps1` or root shims directly?
4. **Power BI reports** — Do any query `qc_pdf_name`, `qc_review_type`, or old state literal strings?
5. **`Publish-QCPipelineCode.ps1` target** — Does production worker copy `Orchestrator.Pipeline.psm1`?

---

*End of audit report.*
