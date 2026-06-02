# appsettings reference

The pipeline reads **`appsettings.json`** at the repo root (strict JSON). Use this document for explanations; use **`appsettings.schema.json`** for hover help in Cursor/VS Code.

## How settings are loaded

| Mechanism | Purpose |
|-----------|---------|
| `Read-QCAppSettings` | All scripts and dashboard load config via `modules/Core.Runtime.psm1`. |
| `-AppSettingsPath` | Override path on `Watch-QCTrigger.ps1`, `Start-QCPipelineDashboard.ps1`, etc. |
| `-DryRun` (CLI) | Forces `dryRun: true` even when JSON says false. |
| `$schema` in JSON | IDE only; ignored at runtime. |
| Profile merge | `appsettings.test.json` loads `appsettings.json` then the profile then `appsettings.test.local.json`. `appsettings.json` also merges `appsettings.local.json`. |

**Testing / local overlays:** see [`docs/testing-config.md`](testing-config.md). Copy `appsettings.test.json.example` → `appsettings.test.json` (gitignored).

Optional: keep **machine-only** secrets in `appsettings.local.json` or `appsettings.test.local.json` — do not commit them.

## Section map

| Section | Role |
|---------|------|
| `projectWise` | PW datasource, credentials, watch roots |
| `dryRun` | Global enqueue/processor safety |
| `filters` | Allow/deny paths for watcher |
| `triggers.rules` | Event → job type mapping |
| `queue` | On-disk job folders and locks |
| `processors` | Job handler map |
| `workers` | Parallel worker pool |
| `qcPrepend` | QC_PREPEND overlay paths |
| `statusSet` | STATUS_SET_GEN merge + manifest + PW upload |
| `qcWorkflow` | Optional PW workflow/attributes after QC |
| `qcReporting` | QC_REPORTING_SCAN metrics |
| `watcher` | Audit vs full-scan orchestration |
| `reconciliation` | When to run full folder scans / startup upload |
| `auditPoller` | `dms_audt` polling and n=20 reconcile |
| `qcCommentSync` | Comment-driven workflow sync |
| `database` | SQL telemetry |
| `notifications` | Email on state transitions |

## Operational scripts (not in JSON)

| Script | When to use |
|--------|-------------|
| `scripts/Combine-StatusSet.ps1` | One Sheets folder: rebuild `_StatusSet.pdf` from PW sheets (`-ForceRebuild`). |
| `scripts/Reconcile-QCStatusSets.ps1` | All workspaces: upload **local** `_StatusSet.pdf` if newer than PW (no sheet re-merge). |
| `scripts/Reconcile-QCSheetOwnership.ps1` | Sync designer/reviewer emails and states across DGN / sheet PDF / QC PDF. |

---

## projectWise

| Key | Type | Description |
|-----|------|-------------|
| `datasourceName` | string | Bentley datasource for `Connect-PW`. |
| `credentialPath` | string | PW username/password file (keep out of git). |
| `clearTriggerTagOnSuccess` | bool | Remove `QC_Archivist` from description after successful QC_PREPEND. |
| `fileOpThrottleMs` | int | Minimum ms between PW file operations. |

### watchList.roots[]

Each entry discovers every `…/{sheetsPathFromProject}` folder under `path`:

| Key | Description |
|-----|-------------|
| `path` | PW root (e.g. `Documents\ProjectName`). |
| `sheetsPathFromProject` | Usually `CADD\Sheets`. |
| `projectDepth` | `1` = projects are direct children of `path`; `2` = one level deeper. |
| `enableQcPrepend` | PDFs with `QC_Archivist` in description → `QC_PREPEND`. |
| `enableQcCommentSync` | Audit path may enqueue `QC_COMMENT_STATUS_SYNC` on `*-qc.pdf`. |
| `enableStatusSet` | Folder-level `STATUS_SET_GEN` (paired PDF + DGN). |

Optional: `environmentEmailAttributes` — see `docs/pw-environment-email-attributes.md`.

---

## dryRun

- **`true`**: Watcher logs `WATCH_ACCEPTED` but does not `Add-QCQueueJob`; processor uses `processors.dryRun`.
- **`false`**: Normal operation.

---

## filters

| Key | Description |
|-----|-------------|
| `whitelist.enabled` + `paths` | If enabled, **only** listed paths are processed. |
| `blacklist.paths` | Skip exact paths or `pw:\\datasource\…` URIs. |
| `blacklist.patterns` | Regex against full path. |

---

## triggers.rules

Rules are evaluated by **priority** (lower number wins). Common `jobType` values:

| jobType | Meaning |
|---------|---------|
| `STATUS_SET_GEN` | Merge sheet PDFs → `_StatusSet.pdf` per folder. |
| `QC_PREPEND` | Redline overlay on archivist-tagged PDF. |
| `QC_COMMENT_STATUS_SYNC` | Update PW state from `*-qc.pdf` comments. |

| Key | Description |
|-----|-------------|
| `triggerType` | `pw` (ProjectWise watcher) or `fs` (local filesystem). |
| `when` | Match criteria (extensions, regex, description text). |
| `requireAll` | All listed `when` keys must match. |
| `grouping.groupBy` | `folder` = one job per folder; `file` = per document. |

---

## queue

| Key | Description |
|-----|-------------|
| `rootDir` | `pending`, `running`, `succeeded`, `failed`, `_locks`, `_watcher`. |
| `recover.maxAttempts` | Failed jobs retried up to this count. |
| `lockAcquireTimeoutMs` | Wait for per-job lock file. |
| `selection.preferJobTypes` | Worker picks preferred types first when multiple pending. |

---

## processors

| Key | Description |
|-----|-------------|
| `processorMap` | Job type → PowerShell function name. |
| | `STATUS_SET_GEN` always maps to `Invoke-StatusSetProcessor` (hard-coded fallback). |
| `dryRun.allowStateChange` | Move jobs to succeeded without handler when global dry run. |
| `dryRun.invokeHandler` | Call handler during dry run (handler should skip PW writes). |

---

## workers

| Key | Description |
|-----|-------------|
| `maxParallel` | Number of `Run-QCProcessor.ps1` processes. |
| `maxJobsPerWorker` | Jobs per process before exit. |
| `leaseSeconds` | Stale `running` job recovery threshold. |
| `idleSleepMs` | Sleep when queue empty (also watcher default if `watcher.idleSleepMs` unset). |

---

## qcPrepend

| Key | Description |
|-----|-------------|
| `historyRoot` / `outputRoot` / `tempRoot` | Overlay working directories. |
| `mode` | `legacyPw` = export/check-in via ProjectWise. |
| `qpdfExePath` | Path to qpdf binary. |

---

## statusSet

| Key | Description |
|-----|-------------|
| `mode` | `native` (production) or `stub` (tests). |
| `writeBackToPW` | Upload `_StatusSet.pdf` after build. |
| `forceRebuild` | Ignore manifest hash; always rebuild. |
| `incrementalMode` | Page-level replace when possible. |
| `localRoot` | Per-folder workspace (manifest + PDF). |
| `manifestFileName` | Default `_statusset.manifest.json`. |
| `statusSetPdfName` | Default `_StatusSet.pdf`. |
| `dryRunOperationReport` | Log planned PW ops without executing. |

**Manifest gate:** Watcher skips `STATUS_SET_GEN` when `folderStateHash` in manifest matches live PW (`WATCH_PW_STATUSSET_SKIP_CURRENT`). Wrong PW copy with matching hash requires `-ForceRebuild` or `Combine-StatusSet.ps1`.

---

## qcWorkflow

Optional ProjectWise **document attribute** writeback (and optional **workflow state** changes) after a successful `QC_PREPEND`. Implemented in `modules/QC.Workflow.psm1`; invoked from `QC.Processors.psm1` when prepend completes.

**Full guide (lifecycle states, review types, assignment, PW admin setup, discovery scripts, rollout):** [`docs/qc-workflow-framework.md`](qc-workflow-framework.md)

**Related:** comment-driven state uses the same PW helpers via `qcCommentSync` — see [`docs/qc-comment-status-sync.md`](qc-comment-status-sync.md). SQL audit rows: `qc_workflow_events` in [`docs/database-telemetry.md`](database-telemetry.md).

### When `enabled` is false (default)

Your repo ships with `"enabled": false` (see `appsettings.json` around line 160). Prepend/overlay still runs; **no** PW attributes or states are written. `Invoke-QCWorkflowWriteback` returns immediately with a skipped result. This is the recommended setting until PW attributes and (optional) states are created and validated.

### Enablement order

1. `enabled: false` — production-safe default.
2. `enabled: true`, `dryRunWriteback: true` — log planned writes (`QC_WORKFLOW_*` codes); no PW changes.
3. Run `tools\discovery\Test-QCWorkflowCapabilities.ps1` (read-only).
4. Pilot one document with `tools\discovery\Test-QCWorkflowWriteback.ps1 -ConfirmWrites`.
5. `dryRunWriteback: false` for a pilot project only.

Global `dryRun: true` does **not** replace `qcWorkflow.dryRunWriteback`; both are independent.

### Core keys

| Key | Default | Description |
|-----|---------|-------------|
| `enabled` | `false` | Master switch. When false, workflow writeback is skipped entirely. |
| `strictMode` | `false` | If true, validation/write failures can fail the `QC_PREPEND` job. |
| `dryRunWriteback` | `true` | When true, plan and log writes only (`QC_WORKFLOW_DRYRUN`, `*_PLANNED`). |
| `mode` | `AttributesOnly` | `AttributesOnly` or `StateAndAttributes`. |
| `autoWriteAttributes` | `true` | Call `Update-PWDocumentAttributes` with mapped QC fields after prepend. |
| `autoSetState` | `false` | When `mode` is `StateAndAttributes`, call `Set-PWDocumentState` if true. |
| `expectedWorkflowName` | `""` | Workflow name for validation (alias: `workflowName`). |
| `attributeMap` | see JSON | Logical key → PW environment attribute name (no `QC_Stage`). |
| `states` | see JSON | ProjectWise lifecycle state names (source of truth). |
| `reviewTypes` | `Production QC`, `Peer Review`, `Independent Check` | Allowed `QC_Review_Type` values. |
| `defaultReviewType` | `Production QC` | Used when document/job has no review type. |

### Lifecycle states (`states.*`)

Names must match your ProjectWise workflow:

| Key | Default |
|-----|---------|
| `production` | `In Production` |
| `readyForQc` | `Ready for QC` |
| `reviewInProgress` | `Review In Progress` |
| `redlinesIssued` | `Redlines Issued` |
| `correctionsInProgress` | `Corrections In Progress` |
| `verificationInProgress` | `Verification In Progress` |
| `complete` | `QC Complete` |
| `error` | `Error Needs Attention` |

| Key | Default | Description |
|-----|---------|-------------|
| `defaultStateAfterPrepend` | `Ready for QC` | Informational default when state writeback is enabled later. |
| `stateAfterSuccessfulPrepend` | `Ready for QC` | Fallback when `autoSetState` runs and prepend trigger is unknown. |
| `stateAfterFailedPrepend` | `Error Needs Attention` | Target state after failed prepend. |
| `defaultPrependTrigger` | `initialQcPdf` | Documented default; set `metadata.prependTrigger` on jobs. |
| `stateAfterPrependByTrigger` | see JSON | Per-trigger target states after successful prepend. |

**Prepend trigger keys** (`metadata.prependTrigger` on `QC_PREPEND` jobs):

| Key | Target state (default) |
|-----|------------------------|
| `initialQcPdf` | `Ready for QC` |
| `reviewerRedlineUpdate` | `Redlines Issued` |
| `designerCorrectionComplete` | `Verification In Progress` (only when explicitly flagged on the job) |

Assignment (`QC_Assigned_To`) is derived from **state + `QC_Review_Type`** via `Resolve-QCWorkflowAssignee` (reviewer vs checker vs designer).

### Deprecated keys (warnings only)

`productionStateName`, `receivedStateName`, `correctionsInProgressStateName`, `backcheckInProgressStateName`, `errorStateName`, and `stageMap` are accepted for backward compatibility but emit deprecation warnings. Prefer `states.*` and `reviewTypes.*`.

Prepend does **not** change ProjectWise state unless `mode` is `StateAndAttributes` and `autoSetState` is `true` (both off by default).

### Tests

- `test/test_qc_workflow.ps1` — unit tests with stub cmdlets.
- `tests/test_qc_workflow_config_defaults.py` — asserts repo defaults stay disabled/attribute-first.

---

## watcher.mode

| Value | Behavior |
|-------|----------|
| `audit_only` | Audit SQL each tick; full folder scan every `auditPoller.reconcileEveryNCycles`. |
| `hybrid` | Audit first; full scan if `reconciliation.downtimeThresholdSeconds` exceeded. |
| `recovery` | Like hybrid for catch-up after outage. |
| `reconciliation` | Full PW folder walk **every** tick (heaviest; repair mode). |

| Key | Description |
|-----|-------------|
| `continuous` | One PW connection, loop until stopped. |
| `idleSleepMs` | Ms between ticks when `continuous: true`. |

---

## reconciliation

| Key | Description |
|-----|-------------|
| `enabled` | Policy in `Get-QCReconciliationPlan` (hybrid downtime). |
| `reconcileStatusSetsOnStart` | Startup `Invoke-StatusSetReconcile` (local PDF → PW only). |
| `downtimeThresholdSeconds` | `0` = off; else audit lag triggers full scan in hybrid/recovery. |

**Note:** Pass 20 full scan (`reconcileEveryNCycles`) is separate from `reconcileStatusSetsOnStart`.

---

## auditPoller

| Key | Default | Description |
|-----|---------|-------------|
| `enabled` | true | Query `dms_audt`. |
| `lookbackSeconds` | 120 | Steady window if watermark missing. |
| `initialLookbackSeconds` | 14400 | First capture only (4h). Delete `queue/_watcher/audit-capture-watermark.txt` to re-run. |
| `reconcileEveryNCycles` | 100 | Full folder scan every N watcher passes; reconciles `sheet_index` EM/QC attributes for paired sheets. |
| `qcPrependAuditActions` | see JSON | On paired sheet PDF audit events (includes `DOCUMENT_ATTR`), re-read description for `QC_Archivist` and enqueue `QC_PREPEND` when matched. |
| `workflowTriggers` | see JSON | `DOCUMENT_STATE` / `DOCUMENT_ATTR` → `sheet_index`, `document_state_history`, `transition_events`, optional email on `*-qc.pdf`. |
| `workflowTriggers.recordFromProcessor` | true | Also record state/attribute changes from `QC_PREPEND` and `QC_COMMENT_STATUS_SYNC` processors. |
| `fallbackToFullScan` | false | Full scan when audit SQL fails. |

Enable `qcCommentSync.enabled` and `enableQcCommentSync` on watch roots to enqueue `QC_COMMENT_STATUS_SYNC` on `*-qc.pdf` for `DOCUMENT_ATTR` / `DOCUMENT_STATE` (see `qcCommentSync.auditActions`).

See `docs/hybrid-polling.md`.

---

## database

| Key | Description |
|-----|-------------|
| `enabled` | Master switch for SQL writes. |
| `connectionString` | SQL Server. |
| `allowWritesInDryRun` | Usually false. |

Tables: `poll_runs`, `sheet_index`, `audit_events`, etc. — `docs/database-telemetry.md`.

---

## notifications

| Key | Description |
|-----|-------------|
| `provider` | `Mock` (log only) or `Graph` (Microsoft 365). |
| `events` | Keyed by PW workflow state name; templates and recipients. |

---

## Related docs

- `docs/qc-workflow-framework.md` — QC workflow writeback (companion to **qcWorkflow** above)
- `docs/hybrid-polling.md` — audit vs n=20 reconcile
- `docs/watcher-architecture-refactor-summary.md` — watcher modes
- `docs/pw-environment-email-attributes.md` — EM_Designer_Email columns
