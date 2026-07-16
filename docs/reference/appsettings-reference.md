# appsettings reference

The pipeline reads **`appsettings.json`** at the repo root (strict JSON). Use this document for explanations; use **`appsettings.schema.json`** for hover help in Cursor/VS Code.

## How settings are loaded

| Mechanism | Purpose |
|-----------|---------|
| `Read-QCAppSettings` | All scripts and dashboard load config via `modules/Core/Core.Runtime.psm1`. |
| `-AppSettingsPath` | Override path on `Watch-QCTrigger.ps1`, `Start-QCPipelineDashboard.ps1`, etc. |
| `-DryRun` (CLI) | Forces `dryRun: true` even when JSON says false. |
| `$schema` in JSON | IDE only; ignored at runtime. |
| Profile merge | `appsettings.test.json` loads `appsettings.json` then the profile then `appsettings.test.local.json` then `appsettings.secrets.json`. `appsettings.json` also merges `appsettings.local.json` and `appsettings.secrets.json`. |

**Testing / local overlays:** see [`docs/reference/testing-config.md`](../reference/testing-config.md). Copy `appsettings.test.json.example` → `appsettings.test.json` (gitignored).

**Dry-run / side effects:** see [`docs/reference/dry-run-policy.md`](../reference/dry-run-policy.md). Watcher and worker log `EFFECTIVE_DRY_RUN_POLICY` at startup.

**Graph credentials:** copy `appsettings.secrets.json.example` → `appsettings.secrets.json` (gitignored). Do not commit Entra client secrets.

Optional: other machine-only paths in `appsettings.local.json` or `appsettings.test.local.json`.

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
| `qcRendition` | QC_RENDITION profiles + Ready for QC notification gate |
| `statusSet` | STATUS_SET_GEN merge + manifest + PW upload |
| `statusSetBatching` | Defer audit STATUS_SET_GEN to dirty-folder batch runs |
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
| `enableQcCommentSync` | Audit path may enqueue `QC_COMMENT_STATUS_SYNC` on lane QC PDFs (`*-prod.pdf`, `*-rev.pdf`, `*-chk.pdf`; legacy `*-qc.pdf` still matched by some rules). |
| `enableStatusSet` | Folder-level `STATUS_SET_GEN` (paired PDF + DGN). |
| `qcRendition` | Per-root rendition profile for `QC_RENDITION` (see below). |

Optional: `environmentEmailAttributes` — see `docs/workflow/pw-environment-email-attributes.md`.

### watchList.roots[].qcRendition

Define the ProjectWise rendition profile for every Sheets folder discovered under that root. Resolution matches the job’s `sourceFolder` to the longest `path` prefix, then applies `qcRendition.folderOverrides` if any.

| Key | Description |
|-----|-------------|
| `profileName` | `New-PWRenditionRequest -ProfileName` value. |
| `sourceDocumentPattern` | `Get-PWDocumentsBySearch` pattern (default `%.dgn`). |
| `useFolderProfileWhenUnspecified` | Fallback to `Get-PWFolderRenditionProfile` when `profileName` is omitted. |
| `outputFolderRelative` / `outputFolderPath` | Where to poll for rendition PDFs after submit. Use `"."` or `""` for the **same folder** as the sheet (not a `Renditions` subfolder). Omit only if you set `outputFolderPath`. |

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
| `QC_COMMENT_STATUS_SYNC` | Update PW state from lane QC PDF comments (`*-prod/-rev/-chk.pdf`; legacy `*-qc.pdf` bridge). |

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

## qcRendition

After a successful `QC_PREPEND` that sets workflow state to **Originated** (TYPSA), the pipeline can enqueue **`QC_RENDITION`** (`modules/Processing/QC.Rendition.psm1`). Profile resolution order: longest matching `folderOverrides[].folderPathPrefix`, then longest matching `projectWise.watchList.roots[].qcRendition`, then legacy `qcRendition.datasources[datasourceName]`, then optional `Get-PWFolderRenditionProfile` when `useFolderProfileWhenUnspecified` is true.

| Key | Description |
|-----|-------------|
| `enabled` | Master switch (default `false` in repo `appsettings.json`). |
| `deferReadyForQcNotification` | `false` (default): **Originated** email when prepend sets workflow state. `true`: wait until prepend **and** rendition complete. |
| `readinessStorePath` | JSON files tracking `prependComplete` / `renditionComplete` per document. |
| `deriveSourceFromQcPdf` | Lane QC PDF → source DGN (e.g. `sheet001-prod.pdf` → `sheet001.dgn`; legacy `sheet001-qc.pdf` still supported). |
| `completion.mode` | `outputFolder` polls `outputFolderPath` or `outputFolderRelative`; `immediate` / `submitOnly` mark complete after `New-PWRenditionRequest`. |
| `completion.fileNamePattern` | Glob with `{stem}` from DGN base name (e.g. `{stem}*.pdf`). |
| `folderOverrides` | Per-folder `folderPathPrefix` + profile/output overrides (longest prefix wins; rare). |
| `datasources` | Deprecated; use `watchList.roots[].qcRendition` instead. |

Register `QC_RENDITION` in `processors.processorMap` and `queue.selection.preferJobTypes` (see `appsettings.json`).

---

## qcPrepend

| Key | Description |
|-----|-------------|
| `historyRoot` / `outputRoot` / `tempRoot` | Overlay working directories (used by `local` mode and overlay staging). |
| `mode` | **`projectWise`** (production): `Invoke-QCPrependProcessor` spawns `scripts\processing\Invoke-QCPrependPw.ps1` for PW export, overlay merge, and lane PDF upload. **`local`**: disk-only path in `QC.Processors.psm1` (tests/staging; no PW file I/O). **`legacyPw`** / **`pw`**: deprecated aliases for `projectWise`. |
| `projectWiseScriptPath` | Optional override for the ProjectWise prepend script (default `scripts\processing\Invoke-QCPrependPw.ps1`). |
| `legacyScriptPath` | Deprecated alias for `projectWiseScriptPath`. |
| `enableOverlay` | Use `qc_overlay_prepend.exe` when available. |
| `overlaySkipIncomingAboveMb` | When `> 0`, skip layered overlay if the **incoming** PDF is larger than this many megabytes (`Mb * 1024 * 1024`); fall back to qpdf simple prepend + stamps (no Old/New/Current layers). Default `5`. Set `0` to disable the size gate. |
| `qpdfExePath` | Path to qpdf binary. |
| `overlayExePath` | Path to `qc_overlay_prepend.exe` (default `dist\qc_overlay_prepend\qc_overlay_prepend.exe`). |

---

## statusSetBatching

Defers **audit-sourced** `STATUS_SET_GEN` work: the watcher marks Sheets folders dirty instead of scanning PW and enqueueing on every audit event. A batch pass runs on a timer while the ProjectWise session is open.

| Key | Default | Description |
|-----|---------|-------------|
| `enabled` | `true` | When `true`, audit status-set triggers call `Mark-StatusSetDirtyFolder` instead of immediate enqueue. Set `false` to restore legacy per-event folder evaluation. |
| `intervalMinutes` | `15` | Minimum minutes between batch runs (`STATUSSET_BATCH_INTERVAL_*` logs). Raise to `60` after rollout validation if desired. |
| `maxFoldersPerRun` | `100` | Cap on folders evaluated per batch. |
| `quietPeriodSeconds` | `120` | Folders with events newer than this are skipped until the quiet window passes. |
| `staleWarningHours` | `24` | Warning-only: emit `STATUSSET_DIRTY_FOLDER_STALE` when `firstSeenUtc` exceeds this age. Does not block or remove folders. |
| `dirtyFolderStorePath` | `_watcher/statusset-dirty-folders.json` | Durable dirty-folder JSON under `queue.rootDir` when set (e.g. `C:\QC_E2E_RealRun\queue\_watcher\statusset-dirty-folders.json`). Absolute paths accepted. Atomic writes. |
| `processOnWatcherStart` | `false` | When `true`, tick 1 runs a batch immediately (still respects `quietPeriodSeconds`). |

**Rollout note:** Audit-triggered `STATUS_SET_GEN` work is now delayed until the batching interval. With default settings, status-set cleanup may occur up to `intervalMinutes` plus `quietPeriodSeconds` after the triggering audit event. Set `statusSetBatching.enabled=false` to restore the old immediate audit-triggered behavior.

**Unchanged:** scheduled full-folder reconciliation scans still evaluate folders directly. Queue dedupe (`Test-QCDuplicateJob`) and `Test-StatusSetWatcherShouldEnqueue` run at batch evaluation time.

**Log codes:** `STATUSSET_BATCHING_SETTINGS`, `STATUSSET_DIRTY_FOLDER_MARKED`, `STATUSSET_DIRTY_FOLDER_STALE`, `STATUSSET_BATCH_INTERVAL_START`, `STATUSSET_BATCH_FOLDER_START`, `STATUSSET_BATCH_FOLDER_EVALUATED`, `STATUSSET_BATCH_FOLDER_FAILED`, `STATUSSET_BATCH_INTERVAL_DONE`.

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

Optional ProjectWise **document attribute** writeback (and optional **workflow state** changes) after a successful `QC_PREPEND`. Implemented in `modules/Workflow/QC.Workflow.psm1`; invoked from `QC.Processors.psm1` when prepend completes.

**Full guide (lifecycle states, review types, assignment, PW admin setup, discovery scripts, rollout):** [`docs/workflow/qc-workflow-framework.md`](../workflow/qc-workflow-framework.md)

**Related:** comment-driven state uses the same PW helpers via `qcCommentSync` — see [`docs/workflow/qc-comment-status-sync.md`](../workflow/qc-comment-status-sync.md). SQL audit rows: `qc_workflow_events` in [`docs/data/database-telemetry.md`](../data/database-telemetry.md).

### When `enabled` is false

When `qcWorkflow.enabled` is false, prepend/overlay still runs; **no** PW attributes or states are written. Committed `appsettings.json` has workflow writeback **enabled** for TYPSA production; use `dryRunWriteback: true` on pilot environments.

### Enablement order

1. Confirm PW attributes and TYPSA workflow states exist.
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
| `states` | see JSON | TYPSA lifecycle state names (source of truth). |
| `processTypes` | `Production`, `Review`, `Check` | Canonical `QC_Process_Type` values. |
| `reviewTypes` | `Production`, `Review`, `Check` | Allowed `QC_Review_Type` values (mirrors process types in TYPSA). |
| `defaultProcessType` | `production` | Used when document/job has no process type. |
| `defaultReviewType` | `Production` | Used when document/job has no review type. |

### Lifecycle states (`states.*`) — TYPSA committed defaults

Names must match your ProjectWise workflow:

| Key | TYPSA default |
|-----|---------------|
| `production` | `In Development` |
| `qcInitiated` | `Initiate Origination` |
| `qcReceived` / `readyForQc` | `Originated` |
| `redlinesReceived` | `Redlines Received` |
| `qcFinalizing` | `Initiate Verification` |
| `readyForVerification` | `Ready for Verification` |
| `complete` | `Verified` |
| `error` | `Error Needs Attention` |

**Legacy labels** (pre-TYPSA): `In Production`, `QC Initiated`, `QC Received`, `Ready for QC`, `Review In Progress`, `Corrections In Progress`, `Verification In Progress`, `QC Complete` — retained only in unmigrated configs.

| Key | TYPSA default | Description |
|-----|---------------|-------------|
| `defaultStateAfterPrepend` | `Originated` | Informational default when state writeback is enabled. |
| `stateAfterSuccessfulPrepend` | `Originated` | Fallback when `autoSetState` runs and prepend trigger is unknown. |
| `stateAfterFailedPrepend` | `Error Needs Attention` | Target state after failed prepend. |
| `defaultPrependTrigger` | `initialQcPdf` | Documented default; set `metadata.prependTrigger` on jobs. |
| `stateAfterPrependByTrigger` | see JSON | Per-trigger target states after successful prepend. |

**Prepend trigger keys** (`metadata.prependTrigger` on `QC_PREPEND` jobs):

| Key | Target state (TYPSA default) |
|-----|------------------------------|
| `initialQcPdf` | `Originated` |
| `finalQcComplete` | `Ready for Verification` |

Assignment (`QC_Assigned_To`) is derived from **state + `QC_Process_Type`** via `Resolve-QCWorkflowAssignee` (reviewer vs checker vs designer).

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
| `audit_only` | Audit SQL each tick; full folder scan at `auditPoller.fullScanSchedule.times` (wall clock). |
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

**Note:** Scheduled full scans (`fullScanSchedule.times`) are separate from `reconcileStatusSetsOnStart`.

---

## auditPoller

| Key | Default | Description |
|-----|---------|-------------|
| `enabled` | true | Query `dms_audt`. |
| `lookbackSeconds` | 120 | Steady window if watermark missing. |
| `initialLookbackSeconds` | 14400 | First capture only (4h). Delete `queue/_watcher/audit-capture-watermark.txt` to re-run. |
| `fullScanSchedule.times` | (none) | Wall-clock times (`HH:mm`) in `runtime.displayTimeZoneId` for full folder scans (once per slot per day). |
| `reconcileEveryNCycles` | — | Legacy fallback when `fullScanSchedule.times` is empty. |
| `qcPrependAuditActions` | see JSON | On paired sheet PDF audit events, enqueue `QC_PREPEND` when workflow state is **Initiate Origination** and/or when description has `QC_Archivist` (state must still be **Initiate Origination** for the tag path). |
| `workflowTriggers` | see JSON | `DOCUMENT_STATE` / `DOCUMENT_ATTR` → `sheet_index`, `document_state_history`, `transition_events`, optional email on lane QC PDFs. |
| `workflowTriggers.recordFromProcessor` | true | Also record state/attribute changes from `QC_PREPEND` and `QC_COMMENT_STATUS_SYNC` processors. |
| `workflowTriggers.ignoreStateChangeFromAutomation` | false | When **true**, skip audit notifications and sibling-sync for `DOCUMENT_STATE` from `automationPwUsernames`. Default **false** so automation-driven **Originated** is visible to audit triggers (dedupe prevents duplicate emails). |
| `workflowTriggers.suppressBaselineIndexStateTransition` | true | Skip `transition_events` / notifications when `sheet_index.pw_state_name` was empty and PW shows `qcWorkflow.states.production` (**In Development**). Safe when expanding QC to folders that already have `sheet_index` rows from status-set scans. |
| `workflowTriggers.baselineStateNames` | `[]` | Extra production-style labels (e.g. legacy **In Production**). |
| `workflowTriggers.processingGoLiveUtc` | (empty) | Optional UTC ISO time. `DOCUMENT_STATE` workflow sync is skipped for older `audit_events` (still ingested + marked processed). Set to **now** when enabling QC on additional watch roots. Clear after backlog drains. |
| `workflowTriggers.automationPwUsernames` | `["srv_typsa_archivist"]` | Service account logins from `dms_user` / `pw_users`. |
| `workflowTriggers.automationPwUserNumbers` | `[]` | Optional numeric `o_userno` list if usernames differ by environment. |
| `fallbackToFullScan` | false | Full scan when audit SQL fails. |

After successful prepend, `QC.Workflow` may call `Sync-PWPostInitialPrependLaneStates` for lane-independent initial prepend, or `Sync-PWAssociatedSheetMembersToWorkflowState` when legacy sibling sync is enabled. The **Originated** email is sent from the workflow hook when `notifications.events['Originated'].enabled` is true (not deferred when `deferReadyForQcNotification` is false).

Enable `qcCommentSync.enabled` and `enableQcCommentSync` on watch roots to enqueue `QC_COMMENT_STATUS_SYNC` on lane QC PDFs (`*-prod.pdf`, `*-rev.pdf`, `*-chk.pdf`) for `DOCUMENT_ATTR` / `DOCUMENT_STATE` (see `qcCommentSync.auditActions` and `triggers.rules`). Legacy `*-qc.pdf` is a compatibility bridge in code paths, not the primary trigger pattern in committed config.

See `docs/architecture/hybrid-polling.md`.

---

## database

| Key | Description |
|-----|-------------|
| `enabled` | Master switch for SQL writes. |
| `connectionString` | SQL Server. |
| `allowWritesInDryRun` | Usually false. |

Tables: `poll_runs`, `sheet_index`, `audit_events`, etc. — `docs/data/database-telemetry.md`.

---

## notifications

| Key | Description |
|-----|-------------|
| `enabled` | Master switch for workflow emails. |
| `provider` | `Mock` (log only) or `MicrosoftGraph`. |
| `dryRun` | When true, Graph builds payload but does not call `sendMail`. |
| `events` | Keyed by PW workflow state name; subject templates and recipients. |
| `attributes` | PW column names for reviewer/designer/cc emails; optional `qcPdfUrlField`, `projectNumberField`, `reviewTypeField`. |

### notifications.email

| Key | Default | Description |
|-----|---------|-------------|
| `bodyFormat` | `Html` | `Text` = plain body only; `Html` uses `email/templates/qc_notification.html`. |
| `subjectTemplate` | `[{ReviewType}] {ProjectName} \| {DocumentName} \| {WorkflowState}` | Default email subject for all workflow states. Per-event `subjectTemplate` overrides this. Placeholders are case-sensitive; supported tokens: `{ReviewType}` / `{reviewType}`, `{ProjectName}` / `{project}`, `{DocumentName}` / `{documentName}`, `{WorkflowState}` / `{currentState}`, `{previousState}`, `{eventType}`, `{documentPath}`. |
| `templatePath` | `email/templates/qc_notification.html` | Repo-relative HTML template. |
| `logoPath` | `email/typsalogo.png.webp` | Inline attachment for Graph HTML sends. |
| `environment` | `Production` | Shown in email footer when set. |
| `qcPdfUrlTemplate` | `""` | Optional URL pattern with `{documentGuid}`, etc. |
| `pwLinkBaseUrl` | Bentley CONNECT pwlink base | Used to build link when GUID known. |
| `pwLinkApp` | `pwe` | `pwe`, `web`, or `webview` query parameter. |

Per-event override: `notifications.events.<State>.emailTemplate` (alternate HTML file).

### notifications.graph (secrets)

Store in `appsettings.secrets.json`: `tenantId`, `clientId`, `clientSecret`, `senderMailbox`.

---

## Related docs

- `docs/workflow/qc-workflow-framework.md` — QC workflow writeback (companion to **qcWorkflow** above)
- `docs/architecture/hybrid-polling.md` — audit-driven watcher and scheduled reconciliation
- `docs/architecture/watcher-architecture-refactor-summary.md` — watcher modes
- `docs/workflow/pw-environment-email-attributes.md` — EM_Designer_Email columns
