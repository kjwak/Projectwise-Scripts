# Hybrid Polling Architecture

## Status: Operational (June 2026)

The pipeline is **audit-driven**: incremental `dms_audt` ingest into SQL, trigger evaluation on `audit_events.processed = 0`, and a durable watermark in `watcher_state` (plus local file mirror). Periodic full folder scans are reconciliation-only, not the steady-state path.

On the QC server (`PXBENTLEY01`), `Watch-QCTrigger.ps1` is spawned by `Start-QCPipelineDashboard.ps1`. Register that dashboard as a boot task with [`scripts/deployment/Register-QCPipelineDashboardTask.ps1`](../../scripts/deployment/Register-QCPipelineDashboardTask.ps1) so the watcher starts after reboot without an interactive logon. The logon operator surface is [`Register-QCPipelineDashboardLogonConsole.ps1`](../../scripts/deployment/Register-QCPipelineDashboardLogonConsole.ps1) → [`Start-QCOpsConsole.ps1`](../../scripts/service/Start-QCOpsConsole.ps1) (WinForms; `-NoGui` is the text tail). The GUI must not start a second dashboard. Pipeline on/off sets `IRegisteredTask.Enabled` on `QC-PipelineDashboard` (Task Scheduler COM, which honors the task DACL). The CIM `Enable-ScheduledTask` / `Disable-ScheduledTask` cmdlets demand admin and are not used. The registrar grants the run-as account a Full Control ACE so the unelevated GUI can toggle it (`-GrantOperatorAccess` for an existing task; owner SID or Read is not enough).

---

## How It Works

### Primary: Audit Trail Scanning

On each watcher tick, `Invoke-AuditTrailScan` (from `PW.AuditPoller.psm1`) queries the `dms_audt` table via `Select-PWSQL`:

1. Read the capture watermark: **`watcher_state`** in SQL when `database.enabled` is true (primary); then `queue/_watcher/audit-capture-watermark.txt` and/or `poll_runs.watermark_after` (mirror/fallback — whichever is latest when the file exists). On first run, look back `initialLookbackSeconds`.
2. Query `dms_audt` with `o_acttime` **newer than the watermark** (`overlapSeconds` optional, default 0). Paginated ASC; rows are ingested to `audit_events`; triggers evaluate **only** `audit_events` rows with `processed = 0`.
3. After a successful scan, persist the **latest ingested `o_acttime`** (or poll end when no rows) to `watcher_state` and `audit-capture-watermark.txt`. Mark evaluated `audit_events` rows `processed = 1`. Document folder paths are cached in `document_activity` and an in-process GUID cache to avoid repeated `Get-PWDocumentsByGUIDs` calls for missing documents.
4. **Timezone:** SQL bounds and the capture file use the **watcher machine clock**. `pw_acttime` in the DB is whatever PW/SQL returns—run the watcher in the **same timezone as the PW datasource/SQL Server** so string comparisons on `o_acttime` match. `maxPwActTime` in logs is informational only and does not advance the capture file.
5. Resolve document GUIDs to folder paths via batched `Get-PWDocumentsByGUIDs` (chunks of 200).
6. Match resolved folders against configured watch roots.
7. Return structured candidates with action name, document info, and folder path.

Relevant audit actions monitored:

| Code | Name | Trigger |
|------|------|---------|
| 1001 | `DOCUMENT_CREATE` | New document — STATUS_SET_GEN rebuild |
| 1002 | `DOCUMENT_MODIFY` | Property change — QC_PREPEND via trigger rules |
| 1003 | `DOCUMENT_ATTR` | Environment attribute change — `sheet_index` refresh + attribute history rows |
| 1006 | `DOCUMENT_FILE_REP` | File replaced (common after rendition) — STATUS_SET_GEN + QC_PREPEND |
| 1007 | `DOCUMENT_CIN` | Check-in — STATUS_SET_GEN rebuild |
| 1012 | `DOCUMENT_STATE` | Workflow state change — live state read, optional fast `QC_PREPEND` enqueue, then sibling/`sheet_index` sync + `document_state_history` / notifications on lane QC PDFs (`*-prod/-rev/-chk.pdf`) |
| 1015 | `DOCUMENT_VERSION` | New version — STATUS_SET_GEN rebuild |
| 1020 | `DOCUMENT_DELETE` | Document deleted — STATUS_SET_GEN rebuild |

### Secondary: Full Folder Reconciliation

At each configured wall-clock time (`auditPoller.fullScanSchedule.times`, e.g. `06:00` and `18:00` in `runtime.displayTimeZoneId`), the watcher runs a full directory scan using the original folder-walking logic. This:

- Catches any events the audit trail missed (retention gaps, edge cases).
- Populates the `sheet_index` database table with document metadata.
- Runs STATUS_SET_GEN manifest comparison for all watched folders.

Full folder reconciliation runs once per schedule slot per calendar day (after the configured time), not on the first audit tick. At the start of each slot the watcher also thins status-set `_history` PDFs and manifest `.bak_*` copies (`statusSet.historyRetention`) and deletes `queue/succeeded` job JSON plus `queue/_logs` files older than `queue.retention` (default 24 hours; once per slot, including preempted scans). `failed/` is not pruned.

**Cooperative preemption (July 2026, updated Aug 2026):** With `fullScanSchedule.preempt.enabled` (default on), each continuous tick processes at most `checkEveryNFolders` folders, then yields so the next tick can run audit again before more folders. The folder list is discovered **once per slot** and stored in `queue/_watcher/full-scan-progress.json` (and `watcher_state` when DB is enabled). Later ticks resume that remaining queue and do **not** re-run `Find-PWSheetsFoldersUnderRoot` or oneLevel expand. `preempt.skipSheetIndex` (default on) skips per-folder `sheet_index` during preempted recon so a large folder cannot stall the watermark for minutes; audit `DOCUMENT_ATTR` / `DOCUMENT_STATE` still update `sheet_index`. The schedule slot is marked complete only when the remaining folder queue is empty.

**Crash fix:** `Invoke-QCWatcherLongRunningWork` no longer uses `System.Threading.Timer` heartbeats (those crashed the watcher ~180s into large `statusset_sheet_index` work). Progress is same-runspace only.

**Operator full scan (ops console):** while the dashboard watcher is live, do **not** spawn a second `Watch-QCTrigger`. The ops GUI writes `queue/_watcher/ops-request.json` (`action=fullScan`, `requestedAtUtc`, `requestedBy`). `Get-QCFullFolderScanReconciliationPlan` treats that as due (`mode=ops_request`); after the scan finishes the watcher deletes the file. “Run audit tick now” is a no-op while the pipeline is on (the watcher is already cycling). If the pipeline is off, a one-shot `Watch-QCTrigger.ps1` is allowed.

Per-folder reconcile noise goes to `Watch-QCTrigger-Reconcile_{yyyy-MM-dd_HH}.jsonl`; lifecycle events (`WATCH_RECONCILE_CYCLE`, `WATCH_FULL_SCAN_PREEMPT` / `RESUME`, slot done/failed) stay in the main watcher JSONL.

### Fallback Behavior

If `Invoke-AuditTrailScan` fails (database unreachable, PW SQL error, etc.) and `auditPoller.fallbackToFullScan` is `true`, the watcher automatically falls back to a full folder scan for that tick.

---

## Module: `PW.AuditPoller.psm1`

### Public Functions

| Function | Purpose |
|----------|---------|
| `Invoke-AuditTrailScan` | Core audit poll: query `dms_audt`, resolve GUIDs, match watch roots, return candidates |
| `Get-AuditTrailPollWindow` | Compute `(since, until]` from last capture + lookback on first run |
| `Get-AuditTrailCaptureWatermark` | Read latest capture time from file and/or `poll_runs` |
| `Set-AuditTrailCaptureWatermark` | Write capture time to `audit-capture-watermark.txt` |
| `Get-AuditTrailHighWaterMark` | Alias for `Get-AuditTrailCaptureWatermark` |
| `Get-AuditPollCycleCounter` | Read current cycle counter (initializes to 0 for first-run reconciliation) |
| `Reset-AuditPollCycleCounter` | Reset cycle counter after reconciliation |

### Internal Helpers

- `_AuditPoller-BuildMatchRoots` — expands watch roots with/without `Documents\` prefix for path matching.
- `_AuditPoller-MatchesWatchRoot` — prefix-matches a resolved folder path against watch roots.
- `_AuditPoller-GetSheetsSubpath` — checks if a folder is under a `sheetsPathFromProject` subdirectory.

---

## Integration in `Watch-QCTrigger.ps1`

The watcher orchestrates audit scanning and scheduled reconciliation (`QC.WatcherOrchestration.psm1`):

```
Each tick (typical production config):
  1. Invoke-QCWatcherAuditTick → Invoke-AuditTrailScan (primary steady-state path; page heartbeats during long ingest)
  2. Process candidates (STATUS_SET_GEN, QC_PREPEND, QC_COMMENT_STATUS_SYNC, etc.)
  3. At configured wall-clock times (auditPoller.fullScanSchedule.times), while slot unpaid:
       → Warm watch-folder GUID cache once per slot (Get-PWFoldersHashTableByGuid + Sheets folders)
       → If no remaining folder queue for this slot: discover Sheets folders + oneLevel expand once
       → Process up to checkEveryNFolders remaining folders (STATUS_SET_GEN; sheet_index skipped when preempt.skipSheetIndex)
       → Checkpoint remaining queue; yield so next tick repeats step 1 (no tree rediscovery)
       → Mark slot complete only when the folder queue is empty
  4. Invoke-QCQueueStartupCheck — queue stats + stale/orphan running\ recovery
  5. Write poll_runs telemetry to database
```

**Legacy fallback:** when `fullScanSchedule.times` is empty, `reconcileEveryNCycles` can still trigger full scans every N ticks (`QC.WatcherOrchestration.psm1`). Committed `appsettings.json` uses scheduled times (`06:00`, `18:00`), not cycle-based reconciliation.

### QC_PREPEND Trigger

On configured audit actions (`auditPoller.qcPrependAuditActions`, default includes `DOCUMENT_MODIFY`, `DOCUMENT_ATTR`, `DOCUMENT_CIN`, `DOCUMENT_FILE_REP`, `DOCUMENT_VERSION`, `DOCUMENT_CREATE`, `DOCUMENT_STATE`), paired sheet PDFs in Sheets folders may enqueue `QC_PREPEND` when workflow state is **Initiate Origination** (non-automation actor), or when the description contains `QC_Archivist` and state is **Initiate Origination**. `DOCUMENT_ATTR` may enqueue after `Sync-PWSheetIndexOwnership` detects a state change. A matching DGN (same stem) is required in Sheets folders. STATUS_SET_GEN skips do not block this check.

When `qcPrepend.fastAuditEnqueue` is **true** (QC server overlay; committed default is **false**), a `DOCUMENT_STATE` candidate is identified, the watcher reads live workflow state once, and if the state is **Initiate Origination** or **Initiate Verification** it enqueues immediately. Sibling maps / attr sync (`Sync-PWAssociatedSheetWorkflowState`) run as a **second pass** after every candidate in the tick has had a chance to enqueue, so a burst of sheets can sit in `pending/` together. Path D (DGN + description) is skipped for that candidate when this tick already attempted state-driven enqueue. Process type may be empty at enqueue. The worker confirms live state, DGN pair, and lane **before** overlay when those reads succeed, and no-op-succeeds (`QC_PREPEND_SKIPPED_*`) if the live state is a real non-trigger value. If live state is empty (no PW session in the parent worker), preflight proceeds to `Invoke-QCPrependPw.ps1` instead of skip-succeeding. Automation-actor events, lane PDFs, and DGNs still do not enqueue from this path. Tag-driven Path D prepend (non-`DOCUMENT_STATE`) is unchanged. When the flag is **false**, `DOCUMENT_STATE` still enqueues via `Sync-PWAssociatedSheetWorkflowState` after sibling/attr work (previous order).

### STATUS_SET_GEN Trigger

Check-in, version, and content-change events on documents in watched Sheets folders trigger folder-level STATUS_SET_GEN candidates. The existing manifest comparison (`Test-StatusSetWatcherShouldEnqueue`) still runs as the correctness gate.

**PW document listing resilience (May 2026):** `Get-StatusSetPWFolderState` no longer relies solely on wildcard `DocumentName=*.pdf` search (often empty on TYPSA datasources). It falls back to full-folder listing and `Get-PWDocumentsInFolder` (same as QC doc scan). If `oneLevelDeep` finds nothing, it retries with `oneLevelDeep: false` on the same path. Logs: `docListingMethod` on `WATCH_PW_STATUSSET_SCAN_DONE`; `WATCH_PW_STATUSSET_NO_DOCS` when still empty; `WATCH_PW_STATUSSET_NO_PAIRS` when PDFs exist but no PDF+DGN pairs.

---

## Configuration (`appsettings.json`)

Committed production excerpt (see `appsettings.json` for full `auditPoller` block):

```json
{
  "auditPoller": {
    "lookbackSeconds": 120,
    "fullScanSchedule": {
      "times": ["06:00", "18:00"],
      "preempt": { "enabled": true, "checkEveryNFolders": 1, "skipSheetIndex": true }
    },
    "fallbackToFullScan": false
  }
}
```

| Key | Committed | Description |
|-----|-----------|-------------|
| `lookbackSeconds` | 120 | How far back to look on first poll when no watermark exists |
| `fullScanSchedule.times` | `06:00`, `18:00` | Wall-clock full reconciliation (preferred over `reconcileEveryNCycles`) |
| `fullScanSchedule.preempt.enabled` | true | Yield between folder chunks so audit/QC events run each tick |
| `fullScanSchedule.preempt.checkEveryNFolders` | 1 | Max Sheets folders processed per continuous tick during an unpaid slot |
| `fullScanSchedule.preempt.skipSheetIndex` | **true** | Skip per-folder `sheet_index` on preempted recon ticks so a large folder cannot stall the watermark. Audit still updates `sheet_index`. |
| `folderGuidCache.warmOnScheduledReconciliation` | **true** | Warm watch-folder GUIDs **once per 06:00 / 18:00 slot**, after that tick's audit scan. Not on watcher start. |
| `folderGuidCache.warmWatchRootsOnStart` | **false** | Deprecated no-op. Audit ticks load `pw_folder_cache` from SQL; they do not walk the PW folder tree. |
| `folderGuidCache.warmSheetsFoldersOnStart` | true | When recon warm runs, also register `CADD\Sheets` folders under each watch root. |
| `folderGuidCache.folderCacheTtlSeconds` | 86400 | SQL TTL for `pw_folder_cache` (covers the gap between 6am and 6pm slots). |
| `reconcileEveryNCycles` | (unset) | **Legacy:** full scan every N ticks when `fullScanSchedule.times` is empty |
| `fallbackToFullScan` | **false** | When **true**, fall back to full folder scan if audit trail query fails for that tick |

---

## Performance

Audit-trail scanning typically processes 2000+ events in 70–90 seconds, compared to the previous full-scan-every-tick approach. API call reduction is ~80–95% during steady state because only changed documents are resolved, not entire folder trees.

GUID resolution is batched (200 GUIDs per `Get-PWDocumentsByGUIDs` call) to minimize PW API round-trips.

**Folder resolution (Jun 2026):** `pw_parentguid` on document audit rows is the containing folder GUID. Resolution: `Get-PWFoldersByGUIDs` (batch with one canonical GUID per parent, then per-GUID retry with brace variants if needed), optional SQL `dms_proj` → `Get-PWFolders -FolderID`, then `Get-PWDocumentsByGUIDs` for document parents. Call `GetFullPath()` before reading path properties. Paths are normalized to `Documents\...` for watch-root matching. `Get-PWFoldersHashTableByGuid` is only for subtree enumeration by path/ID, not arbitrary GUID lists.

**Folder GUID cache warm (Aug 2026):** `Sync-AuditPollerWatchFolderGuidCache` (`Get-PWFoldersHashTableByGuid` + watch-root/`CADD\Sheets` registration) runs **only** from `Invoke-QCAuditFolderGuidCacheWarmForReconciliation` at the start of each unpaid `fullScanSchedule` slot (`06:00` / `18:00`). It is once-per-slot in-process so preempted recon ticks do not re-walk the tree. Watcher startup and audit ticks load `pw_folder_cache` from SQL (`folderCacheTtlSeconds`, default 24h) and never block job discovery on a full-datasource folder index. If SQL cache is empty (first deploy or TTL expiry before the next slot), the parent-GUID filter bypasses (`empty_cache`) until reconciliation warms it; `DOCUMENT_STATE` (1012) remains exempt.

If lookup fails, a **negative cache** row is written (`resolve_failed = 1`, empty `folder_path`, TTL `auditPoller.negativeCacheTtlSeconds`, default 30 minutes). Delete that row (or wait for `expires_at`) before retrying after a fix. Watcher logs: `AUDIT_FOLDER_GUID_NOT_FOUND` (no PW object) vs `AUDIT_FOLDER_GUID_NO_PATH` (object returned but path not extracted).

---

## Sheet Index Population

During full-folder reconciliation, paired sheets can be batch-upserted via `Write-QCSheetIndexBatch`. With `fullScanSchedule.preempt.skipSheetIndex` (default **true**), preempted recon ticks skip that batch so STATUS_SET_GEN cannot stall the next audit tick. Between reconciliation cycles, `Sync-PWSheetIndexOwnership` updates attributes from audit `DOCUMENT_ATTR` events. `Sync-PWAssociatedSheetWorkflowState` runs on audit `DOCUMENT_STATE` events; lane-independent behavior is canonical (`QCProcess.EnableLegacySiblingStateSync: false`). Legacy sibling sync, when enabled, aligns DGN, sheet PDF, and lane QC PDFs (`*-prod/-rev/-chk.pdf`; legacy `*-qc.pdf` bridge).

### Workflow and attribute triggers (`QC.AuditTriggers.psm1`)

When `auditPoller.workflowTriggers.enabled` is true (default):

| Audit action | Runtime behavior |
|--------------|------------------|
| `DOCUMENT_STATE` | Live state read; with `qcPrepend.fastAuditEnqueue`, `QC_PREPEND` enqueue first, then deferred `Sync-PWAssociatedSheetWorkflowState`. Flag off: sibling/attr sync first (aligns siblings when legacy sibling sync is enabled). History/transitions recorded; lane QC PDF notifications when enabled. Events from `workflowTriggers.automationPwUsernames` may skip notify/sync echoes per config. |
| `DOCUMENT_ATTR` | `Sync-PWSheetIndexOwnership` re-reads EM_* / QC_* columns; `QC_Process_Type` and `QC_Review_Type` changes propagate to associated documents; workflow state change to **Initiate Origination** may enqueue `QC_PREPEND` |

Configure under `auditPoller.workflowTriggers` in `appsettings.json`. Notifications still require `notifications.enabled` (separate master switch).

Each row includes:

- Document GUID, name, folder path, extension
- Role emails (`EM_*` preferred over `QC_*`), checker, review type, assigned-to, and `QC_Status`
- Workflow state (from `WorkflowState` or `StateName` property)
- QC PDF pairing (linked via `Update-QCSheetQcPdf` to both PDF and DGN entries sharing the same stem)

See `docs/data/database-telemetry.md` for the `sheet_index` schema.

---

## Related Documentation

- `docs/architecture/audit-trail-architecture.md` — original analysis and design rationale
- `docs/data/database-telemetry.md` — SQL Server schema including `poll_runs`, `watcher_state`, and `sheet_index`
- `docs/workflow/pw-environment-email-attributes.md` — how email attributes are read from ProjectWise
