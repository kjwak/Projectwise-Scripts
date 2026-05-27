# Hybrid Polling Architecture

## Status: Operational (May 2026)

The pipeline uses a **hybrid polling** model where ProjectWise audit-trail events are the primary trigger source, with periodic full folder scanning as a reconciliation fallback.

---

## How It Works

### Primary: Audit Trail Scanning

On each watcher tick, `Invoke-AuditTrailScan` (from `PW.AuditPoller.psm1`) queries the `dms_audt` table via `Select-PWSQL`:

1. Read the capture watermark from `queue/_watcher/audit-capture-watermark.txt` and/or `poll_runs.watermark_after` (whichever is latest). On first run, look back `lookbackSeconds`.
2. `SELECT TOP 500` from `dms_audt` where `o_acttime > @lastCapture` and `o_acttime <= @pollStart` and action is in the relevant set.
3. After a successful scan, persist the new watermark so the next tick only queries and processes events in `(lastCapture, now]`.
4. Resolve document GUIDs to folder paths via batched `Get-PWDocumentsByGUIDs` (chunks of 200).
5. Match resolved folders against configured watch roots.
6. Return structured candidates with action name, document info, and folder path.

Relevant audit actions monitored:

| Code | Name | Trigger |
|------|------|---------|
| 1001 | `DOCUMENT_CREATE` | New document — STATUS_SET_GEN rebuild |
| 1002 | `DOCUMENT_MODIFY` | Property change — QC_PREPEND via trigger rules |
| 1003 | `DOCUMENT_ATTR` | Environment attribute change — QC_PREPEND via trigger rules |
| 1006 | `DOCUMENT_FILE_REP` | File replaced (common after rendition) — STATUS_SET_GEN + QC_PREPEND |
| 1007 | `DOCUMENT_CIN` | Check-in — STATUS_SET_GEN rebuild |
| 1012 | `DOCUMENT_STATE` | Workflow state change |
| 1015 | `DOCUMENT_VERSION` | New version — STATUS_SET_GEN rebuild |
| 1020 | `DOCUMENT_DELETE` | Document deleted — STATUS_SET_GEN rebuild |

### Secondary: Full Folder Reconciliation

Every Nth cycle (configured by `auditPoller.reconcileEveryNCycles`, default 20), the watcher runs a full directory scan using the original folder-walking logic. This:

- Catches any events the audit trail missed (retention gaps, edge cases).
- Populates the `sheet_index` database table with document metadata.
- Runs STATUS_SET_GEN manifest comparison for all watched folders.

Full folder reconciliation runs every `reconcileEveryNCycles` audit ticks (default 20), not on the first audit tick.

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

The watcher orchestrates both modes:

```
Each tick:
  1. Increment cycle counter
  2. If (counter >= reconcileEveryNCycles AND counter % reconcileEveryNCycles == 0):
       → Full folder scan (reconciliation)
     First dashboard tick uses audit scan only (status-set reconcile is separate via `-ReconcileStatusSetsFirst`).
       → Populate sheet_index from paired sheets
       → Link QC PDFs to source documents
     Else:
       → Invoke-AuditTrailScan
       → Process candidates (STATUS_SET_GEN, QC_PREPEND)
       → Populate sheet_index for sheets found in audit events
  3. `Invoke-QCQueueStartupCheck` — queue stats + stale/orphan `running\` recovery (also in `run_prepend_qc -NoDashboard`)
  4. Write poll_runs telemetry to database
```

### QC_PREPEND Trigger

Any audit event on a PDF file triggers a description check. If the document's description contains `QC_Archivist`, a QC_PREPEND job is enqueued. This replaces scanning every document in every folder.

### STATUS_SET_GEN Trigger

Check-in, version, and content-change events on documents in watched Sheets folders trigger folder-level STATUS_SET_GEN candidates. The existing manifest comparison (`Test-StatusSetWatcherShouldEnqueue`) still runs as the correctness gate.

---

## Configuration (`appsettings.json`)

```json
{
  "auditPoller": {
    "lookbackSeconds": 120,
    "reconcileEveryNCycles": 20,
    "fallbackToFullScan": true
  }
}
```

| Key | Default | Description |
|-----|---------|-------------|
| `lookbackSeconds` | 120 | How far back to look on first poll when no watermark exists |
| `reconcileEveryNCycles` | 20 | Run full reconciliation scan every N watcher ticks |
| `fallbackToFullScan` | true | Fall back to full scan if audit trail query fails |

---

## Performance

Audit-trail scanning typically processes 2000+ events in 70–90 seconds, compared to the previous full-scan-every-tick approach. API call reduction is ~80–95% during steady state because only changed documents are resolved, not entire folder trees.

GUID resolution is batched (200 GUIDs per `Get-PWDocumentsByGUIDs` call) to minimize PW API round-trips.

---

## Sheet Index Population

During both scan modes, documents in `CADD\Sheets` subfolders are written to the `sheet_index` database table via `Write-QCSheetIndex`. This includes:

- Document GUID, name, folder path, extension
- Designer and reviewer email attributes (from `Get-PWDocumentsBySearchWithReturnColumns` `.Attributes` bags)
- Workflow state (from `WorkflowState` or `StateName` property)
- QC PDF pairing (linked via `Update-QCSheetQcPdf` to both PDF and DGN entries sharing the same stem)

See `docs/database-telemetry.md` for the `sheet_index` schema.

---

## Related Documentation

- `docs/audit-trail-architecture.md` — original analysis and design rationale
- `docs/database-telemetry.md` — SQL Server schema including `poll_runs` and `sheet_index`
- `docs/pw-environment-email-attributes.md` — how email attributes are read from ProjectWise
