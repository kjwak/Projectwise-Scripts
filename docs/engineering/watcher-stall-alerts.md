# Watcher stall vs session alerts

## Problem

Dashboard stall recovery (`watcher.stallRecovery`) kills and respawns a wedged `Watch-QCTrigger` child when JSONL progress stops. Previously, `sendSessionAlert: true` reused the **ProjectWise session lost** email template, causing false positives during long audit/index work while PW remained connected.

## Current behavior

| Event | Email | Dedupe state file |
| --- | --- | --- |
| PW connect/probe failure (`WATCH_PW_SESSION_STALE`) | `[QC Watcher] URGENT: ProjectWise session lost` | `notifications/session-alerts/last-sent.json` |
| Dashboard stall kill (`DASH_WATCHER_STALL_KILL`) | `[QC Watcher] Watcher child restarted after stall` | `notifications/stall-alerts/last-sent.json` |

Both channels honor `watcher.sessionAlerts.enabled`, recipients, importance, and `dedupeMinutes` independently.

## Configuration (`appsettings.json`)

```json
"watcher": {
  "sessionAlerts": {
    "enabled": true,
    "dedupeMinutes": 60
  },
  "stallRecovery": {
    "enabled": true,
    "noLogActivitySeconds": 3600,
    "sendSessionAlert": false,
    "sendStallAlert": true
  }
}
```

- **`sendSessionAlert`** — deprecated; when true, still sends the old session-lost template on stall kill (avoid in production).
- **`sendStallAlert`** — sends the dedicated stall-recovery template.
- **`noLogActivitySeconds`** — raised to 60 minutes so legitimate long phases are less likely to trip recovery.

## Progress heartbeats

Long silent phases emit throttled `WATCH_PHASE_HEARTBEAT` (default every 3 minutes):

| Phase key | Where |
| --- | --- |
| `audit_folder_guid_cache_warm` | `Sync-AuditPollerWatchFolderGuidCache` |
| `full_reconciliation_scan` | Scheduled full folder scan loop |
| `statusset_sheet_index` | Sheet index reconciliation (`Invoke-QCWatcherLongRunningWork`) |

The dashboard stall detector treats any JSONL bytes (including heartbeats) as activity.

## Modules

| Module | Functions |
| --- | --- |
| `QC.WatcherAlerts.psm1` | `Send-QCWatcherSessionLostAlert`, `Send-QCWatcherStallRecoveryAlert` |
| `QC.WatcherOrchestration.psm1` | `Write-QCWatcherPhaseHeartbeat`, `Invoke-QCWatcherLongRunningWork`, `Test-QCWatcherChildStalled` |

## Verification

```powershell
.\test\powershell\test_qc_watcher_session_alerts.ps1
.\test\powershell\test_qc_watcher_stall_recovery.ps1
```
