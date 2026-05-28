# Watcher Architecture Refactor Summary

## What changed

- Introduced a dedicated orchestration module (`modules/QC.WatcherOrchestration.psm1`) to separate watcher operating-mode decisions and reconciliation scheduling decisions from the main watcher ingestion script.
- Watcher modes are now explicit: `audit_only`, `reconciliation`, `recovery`, `hybrid`.
- Reconciliation planning moved behind `Get-QCReconciliationPlan` with reason + trigger source + downtime/audit-gap metadata.
- Extended poll-run telemetry model for reconciliation lifecycle and service-health context.

## Responsibility boundaries

### Watcher responsibilities
- ProjectWise connection/session lifecycle
- Audit ingestion
- Candidate resolution + trigger/dedupe/enqueue
- Poll-run telemetry publish

### Reconciliation responsibilities
- Decided by orchestration layer (`Get-QCReconciliationPlan`)
- Executed by explicit mode (`reconciliation`) or compatibility mode (`hybrid`)
- Supports disablement through config (`reconciliation.enabled = false`)

### Telemetry responsibilities
- Poll-run records continue to be the source of truth
- Added lifecycle fields:
  - `watcher_mode` (implemented via `run_mode`)
  - `reconciliation_reason`
  - `reconciliation_trigger_source`
  - `downtime_seconds`
  - `audit_gap_detected`
  - `watcher_phase`
  - `throttle_wait_seconds`
  - `queue_depth_snapshot`

## Future orchestration direction
- Move reconciliation execution into a dedicated scheduled/maintenance script (nightly/manual/recovery policy driven).
- Keep watcher in `audit_only` for production steady-state.
- Use `hybrid` only as compatibility fallback while the dedicated scheduler is rolled out.
