# Watcher Architecture Refactor Summary

## What changed

- Introduced a dedicated orchestration module (`modules/QC.WatcherOrchestration.psm1`) to separate watcher operating-mode decisions and reconciliation scheduling decisions from the main watcher ingestion script.
- Watcher modes are now explicit: `audit_only`, `reconciliation`, `recovery`, `hybrid`.
- Reconciliation planning moved behind `Get-QCReconciliationPlan` with reason + trigger source + downtime/audit-gap metadata.
- Extended poll-run telemetry model for reconciliation lifecycle and service-health context.

## Mode contract (source-of-truth behavior)

The watcher must behave differently depending on the selected `run_mode` (canonical mode field). The intent is to make each mode's responsibilities and side-effects explicit and auditable.

### `audit_only`
- Ingest audit events and publish poll-run telemetry.
- May resolve audit candidates and enqueue work **only from audit ingestion**.
- Must **not** execute reconciliation scans.

### `reconciliation`
- Execute reconciliation scans/queries and publish reconciliation lifecycle telemetry.
- May enqueue work **only from reconciliation results**.
- Should not run continuous audit ingestion (one-shot / bounded run is preferred).

### `recovery`
- A bounded run intended to recover from downtime or detected gaps.
- May execute reconciliation scans (as needed) and publish explicit "recovery" rationale in telemetry.
- May optionally ingest audit events if required to re-establish continuity, but reconciliation is the primary mechanism.

### `hybrid` (compatibility fallback)
- Runs audit ingestion and may execute reconciliation **only when orchestration explicitly returns a plan**.
- Must be treated as transitional; production steady-state should target `audit_only` + external scheduler for reconciliation.

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
- Startup status-set push to PW can be disabled via `reconciliation.reconcileStatusSetsOnStart = false` (dashboard no longer passes `-ReconcileStatusSetsFirst` on first pass)

#### Disablement behavior (`reconciliation.enabled = false`)
- Orchestration may still *detect* that reconciliation would be appropriate (so intent is visible),
  but execution must be skipped.
- Telemetry must still record: `reconciliation_reason`, `reconciliation_trigger_source`, and a mode/phase indicating "planned-but-disabled".
- The watcher should continue audit ingestion (if in `audit_only`/`hybrid`) rather than failing the poll-run.

### Telemetry responsibilities
- Poll-run records continue to be the source of truth
- Added lifecycle fields:
  - `run_mode` (canonical; replaces/avoids a separate `watcher_mode` field)
  - `reconciliation_reason`
  - `reconciliation_trigger_source`
  - `downtime_seconds`
  - `audit_gap_detected`
  - `watcher_phase`
  - `throttle_wait_seconds`
  - `queue_depth_snapshot`

#### Telemetry enums / sampling rules (to keep dashboards clean)
- `watcher_phase` should be a small, stable enum (suggested):
  - `startup`
  - `connect_projectwise`
  - `ingest_audit`
  - `plan_reconciliation`
  - `execute_reconciliation`
  - `enqueue_work`
  - `publish_telemetry`
  - `sleep_throttle`
  - `shutdown`
- `queue_depth_snapshot` should state which queue and when sampled (suggested: work-queue depth sampled once per poll-run, immediately before enqueue).
- `throttle_wait_seconds` should represent time actually slept/waited during the poll-run (not the configured limit).

## Orchestration decision table (policy surface)

`Get-QCReconciliationPlan` is the single policy surface that decides whether reconciliation should be executed and why. Suggested inputs and outputs:

### Inputs (examples)
- **run_mode**: one of `audit_only`/`hybrid`/`reconciliation`/`recovery`
- **service health**: last successful poll-run, last audit watermark, etc.
- **gap signals**: `audit_gap_detected`, computed downtime (\( now - last_success \))
- **operator trigger**: manual request, CLI flag, or scheduled trigger (future scheduler)
- **config**: `reconciliation.enabled`, thresholds, backoff/throttle limits

### Outputs (examples)
- **should_reconcile**: boolean
- **reconciliation_reason**: stable enum (e.g. `startup`, `downtime_exceeded`, `audit_gap_detected`, `manual`, `scheduled`, `backlog_pressure`)
- **reconciliation_trigger_source**: stable enum (e.g. `auto`, `operator`, `scheduler`)
- **downtime_seconds** and **audit_gap_detected**: carried through for telemetry
- **plan_window**: optional start/end bounds to keep reconciliation bounded

### Decision rules (high-level)
- **If** `run_mode` is `reconciliation` or `recovery` **then** plan should typically return `should_reconcile = true` (subject to config/guardrails).
- **If** `run_mode` is `audit_only` **then** plan must return `should_reconcile = false` (audit ingestion only).
- **If** `run_mode` is `hybrid` **then** plan may return `should_reconcile = true` when a gap/downtime/explicit trigger is detected, but must remain bounded and observable via telemetry.
- **If** `reconciliation.enabled = false` **then** plan may return intent metadata but execution must be skipped (see disablement behavior).

## Future orchestration direction
- Move reconciliation execution into a dedicated scheduled/maintenance script (nightly/manual/recovery policy driven).
- Keep watcher in `audit_only` for production steady-state.
- Use `hybrid` only as compatibility fallback while the dedicated scheduler is rolled out.
