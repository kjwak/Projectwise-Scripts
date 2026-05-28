# Watcher Polling-Run Telemetry Findings and Changes

## Findings

1. The polling telemetry write path previously used broad silent `catch { }` behavior in `Write-QCPollRunTelemetry`, which could hide SQL failures.
2. A prior mismatch risk (insert columns vs. value parameters) required hard validation before SQL execution.
3. Pass counter reads can degrade metrics quality when helper calls fail; this must be explicit in telemetry and logs.

## What I changed

### 1) Safer telemetry write behavior
- `Write-QCPollRunTelemetry` now:
  - validates SQL insert shape before execution,
  - logs structured errors (`operation`, `runId`, `passNumber`, `watcherName`),
  - returns a **failed QC result object** on validation/DB/exception failures,
  - returns a **success QC result object** on successful insert,
  - returns a success "skipped" result when DB telemetry is disabled.

### 2) Hard validation for poll_runs inserts
- Added `Test-QCPollRunTelemetryInsertShape` helper that verifies:
  - insert column count equals `@param` count in VALUES,
  - required params are present,
  - no duplicate insert columns,
  - every referenced SQL parameter exists in the params hashtable.

### 3) Watcher fail behavior control
- Added watcher handling for `telemetry.failOnWriteError` (default false).
- On telemetry write failure:
  - logs structured watcher error,
  - watcher continues by default,
  - watcher throws only when `telemetry.failOnWriteError = true`.

### 4) Pass number reliability
- Replaced silent numeric fallback (`0`) with explicit `null` + source tracking.
- Added `pass_number_source` telemetry field to distinguish `counter` vs `error`/fallback paths.
- Added warning log when pass counter read fails.

## Validation coverage added

- Static tests for:
  - validation helper presence and mismatch error codes,
  - write path returning error-result semantics/logging hooks,
  - watcher use of `-CounterPath` and fail-on-write-error behavior.

## Expected operational outcome

- Telemetry failures are now visible and diagnosable instead of silently ignored.
- Insert shape regressions are detected by tests and runtime preflight validation.
- Watcher remains resilient by default while supporting strict mode for telemetry-critical runs.
