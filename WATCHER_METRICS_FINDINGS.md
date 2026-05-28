# Watcher Metrics Follow-up Findings

## What was wrong in the previous change

1. **SQL insert mismatch in `Write-QCPollRunTelemetry`**
   - The `INSERT INTO poll_runs (...)` list was expanded with many new columns, but the `VALUES (...)` clause still only supplied the old 11 parameters.
   - This would fail at runtime with a SQL column/value count mismatch and silently drop telemetry because the function catches errors.

2. **Incorrect parameter name for cycle counter helper**
   - `Watch-QCTrigger.ps1` called `Get-AuditPollCycleCounter -Path ...`, but the function expects `-CounterPath`.
   - The bad call is inside `try/catch`, so pass number could be stuck at `0`, reducing telemetry quality.

## What I changed

1. **Fixed poll run insert to include all expanded metrics parameters**
   - Updated `VALUES` in `Write-QCPollRunTelemetry` to pass the full set of new parameters in column order.

2. **Fixed watcher pass counter call-site**
   - Replaced `-Path` with `-CounterPath` when reading watcher pass number.

## Result

- Extended polling-run telemetry now has a valid SQL insert path for the new metrics fields.
- `pass_number` capture uses the correct helper signature and no longer depends on a swallowed parameter-binding failure.
