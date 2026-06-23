# Dry-run and side-effect policy

Production uses **several independent flags** that together determine what a watcher or worker run may do. There is no single `dryRun` switch that means “no side effects.”

At startup, `Watch-QCTrigger.ps1` and `Run-QCProcessor.ps1` emit a structured log with code **`EFFECTIVE_DRY_RUN_POLICY`** (see `Get-QCEffectiveDryRunPolicy` in `modules/Core/Core.Runtime.psm1`).

---

## Configuration layers

| Layer | Config path | Typical production | Effect |
|-------|-------------|-------------------|--------|
| Global | `dryRun` | `false` | Watcher: suppresses `Add-QCQueueJob`. Worker: enters dry-run dispatch path unless overridden below. |
| CLI | `-DryRun` | off | Forces `dryRun: true` for that process. |
| Processor | `processors.dryRun.invokeHandler` | `true` | When global dry run: still call processor functions (with `Config.dryRun = true`). |
| Processor | `processors.dryRun.allowStateChange` | `false` | When global dry run: allow queue lock/move (unusual; default skips lock/move). |
| Workflow | `qcWorkflow.dryRunWriteback` | `false` | When `true`: log planned PW attribute/state writes only. |
| Notifications | `notifications.enabled` + `notifications.dryRun` | enabled / `false` | Graph send gated separately from global dry run. |
| Database | `database.enabled` + `database.allowWritesInDryRun` | enabled / `false` | SQL telemetry off during global dry run unless explicitly allowed. |
| Status set | `statusSet.writeBackToPW` | `true` | Upload `_StatusSet.pdf` to PW after native build (worker only). |

**Independent:** Global `dryRun: true` does **not** imply `qcWorkflow.dryRunWriteback: true` or `notifications.dryRun: true`.

---

## Effective policy matrix

Resolved booleans (see log `data.effectivePolicy`):

| Policy key | Watcher | Worker | When true |
|------------|---------|--------|-----------|
| `discoverPw` | yes | yes | Connect/read ProjectWise (audit poll, folder walk). Watcher never writes PW. |
| `readSql` | yes | yes | `database.enabled` |
| `writeSqlTelemetry` | yes | yes | SQL enabled and (`!dryRun` or `database.allowWritesInDryRun`) |
| `enqueueJobs` | yes | — | `!dryRun` |
| `lockAndMoveQueueJobs` | — | yes | `!dryRun` or `processors.dryRun.allowStateChange` |
| `invokeProcessorHandlers` | — | yes | `!dryRun` or `processors.dryRun.invokeHandler` |
| `writePwFilesViaProcessors` | — | yes | `!dryRun` (handlers receive `dryRun: true` when invoked under global dry run) |
| `writePwWorkflowAttributes` | — | yes | `!dryRun` and `!qcWorkflow.dryRunWriteback` |
| `uploadStatusSetToPw` | — | yes | `writePwFilesViaProcessors` and `statusSet.writeBackToPW` |
| `sendNotificationEmail` | — | yes | `notifications.enabled`, `!notifications.dryRun`, and `!dryRun` |

---

## Common profiles

### Production (`appsettings.json`)

- Full enqueue, lock, dispatch, PW writes, workflow writeback, notifications, SQL telemetry.

### Test profile (`appsettings.test.json`)

- `dryRun: true` → no enqueue; worker dry-run path with `invokeHandler: true` (handlers run read-only/dry).
- `database.enabled: false` → no SQL.
- `notifications` Mock/disabled.

### Safe local enqueue test

```powershell
.\scripts\service\Watch-QCTrigger.ps1 -AppSettingsPath .\appsettings.test.json -DryRun
```

Confirms discovery and trigger evaluation without queue writes.

---

## Related

- [`testing-config.md`](testing-config.md) — test profile merge
- [`appsettings-reference.md`](appsettings-reference.md) — all config sections
- `test/test_effective_dry_run_policy.ps1` — policy resolution unit tests
- `test/test_qc_comment_trigger.ps1` — lane PDF trigger rule tests
