# Testing configuration

Use a **test profile** so queue folders, dry-run behavior, and optional SQL telemetry stay separate from day-to-day `appsettings.json`, without duplicating the full file.

## Quick start

1. Copy the example profile:

   ```powershell
   Copy-Item .\appsettings.test.json.example .\appsettings.test.json
   ```

2. Adjust paths in `appsettings.test.json` if `C:\QC_Test\` is not suitable.

3. Run scripts with the test profile:

   ```powershell
   .\scripts\Watch-QCTrigger.ps1 -AppSettingsPath .\appsettings.test.json -DryRun
   .\scripts\Run-QCProcessor.ps1 -AppSettingsPath .\appsettings.test.json
   .\scripts\Start-QCPipelineDashboard.ps1 -AppSettingsPath .\appsettings.test.json
   ```

The example sets **`database.enabled: false`** so you can test on a laptop with no SQL Server installed. Telemetry is skipped; the JSON queue and processors still run. See `docs/database-telemetry.md`.

`appsettings.test.json` and `appsettings.local.json` are **gitignored**. Only `appsettings.test.json.example` is committed.

## How merge works

`Read-QCAppSettings` (in `modules/Core.Runtime.psm1`) loads files in order; later files override earlier keys (deep merge for nested objects; arrays and scalars are replaced entirely).

| You pass | Load order |
|----------|------------|
| `appsettings.json` | `appsettings.json` → `appsettings.local.json` (if present) |
| `appsettings.test.json` | `appsettings.json` → `appsettings.test.json` → `appsettings.test.local.json` (if present) |
| Any other filename | That file only |

The result’s `Data.mergeChain` lists the files that were merged (useful in logs).

## What the example profile changes

`appsettings.test.json.example` is a **small overlay** merged on top of repo `appsettings.json`:

| Area | Test intent |
|------|-------------|
| `dryRun` | `true` — safer PW / queue behavior |
| `queue` / `qcPrepend` / `statusSet` | Under `C:\QC_Test\` — no collision with `C:\QC_E2E_RealRun\` |
| `statusSet.writeBackToPW` | `false` — local status-set builds without PW upload |
| `database.enabled` | `false` — no SQL required for local testing |
| `watcher.mode` | `audit_only` — lighter watcher loop for local runs |
| `notifications` | Mock / disabled |

Inherited from base `appsettings.json` unless you override them: `projectWise.watchList`, `triggers`, `filters`, `qcWorkflow`, etc. To scope testing to one PW root, add a `projectWise.watchList` block to your local `appsettings.test.json`.

## Optional: enable SQL on your laptop later

When you install SQL Server Express (or already have an instance), set in `appsettings.test.json`:

```json
"database": {
  "enabled": true,
  "connectionString": "Server=localhost\\SQLEXPRESS;Database=QC_Pipeline_Test;Trusted_Connection=True;TrustServerCertificate=True;",
  "allowWritesInDryRun": true,
  "logPlannedEventsInDryRun": true
}
```

Create an empty database `QC_Pipeline_Test` (or let `Initialize-QCDatabaseSchema` run on first pipeline start if your login can create it). Schema is applied automatically by `modules/Core.Database.psm1`.

With global `dryRun: true`, SQL writes stay off unless `allowWritesInDryRun` is `true` — useful when you want PW dry run but still fill `audit_events` / `sheet_index` in a test database.

DB integration tests: `.\test\test_audit_events_db.ps1 -AppSettingsPath .\appsettings.test.json` (skips when `database.enabled` is false).

The `db\DatabaseProjectQC_Pipeline\` folder is an SSDT schema project for publish/review; the runtime does not read it directly. See `docs/database-telemetry.md`.

## Machine-specific overrides

| File | Purpose |
|------|---------|
| `appsettings.local.json` | Secrets and paths when using default `appsettings.json` |
| `appsettings.test.local.json` | Extra overrides on top of `appsettings.test.json` (e.g. credential path) |

Example `appsettings.test.local.json`:

```json
{
  "projectWise": {
    "credentialPath": "C:\\PW_QC_LOCAL\\pw_cred.txt"
  }
}
```

## Python tests

Static tests under `tests/` still read committed `appsettings.json` for default-policy checks. They are not switched by `-AppSettingsPath`. Integration work should use PowerShell tests or discovery scripts with `-AppSettingsPath`.

## Related

- `docs/appsettings-reference.md` — all sections
- `docs/database-telemetry.md` — schema and tables
- `test/test_config_profile_merge.ps1` — merge chain unit test
