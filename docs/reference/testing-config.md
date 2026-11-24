# Testing configuration

## Offline Windows baseline

Use a clean Git checkout and a developer Python environment. Install the PDF and test dependencies once, then launch the baseline twice with separate output directories:

```powershell
python -m pip install -r .\tools\overlay\requirements-dev.txt

$run1 = Join-Path $env:TEMP ('pwqc-baseline-' + [guid]::NewGuid().ToString('N'))
$run2 = Join-Path $env:TEMP ('pwqc-baseline-' + [guid]::NewGuid().ToString('N'))
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\test\powershell\run_offline_baseline.ps1 -OutputDirectory $run1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\test\powershell\run_offline_baseline.ps1 -OutputDirectory $run2
```

Each output path must be new. The runner records the Git revision, tool versions, executed commands and suite outcomes in `summary.txt`. The individual logs contain details and pytest case counts. Compare both outcomes and preserve the logs for review. A failing suite makes the runner exit nonzero.

- Python uses synthetic PDF overlay and annotation fixtures. Tests marked `project_fixture` require existing project PDFs or stamp assets and are deselected.
- Windows PowerShell runs the queue, config, module inventory/import, watcher bootstrap/restore and Status Set publication tests. The publication test substitutes mock export and archive operations. The SQL integration test reports an expected skip using the runner's standalone `offline.json` with `database.enabled: false`.
- The runner reads checked-in defaults for policy assertions and uses scratch paths for test outputs. It does not run services or require ProjectWise, SQL Server, Graph, a scheduled task or a publication target. Use this runner for offline verification. The existing focus runner can invoke the SQL integration test with the checkout's configured settings.

Use the following command to run just the service export regression test.

```powershell
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File .\test\powershell\test_watcher_module_bootstrap.ps1
```

The test runs the watcher and worker bootstrap in separate Windows PowerShell 5.1 processes, with and without Status Set publication preloaded. It checks the watcher's reload and restore path and repeats the worker's bootstrap. A deliberately absent standalone settings file stops each entrypoint before service work. Bentley and scheduled-task stubs fail the test if called.

## Optional integration profiles

There is no checked-in `appsettings.test.json.example`. For separately authorized integration work, create a gitignored `appsettings.test.json` overlay beside the base file and pass it with `-AppSettingsPath`. Set test-only roots and endpoints explicitly. Any omitted setting inherits the base defaults, and the secrets file can override your profile. Global dry-run alone does not isolate external dependencies. Review the [effective dry-run policy](dry-run-policy.md) before using any service or discovery entrypoint.

## How merge works

`Read-QCAppSettings` (in `modules/Core/Core.Runtime.psm1`) loads files in order; later files override earlier keys (deep merge for nested objects; arrays and scalars are replaced entirely).

| You pass | Load order |
|----------|------------|
| `appsettings.json` | `appsettings.json` → `appsettings.local.json` (if present) → `appsettings.secrets.json` (if present) |
| `appsettings.test.json` | `appsettings.json` → `appsettings.test.json` → `appsettings.test.local.json` (if present) → `appsettings.secrets.json` (if present) |
| Any other filename | That file only |

The result’s `Data.mergeChain` lists the files that were merged (useful in logs).

## Optional test SQL Server

The offline baseline skips the SQL integration test. To run it separately against an authorized test instance, provide a test-only profile with `database.enabled: true` and a connection string pointing to an isolated test database such as `QC_Pipeline_Test`.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\test\powershell\test_audit_events_db.ps1 -AppSettingsPath .\appsettings.test.json
```

The SQL integration test initializes the schema and writes disposable rows. The maintained schema initializer is `Initialize-QCDatabaseSchema` in `modules/Database/Core.Database.psm1`. [Database telemetry](../data/database-telemetry.md) documents the schema, and `scripts/sql/` contains checked-in maintenance scripts. This repository has no SSDT project.

## Machine-specific overrides

| File | Purpose |
|------|---------|
| `appsettings.secrets.json` | Microsoft Graph credentials (`notifications.graph`); copy from `appsettings.secrets.json.example` |
| `appsettings.local.json` | Other machine-only overrides (paths, credentials) |
| `appsettings.test.local.json` | Extra overrides on top of `appsettings.test.json` (e.g. credential path) |

Example `appsettings.test.local.json`:

```json
{
  "projectWise": {
    "credentialPath": "C:\\PW_QC_LOCAL\\pw_cred.txt"
  }
}
```

## Related

- [Appsettings reference](appsettings-reference.md)
- [Database telemetry](../data/database-telemetry.md)
- [Profile merge test](../../test/powershell/test_config_profile_merge.ps1)
