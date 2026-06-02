# `test/`

Test and validation scripts for the QC pipeline.

This repo contains both `test/` and `tests/`. They are used for different suites (PowerShell and/or Python) depending on what you’re validating.

For a local SQL test database and isolated queue paths, see [`docs/testing-config.md`](../docs/testing-config.md) (`appsettings.test.json` + `-AppSettingsPath`).

Helpful operational tests also live under `scripts/`:

- `scripts/Test-PWConnection.ps1`
- `scripts/Test-QCNotificationGraph.ps1` — Graph email smoke test (`-To` your address; `-Live` to send)
- `scripts/Show-QCQueueDiag.ps1`

