# `test/`

Unified test directory for the QC pipeline.

| Subfolder | Runner | Examples |
|-----------|--------|----------|
| [`powershell/`](powershell/) | `powershell -File test/powershell/<script>.ps1` | `test_qc_prepend_processor.ps1`, `run_focus_tests.ps1` |
| [`python/`](python/) | `python -m pytest` (see root `pytest.ini`) | `test_qc_overlay_prepend.py`, `test_qc_notifications_config.py` |

Use the [offline Windows baseline](../docs/reference/testing-config.md#offline-windows-baseline) for isolated developer validation. It runs selected PowerShell suites and pytest with synthetic fixtures and saves per-suite evidence to a fresh scratch directory.

The [testing configuration guide](../docs/reference/testing-config.md) also describes how to set up a separate SQL integration profile.

Operational diagnostics (not full tests) also live under `scripts/diagnostics/`.
