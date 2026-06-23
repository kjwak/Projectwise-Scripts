# `test/`

Unified test directory for the QC pipeline.

| Subfolder | Runner | Examples |
|-----------|--------|----------|
| [`powershell/`](powershell/) | `powershell -File test/powershell/<script>.ps1` | `test_qc_prepend_processor.ps1`, `run_focus_tests.ps1` |
| [`python/`](python/) | `python -m pytest` (see root `pytest.ini`) | `test_qc_overlay_prepend.py`, `test_qc_notifications_config.py` |

For local SQL test database and isolated queue paths, see [`docs/reference/testing-config.md`](../docs/reference/testing-config.md) (`appsettings.test.json` + `-AppSettingsPath`).

**Quick commands (from repo root):**

```powershell
python -m pytest -q
.\test\powershell\run_focus_tests.ps1
.\test\powershell\test_module_inventory.ps1
```

Operational diagnostics (not full tests) also live under `scripts/diagnostics/`.
