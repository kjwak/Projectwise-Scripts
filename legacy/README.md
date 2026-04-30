This folder holds **legacy scripts** from the pre-modular QC pipeline.

- Keep these here temporarily while validating parity with the new framework.
- The new entrypoints are in repo root (`Watch-QCTrigger.ps1`, `Run-QCProcessor.ps1`, `run_prepend_qc.ps1`).
- The `run_prepend_qc.ps1 -Legacy` switch calls into `legacy\prepend_qc_on_trigger.ps1`.

