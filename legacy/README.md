# `legacy/` — Compatibility scripts

This folder contains older monolithic scripts retained for parity and fallback.

## What’s in here

- **`prepend_qc*.ps1`**: legacy QC prepend workflows (ProjectWise-triggered processing).
- **`combine_status_set.ps1`**: legacy Status Set generation workflow.
- **`watchlist.json`** (if present): legacy watch configuration.

## How it relates to the new pipeline

The modern pipeline is modular (`modules/`) and queue-based (`modules/QC.Queue.Json.psm1`), driven by:

- `scripts/Watch-QCTrigger.ps1` (enqueue)
- `scripts/Run-QCProcessor.ps1` (dequeue/process)
- `scripts/Start-QCPipelineDashboard.ps1` (orchestrated dashboard run)

Some processors still support a **legacy execution mode** for safety/parity:

- `QC_PREPEND` can run legacy prepend via `qcPrepend.mode = "legacyPw"` (uses legacy prepend script).
- `STATUS_SET_GEN` can run legacy combine via `statusSet.mode = "legacy"`.

This folder holds **legacy scripts** from the pre-modular QC pipeline.

- Keep these here temporarily while validating parity with the new framework.
- The new entrypoints are in repo root (`Watch-QCTrigger.ps1`, `Run-QCProcessor.ps1`, `run_prepend_qc.ps1`).
- The `run_prepend_qc.ps1 -Legacy` switch calls into `legacy\prepend_qc_on_trigger.ps1`.

