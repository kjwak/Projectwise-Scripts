# `legacy/` — Compatibility scripts

This folder contains scripts retained for **production prepend** and parity fallback while native paths are validated.

## Production relevance

Committed `appsettings.json` sets **`qcPrepend.mode: "legacyPw"`**. `QC_PREPEND` jobs therefore route through **`legacy/prepend_qc.ps1`** for ProjectWise export, overlay merge, lane PDF upload, and attribute sync. **Do not remove this folder** while that mode remains in use.

Native prepend logic exists in `modules/QC.Processors.psm1` (`mode` other than `legacyPw`) but is not the production default. See [`docs/phase-2-native-prepend-parity-plan.md`](../docs/phase-2-native-prepend-parity-plan.md).

## What's in here

- **`prepend_qc.ps1`**: production QC prepend (PW export + overlay + lane PDF history upload).
- **`prepend_qc_on_trigger.ps1`**: alternate entry used by `run_prepend_qc.ps1 -Legacy`.
- **`combine_status_set.ps1`**: legacy Status Set generation when `statusSet.mode = "legacy"`.
- **`Resolve-OverlayExe.ps1`**: overlay executable resolution (still referenced).
- **`watchlist.json`** (if present): legacy watch configuration.

## How it relates to the modular pipeline

The modern pipeline is modular (`modules/`) and queue-based (`modules/QC.Queue.Json.psm1`), driven by:

- `scripts/Watch-QCTrigger.ps1` (enqueue)
- `scripts/Run-QCProcessor.ps1` (dequeue/process)
- `scripts/Start-QCPipelineDashboard.ps1` (orchestrated dashboard run)

Processor legacy modes:

| Job type | Config key | Legacy script |
|----------|------------|---------------|
| `QC_PREPEND` | `qcPrepend.mode = "legacyPw"` | `legacy/prepend_qc.ps1` |
| `STATUS_SET_GEN` | `statusSet.mode = "legacy"` | `legacy/combine_status_set.ps1` |

Root shims (`Watch-QCTrigger.ps1`, `Run-QCProcessor.ps1`, `run_prepend_qc.ps1`) forward to `scripts/`.

## Lane PDF naming

When invoked from `Invoke-QCPrependProcessor` with strict lane params, prepend targets **`{stem}-prod.pdf`**, **`{stem}-rev.pdf`**, or **`{stem}-chk.pdf`** per `QC_Process_Type`.

Standalone legacy invocation without lane params may still create **`{stem}-qc.pdf`** (deprecated bridge). See `docs/qc-workflow-framework.md`.

## Three-lane workflow

Current TYPSA workflow states and process types are documented in [`docs/qc-workflow-framework.md`](../docs/qc-workflow-framework.md).
