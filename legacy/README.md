# `legacy/` — Compatibility scripts

This folder contains **compatibility shims and shared helpers** — not the production QC prepend implementation.

## Production prepend (not in this folder)

Committed `appsettings.json` sets **`qcPrepend.mode: "projectWise"`**. `QC_PREPEND` jobs run through:

1. `Invoke-QCPrependProcessor` — `modules/Processing/QC.Processors.psm1`
2. `scripts/processing/Invoke-QCPrependPw.ps1` — PW export, overlay merge, lane PDF upload

`legacy/prepend_qc.ps1` forwards to that script with a deprecation warning. Do not point new automation at it.

Disk-only prepend (`qcPrepend.mode: "local"`) lives in `QC.Processors.psm1` and is for tests — not production. See [`docs/engineering/phase-2-native-prepend-parity-plan.md`](../docs/engineering/phase-2-native-prepend-parity-plan.md).

## What's in here

- **`prepend_qc.ps1`**: deprecated shim → `scripts/processing/Invoke-QCPrependPw.ps1`.
- **`prepend_qc_on_trigger.ps1`**: pre-queue monolith; used only by `run_prepend_qc.ps1 -Legacy`.
- **`combine_status_set.ps1`**: legacy Status Set generation when `statusSet.mode = "legacy"`.
- **`Logging.ps1`**, **`StaMtaRelaunch.ps1`**, **`Resolve-OverlayExe.ps1`**: dot-sourced by `Invoke-QCPrependPw.ps1` and related tools.
- **`watchlist.json`** (if present): legacy watch configuration.

## How it relates to the modular pipeline

The modern pipeline is modular (`modules/`) and queue-based (`modules/Queue/QC.Queue.Json.psm1`), driven by:

- `scripts/service/Watch-QCTrigger.ps1` (enqueue)
- `scripts/service/Run-QCProcessor.ps1` (dequeue/process)
- `scripts/Start-QCPipelineDashboard.ps1` (orchestrated dashboard run)

Processor paths:

| Job type | Config key | Script |
|----------|------------|--------|
| `QC_PREPEND` | `qcPrepend.mode = "projectWise"` | `scripts/processing/Invoke-QCPrependPw.ps1` |
| `STATUS_SET_GEN` | `statusSet.mode = "legacy"` | `legacy/combine_status_set.ps1` |

Compatibility wrappers under `scripts/` may forward to `scripts/service/`. Canonical entrypoints are `scripts/service/Watch-QCTrigger.ps1`, `scripts/service/Run-QCProcessor.ps1`, and `scripts/service/run_prepend_qc.ps1`. See [`AGENTS.md`](../AGENTS.md).

## Lane PDF naming

When invoked from `Invoke-QCPrependProcessor` with strict lane params, prepend targets **`{stem}-prod.pdf`**, **`{stem}-rev.pdf`**, or **`{stem}-chk.pdf`** per `QC_Process_Type`.

Standalone legacy invocation without lane params may still create **`{stem}-qc.pdf`** (deprecated bridge). See `docs/workflow/qc-workflow-framework.md`.

## Three-lane workflow

Current TYPSA workflow states and process types are documented in [`docs/workflow/qc-workflow-framework.md`](../docs/workflow/qc-workflow-framework.md).
