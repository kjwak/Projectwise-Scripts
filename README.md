# Prepend PDF QC — Repository Overview

This repo implements a **queue-based QC pipeline** with:

- **Watcher**: discovers candidates, applies filters + triggers, and enqueues jobs.
- **Worker(s)**: dequeue jobs, dispatch to processors, and transition job state.
- **QC comment sync**: `QC_COMMENT_STATUS_SYNC` jobs react to lane QC PDF updates (`*-prod.pdf`, `*-rev.pdf`, `*-chk.pdf` — comment extraction, workflow state, telemetry). See [`docs/workflow/qc-comment-status-sync.md`](docs/workflow/qc-comment-status-sync.md).
- **Dashboard**: runs watcher + worker pool and renders a live terminal UI.

Most logic is in `modules/`. Most runnable entrypoints are in `scripts/`.

## Quick start (recommended)

- **Run the live dashboard**:
  - `.\scripts\service\Start-QCPipelineDashboard.ps1 -AppSettingsPath .\appsettings.json`

- **Run one watch tick (enqueue only)**:
  - `.\scripts\service\Watch-QCTrigger.ps1 -AppSettingsPath .\appsettings.json`

- **Run one worker tick (dequeue/process one job)**:
  - `.\scripts\service\Run-QCProcessor.ps1 -AppSettingsPath .\appsettings.json`

## Root files

- **`appsettings.json`**: primary configuration (ProjectWise watch list, filters, triggers, queue root, worker pool, processor modes).
- **`qc_overlay_prepend.spec`**: PyInstaller spec for building the overlay executable (`dist/qc_overlay_prepend/...`).
- **`pytest.ini`**: Python test configuration (`test/python/`).

## Root directories (high level)

- **`modules/`**: PowerShell modules implementing config/path normalization, filtering, triggers, job factory, JSON queue, processors, and native status set generation.
- **`scripts/`**: operational scripts (watcher, worker, dashboard, diagnostics, maintenance).
- **`legacy/`**: legacy monolith scripts kept for compatibility (`prepend_qc*.ps1`, `combine_status_set.ps1`).
- **`tools/`**: third-party utilities (`qpdf`) and overlay Python sources (`tools/overlay/`, built to `dist/qc_overlay_prepend/`).
- **`docs/`**: documentation — start at [`docs/README.md`](docs/README.md); agent guidance in [`AGENTS.md`](AGENTS.md).
- **`test/`**: unified tests — `test/powershell/` (PowerShell) and `test/python/` (pytest).
- **`dist/`**: packaged overlay exe (`dist/qc_overlay_prepend/`) used by production prepend and review stamps.
- **`tools/overlay/build/`**: PyInstaller intermediate output (gitignored); produced by `.\tools\overlay\build_overlay_exe.ps1`.

