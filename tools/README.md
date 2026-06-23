# `tools/` — Utilities and build tooling

This folder contains third-party binaries and Python build sources used by the pipeline.

## `qpdf`

- The pipeline uses **`qpdf.exe`** for PDF page concatenation/merging:
  - `modules/Processing/QC.Processors.psm1` (QC history prepend / merge)
  - `modules/Processing/QC.StatusSet.psm1` (status set PDF assembly)

Default expected path (unless overridden in `appsettings.json`):

- `tools\qpdf\bin\qpdf.exe`

## `overlay`

Python sources for the QC overlay / prepend tool. **Production workers use the built exe**, not Python:

- Build: `.\tools\overlay\build_overlay_exe.ps1` → `dist\qc_overlay_prepend\`
- Runtime: `dist\qc_overlay_prepend\qc_overlay_prepend.exe` (prepend, review stamps)

See [`tools/overlay/README.md`](overlay/README.md).
