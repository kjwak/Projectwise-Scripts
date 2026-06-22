# `tools/` — Third-party utilities

This folder contains third-party tools required by the pipeline.

## `qpdf`

- The pipeline uses **`qpdf.exe`** for PDF page concatenation/merging:
  - `modules/Processing/QC.Processors.psm1` (QC history prepend / merge)
  - `modules/Processing/QC.StatusSet.psm1` (status set PDF assembly)

Default expected path (unless overridden in `appsettings.json`):

- `tools\qpdf\bin\qpdf.exe`

