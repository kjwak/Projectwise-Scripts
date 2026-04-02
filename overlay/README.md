# PDF Overlay – Old/New Comparison Layers

Creates a PDF with three toggleable layers: **Old** (red), **New** (green), **Current** (original color). Uses PyMuPDF + pikepdf for vector preservation and OCG support.

## Requirements

- Python 3.8+
- pymupdf >= 1.24
- pikepdf >= 8.0.0

```bash
pip install -r overlay/requirements.txt
```

## Usage

### One-command (recommended)

```bash
python overlay/run_complete_overlay.py old.pdf new.pdf output.pdf --colors-only
```

### Options

- `--fit` – Scale pages to match size
- `--old-color R,G,B` – Old layer color (default: 255,0,0)
- `--new-color R,G,B` – New layer color (default: 0,255,0)
- `--alpha 0.0-1.0` – Transparency for colored layers
- `--keep-intermediate` – Keep intermediate PDFs for debugging
- `--verbose` – More detailed output

### Two-step (advanced)

```bash
python overlay/overlay_build.py old.pdf new.pdf temp_overlay.pdf --fit
python overlay/overlay_layerize.py temp_overlay.pdf final.pdf --old-color 255,0,0 --new-color 0,255,0
```

## Output layers

| Layer   | Content                     | Default color |
|---------|-----------------------------|---------------|
| Old     | Old PDF linework            | Red           |
| New     | New PDF linework            | Green         |
| Current | New PDF in original color   | Black         |

Toggle layers in your PDF viewer (Bluebeam, Adobe Acrobat, etc.) using the Layers panel.

## QC History Prepend

For multi-sheet QC history PDFs, use `qc_overlay_prepend.py` to compare incoming vs **page 1 only** (the previous current sheet), then prepend the overlay to the history:

```bash
# Update QC history in place (page 1 of history = Old/red, incoming = New/green + Current/black):
python overlay/qc_overlay_prepend.py incoming.pdf sheet-qc.pdf

# Write to different output:
python overlay/qc_overlay_prepend.py incoming.pdf sheet-qc.pdf -o result.pdf
```

Result: overlay page (with 3 layers) becomes new page 1; previous page 1 shifts to page 2, etc.

**Authoring layers (Civil/CAD):** By default there is **no** preprocessing: the pipeline matches `run_complete_overlay.py` — vector artwork is embedded under Old / New / Current via `show_pdf_page`. If nested CAD layers still appear inside those three, pass **`--flatten-sources`** to merge authoring layers first (vector merge + strip). For stubborn files, add **`--flatten-raster`** and optionally **`--flatten-dpi 200`** (full-page images; vectors are lost on the source side).

## Standalone executable (no Python on the machine that runs QC)

Build **once** on a machine that has Python; deploy **`qc_overlay_prepend.exe`** next to your PowerShell scripts (repo root next to `prepend_qc.ps1`) or another path you pass as `-QcOverlayExe`. Those hosts do not need Python—PowerShell or any scheduler should call the `.exe` with PDF paths.

```powershell
# Build (developer / build agent only):
.\overlay\build_overlay_exe.ps1
```

This uses [`qc_overlay_prepend.spec`](../qc_overlay_prepend.spec) (portable paths) and writes `dist\qc_overlay_prepend\` (one-folder build). Copy `qc_overlay_prepend.exe` to the repo root beside `prepend_qc.ps1` when using a single-file build, or point `-QcOverlayExe` at `dist\qc_overlay_prepend\qc_overlay_prepend.exe` if you deploy the full folder.

**Typical automation (target machine, no Python):** from your trigger script, invoke the exe with full paths:

```powershell
$exe = "C:\Tools\qc_overlay_prepend.exe"   # wherever you deployed the single file
& $exe $incomingPdf $qcHistoryPdf -o $outputPdf
if ($LASTEXITCODE -ne 0) { throw "qc_overlay_prepend failed with exit code $LASTEXITCODE" }
```

Arguments match the CLI: `incoming.pdf`, `qc_history.pdf` (path may not exist on first run), optional `-o` output. See `python overlay/qc_overlay_prepend.py --help` for options (`--fit`, `--alpha`, etc.)—the same flags work on the exe.
