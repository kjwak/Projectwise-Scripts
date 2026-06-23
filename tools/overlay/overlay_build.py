#!/usr/bin/env python3
"""Build PDF overlay using PyMuPDF - creates vector XObjects without color/layer controls."""

import argparse
import logging
import re
import shutil
import sys
import tempfile
import time
from pathlib import Path
from typing import Optional

from flatten_source_layers import flatten_authoring_layers

try:
    import pymupdf as fitz  # type: ignore
except Exception:
    import fitz  # type: ignore

try:
    import pikepdf
except ImportError:
    print("Error: pikepdf is required for layer configurations. Install with: pip install pikepdf")
    sys.exit(1)

__version__ = "0.1.0"
LOGGER = logging.getLogger("overlay_build")


def _save_with_retry(doc, out_path: Path, *, deflate: bool = True, attempts: int = 8) -> None:
    """Save with short retries to ride out endpoint-security transient locks."""
    last_exc: Exception | None = None
    for i in range(attempts):
        try:
            doc.save(str(out_path), deflate=deflate)
            return
        except Exception as exc:  # pragma: no cover - environment dependent
            last_exc = exc
            # Increasing backoff: 100ms, 200ms, ... up to 800ms
            time.sleep(0.1 * (i + 1))
    if last_exc is not None:
        raise last_exc

def add_layer_configurations(pdf_path: Path) -> None:
    """Add layer view configurations to the PDF using pikepdf."""
    LOGGER.info("Adding layer view configurations...")
    
    # Create a temporary file to avoid overwriting issues
    temp_path = pdf_path.with_suffix('.tmp.pdf')
    
    with pikepdf.open(pdf_path) as pdf:
        ocprops = pdf.Root.get('/OCProperties')
        if not ocprops:
            LOGGER.warning("No OCProperties found - skipping layer configurations")
            return
        
        ocgs = ocprops.get('/OCGs')
        if not ocgs:
            LOGGER.warning("No OCGs found - skipping layer configurations")
            return
        
        # Find the three layers
        old_ocg = None
        new_ocg = None
        current_ocg = None
        
        for ocg in ocgs:
            name = str(ocg.get('/Name', ''))
            if 'Old' in name:
                old_ocg = ocg
            elif 'New' in name:
                new_ocg = ocg
            elif 'Current' in name:
                current_ocg = ocg
        
        if not all([old_ocg, new_ocg, current_ocg]):
            LOGGER.warning("Could not find all three layers - skipping layer configurations")
            return
        
        LOGGER.info("✅ Found all three layers: Old, New, Current")
        
        # Keep Usage properties simple to avoid viewer compatibility issues
        # Just set basic CreatorInfo without complex Print states
        for layer_ocg in [old_ocg, new_ocg, current_ocg]:
            if '/Usage' not in layer_ocg:
                layer_ocg['/Usage'] = pikepdf.Dictionary()
            usage = layer_ocg['/Usage']
            
            if '/CreatorInfo' not in usage:
                usage['/CreatorInfo'] = pikepdf.Dictionary({
                    '/Creator': 'PyMuPDF',
                    '/Subtype': '/Artwork'
                })
        
        # Create default configuration
        if '/D' not in ocprops:
            ocprops['/D'] = pikepdf.Dictionary()
        default_config = ocprops['/D']
        
        # Clear existing config
        for key in list(default_config.keys()):
            del default_config[key]
        
        # Set layer order
        default_config['/Order'] = pikepdf.Array([old_ocg, new_ocg, current_ocg])
        
        # Default config: Old and New OFF, Current ON (implicit)
        default_config['/OFF'] = pikepdf.Array([old_ocg, new_ocg])
        
        # Skip AS configuration - it often causes viewer compatibility issues
        # Just use simple layer ordering and OFF settings
        
        # Create alternative configuration for "Difference" view (simplified)
        difference_config = pikepdf.Dictionary()
        difference_config['/Name'] = 'Difference'
        difference_config['/OFF'] = pikepdf.Array([current_ocg])
        
        # Add configurations to OCProperties
        ocprops['/Configs'] = pikepdf.Array([difference_config])
        
        LOGGER.info("✅ Enhanced layer configurations added successfully!")
        LOGGER.info("📊 Default view: Shows only 'Current' layer")
        LOGGER.info("🎨 Difference view: Shows only 'Old' and 'New' layers")
        
        # Save the updated PDF to temporary file
        pdf.save(temp_path)
    
    # Replace the original file with the updated one
    temp_path.replace(pdf_path)

def fit_rect(src: fitz.Rect, dst: fitz.Rect) -> fitz.Rect:
    """Scale source rectangle to fit within destination while preserving aspect ratio."""
    sx = dst.width / src.width if src.width else 1.0
    sy = dst.height / src.height if src.height else 1.0
    scale = min(sx, sy)
    new_width = src.width * scale
    new_height = src.height * scale
    x0 = dst.x0 + (dst.width - new_width) / 2.0
    y0 = dst.y0 + (dst.height - new_height) / 2.0
    return fitz.Rect(x0, y0, x0 + new_width, y0 + new_height)

def _fitz_enable_all_optional_content(doc) -> None:
    """Turn all OCGs on in the default layer config so show_pdf_page embeds full visible art.

    Layer configs (e.g. Old/New OFF, Current ON) can otherwise make the Old slot look empty
    when the source is a full QC page with internal OCGs.
    """
    if not doc.is_pdf:
        return
    ocgs = doc.get_ocgs()
    if not ocgs:
        return
    xrefs = list(ocgs.keys())
    try:
        doc.set_layer(-1, on=xrefs)
    except Exception as exc:
        LOGGER.debug("Could not enable all OCGs on source: %s", exc)


def _fitz_qc_triplet_xrefs(ocgs: dict) -> dict[str, int]:
    """Map Old / New / Current -> xref (same rules as qc_overlay_prepend.extract_page_1_current_as_old)."""
    out: dict[str, int] = {}
    for xref, info in ocgs.items():
        nm = str(info.get("name", "")).strip().strip("/")
        key = nm.lower()
        if key == "old":
            out["Old"] = xref
        elif key == "new":
            out["New"] = xref
        elif key == "current":
            out["Current"] = xref
    return out


def _fitz_apply_old_source_layer_visibility(doc) -> None:
    """Match page-1 extract: QC Old+New off, everything else on (previous Current visible in Old slot).

    Re-applying here matters because show_pdf_page can embed optional content such that the
    reopened extract's default /D state is not what MuPDF uses when grafting into the overlay,
    which produced an empty Old layer while New/Current still drew.
    """
    if not doc.is_pdf:
        return
    ocgs = doc.get_ocgs()
    if not ocgs:
        return
    qc = _fitz_qc_triplet_xrefs(ocgs)
    if not (qc.get("Old") and qc.get("New") and qc.get("Current")):
        return
    all_x = list(ocgs.keys())
    off_list = [qc["Old"], qc["New"]]
    on_list = [x for x in all_x if x not in off_list]
    try:
        doc.set_layer(-1, on=on_list, off=off_list)
        LOGGER.info(
            "Old source: set_layer (QC Old/New off, Current + base layers on) before overlay embed"
        )
    except Exception as exc:
        LOGGER.warning("Old source set_layer failed (Old slot may be empty): %s", exc)


# QC compare form XObjects invoked from page Contents (PyMuPDF overlay / layerize names).
_QC_COMPARE_DO_RE = re.compile(
    rb"/(?:fzFrm\d+|BBL_Current|BBL1|BBL)\s+Do"
)


def _fitz_page_contents_bytes(page) -> bytes:
    """Concatenated page content streams (PyMuPDF)."""
    try:
        return page.read_contents()
    except Exception:
        pass
    try:
        c = page.get_contents()
        if isinstance(c, bytes):
            return c
    except Exception:
        pass
    return b""


def _fitz_qc_compare_form_do_count(page) -> int:
    """How many QC compare form Do operators appear in page Contents."""
    return len(_QC_COMPARE_DO_RE.findall(_fitz_page_contents_bytes(page)))


def _fitz_enable_cad_for_trimmed_old_source(doc) -> None:
    """When Old input draws exactly one QC compare form in Contents (pikepdf trim), Civil/DGN
    OCGs nested inside that form may still default off — show_pdf_page then embeds an empty Old.

    Turning all OCGs on is safe here: other QC forms (Old/New slots) are not invoked by Contents,
    so they do not draw. Do not run when multiple QC compare Dos appear (full page — would stack).

    Runs after _fitz_apply_old_source_layer_visibility; may override layer state so nested CAD art
    inside the single form is visible.
    """
    if not doc.is_pdf or doc.page_count < 1:
        return
    ocgs = doc.get_ocgs()
    if not ocgs:
        return
    if _fitz_qc_compare_form_do_count(doc[0]) != 1:
        return
    _fitz_enable_all_optional_content(doc)
    LOGGER.info(
        "Old source: single QC compare form in Contents — enabled all OCGs for overlay embed "
        "(nested Civil/DGN layers inside that form)"
    )


def parse_page_ranges(spec: Optional[str], max_pages: int) -> list[int]:
    """Parse 1-based page range specification."""
    if max_pages < 0:
        raise ValueError("max_pages must be non-negative")
    if not spec:
        return list(range(max_pages))

    pages: list[int] = []
    tokens = [part.strip() for part in spec.split(",") if part.strip()]
    for token in tokens:
        if "-" in token:
            start_str, end_str = token.split("-", 1)
            start = 1 if not start_str else int(start_str)
            end = max_pages if not end_str else int(end_str)
            if start < 1:
                raise ValueError(f"Invalid page range start: {token}")
            if end < start:
                raise ValueError(f"Invalid page range: {token}")
            for page in range(start, min(end, max_pages) + 1):
                index = page - 1
                if index not in pages and index < max_pages:
                    pages.append(index)
        else:
            page = int(token)
            if page < 1:
                raise ValueError(f"Invalid page number: {token}")
            index = page - 1
            if index < max_pages and index not in pages:
                pages.append(index)
    return pages

def build_overlay(old_path: Path, new_path: Path, out_path: Path, 
                 pages_spec: Optional[str] = None, fit: bool = False, 
                 canvas: str = "new", add_configs: bool = True,
                 flatten_sources: bool = False,
                 flatten_dpi: float = 144.0,
                 flatten_raster: bool = False) -> None:
    """Build basic PDF overlay with vector XObjects.

    Default (flatten_sources=False): direct show_pdf_page — same as run_complete_overlay.py.
    Optional flatten_sources: strip/merge authoring layers before overlay; use flatten_raster
    for full-page images (CAD fallback); otherwise vector merge + strip.
    """
    
    LOGGER.info(f"Building overlay: {old_path} + {new_path} -> {out_path}")
    
    tmp_flat: Optional[Path] = None
    open_old = old_path
    open_new = new_path
    if flatten_sources:
        tmp_flat = Path(tempfile.mkdtemp(prefix="overlay_flat_"))
        try:
            open_old = tmp_flat / "_old_flat.pdf"
            open_new = tmp_flat / "_new_flat.pdf"
            mode = "raster" if flatten_raster else "vector"
            LOGGER.info(
                "Flattening source PDFs for overlay (mode=%s, dpi=%s)...",
                mode,
                flatten_dpi,
            )
            flatten_authoring_layers(
                old_path,
                open_old,
                dpi=flatten_dpi,
                mode=mode,
            )
            flatten_authoring_layers(
                new_path,
                open_new,
                dpi=flatten_dpi,
                mode=mode,
            )
        except Exception as exc:
            LOGGER.warning("Source flatten failed, using originals: %s", exc)
            open_old, open_new = old_path, new_path
    
    # Open source documents (optionally pre-flattened)
    old_doc = None
    new_doc = None
    try:
        old_doc = fitz.open(str(open_old))
        new_doc = fitz.open(str(open_new))
    except Exception:
        if tmp_flat is not None and tmp_flat.exists():
            shutil.rmtree(tmp_flat, ignore_errors=True)
        raise

    try:
        # Old input is often a page-1 extract with QC layer defaults (Current on, Old/New off).
        # Do not force all OCGs on here — that stacks all three QC compare layers into the Old slot.
        # Re-apply the same visibility as extract_page_1_current_as_old() so show_pdf_page embeds
        # the previous Current (not empty nested OCG).
        _fitz_apply_old_source_layer_visibility(old_doc)
        _fitz_enable_cad_for_trimmed_old_source(old_doc)

        _fitz_enable_all_optional_content(new_doc)

        max_pairs = min(len(old_doc), len(new_doc))
        if max_pairs == 0:
            raise ValueError("Both PDFs must contain at least one page.")
        
        selected_pages = parse_page_ranges(pages_spec, max_pairs)
        if not selected_pages:
            raise ValueError("No pages selected for overlay.")
        
        LOGGER.info(f"Processing {len(selected_pages)} page pairs")
        
        # Create output document
        out_path.parent.mkdir(parents=True, exist_ok=True)
        output_doc = fitz.open()
        
        try:
            for page_index in selected_pages:
                LOGGER.debug(f"Processing page {page_index + 1}")
                
                old_page = old_doc[page_index]
                new_page = new_doc[page_index]
                
                # Capture rects and rotation before any modification (rect = display size)
                old_rect = old_page.rect
                new_rect = new_page.rect
                rot_old = getattr(old_page, 'rotation', 0)
                rot_new = getattr(new_page, 'rotation', 0)
                
                # Flatten source rotation: set_rotation(0) so show_pdf_page works predictably,
                # then pass rotate=-rot to preserve visual (PyMuPDF #1378)
                if rot_old:
                    old_page.set_rotation(0)
                if rot_new:
                    new_page.set_rotation(0)
                
                # Determine canvas size (use original display rects)
                canvas_rect = fitz.Rect(0, 0, new_rect.width, new_rect.height) if canvas == "new" else fitz.Rect(0, 0, old_rect.width, old_rect.height)
                output_page = output_doc.new_page(width=canvas_rect.width, height=canvas_rect.height)
                
                # Calculate destination rectangles (use original rects for display size)
                if fit:
                    old_dest = fit_rect(old_rect, canvas_rect)
                    new_dest = fit_rect(new_rect, canvas_rect)
                else:
                    old_dest = fitz.Rect(0, 0, old_rect.width, old_rect.height)
                    new_dest = fitz.Rect(0, 0, new_rect.width, new_rect.height)
                
                # Create separate XObjects with embedded OCGs (like layered_overlay.py)
                ocg_old = output_doc.add_ocg(f"Old", on=True)
                ocg_new = output_doc.add_ocg(f"New", on=True)
                ocg_current = output_doc.add_ocg(f"Current", on=False)
                
                # Place pages; rotate=-rot preserves visual after set_rotation(0)
                output_page.show_pdf_page(old_dest, old_doc, page_index, oc=ocg_old, rotate=-rot_old)
                output_page.show_pdf_page(new_dest, new_doc, page_index, oc=ocg_new, rotate=-rot_new)
                output_page.show_pdf_page(new_dest, new_doc, page_index, oc=ocg_current, rotate=-rot_new)
                
                LOGGER.debug(f"Placed 3 XObjects (Old, New, Current) for page {page_index + 1}")
            
            # Save the basic overlay
            LOGGER.info(f"Saving basic overlay to {out_path}")
            _save_with_retry(output_doc, out_path, deflate=True)
            
        finally:
            output_doc.close()

        # Close inputs and remove flatten temp *before* pikepdf touches the overlay (Windows file locks).
        if old_doc is not None and not old_doc.is_closed:
            old_doc.close()
        if new_doc is not None and not new_doc.is_closed:
            new_doc.close()
        if tmp_flat is not None and tmp_flat.exists():
            shutil.rmtree(tmp_flat, ignore_errors=True)
            tmp_flat = None

        if add_configs:
            LOGGER.info("Adding layer view configurations...")
            add_layer_configurations(out_path)

    finally:
        if old_doc is not None and not old_doc.is_closed:
            old_doc.close()
        if new_doc is not None and not new_doc.is_closed:
            new_doc.close()
        if tmp_flat is not None and tmp_flat.exists():
            shutil.rmtree(tmp_flat, ignore_errors=True)

def build_parser() -> argparse.ArgumentParser:
    """Build command line argument parser."""
    parser = argparse.ArgumentParser(
        description="Build basic PDF overlay with vector XObjects",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("old_pdf", type=Path, help="Path to OLD PDF")
    parser.add_argument("new_pdf", type=Path, help="Path to NEW PDF")
    parser.add_argument("out_pdf", type=Path, help="Output PDF path")
    parser.add_argument("--pages", dest="pages_spec", help="1-based page ranges, e.g. 1-3,7,9-", default=None)
    parser.add_argument("--fit", action="store_true", help="Scale pages to fit output canvas")
    parser.add_argument("--canvas", choices=("new", "old"), default="new", help="Which PDF defines the output page size")
    parser.add_argument("--no-configs", action="store_true", help="Skip adding layer view configurations")
    parser.add_argument("--flatten-sources", action="store_true",
                        help="Preprocess inputs to merge/strip authoring layers before overlay (vector merge by default)")
    parser.add_argument("--flatten-raster", action="store_true",
                        help="With --flatten-sources, rasterize each page instead of vector merge")
    parser.add_argument("--flatten-dpi", type=float, default=144.0,
                        help="Resolution when using --flatten-raster (ignored for vector merge)")
    parser.add_argument("--verbose", action="store_true", help="Enable verbose logging")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser

def configure_logging(verbose: bool) -> None:
    """Configure logging level."""
    level = logging.INFO if verbose else logging.WARNING
    logging.basicConfig(level=level, format="%(levelname)s: %(message)s")

def main(argv: Optional[list[str]] = None) -> int:
    """Main entry point."""
    parser = build_parser()
    args = parser.parse_args(argv)
    
    configure_logging(args.verbose)
    
    try:
        build_overlay(
            args.old_pdf, 
            args.new_pdf, 
            args.out_pdf,
            pages_spec=args.pages_spec,
            fit=args.fit,
            canvas=args.canvas,
            add_configs=not args.no_configs,
            flatten_sources=args.flatten_sources,
            flatten_dpi=args.flatten_dpi,
            flatten_raster=args.flatten_raster,
        )
        LOGGER.info("Overlay build completed successfully")
        return 0
    except Exception as exc:
        LOGGER.error(f"Overlay build failed: {exc}")
        return 1

if __name__ == "__main__":
    sys.exit(main())
