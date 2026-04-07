#!/usr/bin/env python3
"""Add layers and color controls to PDF overlay using the complete Bluebeam approach."""

import argparse
import logging
import re
import sys
from pathlib import Path
from typing import Any, Optional, Tuple

try:
    import pikepdf
except ImportError:
    print("Error: pikepdf is required. Install with: pip install pikepdf")
    sys.exit(1)

__version__ = "0.5.0"
LOGGER = logging.getLogger("overlay_layerize")

# /fullpage streams shorter than this are treated as stubs (not real vector paint).
_FULLPAGE_STUB_MAX = 512


def _find_nested_qc_current_form(fullpage_obj) -> Any:
    """When PyMuPDF embeds a prior QC page, /fullpage may be a stub whose Resources hold BBL_Current."""
    res = fullpage_obj.get("/Resources")
    if not res or "/XObject" not in res:
        return None
    xo = res["/XObject"]
    for key, child in xo.items():
        nk = str(key).strip("'").lstrip("/")
        if nk in ("BBL_Current", "fzFrm2"):
            return child
    return None


def _try_promote_stub_fullpage_to_nested_current(xobj, pdf: pikepdf.Pdf, xobjects: Any, name) -> bool:
    """Replace Old (etc.) form with nested BBL_Current when /fullpage is a tiny stub wrapping prior QC art."""
    try:
        if "/Resources" not in xobj:
            return False
        res = xobj["/Resources"]
        if "/XObject" not in res:
            return False
        xo = res["/XObject"]
        if "/fullpage" not in xo:
            return False
        fp = xo["/fullpage"]
        fp_b = fp.read_bytes()
    except Exception:
        return False

    if len(fp_b) >= _FULLPAGE_STUB_MAX:
        return False

    cur = _find_nested_qc_current_form(fp)
    if cur is None:
        return False
    try:
        paint_b = cur.read_bytes()
    except Exception:
        return False
    if len(paint_b) < _FULLPAGE_STUB_MAX:
        return False

    ocg = xobj.get("/OC")
    new_stream = pikepdf.Stream(pdf, paint_b)
    for key, value in cur.items():
        if key not in ("/Length", "/Filter"):
            new_stream[key] = value
    if ocg is not None:
        new_stream["/OC"] = ocg

    xobjects[name] = new_stream
    LOGGER.info(
        "Promoted nested BBL_Current into %s (stub fullpage %d B -> paint %d B)",
        name,
        len(fp_b),
        len(paint_b),
    )
    return True


def _merge_pdf_resources(dst: pikepdf.Dictionary, src: pikepdf.Dictionary) -> None:
    """Merge src into dst without overwriting existing keys (parent wins on conflict)."""
    for key in src.keys():
        if key not in dst:
            dst[key] = src[key]
            continue
        if key in ("/XObject", "/Font", "/ExtGState", "/ColorSpace", "/Pattern", "/Shading", "/Properties"):
            if isinstance(dst[key], pikepdf.Dictionary) and isinstance(src[key], pikepdf.Dictionary):
                for sk in src[key].keys():
                    if sk not in dst[key]:
                        dst[key][sk] = src[key][sk]


def _try_inline_fullpage_wrapper(
    xobj,
    pdf: pikepdf.Pdf,
    xobjects: Any,
    name,
) -> bool:
    """If form is a short wrapper that invokes /fullpage and fullpage holds real paint, inline it.

    PyMuPDF often puts a tiny or short wrapper in the parent Form and the real vectors in
    /Resources/XObject/fullpage. Layerize used to prefer /fullpage bytes — but when /fullpage was
    a stub, the ratio heuristic picked the *wrapper* stream, recolored it, deleted /fullpage, and
    wiped the Old layer. Inlining merges fullpage paint + resources into the parent Form first.
    """
    try:
        if "/Resources" not in xobj:
            return False
        res = xobj["/Resources"]
        if "/XObject" not in res:
            return False
        xo = res["/XObject"]
        if "/fullpage" not in xo:
            return False
        fp = xo["/fullpage"]
        parent_b = xobj.read_bytes()
        fp_b = fp.read_bytes()
    except Exception:
        return False

    if len(fp_b) < _FULLPAGE_STUB_MAX:
        return False
    if len(parent_b) > max(65536, len(fp_b) * 2):
        return False
    low = parent_b.lower()
    if b"/fullpage" not in low and b"fullpage" not in low:
        return False

    if "/Resources" in fp:
        if "/Resources" not in xobj:
            xobj["/Resources"] = pikepdf.Dictionary()
        _merge_pdf_resources(xobj["/Resources"], fp["/Resources"])

    if "/Resources" in xobj and "/XObject" in xobj["/Resources"]:
        xo2 = xobj["/Resources"]["/XObject"]
        if "/fullpage" in xo2:
            del xo2["/fullpage"]

    new_stream = pikepdf.Stream(pdf, fp_b)
    for key, value in xobj.items():
        if key not in ("/Length", "/Filter"):
            new_stream[key] = value

    xobjects[name] = new_stream
    LOGGER.info(
        "Inlined /fullpage paint into form %s (parent %d B -> %d B stream)",
        name,
        len(parent_b),
        len(fp_b),
    )
    return True


def _inline_fullpage_wrappers_on_page(page, pdf: pikepdf.Pdf) -> None:
    if "/Resources" not in page:
        return
    resources = page["/Resources"]
    if "/XObject" not in resources:
        return
    xobjects = resources["/XObject"]
    for name in list(xobjects.keys()):
        try:
            if _try_promote_stub_fullpage_to_nested_current(xobjects[name], pdf, xobjects, name):
                continue
            _try_inline_fullpage_wrapper(xobjects[name], pdf, xobjects, name)
        except Exception as exc:
            LOGGER.debug("inline fullpage skip %s: %s", name, exc)


def _pick_form_paint_content_bytes(xobj, fullpage_xobj) -> Optional[bytes]:
    """Choose PDF form stream bytes for recoloring.

    PyMuPDF sometimes embeds show_pdf_page as a Form XObject with a tiny /fullpage stub and the
    real paint on the parent stream (or the reverse). Preferring /fullpage unconditionally can
    replace the Old layer with ~tens of bytes and leave the layer empty on round 2+ QC history.
    """
    direct: Optional[bytes] = None
    fullpage: Optional[bytes] = None
    try:
        direct = xobj.read_bytes()
    except Exception:
        pass
    if fullpage_xobj is not None:
        try:
            fullpage = fullpage_xobj.read_bytes()
        except Exception:
            pass
    if not direct and not fullpage:
        return None
    if not fullpage:
        return direct
    if not direct:
        return fullpage

    # Both streams are tiny stubs — real paint may be nested (handled by promote pass); do not recolor.
    if max(len(direct), len(fullpage)) < _FULLPAGE_STUB_MAX:
        LOGGER.warning(
            "Direct and /fullpage are both tiny (%d and %d B); skip stream recolor",
            len(direct),
            len(fullpage),
        )
        return None

    # Tiny /fullpage: never treat as the main paint body by itself.
    if len(fullpage) < _FULLPAGE_STUB_MAX:
        if direct and len(direct) > len(fullpage):
            low = direct.lower()
            if (b"/fullpage" in low or b"fullpage" in low) and len(direct) < 16384:
                # Wrapper around a stub — safe recolor needs inline step first; do not replace stream.
                LOGGER.warning(
                    "Form has stub /fullpage (%d B) and short wrapper parent (%d B); "
                    "skip destructive recolor (run inline pass or use vector flatten on sources).",
                    len(fullpage),
                    len(direct),
                )
                return None

    # One stream is a stub: take the larger paint payload.
    if max(len(direct), len(fullpage)) > 8 * min(len(direct), len(fullpage)) + 64:
        return direct if len(direct) >= len(fullpage) else fullpage
    return fullpage


def add_colors_like_bluebeam(page: pikepdf.Object,
                            old_color: Tuple[float, float, float],
                            new_color: Tuple[float, float, float],
                            old_alpha: float,
                            new_alpha: float,
                            pdf: pikepdf.Pdf) -> None:
    """Add colors using the complete Bluebeam approach with proper content stream manipulation."""
    
    LOGGER.info(f"Adding colors using complete Bluebeam approach: OLD={old_color} (α={old_alpha}), NEW={new_color} (α={new_alpha})")
    
    # Keep original OCG settings - focus on color application
    
    # Get or create page resources
    if '/Resources' not in page:
        page['/Resources'] = pikepdf.Dictionary()
    resources = page['/Resources']
    
    # Add ExtGStates to page resources (like Bluebeam's /BBGS)
    if '/ExtGState' not in resources:
        resources['/ExtGState'] = pikepdf.Dictionary()
    extgstate = resources['/ExtGState']
    
    # Use Darken so Old/New overlap trends dark while reducing top-layer color dominance.
    # Old layer graphics state
    bbgs_old = pikepdf.Dictionary({
        '/Type': '/ExtGState',
        '/BM': '/Darken',
        '/CA': old_alpha,  # Stroke alpha
        '/ca': old_alpha   # Fill alpha
    })
    extgstate['/BBGS_Old'] = bbgs_old
    
    # New layer graphics state  
    bbgs_new = pikepdf.Dictionary({
        '/Type': '/ExtGState',
        '/BM': '/Darken',
        '/CA': new_alpha,  # Stroke alpha
        '/ca': new_alpha   # Fill alpha
    })
    extgstate['/BBGS_New'] = bbgs_new
    
    # Check if we have XObjects
    if '/XObject' not in resources:
        LOGGER.warning("Page has no XObjects")
        return
    
    xobjects = resources['/XObject']

    # Hoist PyMuPDF /fullpage paint into parent forms before recolor (avoids empty Old/New).
    _inline_fullpage_wrappers_on_page(page, pdf)

    # Process existing XObjects to add Bluebeam-style graphics states
    for xobj_name, xobj_ref in xobjects.items():
        try:
            xobj = xobj_ref
            if '/OC' in xobj:
                ocg_name = str(xobj['/OC'].get('/Name', ''))
                LOGGER.info(f"Processing XObject {xobj_name} (OCG: {ocg_name})")
                
                # Determine color and alpha for this layer
                if 'Old' in ocg_name:
                    layer_color = old_color
                    layer_alpha = old_alpha
                    LOGGER.info(f"Applying red color (α={layer_alpha}) to {xobj_name}")
                elif 'New' in ocg_name:
                    layer_color = new_color
                    layer_alpha = new_alpha
                    LOGGER.info(f"Applying green color (α={layer_alpha}) to {xobj_name}")
                elif 'Current' in ocg_name:
                    # Skip coloring for Current layer - keep original black
                    LOGGER.info(f"Preserving original colors for {xobj_name} (Current layer)")
                    continue
                else:
                    LOGGER.warning(f"Unknown layer type for {xobj_name}: {ocg_name}")
                    continue
                
                # Extract content from XObject (direct stream and/or /fullpage — see _pick_form_paint_content_bytes)
                fullpage_xobj = None
                if '/Resources' in xobj and '/XObject' in xobj['/Resources']:
                    xres = xobj['/Resources']['/XObject']
                    if '/fullpage' in xres:
                        fullpage_xobj = xres['/fullpage']
                        LOGGER.info("Found /fullpage reference alongside parent form stream")
                content_to_process = _pick_form_paint_content_bytes(xobj, fullpage_xobj)
                if content_to_process is not None:
                    LOGGER.info(
                        "Using %d bytes for %s recolor (direct vs /fullpage resolved)",
                        len(content_to_process),
                        xobj_name,
                    )
                if not content_to_process:
                    LOGGER.warning(f"Could not extract form bytes for {xobj_name}")
                    continue

                # Inject colors into the content
                content_str = content_to_process.decode('latin-1')

                # Replace all grayscale colors with our layer color
                # This includes black (0 0 0) and all gray shades (where R=G=B)
                def replace_grayscale_rg(match):
                    r, g, b = match.groups()
                    if r == g == b:  # Grayscale color
                        return f'{layer_color[0]} {layer_color[1]} {layer_color[2]} RG'
                    return match.group(0)  # Keep original if not grayscale

                content_str = re.sub(r'([0-9.]+)\s+([0-9.]+)\s+([0-9.]+)\s+RG', replace_grayscale_rg, content_str)

                # Replace grayscale fill colors (rg)
                def replace_grayscale_rg_fill(match):
                    r, g, b = match.groups()
                    if r == g == b:  # Grayscale color
                        return f'{layer_color[0]} {layer_color[1]} {layer_color[2]} rg'
                    return match.group(0)  # Keep original if not grayscale

                content_str = re.sub(r'([0-9.]+)\s+([0-9.]+)\s+([0-9.]+)\s+rg', replace_grayscale_rg_fill, content_str)

                # Create new stream with the modified content
                new_content = content_str.encode('latin-1')

                # Update the XObject with the new content
                # Remove any existing filter to avoid compression issues
                if '/Filter' in xobj:
                    del xobj['/Filter']

                # Create the new stream
                new_stream = pikepdf.Stream(pdf, new_content)

                # Copy other properties from original XObject
                for key, value in xobj.items():
                    if key not in ['/Length', '/Filter']:
                        new_stream[key] = value

                # Keep XObject simple - transparency is handled at page level

                # Replace the XObject content
                xobjects[xobj_name] = new_stream
                placed = xobjects[xobj_name]

                # Remove the /fullpage reference from resources if it exists (use placed, not stale xobj)
                if '/Resources' in placed and '/XObject' in placed['/Resources']:
                    xobj_resources = placed['/Resources']['/XObject']
                    if '/fullpage' in xobj_resources:
                        del xobj_resources['/fullpage']

                LOGGER.info(f"✅ Modified {xobj_name} content with {['red', 'green'][layer_color == new_color]} color ({layer_color}) and transparency (α={layer_alpha}) - {len(new_content)} bytes")

        except Exception as e:
            LOGGER.error(f"Error processing XObject {xobj_name}: {e}")
            continue
    
    # Handle Current layer (fzFrm2 -> BBL_Current) separately
    if '/fzFrm2' in xobjects:
        LOGGER.info("Processing Current layer (/fzFrm2)")
        current_xobj = xobjects['/fzFrm2']
        current_resources = None
        if '/Resources' in current_xobj and '/XObject' in current_xobj['/Resources']:
            current_resources = current_xobj['/Resources']['/XObject']
        if current_resources is not None and '/fullpage' in current_resources:
            LOGGER.info("Found /fullpage reference in Current layer - resolving paint stream")
            fullpage_xobj = current_resources['/fullpage']

            try:
                fullpage_content = _pick_form_paint_content_bytes(current_xobj, fullpage_xobj)
                if not fullpage_content:
                    raise ValueError("no form bytes for Current layer")
                LOGGER.info(f"Using {len(fullpage_content)} bytes for Current layer (direct vs /fullpage resolved)")

                if '/Filter' in current_xobj:
                    del current_xobj['/Filter']

                new_stream = pikepdf.Stream(pdf, fullpage_content)

                for key, value in current_xobj.items():
                    if key not in ['/Length', '/Filter']:
                        new_stream[key] = value

                xobjects['/BBL_Current'] = new_stream
                del xobjects['/fzFrm2']

                if '/fullpage' in current_resources:
                    del current_resources['/fullpage']

                LOGGER.info(f"✅ Created /BBL_Current layer (original black colors) - {len(fullpage_content)} bytes")

            except Exception as e:
                LOGGER.error(f"Error processing Current layer content: {e}")

        if '/fzFrm2' in xobjects and '/BBL_Current' not in xobjects:
            xobjects['/BBL_Current'] = xobjects['/fzFrm2']
            del xobjects['/fzFrm2']
            LOGGER.info("Renamed /fzFrm2 -> /BBL_Current (paint already on form stream)")

    # Rename XObjects to match Bluebeam naming (fzFrm0 -> BBL, fzFrm1 -> BBL1)
    if '/fzFrm0' in xobjects:
        xobjects['/BBL'] = xobjects['/fzFrm0']
        del xobjects['/fzFrm0']
    
    if '/fzFrm1' in xobjects:
        xobjects['/BBL1'] = xobjects['/fzFrm1'] 
        del xobjects['/fzFrm1']
    
    # Update page content streams to use the renamed XObjects
    if '/Contents' in page:
        contents = page['/Contents']
        new_contents = []
        
        # Add /BBL (Old layer) with transparency graphics state
        bbl_content = pikepdf.Stream(pdf, b'q /BBGS_Old gs 1 0 0 1 0 0 cm q 1 0 0 1 0 0 cm /BBL Do Q Q ')
        new_contents.append(bbl_content)
        
        # Add /BBL1 (New layer) with transparency graphics state
        bbl1_content = pikepdf.Stream(pdf, b'q /BBGS_New gs 1 0 0 1 0 0 cm q 1 0 0 1 0 0 cm /BBL1 Do Q Q ')
        new_contents.append(bbl1_content)
        
        # Add /BBL_Current if it exists
        if '/BBL_Current' in xobjects:
            bbl_current_content = pikepdf.Stream(pdf, b'q 1 0 0 1 0 0 cm q 1 0 0 1 0 0 cm /BBL_Current Do Q Q ')
            new_contents.append(bbl_current_content)
            LOGGER.info("✅ Added Current layer to page content streams")
        
        # Replace the page contents with our new array
        page['/Contents'] = pikepdf.Array(new_contents)
        layer_count = len(new_contents)
        LOGGER.info(f"✅ Updated page contents to use {layer_count} layers (/BBL, /BBL1" + (", /BBL_Current)" if '/BBL_Current' in xobjects else ")"))   
    
    # The colors and transparency are now applied at the page level for proper blending
    LOGGER.info(f"✅ Color and transparency injection completed - colors (α={old_alpha}, α={new_alpha}) are applied at page level for proper blending")

def layerize_overlay(
    input_path: Path,
    output_path: Path,
    old_color: Tuple[float, float, float] = (1.0, 0.0, 0.0),
    new_color: Tuple[float, float, float] = (0.0, 1.0, 0.0),
    old_alpha: float = 0.2,
    new_alpha: float = 0.2,
) -> None:
    """Main function to layerize overlay using complete Bluebeam approach."""
    LOGGER.info(f"Layerizing overlay (complete Bluebeam approach): {input_path} -> {output_path}")
    
    with pikepdf.open(input_path, allow_overwriting_input=True) as pdf:
        for page_index, page in enumerate(pdf.pages):
            add_colors_like_bluebeam(page, old_color, new_color, old_alpha, new_alpha, pdf)
        
        output_path.parent.mkdir(parents=True, exist_ok=True)
        pdf.save(output_path)
        LOGGER.info(f"Layerized overlay saved to {output_path}")

def parse_color(color_str: str) -> Tuple[float, float, float]:
    """Parse color string in format R,G,B (0-255) to float tuple (0.0-1.0)."""
    parts = color_str.split(',')
    if len(parts) != 3:
        raise ValueError("Color must be in format R,G,B (e.g., 255,0,0)")
    
    try:
        r, g, b = [int(part.strip()) for part in parts]
        if not all(0 <= val <= 255 for val in [r, g, b]):
            raise ValueError("Color values must be between 0 and 255")
        return (r/255.0, g/255.0, b/255.0)
    except ValueError as e:
        raise ValueError(f"Invalid color format: {e}")

def build_parser() -> argparse.ArgumentParser:
    """Build command line argument parser."""
    parser = argparse.ArgumentParser(
        description="Add colors using the complete Bluebeam approach to PDF overlay",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("input_pdf", type=Path, help="Input PDF from overlay_build.py")
    parser.add_argument("output_pdf", type=Path, help="Output layerized PDF")
    parser.add_argument("--old-color", type=parse_color, default="255,0,0", help="OLD layer color (R,G,B)")
    parser.add_argument("--new-color", type=parse_color, default="0,255,0", help="NEW layer color (R,G,B)")
    parser.add_argument("--old-alpha", type=float, default=0.4, help="OLD layer transparency (0.0-1.0)")
    parser.add_argument("--new-alpha", type=float, default=0.4, help="NEW layer transparency (0.0-1.0)")
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
        layerize_overlay(
            args.input_pdf,
            args.output_pdf,
            old_color=args.old_color,
            new_color=args.new_color,
            old_alpha=args.old_alpha,
            new_alpha=args.new_alpha,
        )
        LOGGER.info("Overlay layerization completed successfully")
        return 0
    except Exception as e:
        LOGGER.error(f"Overlay layerization failed: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main())
