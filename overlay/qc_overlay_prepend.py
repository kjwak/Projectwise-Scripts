#!/usr/bin/env python3
"""
QC Overlay Prepend: Compare incoming PDF vs page 1 of QC history, create overlay, prepend to history.

Flow:
  1. Determine Old input: use --current-master when provided (preferred, stable vector path);
     otherwise extract from page 1 of QC history.
  2. Create overlay: Old=extracted page 1 (red), New=incoming (green), Current=incoming (black)
  3. Prepend: PyMuPDF insert_pdf grafts history after the overlay page (better OC preservation
     than pikepdf append for multi-layer sheets)

Usage:
  python qc_overlay_prepend.py incoming.pdf qc_history.pdf [--output qc_history.pdf]
"""

import argparse
import hashlib
import logging
import os
import sys
import tempfile
from pathlib import Path
try:
    import pikepdf
except ImportError:
    print("Error: pikepdf is required. Install with: pip install pikepdf")
    sys.exit(1)

# Add script dir for imports (works when run as script or frozen exe)
if getattr(sys, "frozen", False):
    SCRIPT_DIR = Path(sys._MEIPASS)
else:
    SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

LOGGER = logging.getLogger("qc_overlay_prepend")


def _file_sha256(path: Path) -> str:
    """SHA-256 of file contents (for duplicate Old/New detection)."""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _get_fitz():
    try:
        import pymupdf as fitz  # type: ignore
    except Exception:
        import fitz  # type: ignore
    return fitz


def _qc_triplet_xrefs(ocgs: dict) -> dict[str, int]:
    """Map Old / New / Current -> xref. Match names case-insensitively (PDFs vary)."""
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


def _page0_has_qc_overlay_forms(pdf_path: Path) -> bool:
    """True if page 1 has our overlay/layerize XObjects (QC lives on-page, not in Civil OCG list).

    Civil exports register hundreds of global OCG names (DGN layers); Old/New/Current from
    PyMuPDF appear as forms BBL/BBL1/BBL_Current (or fzFrm*) on page 1, not as top-level
    names in get_ocgs() alongside those layers.
    """
    try:
        with pikepdf.open(pdf_path) as pdf:
            if len(pdf.pages) < 1:
                return False
            res = pdf.pages[0].get("/Resources")
            if not res:
                return False
            xo = res.get("/XObject")
            if not xo:
                return False
            for key in xo.keys():
                nk = str(key).strip("'").lstrip("/")
                if nk in ("BBL", "BBL1", "BBL_Current", "fzFrm0", "fzFrm1", "fzFrm2"):
                    return True
                if "BBL" in nk or nk.startswith("fzFrm"):
                    return True
            return False
    except Exception:
        return False


def _run_overlay_pipeline(
    old_pdf: Path,
    new_pdf: Path,
    out_pdf: Path,
    *,
    fit: bool = False,
    alpha: float = 0.6,
    verbose: bool = False,
    flatten_sources: bool = False,
    flatten_dpi: float = 144.0,
    flatten_raster: bool = False,
) -> bool:
    """Build overlay (Old red, New green, Current black) using direct imports. No subprocess."""
    from overlay_build import build_overlay
    from overlay_layerize import layerize_overlay

    if verbose:
        logging.getLogger("overlay_build").setLevel(logging.INFO)
        logging.getLogger("overlay_layerize").setLevel(logging.INFO)
    overlay_stage = out_pdf.parent / (out_pdf.stem + "__stage.pdf")
    build_overlay(
        old_pdf,
        new_pdf,
        overlay_stage,
        pages_spec=None,
        fit=fit,
        canvas="new",
        add_configs=True,
        flatten_sources=flatten_sources,
        flatten_dpi=flatten_dpi,
        flatten_raster=flatten_raster,
    )
    old_alpha = max(0.0, min(1.0, alpha))
    new_alpha = max(0.0, min(1.0, alpha * 0.8))
    layerize_overlay(
        overlay_stage,
        out_pdf,
        old_color=(1.0, 0.0, 0.0),
        new_color=(0.0, 0.75, 0.0),
        old_alpha=old_alpha,
        new_alpha=new_alpha,
    )
    try:
        overlay_stage.unlink()
    except OSError:
        pass
    return out_pdf.exists()


def _find_current_layer_xobject(page) -> tuple:
    """Return (name, xobj) for the Current-compare layer on a QC sheet, or (None, None).

    After overlay_layerize.py, Current is the form /BBL_Current (was /fzFrm2); /OC may be missing
    on the flattened stream. pikepdf keys may be Name objects — check by string, not `in` dict.
    """
    resources = page.get("/Resources")
    if not resources:
        return None, None
    xobjects = resources.get("/XObject")
    if not xobjects:
        return None, None

    # Layerized QC page (round 2+): canonical names from overlay_layerize.py
    for key, xobj in xobjects.items():
        nk = str(key).strip("'").lstrip("/")
        if nk in ("BBL_Current", "fzFrm2"):
            return key, xobj

    # PyMuPDF overlay: forms tagged with optional content (OCG name contains Current)
    for name, xobj in xobjects.items():
        oc = xobj.get("/OC")
        if oc is not None:
            ocg_name = str(oc.get("/Name", "")) if hasattr(oc, "get") else ""
            if "Current" in ocg_name:
                return name, xobj

    return None, None


def _pdf_name_token_for_content_stream(key) -> bytes:
    """PDF content stream name token for an XObject key (e.g. /fzFrm2)."""
    s = str(key).strip("'")
    if not s.startswith("/"):
        s = "/" + s.lstrip("/")
    return s.encode("latin-1")


def _strip_oc_from_xobject_tree(xobj, max_depth: int = 12) -> None:
    """Remove /OC so optional content state cannot hide art when re-embedding as Old."""
    if max_depth <= 0 or xobj is None:
        return
    try:
        if hasattr(xobj, "get") and xobj.get("/OC") is not None:
            del xobj["/OC"]
    except Exception:
        return
    try:
        res = xobj.get("/Resources") if hasattr(xobj, "get") else None
        if not res:
            return
        xo = res.get("/XObject")
        if not xo:
            return
        for _k, child in xo.items():
            _strip_oc_from_xobject_tree(child, max_depth - 1)
    except Exception:
        return


def _extract_page_1_current_form_only_full_page_impl(pdf: pikepdf.Pdf, out_path: Path) -> bool:
    """Copy page 1 with full resource graph, then Contents = draw Current form only.

    Copying the Current XObject alone drops nested /fullpage and other dependencies — empty Old.
    """
    if len(pdf.pages) < 1:
        return False
    page0 = pdf.pages[0]
    key, cur = _find_current_layer_xobject(page0)
    if key is None or cur is None:
        return False
    try:
        dst = pikepdf.Pdf.new()
        dst.pages.append(dst.copy_foreign(page0))
        page = dst.pages[0]
        key2, cur2 = _find_current_layer_xobject(page)
        if key2 is None:
            LOGGER.warning("Current form not found on page after copy_foreign; cannot build Old extract")
            return False
        _strip_oc_from_xobject_tree(cur2)
        name = _pdf_name_token_for_content_stream(key2)
        content = b"q 1 0 0 1 0 0 cm " + name + b" Do Q"
        page.obj["/Contents"] = pikepdf.Stream(dst, content)
        dst.save(out_path)
        LOGGER.info("Extracted page 1 Current via full page copy + Contents trim -> %s", out_path)
        return True
    except Exception as exc:
        LOGGER.warning("Full-page Current-form extract failed: %s", exc)
        return False


def _extract_page_1_current_form_only_full_page(pdf_path: Path, out_path: Path) -> bool:
    try:
        with pikepdf.open(pdf_path) as pdf:
            return _extract_page_1_current_form_only_full_page_impl(pdf, out_path)
    except Exception as exc:
        LOGGER.warning("Could not open PDF for Current-form extract: %s", exc)
        return False


def _extract_page_1_pikepdf_current_form(pdf_path: Path, out_path: Path) -> bool:
    """Fallback when fitz fails: full page + Current-only contents, else full page 1."""
    with pikepdf.open(pdf_path) as pdf:
        if len(pdf.pages) < 1:
            LOGGER.error("QC history has no pages")
            return False
        page = pdf.pages[0]
        _name, current_xobj = _find_current_layer_xobject(page)

        if _name is None or current_xobj is None:
            pdf_out = pikepdf.Pdf.new()
            pdf_out.pages.append(pdf_out.copy_foreign(page))
            pdf_out.save(out_path)
            LOGGER.info("Extracted page 1 (full, pikepdf fallback) to %s", out_path)
            return True

        return _extract_page_1_current_form_only_full_page_impl(pdf, out_path)


def extract_page_1_current_as_old(pdf_path: Path, out_path: Path) -> bool:
    """Extract page 1 for the Old overlay input: visible 'Current' QC art (previous approved).

    When page 1 has QC overlay forms (BBL_*/fzFrm*), prefer copying the full page graph with
    pikepdf then trimming Contents to only the Current form — copying the form XObject alone
    drops nested /fullpage resources and yields an empty Old layer.

    Otherwise use PyMuPDF (global Old/New/Current OCG triplet + select page 0), then pikepdf fallbacks.
    """
    if _extract_page_1_current_form_only_full_page(pdf_path, out_path):
        return True

    fitz = _get_fitz()
    doc = None
    try:
        doc = fitz.open(str(pdf_path))
        if doc.page_count < 1:
            LOGGER.error("QC history has no pages")
            return False

        ocgs = doc.get_ocgs()
        if ocgs:
            qc = _qc_triplet_xrefs(ocgs)
            if LOGGER.isEnabledFor(logging.DEBUG):
                LOGGER.debug(
                    "Page-1 extract OCG names: %s",
                    {xref: info.get("name") for xref, info in ocgs.items()},
                )
            if qc.get("Old") and qc.get("New") and qc.get("Current"):
                all_x = list(ocgs.keys())
                off_list = [qc["Old"], qc["New"]]
                on_list = [x for x in all_x if x not in off_list]
                try:
                    doc.set_layer(-1, on=on_list, off=off_list)
                except Exception as exc:
                    LOGGER.warning("set_layer (show Current, hide Old/New) failed: %s", exc)
            elif _page0_has_qc_overlay_forms(pdf_path):
                # Civil + QC: global OCG list is DGN layers; QC compare uses page forms BBL_* / fzFrm*
                LOGGER.info(
                    "Page 1 has QC overlay forms (BBL/fzFrm); no global Old/New/Current names. "
                    "Using PDF default layer config (/D from add_layer_configurations) + select([0])."
                )
            else:
                n = len(ocgs)
                sample = [str(info.get("name", "")) for info in list(ocgs.values())[:15]]
                LOGGER.warning(
                    "No global Old/New/Current OCG triplet (%d OCGs in catalog). "
                    "Sample names: %s%s — using default layer config as-is.",
                    n,
                    sample,
                    " ..." if n > 15 else "",
                )

        # Keep full document OC catalog; insert_pdf into an empty doc often drops /OCProperties
        # and breaks form visibility (empty Old). select() uses pdf_rearrange_pages2(..., KEEP).
        if doc.page_count > 1:
            doc.select([0])
        doc.save(str(out_path))

        LOGGER.info("Extracted page 1 for Old input (fitz, select page 0, Current visible)")
        return True
    except Exception as exc:
        LOGGER.warning("fitz extract failed (%s); trying pikepdf Current form", exc)
        try:
            return _extract_page_1_pikepdf_current_form(pdf_path, out_path)
        except Exception as exc2:
            LOGGER.error("Failed to extract page 1: %s", exc2)
            return False
    finally:
        if doc is not None and not doc.is_closed:
            doc.close()


def prepend_overlay_to_history(overlay_path: Path, history_path: Path, output_path: Path) -> bool:
    """Prepend overlay page and append history pages verbatim with pikepdf.

    After adopting --current-master, we no longer need fitz-based history grafting to stabilize
    page-1 Old extraction. Using pikepdf append avoids MuPDF rewrites that can flatten optional
    content behavior on appended pages (page 2+ in subsequent runs).
    """
    def _find_current_ocg_in_output(pdf: pikepdf.Pdf):
        ocprops = pdf.Root.get("/OCProperties")
        if not ocprops:
            return None
        ocgs = ocprops.get("/OCGs")
        if not ocgs:
            return None
        for ocg in ocgs:
            name = str(ocg.get("/Name", ""))
            if name == "Current" or "Current" in name:
                return ocg
        return None

    def _find_qc_ocgs_in_output(pdf: pikepdf.Pdf) -> dict[str, object]:
        """Find canonical Old/New/Current OCG objects in output catalog."""
        out: dict[str, object] = {}
        ocprops = pdf.Root.get("/OCProperties")
        if not ocprops:
            return out
        ocgs = ocprops.get("/OCGs")
        if not ocgs:
            return out
        for ocg in ocgs:
            name = str(ocg.get("/Name", ""))
            if "Old" in name and "Old" not in out:
                out["Old"] = ocg
            elif "New" in name and "New" not in out:
                out["New"] = ocg
            elif "Current" in name and "Current" not in out:
                out["Current"] = ocg
        return out

    def _page_has_oc_content(page: pikepdf.Page) -> bool:
        res = page.get("/Resources")
        if not res:
            return False
        xo = res.get("/XObject")
        if not xo:
            return False
        for _, xobj in xo.items():
            try:
                if xobj.get("/OC") is not None:
                    return True
            except Exception:
                continue
        return False

    def _wrap_page_contents_in_current(page: pikepdf.Page, current_ocg, pdf: pikepdf.Pdf) -> None:
        """Wrap existing page content in /OC ... BDC ... EMC so viewer layer toggles still apply."""
        if "/Resources" not in page.obj:
            page.obj["/Resources"] = pikepdf.Dictionary()
        resources = page.obj["/Resources"]
        if "/Properties" not in resources:
            resources["/Properties"] = pikepdf.Dictionary()
        props = resources["/Properties"]
        props["/BBL_Current"] = current_ocg

        contents = page.obj.get("/Contents")
        if contents is None:
            page.obj["/Contents"] = pikepdf.Stream(pdf, b"q /OC /BBL_Current BDC EMC Q")
            return
        if isinstance(contents, pikepdf.Stream):
            original = bytes(contents.read_bytes())
        elif isinstance(contents, pikepdf.Array):
            parts = []
            for item in contents:
                if isinstance(item, pikepdf.Stream):
                    parts.append(bytes(item.read_bytes()))
            original = b"\n".join(parts)
        else:
            return
        wrapped = b"q\n/OC /BBL_Current BDC\n" + original + b"\nEMC\nQ\n"
        page.obj["/Contents"] = pikepdf.Stream(pdf, wrapped)

    def _normalize_page_qc_ocgs(page: pikepdf.Page, canonical: dict[str, object]) -> None:
        """Rebind page XObject /OC refs to the canonical triplet from page 1."""
        if not canonical:
            return
        res = page.get("/Resources")
        if not res:
            return
        xo = res.get("/XObject")
        if not xo:
            return
        for _, xobj in xo.items():
            try:
                oc = xobj.get("/OC")
                if oc is None:
                    continue
                name = str(oc.get("/Name", "")) if hasattr(oc, "get") else ""
                if "Old" in name and "Old" in canonical:
                    xobj["/OC"] = canonical["Old"]
                elif "New" in name and "New" in canonical:
                    xobj["/OC"] = canonical["New"]
                elif "Current" in name and "Current" in canonical:
                    xobj["/OC"] = canonical["Current"]
            except Exception:
                continue

    try:
        with pikepdf.open(overlay_path, allow_overwriting_input=True) as overlay_doc:
            with pikepdf.open(history_path) as history_doc:
                output_doc = overlay_doc
                current_ocg = _find_current_ocg_in_output(output_doc)
                canonical_qc = _find_qc_ocgs_in_output(output_doc)
                for page in history_doc.pages:
                    output_doc.pages.append(page)
                    appended = output_doc.pages[-1]
                    _normalize_page_qc_ocgs(appended, canonical_qc)
                    if current_ocg is not None:
                        if not _page_has_oc_content(appended):
                            _wrap_page_contents_in_current(appended, current_ocg, output_doc)

                if output_path.resolve() == history_path.resolve():
                    fd, tmp = tempfile.mkstemp(suffix=".pdf", dir=str(output_path.parent))
                    os.close(fd)
                    tmp_path = Path(tmp)
                    try:
                        output_doc.save(tmp_path)
                        tmp_path.replace(output_path)
                    except Exception:
                        if tmp_path.exists():
                            tmp_path.unlink()
                        raise
                else:
                    output_doc.save(str(output_path))

        LOGGER.info("Prepended overlay to history (pikepdf append)")
        return True
    except Exception as e:
        LOGGER.error("Failed to prepend: %s", e)
        return False


def write_current_master_from_incoming(incoming_path: Path, current_master_path: Path) -> bool:
    """Update current-master with page 1 of incoming (vector-preserving)."""
    fitz = _get_fitz()
    src = fitz.open(str(incoming_path))
    try:
        if src.page_count < 1:
            LOGGER.error("Incoming PDF has no pages")
            return False
        out = fitz.open()
        try:
            out.insert_pdf(src, from_page=0, to_page=0)
            current_master_path.parent.mkdir(parents=True, exist_ok=True)
            out.save(str(current_master_path))
        finally:
            out.close()
        LOGGER.info("Updated current master: %s", current_master_path)
        return True
    except Exception as exc:
        LOGGER.error("Failed to update current master: %s", exc)
        return False
    finally:
        src.close()


def main():
    parser = argparse.ArgumentParser(
        description="QC overlay prepend: compare incoming vs page 1 of history, create overlay, prepend",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
        epilog="""
Examples:
  # Update QC history in place (default):
  python qc_overlay_prepend.py incoming.pdf sheet-qc.pdf

  # Write to different output:
  python qc_overlay_prepend.py incoming.pdf sheet-qc.pdf --output result.pdf

  # First run (no history yet): creates qc history from incoming only
        """
    )
    parser.add_argument("incoming_pdf", type=Path, help="New incoming PDF (becomes New + Current layers)")
    parser.add_argument("qc_history_pdf", type=Path, help="QC history PDF (page 1 becomes Old layer)")
    parser.add_argument("--output", "-o", type=Path, default=None,
                        help="Output path (default: overwrite qc_history_pdf)")
    parser.add_argument("--current-master", type=Path, default=None,
                        help="Preferred OLD input source (single-page current sheet). "
                             "If set and exists, run uses this instead of extracting from history page 1.")
    parser.add_argument("--alpha", type=float, default=0.6, help="Layer transparency")
    parser.add_argument("--fit", action="store_true", help="Fit pages to match size")
    parser.add_argument("--keep-temp", action="store_true", help="Keep temporary files")
    parser.add_argument("--flatten-sources", action="store_true",
                        help="Preprocess inputs to merge/strip authoring layers (vector merge by default)")
    parser.add_argument("--flatten-raster", action="store_true",
                        help="With --flatten-sources, rasterize each page (CAD fallback; loses vectors)")
    parser.add_argument("--flatten-dpi", type=float, default=144.0,
                        help="Resolution for --flatten-raster (higher = sharper, larger)")
    parser.add_argument("--verbose", "-v", action="store_true", help="Verbose output")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.INFO if args.verbose else logging.WARNING,
        format="%(levelname)s: %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)]
    )

    incoming = args.incoming_pdf.resolve()
    history = args.qc_history_pdf.resolve()
    output = (args.output or history).resolve()
    current_master = args.current_master.resolve() if args.current_master else None

    if not incoming.exists():
        LOGGER.error(f"Incoming PDF not found: {incoming}")
        sys.exit(1)

    if not history.exists():
        LOGGER.info("QC history does not exist - creating from incoming only (no overlay)")
        with pikepdf.open(incoming) as doc:
            doc.save(output)
        if current_master is not None:
            if not write_current_master_from_incoming(incoming, current_master):
                sys.exit(1)
        LOGGER.info(f"Created {output}")
        sys.exit(0)

    tmp_base = output.parent / ".qc_overlay_tmp"
    tmp_base.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="run_", dir=str(tmp_base)) as tmpdir:
        tmp = Path(tmpdir)
        page1_pdf = tmp / "page1.pdf"
        overlay_pdf = tmp / "overlay.pdf"

        used_current_master_for_old = False
        if current_master is not None and current_master.exists():
            old_source_pdf = current_master
            used_current_master_for_old = True
            LOGGER.info("Using current master as OLD source: %s", old_source_pdf)
        else:
            if current_master is not None and not current_master.exists():
                LOGGER.warning("Current master not found, falling back to history page 1 extraction: %s", current_master)
            if not extract_page_1_current_as_old(history, page1_pdf):
                sys.exit(1)
            old_source_pdf = page1_pdf

        # If current-master is the same file bytes as incoming, build_overlay would embed identical art for Old and New
        # (only colors differ). Typical causes: same revision resubmitted, or a seed file that matched incoming.
        if used_current_master_for_old and current_master is not None:
            try:
                if _file_sha256(current_master) == _file_sha256(incoming):
                    LOGGER.warning(
                        "current-master is byte-identical to incoming; using history page-1 extract for Old instead "
                        "(avoids duplicate Old/New content). If this is unexpected, fix or delete the current-master file."
                    )
                    if not extract_page_1_current_as_old(history, page1_pdf):
                        sys.exit(1)
                    old_source_pdf = page1_pdf
            except OSError as exc:
                LOGGER.warning("Could not compare current-master to incoming: %s", exc)

        LOGGER.info("[*] Building overlay (Old=page1 Current layer red, New=incoming green, Current=incoming black)")
        if not _run_overlay_pipeline(
            old_source_pdf,
            incoming,
            overlay_pdf,
            fit=args.fit,
            alpha=args.alpha,
            verbose=args.verbose,
            flatten_sources=args.flatten_sources,
            flatten_dpi=args.flatten_dpi,
            flatten_raster=args.flatten_raster,
        ):
            LOGGER.error("[FAIL] Overlay build failed")
            sys.exit(1)
        LOGGER.info("[OK] Overlay built successfully")

        if not overlay_pdf.exists():
            LOGGER.error("Overlay was not created")
            sys.exit(1)

        if not prepend_overlay_to_history(overlay_pdf, history, output):
            sys.exit(1)

        if current_master is not None:
            if not write_current_master_from_incoming(incoming, current_master):
                sys.exit(1)

        if args.keep_temp:
            import shutil
            keep_dir = output.parent / (output.stem + "_overlay_temp")
            keep_dir.mkdir(exist_ok=True)
            if page1_pdf.exists():
                shutil.copy(page1_pdf, keep_dir / "page1.pdf")
            shutil.copy(overlay_pdf, keep_dir / "overlay.pdf")
            LOGGER.info(f"Temp files kept in {keep_dir}")

    LOGGER.info(f"Done. QC history updated: {output}")
    sys.exit(0)


if __name__ == "__main__":
    main()
