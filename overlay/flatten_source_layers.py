#!/usr/bin/env python3
"""Optional preprocessing of source PDFs before overlay (see overlay_build / qc_overlay_prepend).

Original pipeline (default): no preprocessing — PyMuPDF show_pdf_page embeds vector artwork under
Old/New/Current OCGs (same as run_complete_overlay.py).

Optional --flatten-sources:
  • mode "vector" (default): MuPDF layer merge + pikepdf strip + clean_contents — keeps vectors.
  • mode "raster": full-page pixmap — use only if vector merge still shows nested CAD layers.

On Windows, pikepdf writes via a sidecar file then replaces (avoids in-place rename errors)."""

from __future__ import annotations

import logging
import os
from pathlib import Path

import pikepdf

try:
    import pymupdf as fitz  # type: ignore
except Exception:
    import fitz  # type: ignore

LOGGER = logging.getLogger("flatten_source_layers")


def _pikepdf_deep_strip(path: Path) -> None:
    """Remove OC catalog, /OC on objects, /Properties in Resources (vector-mode)."""
    sidecar = path.with_name(path.stem + "_strip_work.pdf")
    try:
        with pikepdf.open(path) as pdf:
            if pdf.Root.get("/OCProperties") is not None:
                del pdf.Root["/OCProperties"]

            visited: set[tuple[int, int]] = set()

            def walk(obj: object) -> None:
                if isinstance(obj, pikepdf.Object) and obj.is_indirect:
                    og = obj.objgen
                    if og in visited:
                        return
                    visited.add(og)
                if isinstance(obj, pikepdf.Dictionary):
                    if "/OC" in obj:
                        del obj["/OC"]
                    if "/Resources" in obj:
                        res = obj["/Resources"]
                        if isinstance(res, pikepdf.Dictionary) and "/Properties" in res:
                            del res["/Properties"]
                    for v in list(obj.values()):
                        walk(v)
                elif isinstance(obj, pikepdf.Array):
                    for v in obj:
                        walk(v)

            for page in pdf.pages:
                walk(page)

            pdf.save(sidecar)
    except Exception:
        if sidecar.exists():
            try:
                sidecar.unlink()
            except OSError:
                pass
        raise

    try:
        os.replace(str(sidecar), str(path))
    except OSError:
        if path.exists():
            path.unlink()
        os.replace(str(sidecar), str(path))


def _flatten_vector_merge(src: Path, dst: Path) -> None:
    """Vector path: MuPDF layer merge + pikepdf strip + clean_contents."""
    dst.parent.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(str(src))
    had_ocgs = False
    try:
        if not doc.is_pdf:
            doc.save(str(dst))
            return
        ocgs = doc.get_ocgs()
        if not ocgs:
            doc.save(str(dst), garbage=4, deflate=True)
            return

        had_ocgs = True
        xrefs = list(ocgs.keys())
        try:
            doc.set_layer(-1, on=xrefs)
        except Exception as exc:
            LOGGER.warning("set_layer (all ON) failed: %s", exc)

        layers = doc.get_layers()
        if layers:
            for layer in layers:
                if not doc.get_ocgs():
                    break
                try:
                    doc.switch_layer(layer["number"], as_default=True)
                except Exception as exc:
                    LOGGER.warning("switch_layer(%s) failed: %s", layer.get("number"), exc)
        doc.save(str(dst), garbage=4, deflate=True)
    finally:
        doc.close()

    recheck = fitz.open(str(dst))
    try:
        still_oc = recheck.is_pdf and bool(recheck.get_ocgs())
    finally:
        recheck.close()

    if had_ocgs or still_oc:
        if still_oc:
            LOGGER.info("Vector flatten: deep-stripping OCG metadata: %s", dst.name)
        _pikepdf_deep_strip(dst)

    doc2 = fitz.open(str(dst))
    try:
        for i in range(len(doc2)):
            try:
                doc2[i].clean_contents(sanitize=True)
            except Exception as exc:
                LOGGER.warning("clean_contents page %s: %s", i + 1, exc)
        doc2.save(str(dst), garbage=4, deflate=True)
    finally:
        doc2.close()


def _flatten_raster_combined(src: Path, dst: Path, dpi: float) -> None:
    """Raster path: one image per page (no vectors). Optional fallback for stubborn CAD files."""
    dst.parent.mkdir(parents=True, exist_ok=True)
    src_doc = fitz.open(str(src))
    try:
        if not src_doc.is_pdf:
            src_doc.save(str(dst))
            return

        if src_doc.get_ocgs():
            xrefs = list(src_doc.get_ocgs().keys())
            try:
                src_doc.set_layer(-1, on=xrefs)
            except Exception as exc:
                LOGGER.warning("set_layer (all ON) failed: %s", exc)
            layers = src_doc.get_layers()
            if layers:
                for layer in layers:
                    if not src_doc.get_ocgs():
                        break
                    try:
                        src_doc.switch_layer(layer["number"], as_default=True)
                    except Exception as exc:
                        LOGGER.warning("switch_layer(%s) failed: %s", layer.get("number"), exc)

        out_doc = fitz.open()
        try:
            mat = fitz.Matrix(dpi / 72.0, dpi / 72.0)
            for page_index in range(len(src_doc)):
                page = src_doc[page_index]
                pix = page.get_pixmap(matrix=mat, alpha=False)
                new_page = out_doc.new_page(width=page.rect.width, height=page.rect.height)
                new_page.insert_image(new_page.rect, pixmap=pix)
            out_doc.save(str(dst), garbage=4, deflate=True)
        finally:
            out_doc.close()
    finally:
        src_doc.close()

    LOGGER.info("Raster-flattened source (dpi=%s): %s", dpi, dst.name)


def flatten_authoring_layers(
    src: Path,
    dst: Path,
    *,
    dpi: float = 144.0,
    mode: str = "vector",
) -> None:
    """mode: 'vector' (default when flattening) or 'raster'."""
    if mode == "raster":
        _flatten_raster_combined(src, dst, dpi)
    else:
        _flatten_vector_merge(src, dst)
