#!/usr/bin/env python3
"""
Copy a stamp template PDF onto a target PDF page.

Default behavior matches pasting the stamp in Bluebeam: every annotation on the
stamp page (linework, FreeText, widgets) is copied onto the target with only
position/scaling adjusted. Page size is not changed.

Optional --flatten-stamp-annotation rasterizes into one Stamp annotation (legacy).
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import shutil
import sys
import tempfile
from datetime import date
from pathlib import Path
from typing import Any

LOGGER = logging.getLogger("qc_review_stamp")

DEFAULT_STAMP_HEIGHT_PT = 200.0
DEFAULT_MARGIN_INSET_PT = 12.0
STAMP_APPEARANCE_SUPERSAMPLE = 2.0


def _get_fitz():
    try:
        import pymupdf as fitz  # type: ignore
    except Exception:
        import fitz  # type: ignore
    return fitz


def _parse_field_values(raw: dict[str, Any] | None) -> dict[str, str]:
    if not raw:
        return {}
    out: dict[str, str] = {}
    for k, v in raw.items():
        if v is None:
            continue
        s = str(v).strip()
        if s:
            out[str(k)] = s
    return out


def compute_stamp_rect_overlay_top_left(
    page_rect: Any,
    stamp_width: float,
    stamp_height: float,
    margin_inset: float,
) -> Any:
    return compute_stamp_rect_at_offset(
        page_rect, stamp_width, stamp_height, float(margin_inset), float(margin_inset)
    )


def compute_stamp_rect_at_offset(
    page_rect: Any,
    stamp_width: float,
    stamp_height: float,
    x_pt: float,
    y_pt: float,
) -> Any:
    """Stamp bounds in page view space; offsets are from the visible top-left."""
    fitz = _get_fitz()
    px0, py0, _px1, _py1 = (
        float(page_rect.x0),
        float(page_rect.y0),
        float(page_rect.x1),
        float(page_rect.y1),
    )
    x_off = float(x_pt)
    y_off = float(y_pt)
    sw = float(stamp_width)
    sh = float(stamp_height)
    return fitz.Rect(
        px0 + x_off,
        py0 + y_off,
        px0 + x_off + sw,
        py0 + y_off + sh,
    )


def _fill_stamp_template_widgets(stamp_page: Any, values: dict[str, str]) -> None:
    for w in stamp_page.widgets() or []:
        name = (w.field_name or "").strip()
        if not name:
            continue
        if name in values:
            w.field_value = values[name]
        try:
            w.update()
        except Exception:
            LOGGER.debug("widget update failed for %s", name, exc_info=True)


def _get_stamp_source_bounds(stamp_path: Path) -> tuple[float, float, float, float]:
    """PDF user-space bounds containing stamp markup (CropBox when present)."""
    import pikepdf

    with pikepdf.open(stamp_path) as doc:
        page = doc.pages[0]
        box = page.CropBox if "/CropBox" in page else page.MediaBox
        return tuple(float(v) for v in box)


def _stamp_dimensions_from_view_rect(stamp_view_rect: Any, stamp_height_pt: float) -> tuple[float, float]:
    sh = float(stamp_view_rect.height) or 1.0
    scale = float(stamp_height_pt) / sh
    return float(stamp_view_rect.width) * scale, float(stamp_height_pt)


def _compute_target_stamp_rect_view(
    page: Any,
    stamp_width: float,
    stamp_height: float,
    *,
    margin_inset_pt: float,
    position_x_pt: float | None,
    position_y_pt: float | None,
) -> Any:
    """Target stamp bounds in page view space (what Bluebeam shows on screen)."""
    if position_x_pt is not None and position_y_pt is not None:
        return compute_stamp_rect_at_offset(
            page.rect, stamp_width, stamp_height, float(position_x_pt), float(position_y_pt)
        )
    return compute_stamp_rect_overlay_top_left(
        page.rect, stamp_width, stamp_height, float(margin_inset_pt)
    )


def _view_rect_from_pdf_rect(transformation_matrix: Any, pdf_rect: Any) -> Any:
    fitz = _get_fitz()
    x0, y0, x1, y1 = [float(v) for v in pdf_rect]
    p1 = fitz.Point(x0, y0) * transformation_matrix
    p2 = fitz.Point(x1, y1) * transformation_matrix
    return fitz.Rect(
        min(p1.x, p2.x),
        min(p1.y, p2.y),
        max(p1.x, p2.x),
        max(p1.y, p2.y),
    )


def _pdf_rect_from_view_rect(transformation_matrix: Any, view_rect: Any) -> list[float]:
    fitz = _get_fitz()
    r = fitz.Rect(view_rect)
    inv = ~transformation_matrix
    pts = [
        fitz.Point(r.x0, r.y0) * inv,
        fitz.Point(r.x1, r.y0) * inv,
        fitz.Point(r.x0, r.y1) * inv,
        fitz.Point(r.x1, r.y1) * inv,
    ]
    xs = [float(p.x) for p in pts]
    ys = [float(p.y) for p in pts]
    return [min(xs), min(ys), max(xs), max(ys)]


def _view_rect_from_pdf_rect_page(page: Any, pdf_rect: Any) -> Any:
    return _view_rect_from_pdf_rect(page.transformation_matrix, pdf_rect)


def _pdf_rect_from_view_rect_page(page: Any, view_rect: Any) -> list[float]:
    return _pdf_rect_from_view_rect(page.transformation_matrix, view_rect)


def _map_pdf_rect_in_bounds(
    src_rect: list[float],
    src_bounds: list[float],
    dst_bounds: list[float],
) -> list[float]:
    x0, y0, x1, y1 = (float(v) for v in src_rect)
    sx0, sy0, sx1, sy1 = (float(v) for v in src_bounds)
    dx0, dy0, dx1, dy1 = (float(v) for v in dst_bounds)
    sw = float(sx1 - sx0) or 1.0
    sh = float(sy1 - sy0) or 1.0
    dw = float(dx1 - dx0)
    dh = float(dy1 - dy0)
    return [
        dx0 + ((x0 - sx0) / sw) * dw,
        dy0 + ((y0 - sy0) / sh) * dh,
        dx0 + ((x1 - sx0) / sw) * dw,
        dy0 + ((y1 - sy0) / sh) * dh,
    ]


def _map_view_rect_in_bounds(src_rect: Any, src_bounds: Any, dst_bounds: Any) -> Any:
    fitz = _get_fitz()
    src = fitz.Rect(src_rect)
    sb = fitz.Rect(src_bounds)
    db = fitz.Rect(dst_bounds)
    if sb.width <= 0 or sb.height <= 0:
        raise ValueError("Invalid stamp source bounds")
    return fitz.Rect(
        db.x0 + (src.x0 - sb.x0) / sb.width * db.width,
        db.y0 + (src.y0 - sb.y0) / sb.height * db.height,
        db.x0 + (src.x1 - sb.x0) / sb.width * db.width,
        db.y0 + (src.y1 - sb.y0) / sb.height * db.height,
    )


def _pdf_rect_from_view_rect_rotation_matrix(page: Any, view_rect: Any) -> list[float]:
    """Map a view/display rect to PDF user-space /Rect (Bluebeam paste model)."""
    fitz = _get_fitz()
    r = fitz.Rect(view_rect)
    rm = page.rotation_matrix
    pts = [
        fitz.Point(r.x0, r.y0) * rm,
        fitz.Point(r.x1, r.y0) * rm,
        fitz.Point(r.x0, r.y1) * rm,
        fitz.Point(r.x1, r.y1) * rm,
    ]
    xs = [float(p.x) for p in pts]
    ys = [float(p.y) for p in pts]
    return [min(xs), min(ys), max(xs), max(ys)]


def _stamp_source_block_bounds_pikepdf(stamp_path: Path) -> list[float]:
    """Union /Rect of stamp markup in stamp-page PDF user space."""
    import pikepdf

    stamp_subtypes = {"/Square", "/Stamp", "/FreeText"}
    rects: list[list[float]] = []
    with pikepdf.open(stamp_path) as doc:
        page = doc.pages[0]
        for ref in page.Annots or []:
            if str(ref.get("/Subtype", "")) not in stamp_subtypes:
                continue
            rects.append([float(v) for v in ref["/Rect"]])
    if not rects:
        raise ValueError(f"No stamp markup bounds in {stamp_path}")
    xs0 = [r[0] for r in rects]
    ys0 = [r[1] for r in rects]
    xs1 = [r[2] for r in rects]
    ys1 = [r[3] for r in rects]
    return [min(xs0), min(ys0), max(xs1), max(ys1)]


def _norm_point_in_rotated_block(u: float, v: float, page_rotation: int) -> tuple[float, float]:
    """Map normalized stamp-block coords for a target page /Rotate."""
    rot = int(page_rotation or 0) % 360
    if rot == 0:
        return u, v
    if rot == 90:
        return 1.0 - v, u
    if rot == 180:
        return 1.0 - u, 1.0 - v
    if rot == 270:
        return v, 1.0 - u
    raise ValueError(f"Unsupported page rotation: {page_rotation}")


def _pdf_rect_from_stamp_block_map(
    src_rect: list[float],
    src_bounds: list[float],
    dst_bounds: list[float],
    page_rotation: int,
) -> list[float]:
    """Map a stamp-template /Rect into target PDF space via block rotation + scale."""
    sx0, sy0, sx1, sy1 = (float(v) for v in src_bounds)
    tx0, ty0, tx1, ty1 = (float(v) for v in dst_bounds)
    sw = float(sx1 - sx0) or 1.0
    sh = float(sy1 - sy0) or 1.0
    tw = float(tx1 - tx0)
    th = float(ty1 - ty0)
    x0, y0, x1, y1 = (float(v) for v in src_rect)
    mapped: list[tuple[float, float]] = []
    for x, y in ((x0, y0), (x1, y0), (x0, y1), (x1, y1)):
        u = (x - sx0) / sw
        v = (y - sy0) / sh
        tu, tv = _norm_point_in_rotated_block(u, v, page_rotation)
        mapped.append((tx0 + tu * tw, ty0 + tv * th))
    xs = [p[0] for p in mapped]
    ys = [p[1] for p in mapped]
    return [min(xs), min(ys), max(xs), max(ys)]


def _stamp_ap_matrix_for_rotation(appearance_rot_deg: int, bbox: list[float]) -> list[float]:
    """Appearance /Matrix for Stamp /AP /N, matching Bluebeam counter-rotation."""
    x0, y0, x1, y1 = (float(v) for v in bbox)
    rot = int(appearance_rot_deg) % 360
    if rot == 0:
        return [1.0, 0.0, 0.0, 1.0, -x0, -y0]
    if rot == 90:
        return [0.0, -1.0, 1.0, 0.0, -y0, x1]
    if rot == 180:
        return [-1.0, 0.0, 0.0, -1.0, x1, y1]
    if rot == 270:
        return [0.0, 1.0, -1.0, 0.0, y1, -x0]
    raise ValueError(f"Unsupported stamp appearance rotation: {appearance_rot_deg}")


def _normalize_rect_ap_stream(ap_n: Any, pdf_rect: list[float]) -> None:
    """Rewrite Square/FreeText /AP /N /BBox and /Matrix to match /Rect."""
    x0, y0, x1, y1 = (float(v) for v in pdf_rect)
    ap_n["/BBox"] = [x0, y0, x1, y1]
    ap_n["/Matrix"] = [1.0, 0.0, 0.0, 1.0, -x0, -y0]


def _update_copied_annotation_ap(
    copied: Any,
    *,
    subtype: str,
    pdf_rect: list[float],
    page_rotation: int,
) -> None:
    """Adjust copied /AP metadata per Bluebeam subtype behavior."""
    if "/AP" not in copied or copied["/AP"].get("/N") is None:
        return
    ap_n = copied["/AP"]["/N"]
    if subtype in ("/Square", "/FreeText"):
        _normalize_rect_ap_stream(ap_n, pdf_rect)
        return
    if subtype == "/Stamp":
        if "/BBox" not in ap_n:
            return
        bbox = [float(v) for v in ap_n["/BBox"]]
        appearance_rot = (360 - int(page_rotation or 0)) % 360
        ap_n["/Matrix"] = _stamp_ap_matrix_for_rotation(appearance_rot, bbox)


def _int_color_to_rgb(color: int) -> tuple[float, float, float]:
    value = int(color or 0)
    return (
        ((value >> 16) & 255) / 255.0,
        ((value >> 8) & 255) / 255.0,
        (value & 255) / 255.0,
    )


def _freetext_style_from_annot(annot: Any) -> tuple[str, int, float]:
    text = (annot.get_text() or "").strip()
    td = annot.get_text("dict")
    color = 0
    size = 8.0
    for block in td.get("blocks") or []:
        for line in block.get("lines") or []:
            for span in line.get("spans") or []:
                span_color = int(span.get("color") or 0)
                if span_color not in (0, 16777215):
                    color = span_color
                size = float(span.get("size") or size)
    return text, color, size


def _render_stamp_annot_pixmap(stamp_page: Any, stamp_annot: Any, target_rect: Any) -> Any:
    fitz = _get_fitz()
    src = fitz.Rect(stamp_annot.rect)
    if src.width <= 0 or src.height <= 0:
        raise ValueError("Invalid stamp annotation bounds")
    scale = max(
        float(target_rect.width) / float(src.width),
        float(target_rect.height) / float(src.height),
    )
    matrix = fitz.Matrix(scale * STAMP_APPEARANCE_SUPERSAMPLE, scale * STAMP_APPEARANCE_SUPERSAMPLE)
    return stamp_page.get_pixmap(clip=src, matrix=matrix, alpha=True)


def _stamp_markup_count(stamp_path: Path) -> int:
    import pikepdf

    with pikepdf.open(stamp_path) as doc:
        page = doc.pages[0]
        if "/Annots" not in page:
            return 0
        return len(page.Annots)


def _copy_stamp_widgets_to_page(
    stamp_path: Path,
    target_path: Path,
    *,
    page_index: int,
    stamp_view_bounds: list[float],
    target_view_bounds: list[float],
    values: dict[str, str],
    populate_text_fields: bool,
) -> int:
    """Copy AcroForm widget annotations and reposition in view space."""
    import pikepdf
    from pikepdf import Array

    fitz = _get_fitz()
    widget_rects: dict[str, Any] = {}
    with fitz.open(stamp_path) as sdoc:
        spage = sdoc[0]
        sb = fitz.Rect(stamp_view_bounds)
        db = fitz.Rect(target_view_bounds)
        for widget in spage.widgets() or []:
            name = (widget.field_name or "").strip()
            if not name:
                continue
            widget_rects[name] = _map_view_rect_in_bounds(widget.rect, sb, db)

    if not widget_rects:
        return 0

    added = 0
    with pikepdf.open(stamp_path) as src, pikepdf.open(target_path, allow_overwriting_input=True) as dst:
        dst_page = dst.pages[page_index]
        if "/Annots" not in dst_page:
            dst_page.Annots = Array()
        for ref in src.pages[0].Annots or []:
            if str(ref.get("/Subtype", "")) != "/Widget":
                continue
            copied = dst.copy_foreign(ref)
            name = str(copied.get("/T", "")).strip()
            if populate_text_fields and name and name in values:
                copied["/V"] = values[name]
            if "/P" in copied:
                del copied["/P"]
            dst_page.Annots.append(dst.make_indirect(copied))
            added += 1
        dst.save(target_path)

    if not added:
        return 0

    with fitz.open(target_path) as tdoc:
        tpage = tdoc[page_index]
        for widget in tpage.widgets() or []:
            name = (widget.field_name or "").strip()
            if name not in widget_rects:
                continue
            widget.field_rect = widget_rects[name]
            if populate_text_fields and name in values:
                widget.field_value = values[name]
            widget.update()
        tdoc.saveIncr()

    return added


def _place_stamp_markup_pikepdf(
    stamp_path: Path,
    target_path: Path,
    *,
    page_index: int,
    stamp_view_bounds: list[float],
    target_view_bounds: list[float],
    values: dict[str, str],
    populate_text_fields: bool,
) -> int:
    """Copy stamp annotations with Bluebeam paste semantics (pikepdf + block map)."""
    import pikepdf
    from pikepdf import Array

    fitz = _get_fitz()
    src_bounds = _stamp_source_block_bounds_pikepdf(stamp_path)

    with fitz.open(target_path) as tdoc:
        tpage = tdoc[page_index]
        page_rotation = int(tpage.rotation or 0)
        dst_bounds = _pdf_rect_from_view_rect_rotation_matrix(tpage, target_view_bounds)

    added = 0
    with pikepdf.open(stamp_path) as src, pikepdf.open(target_path, allow_overwriting_input=True) as dst:
        dst_page = dst.pages[page_index]
        if "/Annots" not in dst_page:
            dst_page.Annots = Array()

        src_annots = [
            ref
            for ref in src.pages[0].Annots or []
            if str(ref.get("/Subtype", "")) != "/Widget"
        ]
        if not src_annots:
            return 0

        for ref in src_annots:
            src_rect = [float(v) for v in ref["/Rect"]]
            pdf_rect = _pdf_rect_from_stamp_block_map(
                src_rect, src_bounds, dst_bounds, page_rotation
            )
            subtype = str(ref.get("/Subtype", ""))

            copied = dst.copy_foreign(ref)
            if "/P" in copied:
                del copied["/P"]
            copied["/Rect"] = pdf_rect

            if subtype == "/FreeText":
                if page_rotation:
                    copied["/Rotate"] = page_rotation
                elif "/Rotate" in copied:
                    del copied["/Rotate"]

            _update_copied_annotation_ap(
                copied,
                subtype=subtype,
                pdf_rect=pdf_rect,
                page_rotation=page_rotation,
            )

            dst_page.Annots.append(dst.make_indirect(copied))
            added += 1

        dst.save(target_path)

    added += _copy_stamp_widgets_to_page(
        stamp_path,
        target_path,
        page_index=page_index,
        stamp_view_bounds=stamp_view_bounds,
        target_view_bounds=target_view_bounds,
        values=values,
        populate_text_fields=populate_text_fields,
    )
    return added


def _render_stamp_pixmap(stamp_page: Any, stamp_w: float, stamp_h: float) -> Any:
    sw = float(stamp_page.rect.width) or 1.0
    sh = float(stamp_page.rect.height) or 1.0
    fitz = _get_fitz()
    base = float(stamp_w) / sw
    matrix = fitz.Matrix(base * STAMP_APPEARANCE_SUPERSAMPLE, (stamp_h / sh) * STAMP_APPEARANCE_SUPERSAMPLE)
    return stamp_page.get_pixmap(matrix=matrix, alpha=True)


def _stamp_annot_metadata(values: dict[str, str]) -> tuple[str, str]:
    parts = []
    if values.get("qc_originator_name"):
        parts.append(f"Originator: {values['qc_originator_name']}")
    if values.get("qc_checker_name"):
        parts.append(f"Checker: {values['qc_checker_name']}")
    if values.get("qc_originator_date"):
        parts.append(f"Date: {values['qc_originator_date']}")
    content = "; ".join(parts) if parts else "Peer Review"
    return "Peer Review", content


def _apply_flattened_stamp_annotation(
    pdf_path: Path,
    stamp_path: Path,
    values: dict[str, str],
    *,
    page_index: int,
    stamp_height_pt: float,
    margin_inset_pt: float,
    position_x_pt: float | None,
    position_y_pt: float | None,
    populate_text_fields: bool,
    in_place: bool,
    out_path: Path,
) -> dict[str, Any]:
    """Legacy path: one rasterized Stamp annotation."""
    fitz = _get_fitz()
    with fitz.open(stamp_path) as stamp_doc:
        stamp_page = stamp_doc[0]
        if stamp_page.rect.height <= 0:
            raise ValueError("Invalid stamp page size")
        stamp_w, stamp_h = _stamp_dimensions_from_view_rect(stamp_page.rect, stamp_height_pt)
        if populate_text_fields:
            _fill_stamp_template_widgets(stamp_page, values)

        with fitz.open(pdf_path) as doc:
            page = doc[page_index]
            orig_mediabox = fitz.Rect(page.mediabox)
            orig_cropbox = fitz.Rect(page.cropbox)
            orig_rotation = int(page.rotation or 0)

            stamp_rect_view = _compute_target_stamp_rect_view(
                page,
                stamp_w,
                stamp_h,
                margin_inset_pt=float(margin_inset_pt),
                position_x_pt=position_x_pt,
                position_y_pt=position_y_pt,
            )
            stamp_rect_pdf = _pdf_rect_from_view_rect_rotation_matrix(page, stamp_rect_view)

            pixmap = _render_stamp_pixmap(stamp_page, stamp_rect_view.width, stamp_rect_view.height)
            stamp_annot = page.add_stamp_annot(stamp_rect_view, stamp=pixmap)
            title, content = _stamp_annot_metadata(values)
            stamp_annot.set_info(title=title, subject="QC Review Stamp", content=content)
            stamp_annot.set_flags(stamp_annot.flags | fitz.PDF_ANNOT_IS_PRINT)
            try:
                stamp_annot.set_metadata({"qc_review_fields": json.dumps(values, sort_keys=True)})
            except Exception:
                LOGGER.debug("set_metadata skipped", exc_info=True)
            stamp_annot.update()

            if tuple(page.mediabox) != tuple(orig_mediabox):
                raise RuntimeError("Review stamp modified page MediaBox (not allowed).")
            if tuple(page.cropbox) != tuple(orig_cropbox):
                raise RuntimeError("Review stamp modified page CropBox (not allowed).")
            if int(page.rotation or 0) != orig_rotation:
                raise RuntimeError("Review stamp modified page rotation (not allowed).")

            if in_place:
                doc.saveIncr()
            else:
                doc.save(out_path, garbage=4, deflate=True)

    return {
        "pdf": str(out_path if not in_place else pdf_path),
        "stamp_rect": stamp_rect_pdf,
        "stamp_rect_view": [
            stamp_rect_view.x0,
            stamp_rect_view.y0,
            stamp_rect_view.x1,
            stamp_rect_view.y1,
        ],
        "stamp_annotation": True,
        "copy_markup": False,
        "stamp_markup_count": 1,
        "field_values": values,
        "page_mediabox": [orig_mediabox.x0, orig_mediabox.y0, orig_mediabox.x1, orig_mediabox.y1],
        "stamp_height_pt": stamp_height_pt,
    }


def apply_review_stamp(
    pdf_path: Path,
    stamp_path: Path,
    field_values: dict[str, str] | None = None,
    *,
    page_index: int = 0,
    stamp_height_pt: float = DEFAULT_STAMP_HEIGHT_PT,
    margin_inset_pt: float = DEFAULT_MARGIN_INSET_PT,
    position_x_pt: float | None = None,
    position_y_pt: float | None = None,
    populate_text_fields: bool = True,
    flatten_stamp_annotation: bool = False,
    in_place: bool = True,
    output_path: Path | None = None,
) -> dict[str, Any]:
    """
    Place stamp template markup on the target PDF.

    Default: copy all stamp-page annotations onto the target (Bluebeam paste).
    Optional flatten_stamp_annotation: legacy single raster Stamp annotation.
    """
    fitz = _get_fitz()
    pdf_path = Path(pdf_path)
    stamp_path = Path(stamp_path)
    values = _parse_field_values(field_values) if populate_text_fields else {}

    if not pdf_path.is_file():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")
    if not stamp_path.is_file():
        raise FileNotFoundError(f"Stamp template not found: {stamp_path}")

    out_path = Path(output_path) if output_path else pdf_path
    if not in_place and out_path.resolve() == pdf_path.resolve():
        raise ValueError("output_path required when in_place is False")

    if flatten_stamp_annotation:
        return _apply_flattened_stamp_annotation(
            pdf_path,
            stamp_path,
            values,
            page_index=page_index,
            stamp_height_pt=stamp_height_pt,
            margin_inset_pt=margin_inset_pt,
            position_x_pt=position_x_pt,
            position_y_pt=position_y_pt,
            populate_text_fields=populate_text_fields,
            in_place=in_place,
            out_path=out_path,
        )

    with fitz.open(stamp_path) as stamp_doc, fitz.open(pdf_path) as doc:
        if stamp_doc.page_count < 1:
            raise ValueError(f"Stamp PDF has no pages: {stamp_path}")
        stamp_page = stamp_doc[0]
        if stamp_page.rect.height <= 0:
            raise ValueError("Invalid stamp page size")
        stamp_w, stamp_h = _stamp_dimensions_from_view_rect(stamp_page.rect, stamp_height_pt)

        if page_index < 0 or page_index >= doc.page_count:
            raise ValueError(f"page_index {page_index} out of range (pages={doc.page_count})")
        page = doc[page_index]
        orig_mediabox = fitz.Rect(page.mediabox)
        orig_cropbox = fitz.Rect(page.cropbox)
        orig_rotation = int(page.rotation or 0)

        stamp_rect_view = _compute_target_stamp_rect_view(
            page,
            stamp_w,
            stamp_h,
            margin_inset_pt=float(margin_inset_pt),
            position_x_pt=position_x_pt,
            position_y_pt=position_y_pt,
        )
        stamp_rect_pdf = _pdf_rect_from_view_rect_rotation_matrix(page, stamp_rect_view)
        stamp_view_bounds = [
            float(stamp_page.rect.x0),
            float(stamp_page.rect.y0),
            float(stamp_page.rect.x1),
            float(stamp_page.rect.y1),
        ]
        target_view_bounds = [
            float(stamp_rect_view.x0),
            float(stamp_rect_view.y0),
            float(stamp_rect_view.x1),
            float(stamp_rect_view.y1),
        ]

    target_path = out_path if not in_place else pdf_path
    if not in_place:
        shutil.copy2(pdf_path, target_path)

    stamp_markup_count = _place_stamp_markup_pikepdf(
        stamp_path,
        target_path,
        page_index=page_index,
        stamp_view_bounds=stamp_view_bounds,
        target_view_bounds=target_view_bounds,
        values=values,
        populate_text_fields=populate_text_fields,
    )
    expected = _stamp_markup_count(stamp_path)
    if stamp_markup_count != expected:
        raise RuntimeError(
            f"Stamp markup count mismatch in {stamp_path}: expected={expected} added={stamp_markup_count}"
        )
    if stamp_markup_count <= 0:
        raise RuntimeError(f"Stamp template has no annotations to copy: {stamp_path}")

    with fitz.open(target_path) as doc:
        page = doc[page_index]
        if tuple(page.mediabox) != tuple(orig_mediabox):
            raise RuntimeError("Review stamp modified page MediaBox (not allowed).")
        if tuple(page.cropbox) != tuple(orig_cropbox):
            raise RuntimeError("Review stamp modified page CropBox (not allowed).")
        if int(page.rotation or 0) != orig_rotation:
            raise RuntimeError("Review stamp modified page rotation (not allowed).")

    return {
        "pdf": str(target_path),
        "stamp_rect": stamp_rect_pdf,
        "stamp_rect_view": [
            stamp_rect_view.x0,
            stamp_rect_view.y0,
            stamp_rect_view.x1,
            stamp_rect_view.y1,
        ],
        "stamp_annotation": False,
        "copy_markup": True,
        "stamp_markup_count": stamp_markup_count,
        "field_values": values,
        "page_mediabox": [orig_mediabox.x0, orig_mediabox.y0, orig_mediabox.x1, orig_mediabox.y1],
        "stamp_height_pt": stamp_height_pt,
    }


def build_peer_review_field_values(
    *,
    originator: str = "",
    checker: str = "",
    backchecker: str = "",
    originator_date: str | None = None,
) -> dict[str, str]:
    originator = (originator or "").strip()
    checker = (checker or "").strip()
    backchecker = (backchecker or "").strip()
    if originator_date is None:
        originator_date = date.today().strftime("%m/%d/%Y")
    out: dict[str, str] = {
        "qc_originator_name": originator,
        "qc_checker_name": checker,
        "qc_backchecker_name": backchecker,
        "qc_updater_name": originator,
        "qc_originator_date": originator_date,
    }
    return {k: v for k, v in out.items() if v}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Copy stamp template markup onto a target PDF page."
    )
    parser.add_argument("pdf", type=Path, help="Target PDF (updated in place unless -o)")
    parser.add_argument("stamp", type=Path, help="Stamp template PDF")
    parser.add_argument("-o", "--output", type=Path, default=None, help="Output PDF (default: update pdf in place)")
    parser.add_argument("--fields-json", type=Path, default=None, help="JSON object of AcroForm field name -> value")
    parser.add_argument("--originator", default="", help="Originator / updater display (e.g. email)")
    parser.add_argument("--checker", default="", help="Checker display")
    parser.add_argument("--backchecker", default="", help="Backchecker display")
    parser.add_argument("--originator-date", default=None, help="Originator date (default: today MM/DD/YYYY)")
    parser.add_argument("--stamp-height-pt", type=float, default=DEFAULT_STAMP_HEIGHT_PT)
    parser.add_argument(
        "--margin-inset-pt",
        "--margin-outside-pt",
        dest="margin_inset_pt",
        type=float,
        default=DEFAULT_MARGIN_INSET_PT,
        help="Uniform offset from page top-left (ignored when --stamp-x-pt and --stamp-y-pt are set)",
    )
    parser.add_argument("--stamp-x-pt", type=float, default=None)
    parser.add_argument("--stamp-y-pt", type=float, default=None)
    parser.add_argument("--page-index", type=int, default=0)
    parser.add_argument(
        "--no-populate-text-fields",
        action="store_true",
        help="Do not fill AcroForm widget values from role fields",
    )
    parser.add_argument(
        "--flatten-stamp-annotation",
        action="store_true",
        help="Legacy: rasterize stamp into one Stamp annotation instead of copying markup",
    )
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args(argv)

    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO)

    fields: dict[str, str] = {}
    if not args.no_populate_text_fields:
        if args.fields_json:
            fields = _parse_field_values(json.loads(args.fields_json.read_text(encoding="utf-8")))
        else:
            fields = build_peer_review_field_values(
                originator=args.originator,
                checker=args.checker,
                backchecker=args.backchecker,
                originator_date=args.originator_date,
            )

    in_place = args.output is None
    stamp_kwargs: dict[str, Any] = {
        "page_index": args.page_index,
        "stamp_height_pt": args.stamp_height_pt,
        "populate_text_fields": not args.no_populate_text_fields,
        "flatten_stamp_annotation": bool(args.flatten_stamp_annotation),
        "in_place": in_place,
        "output_path": args.output,
    }
    if args.stamp_x_pt is not None and args.stamp_y_pt is not None:
        stamp_kwargs["position_x_pt"] = args.stamp_x_pt
        stamp_kwargs["position_y_pt"] = args.stamp_y_pt
    else:
        stamp_kwargs["margin_inset_pt"] = args.margin_inset_pt
    result = apply_review_stamp(args.pdf, args.stamp, fields, **stamp_kwargs)
    LOGGER.info("Review stamp applied: %s", result)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        LOGGER.error("%s", exc)
        raise SystemExit(1) from exc
