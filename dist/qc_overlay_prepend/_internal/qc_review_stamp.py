#!/usr/bin/env python3
"""
Apply an editable AcroForm review stamp outside the top-left of page 1.

Typical sheet PDFs (e.g. 22x34 in) keep the plot on the mediabox; the stamp is placed
above and to the left of the page corner, then MediaBox/CropBox are expanded so viewers
show the stamp in the margin.
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import date
from pathlib import Path
from typing import Any

LOGGER = logging.getLogger("qc_review_stamp")

# Letter template default; scaled to stamp_height_pt on apply.
DEFAULT_STAMP_HEIGHT_PT = 200.0
DEFAULT_MARGIN_OUTSIDE_PT = 12.0


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


def compute_stamp_rect_outside_top_left(
    page_rect: Any,
    stamp_width: float,
    stamp_height: float,
    margin_outside: float,
) -> tuple[Any, Any, float, float]:
    """
    Return (stamp_rect, expanded_mediabox, translate_x, translate_y) in PDF coordinates (y up).

    Keeps expanded mediabox origin at the page's lower-left (no negative x0). Existing page
    content is shifted right by translate_x so the stamp can sit outside the sheet's top-left
    corner at positive coordinates (viewers render reliably).
    """
    fitz = _get_fitz()
    px0, py0, px1, py1 = float(page_rect.x0), float(page_rect.y0), float(page_rect.x1), float(page_rect.y1)

    left_pad = float(margin_outside) + float(stamp_width)
    top_pad = float(margin_outside) + float(stamp_height)
    translate_x = left_pad
    translate_y = 0.0

    stamp_rect = fitz.Rect(
        px0,
        py1 + float(margin_outside),
        px0 + float(stamp_width),
        py1 + float(margin_outside) + float(stamp_height),
    )
    expanded = fitz.Rect(
        px0,
        py0,
        px1 + left_pad,
        py1 + top_pad,
    )
    return stamp_rect, expanded, translate_x, translate_y


def _shift_page_content(page: Any, dx: float, dy: float, page_rect: Any) -> None:
    """Re-draw existing page content shifted by (dx, dy)."""
    if abs(dx) < 1e-6 and abs(dy) < 1e-6:
        return
    fitz = _get_fitz()
    doc = page.parent
    idx = page.number
    snap = fitz.open()
    try:
        snap.insert_pdf(doc, from_page=idx, to_page=idx)
        page.clean_contents()
        dest = fitz.Rect(
            float(page_rect.x0) + dx,
            float(page_rect.y0) + dy,
            float(page_rect.x1) + dx,
            float(page_rect.y1) + dy,
        )
        page.show_pdf_page(dest, snap, 0)
    finally:
        snap.close()


def _normalize_page_rotation(page: Any) -> int:
    """Bake page rotation into content so rect/mediabox match the viewed sheet."""
    fitz = _get_fitz()
    original = int(page.rotation or 0)
    if original % 360 == 0:
        return 0
    page.remove_rotation()
    if int(page.rotation or 0) % 360 != 0:
        raise ValueError(f"Could not normalize page rotation (was {original}, now {page.rotation})")
    return original


def _scale_rect_from_stamp_page(src_rect: Any, stamp_page_rect: Any, dest_rect: Any) -> Any:
    fitz = _get_fitz()
    sw = float(stamp_page_rect.width) or 1.0
    sh = float(stamp_page_rect.height) or 1.0
    sx = float(dest_rect.width) / sw
    sy = float(dest_rect.height) / sh
    rx0 = float(src_rect.x0) - float(stamp_page_rect.x0)
    ry0 = float(src_rect.y0) - float(stamp_page_rect.y0)
    rx1 = float(src_rect.x1) - float(stamp_page_rect.x0)
    ry1 = float(src_rect.y1) - float(stamp_page_rect.y0)
    return fitz.Rect(
        dest_rect.x0 + rx0 * sx,
        dest_rect.y0 + ry0 * sy,
        dest_rect.x0 + rx1 * sx,
        dest_rect.y0 + ry1 * sy,
    )


def _copy_editable_widgets(
    stamp_page: Any,
    target_page: Any,
    stamp_page_rect: Any,
    dest_rect: Any,
    field_values: dict[str, str],
) -> int:
    """Copy stamp widgets onto target_page (editable). Returns count added."""
    fitz = _get_fitz()
    sh = float(stamp_page_rect.height) or 1.0
    sy = float(dest_rect.height) / sh
    added = 0
    for w in stamp_page.widgets() or []:
        name = (w.field_name or "").strip()
        if not name:
            continue
        value = field_values.get(name)
        if value is None:
            value = w.field_value
        if value is None:
            value = ""
        else:
            value = str(value)

        nr = _scale_rect_from_stamp_page(w.rect, stamp_page_rect, dest_rect)
        nw = fitz.Widget()
        nw.field_name = name
        nw.field_type = w.field_type
        nw.field_value = value
        nw.rect = nr
        try:
            if w.text_fontsize:
                nw.text_fontsize = max(4.0, float(w.text_fontsize) * sy)
        except Exception:
            pass
        try:
            if w.text_color:
                nw.text_color = w.text_color
        except Exception:
            pass
        try:
            if w.border_color:
                nw.border_color = w.border_color
        except Exception:
            pass
        try:
            if w.fill_color:
                nw.fill_color = w.fill_color
        except Exception:
            pass
        target_page.add_widget(nw)
        added += 1
    return added


def apply_review_stamp(
    pdf_path: Path,
    stamp_path: Path,
    field_values: dict[str, str] | None = None,
    *,
    page_index: int = 0,
    stamp_height_pt: float = DEFAULT_STAMP_HEIGHT_PT,
    margin_outside_pt: float = DEFAULT_MARGIN_OUTSIDE_PT,
    in_place: bool = True,
    output_path: Path | None = None,
) -> dict[str, Any]:
    """
    Apply stamp template to page_index; expand mediabox/cropbox to include margin area.
    Widgets remain editable on the target page.
    """
    fitz = _get_fitz()
    pdf_path = Path(pdf_path)
    stamp_path = Path(stamp_path)
    values = _parse_field_values(field_values)

    if not pdf_path.is_file():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")
    if not stamp_path.is_file():
        raise FileNotFoundError(f"Stamp template not found: {stamp_path}")

    out_path = Path(output_path) if output_path else pdf_path
    if not in_place and out_path.resolve() == pdf_path.resolve():
        raise ValueError("output_path required when in_place is False")

    with fitz.open(stamp_path) as stamp_doc:
        if stamp_doc.page_count < 1:
            raise ValueError(f"Stamp PDF has no pages: {stamp_path}")
        stamp_page = stamp_doc[0]
        stamp_page_rect = stamp_page.rect
        sw, sh = float(stamp_page_rect.width), float(stamp_page_rect.height)
        if sh <= 0:
            raise ValueError("Invalid stamp page size")
        scale = float(stamp_height_pt) / sh
        stamp_w = sw * scale
        stamp_h = sh * scale

        # Fill template widgets in-memory (for appearance when copying).
        for w in stamp_page.widgets() or []:
            name = (w.field_name or "").strip()
            if name in values:
                w.field_value = values[name]
                w.update()

        with fitz.open(pdf_path) as doc:
            if page_index < 0 or page_index >= doc.page_count:
                raise ValueError(f"page_index {page_index} out of range (pages={doc.page_count})")

            page = doc[page_index]
            _normalize_page_rotation(page)
            page_rect = page.rect
            stamp_rect, expanded, translate_x, translate_y = compute_stamp_rect_outside_top_left(
                page_rect, stamp_w, stamp_h, float(margin_outside_pt)
            )

            _shift_page_content(page, translate_x, translate_y, page_rect)

            page.set_mediabox(expanded)
            try:
                page.set_cropbox(expanded)
            except Exception:
                LOGGER.debug("set_cropbox skipped", exc_info=True)

            page.show_pdf_page(stamp_rect, stamp_doc, 0)
            added = _copy_editable_widgets(
                stamp_page, page, stamp_page_rect, stamp_rect, values
            )

            if in_place:
                doc.saveIncr()
            else:
                doc.save(out_path, garbage=4, deflate=True)

    return {
        "pdf": str(out_path if not in_place else pdf_path),
        "stamp_rect": [stamp_rect.x0, stamp_rect.y0, stamp_rect.x1, stamp_rect.y1],
        "expanded_mediabox": [expanded.x0, expanded.y0, expanded.x1, expanded.y1],
        "content_translate": [translate_x, translate_y],
        "widgets_added": added,
        "stamp_height_pt": stamp_height_pt,
    }


def build_peer_review_field_values(
    *,
    originator: str = "",
    checker: str = "",
    backchecker: str = "",
    originator_date: str | None = None,
) -> dict[str, str]:
    """Map role emails to Peer Review stamp AcroForm field names."""
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
    parser = argparse.ArgumentParser(description="Apply editable QC review stamp outside page 1 top-left.")
    parser.add_argument("pdf", type=Path, help="QC/history PDF to stamp (updated in place unless -o)")
    parser.add_argument("stamp", type=Path, help="Stamp template PDF (e.g. Peer_Review_Stamp.pdf)")
    parser.add_argument("-o", "--output", type=Path, default=None, help="Output PDF (default: update pdf in place)")
    parser.add_argument("--fields-json", type=Path, default=None, help="JSON object of AcroForm field name -> value")
    parser.add_argument("--originator", default="", help="Originator / updater display (e.g. email)")
    parser.add_argument("--checker", default="", help="Checker display")
    parser.add_argument("--backchecker", default="", help="Backchecker display")
    parser.add_argument("--originator-date", default=None, help="Originator date (default: today MM/DD/YYYY)")
    parser.add_argument("--stamp-height-pt", type=float, default=DEFAULT_STAMP_HEIGHT_PT)
    parser.add_argument("--margin-outside-pt", type=float, default=DEFAULT_MARGIN_OUTSIDE_PT)
    parser.add_argument("--page-index", type=int, default=0)
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args(argv)

    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO)

    fields: dict[str, str] = {}
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
    result = apply_review_stamp(
        args.pdf,
        args.stamp,
        fields,
        page_index=args.page_index,
        stamp_height_pt=args.stamp_height_pt,
        margin_outside_pt=args.margin_outside_pt,
        in_place=in_place,
        output_path=args.output,
    )
    LOGGER.info("Review stamp applied: %s", result)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        LOGGER.error("%s", exc)
        raise SystemExit(1) from exc
