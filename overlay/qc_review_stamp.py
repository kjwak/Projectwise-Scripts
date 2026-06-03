#!/usr/bin/env python3
"""
Apply a peer-review stamp as a single PDF Stamp annotation (Bluebeam-style).

The sheet page is never modified (MediaBox, CropBox, rotation, content streams).
Field values are filled on the stamp template, then rendered into the stamp
annotation appearance so the whole stamp moves as one object in the viewer.
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

DEFAULT_STAMP_HEIGHT_PT = 200.0
DEFAULT_MARGIN_INSET_PT = 12.0
# Supersample stamp appearance so text stays sharp when scaled into the annot rect.
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
    """
    Stamp rectangle at the visual top-left of the page (legacy margin inset).

    Same as compute_stamp_rect_at_offset with x=y=margin_inset.
    """
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
    """
    Stamp rectangle from offsets relative to the page's visual top-left.

    x_pt: distance from the page's left edge to the stamp's left edge (negative = outside left).
    y_pt: distance from the page's top edge to the stamp's top edge (negative = above the page).

    Uses PDF coordinates (y increases upward). Does not expand the page.
    """
    fitz = _get_fitz()
    px0, _py0, _px1, py1 = (
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
        py1 - y_off - sh,
        px0 + x_off + sw,
        py1 - y_off,
    )


def _fill_stamp_template_widgets(stamp_page: Any, values: dict[str, str]) -> None:
    """Apply role values to the stamp template so they appear in the rendered appearance."""
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


def _render_stamp_pixmap(stamp_page: Any, stamp_w: float, stamp_h: float) -> Any:
    """Rasterize filled stamp template for use as stamp annotation appearance."""
    sw = float(stamp_page.rect.width) or 1.0
    sh = float(stamp_page.rect.height) or 1.0
    fitz = _get_fitz()
    base = float(stamp_w) / sw
    matrix = fitz.Matrix(base * STAMP_APPEARANCE_SUPERSAMPLE, (stamp_h / sh) * STAMP_APPEARANCE_SUPERSAMPLE)
    return stamp_page.get_pixmap(matrix=matrix, alpha=True)


def _stamp_annot_metadata(values: dict[str, str]) -> tuple[str, str]:
    """Popup title/subject and a short content line for the stamp annotation."""
    parts = []
    if values.get("qc_originator_name"):
        parts.append(f"Originator: {values['qc_originator_name']}")
    if values.get("qc_checker_name"):
        parts.append(f"Checker: {values['qc_checker_name']}")
    if values.get("qc_originator_date"):
        parts.append(f"Date: {values['qc_originator_date']}")
    content = "; ".join(parts) if parts else "Peer Review"
    return "Peer Review", content


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
    in_place: bool = True,
    output_path: Path | None = None,
) -> dict[str, Any]:
    """
    Add one rubber-stamp annotation with pre-filled appearance (moves as a unit in viewers).

    Page geometry and content are preserved.
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

        if populate_text_fields:
            _fill_stamp_template_widgets(stamp_page, values)

        with fitz.open(pdf_path) as doc:
            if page_index < 0 or page_index >= doc.page_count:
                raise ValueError(f"page_index {page_index} out of range (pages={doc.page_count})")

            page = doc[page_index]
            orig_mediabox = fitz.Rect(page.mediabox)
            orig_cropbox = fitz.Rect(page.cropbox)
            orig_rotation = int(page.rotation or 0)

            if position_x_pt is not None and position_y_pt is not None:
                stamp_rect = compute_stamp_rect_at_offset(
                    page.rect, stamp_w, stamp_h, float(position_x_pt), float(position_y_pt)
                )
            else:
                stamp_rect = compute_stamp_rect_overlay_top_left(
                    page.rect, stamp_w, stamp_h, float(margin_inset_pt)
                )

            pixmap = _render_stamp_pixmap(stamp_page, stamp_w, stamp_h)
            stamp_annot = page.add_stamp_annot(stamp_rect, stamp=pixmap)
            title, content = _stamp_annot_metadata(values)
            stamp_annot.set_info(title=title, subject="QC Review Stamp", content=content)
            stamp_annot.set_flags(stamp_annot.flags | fitz.PDF_ANNOT_IS_PRINT)
            # Store field payload for viewers/tools that read annotation metadata.
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
        "stamp_rect": [stamp_rect.x0, stamp_rect.y0, stamp_rect.x1, stamp_rect.y1],
        "stamp_annotation": True,
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
    parser = argparse.ArgumentParser(
        description="Apply QC review stamp as PDF markup (does not change page size)."
    )
    parser.add_argument("pdf", type=Path, help="QC/history PDF to stamp (updated in place unless -o)")
    parser.add_argument("stamp", type=Path, help="Stamp template PDF (e.g. Peer_Review_Stamp.pdf)")
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
        help="Uniform offset from page top-left (legacy; ignored when --stamp-x-pt and --stamp-y-pt are set)",
    )
    parser.add_argument(
        "--stamp-x-pt",
        type=float,
        default=None,
        help="Horizontal offset from page left to stamp left (PDF pt; negative = outside)",
    )
    parser.add_argument(
        "--stamp-y-pt",
        type=float,
        default=None,
        help="Vertical offset from page top to stamp top (PDF pt; negative = above page)",
    )
    parser.add_argument("--page-index", type=int, default=0)
    parser.add_argument(
        "--no-populate-text-fields",
        action="store_true",
        help="Apply blank stamp appearance without filling AcroForm role/date fields",
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
