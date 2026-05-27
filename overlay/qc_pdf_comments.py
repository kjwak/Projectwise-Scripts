#!/usr/bin/env python3
"""Extract PDF annotations/comments to normalized JSON (Bluebeam logic isolated here)."""
from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

PARSER_VERSION = "1.1.0"

try:
    import pymupdf as fitz  # type: ignore
except ImportError:
    fitz = None  # type: ignore


def _pdf_date_to_iso(value: str | None) -> str | None:
    if not value:
        return None
    s = str(value).strip()
    if s.startswith("D:"):
        s = s[2:]
    s = re.sub(r"[^0-9]", "", s)[:14]
    if len(s) < 8:
        return None
    try:
        y, m, d = int(s[0:4]), int(s[4:6]), int(s[6:8])
        hh = int(s[8:10]) if len(s) >= 10 else 0
        mm = int(s[10:12]) if len(s) >= 12 else 0
        ss = int(s[12:14]) if len(s) >= 14 else 0
        return datetime(y, m, d, hh, mm, ss, tzinfo=timezone.utc).isoformat()
    except (ValueError, TypeError):
        return None


def _xref_key(doc: Any, xref: int, key: str) -> str | None:
    try:
        pair = doc.xref_get_key(xref, key)
        if pair and len(pair) >= 2 and pair[1]:
            v = str(pair[1]).strip()
            if v.startswith("(") and v.endswith(")"):
                v = v[1:-1]
            if v.startswith("/"):
                v = v[1:]
            return v or None
    except Exception:
        pass
    return None


def _xref_object(doc: Any, xref: int) -> str | None:
    try:
        s = doc.xref_object(xref, compressed=True)
        if s is None:
            return None
        s = str(s).strip()
        return s or None
    except Exception:
        return None


def _safe_rect(annot: Any) -> dict[str, float] | None:
    try:
        r = annot.rect
        if not r:
            return None
        return {"x0": float(r.x0), "y0": float(r.y0), "x1": float(r.x1), "y1": float(r.y1)}
    except Exception:
        return None


def _safe_vertices(annot: Any) -> list[float] | None:
    try:
        v = getattr(annot, "vertices", None)
        if not v:
            return None
        out: list[float] = []
        for pt in v:
            try:
                out.extend([float(pt.x), float(pt.y)])
            except Exception:
                try:
                    out.extend([float(pt[0]), float(pt[1])])
                except Exception:
                    continue
        return out or None
    except Exception:
        return None


def _collect_irt_status(doc: Any, page: Any) -> dict[int, dict[str, Any]]:
    """Map parent annot xref -> status reply metadata (IRT /State pattern)."""
    out: dict[int, dict[str, Any]] = {}
    try:
        for annot in page.annots() or []:
            irt = getattr(annot, "irt_xref", 0) or 0
            if not irt:
                continue
            state = _xref_key(doc, annot.xref, "State") or _xref_key(doc, annot.xref, "Name")
            author = _xref_key(doc, annot.xref, "T")
            mod = _xref_key(doc, annot.xref, "M")
            out[int(irt)] = {
                "status": state or "Unknown",
                "status_author": author,
                "status_timestamp_utc": _pdf_date_to_iso(mod),
                "status_xref": annot.xref,
            }
    except Exception:
        pass
    return out


def _annot_color(annot: Any) -> str | None:
    try:
        c = annot.colors
        if not c:
            return None
        stroke = c.get("stroke") if isinstance(c, dict) else None
        if stroke:
            return ",".join(str(x) for x in stroke)
    except Exception:
        pass
    return None


def extract_comments(pdf_path: Path) -> dict[str, Any]:
    if fitz is None:
        return {
            "parser_version": PARSER_VERSION,
            "parser_status": "error",
            "error": "pymupdf not installed",
            "annotations": [],
        }

    warnings: list[str] = []
    annotations: list[dict[str, Any]] = []

    doc = fitz.open(str(pdf_path))
    try:
        for pno in range(doc.page_count):
            page = doc[pno]
            irt_map = _collect_irt_status(doc, page)
            for annot in page.annots() or []:
                atype = annot.type[1] if annot.type else "Unknown"
                if atype == "Popup":
                    continue
                if getattr(annot, "irt_xref", 0):
                    continue

                info = annot.info or {}
                xref = annot.xref
                status_info = irt_map.get(int(xref), {})
                status = status_info.get("status") or "Unknown"

                raw: dict[str, Any] = {
                    "xref": xref,
                    "type": atype,
                    "info": dict(info),
                }
                raw_obj = _xref_object(doc, xref)
                if raw_obj:
                    raw["xref_object"] = raw_obj
                rect = _safe_rect(annot)
                if rect:
                    raw["rect"] = rect
                verts = _safe_vertices(annot)
                if verts:
                    raw["vertices"] = verts

                # Frequently useful raw keys (Bluebeam & general PDF annotation plumbing).
                # These are redundant with xref_object but much easier to query in PowerBI later.
                for k in (
                    "NM",  # unique name
                    "Subj",
                    "Contents",
                    "T",
                    "M",
                    "CreationDate",
                    "CA",  # opacity
                    "C",   # color array
                    "RC",  # rich content
                    "IT",  # intent
                    "Name",
                    "State",
                    "StateModel",
                    "InReplyTo",
                ):
                    v = _xref_key(doc, xref, k)
                    if v is not None:
                        raw[f"xref_{k}"] = v
                if status_info:
                    raw["irt_status"] = status_info

                annotations.append(
                    {
                        "annotation_id": str(xref),
                        "page_number": pno + 1,
                        "author": info.get("title") or info.get("author"),
                        "subject": info.get("subject"),
                        "comment_text": info.get("content"),
                        "color": _annot_color(annot),
                        "status": status,
                        "status_author": status_info.get("status_author"),
                        "status_timestamp_utc": status_info.get("status_timestamp_utc"),
                        "created_utc": _pdf_date_to_iso(info.get("creationDate")),
                        "modified_utc": _pdf_date_to_iso(info.get("modDate")),
                        "parent_annotation_id": None,
                        "raw": raw,
                    }
                )
    finally:
        doc.close()

    parser_status = "ok" if annotations else "empty"
    return {
        "parser_version": PARSER_VERSION,
        "parser_status": parser_status,
        "warnings": warnings,
        "annotations": annotations,
    }


def main() -> int:
    ap = argparse.ArgumentParser(description="Extract PDF comments to JSON")
    ap.add_argument("--input", "-i", required=True, help="Input PDF path")
    ap.add_argument("--output", "-o", default="-", help="Output JSON file or - for stdout")
    args = ap.parse_args()

    result = extract_comments(Path(args.input))
    text = json.dumps(result, indent=2)
    if args.output == "-":
        sys.stdout.write(text)
    else:
        Path(args.output).write_text(text, encoding="utf-8")
    return 0 if result.get("parser_status") != "error" else 1


if __name__ == "__main__":
    raise SystemExit(main())
