#!/usr/bin/env python3
"""Forensic comparison of stamp annotations: source, Bluebeam manual, generated."""

from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT / "overlay") not in sys.path:
    sys.path.insert(0, str(ROOT / "overlay"))

import qc_review_stamp  # noqa: E402


def _get_fitz():
    return qc_review_stamp._get_fitz()


def _pikepdf_dict(obj: Any, depth: int = 0, max_depth: int = 8) -> Any:
    import pikepdf

    if depth > max_depth:
        return str(obj)
    if obj is None:
        return None
    if isinstance(obj, (str, int, float, bool)):
        return obj
    if isinstance(obj, pikepdf.Name):
        return str(obj)
    if isinstance(obj, pikepdf.Array):
        return [_pikepdf_dict(x, depth + 1, max_depth) for x in obj]
    if isinstance(obj, pikepdf.Dictionary):
        out: dict[str, Any] = {}
        for k, v in obj.items():
            key = str(k)
            if key == "/Contents" and isinstance(v, pikepdf.Stream):
                try:
                    out[key] = {
                        "_stream_objnum": v.objgen[0],
                        "decoded": v.read_bytes().decode("latin-1", errors="replace")[:4000],
                    }
                except Exception as exc:
                    out[key] = {"_error": str(exc)}
            elif isinstance(v, pikepdf.Stream):
                stream_info: dict[str, Any] = {"_stream_objnum": v.objgen[0]}
                for sk in ("/Subtype", "/BBox", "/Matrix", "/Filter"):
                    if sk in v:
                        stream_info[str(sk)] = _pikepdf_dict(v[sk], depth + 1, max_depth)
                try:
                    data = v.read_bytes().decode("latin-1", errors="replace")
                    stream_info["has_text_operators"] = any(
                        tok in data for tok in (" Tj", " TJ", " Td", " BT", " Tf")
                    )
                    stream_info["has_image"] = "/Im" in data
                    stream_info["preview"] = data[:2000]
                except Exception as exc:
                    stream_info["_read_error"] = str(exc)
                out[key] = stream_info
            else:
                out[key] = _pikepdf_dict(v, depth + 1, max_depth)
        return out
    if isinstance(obj, pikepdf.Object):
        try:
            if isinstance(obj, pikepdf.Stream):
                return _pikepdf_dict(obj, depth, max_depth)
            return str(obj)
        except Exception:
            return str(obj)
    return str(obj)


def _rect_metrics(rect: list[float]) -> dict[str, float]:
    x0, y0, x1, y1 = (float(v) for v in rect)
    w = abs(x1 - x0)
    h = abs(y1 - y0)
    return {
        "width": w,
        "height": h,
        "aspect_ratio": w / h if h else 0.0,
        "center_x": (x0 + x1) / 2.0,
        "center_y": (y0 + y1) / 2.0,
    }


def _page_info(path: Path, page_index: int = 0) -> dict[str, Any]:
    import pikepdf

    fitz = _get_fitz()
    with pikepdf.open(path) as doc:
        page = doc.pages[page_index]
        mb = [float(v) for v in page.MediaBox]
        cb = [float(v) for v in (page.CropBox if "/CropBox" in page else page.MediaBox)]
        rot = int(page.get("/Rotate", 0) or 0)
        uu = float(page.get("/UserUnit", 1) or 1)
        annot_count = len(page.Annots) if "/Annots" in page else 0
    with fitz.open(path) as doc:
        pg = doc[page_index]
        return {
            "path": str(path),
            "MediaBox": mb,
            "CropBox": cb,
            "Rotate": rot,
            "UserUnit": uu,
            "annotation_count": annot_count,
            "fitz_page_rect": [pg.rect.x0, pg.rect.y0, pg.rect.x1, pg.rect.y1],
            "fitz_rotation": int(pg.rotation or 0),
            "fitz_transformation_matrix": [float(v) for v in pg.transformation_matrix],
            "fitz_rotation_matrix": [float(v) for v in pg.rotation_matrix],
            "fitz_derotation_matrix": [float(v) for v in pg.derotation_matrix],
        }


def _extract_stamp_annots(path: Path, page_index: int = 0) -> list[dict[str, Any]]:
    import pikepdf

    stamp_subtypes = {"/Square", "/Stamp", "/FreeText"}
    out: list[dict[str, Any]] = []
    with pikepdf.open(path) as doc:
        page = doc.pages[page_index]
        annots = list(page.Annots) if "/Annots" in page else []
        for idx, ref in enumerate(annots):
            subtype = str(ref.get("/Subtype", ""))
            if subtype not in stamp_subtypes:
                continue
            rect = [float(v) for v in ref["/Rect"]]
            metrics = _rect_metrics(rect)
            ap_dump: dict[str, Any] | None = None
            ap_n_objnum = None
            if "/AP" in ref and ref["/AP"].get("/N") is not None:
                n = ref["/AP"]["/N"]
                ap_n_objnum = n.objgen[0]
                ap_dump = _pikepdf_dict({"/N": n})
            entry: dict[str, Any] = {
                "annotation_index": idx,
                "pdf_object_number": ref.objgen[0],
                "Subtype": subtype,
                "Rect": rect,
                "metrics": metrics,
                "Rotate": int(ref.get("/Rotate", 0) or 0),
                "Contents": str(ref.get("/Contents", "")),
                "T": str(ref.get("/T", "")),
                "Subj": str(ref.get("/Subj", "")),
                "NM": str(ref.get("/NM", "")),
                "M": str(ref.get("/M", "")),
                "F": int(ref.get("/F", 0) or 0),
                "C": _pikepdf_dict(ref.get("/C")),
                "IC": _pikepdf_dict(ref.get("/IC")),
                "CA": float(ref["/CA"]) if "/CA" in ref else None,
                "DA": str(ref.get("/DA", "")),
                "DS": str(ref.get("/DS", "")),
                "Q": int(ref.get("/Q", 0) or 0) if "/Q" in ref else None,
                "IT": str(ref.get("/IT", "")),
                "RD": _pikepdf_dict(ref.get("/RD")),
                "BS": _pikepdf_dict(ref.get("/BS")),
                "Border": _pikepdf_dict(ref.get("/Border")),
                "MK": _pikepdf_dict(ref.get("/MK")),
                "AP": ap_dump,
                "AP_N_object_number": ap_n_objnum,
                "full_object": _pikepdf_dict(ref, max_depth=4),
            }
            out.append(entry)
    return out


def _union_rect(rects: list[list[float]]) -> list[float]:
    xs0 = [r[0] for r in rects]
    ys0 = [r[1] for r in rects]
    xs1 = [r[2] for r in rects]
    ys1 = [r[3] for r in rects]
    return [min(xs0), min(ys0), max(xs1), max(ys1)]


def _relative_geometry(annots: list[dict[str, Any]]) -> dict[str, Any]:
    if not annots:
        return {}
    union = _union_rect([a["Rect"] for a in annots])
    ux0, uy0, ux1, uy1 = union
    uw = ux1 - ux0 or 1.0
    uh = uy1 - uy0 or 1.0
    rel: list[dict[str, Any]] = []
    for a in annots:
        x0, y0, x1, y1 = a["Rect"]
        m = a["metrics"]
        rel.append(
            {
                "Subtype": a["Subtype"],
                "Contents": a.get("Contents", ""),
                "normalized": {
                    "left": (x0 - ux0) / uw,
                    "top": (y1 - uy0) / uh,
                    "width": (x1 - x0) / uw,
                    "height": (y1 - y0) / uh,
                    "center_x": (m["center_x"] - ux0) / uw,
                    "center_y": (m["center_y"] - uy0) / uh,
                    "aspect_ratio": m["aspect_ratio"],
                },
            }
        )
    return {
        "union_rect": union,
        "union_width": uw,
        "union_height": uh,
        "union_aspect_ratio": uw / uh if uh else 0.0,
        "annotations": rel,
    }


def _fit_affine(src_centers: list[tuple[float, float]], dst_centers: list[tuple[float, float]]) -> dict[str, Any]:
    import numpy as np

    if len(src_centers) != len(dst_centers) or not src_centers:
        return {"error": "center count mismatch"}
    src = np.array(src_centers)
    dst = np.array(dst_centers)
    A = np.hstack([src, np.ones((len(src), 1))])
    M = np.vstack(
        [
            np.linalg.lstsq(A, dst[:, 0], rcond=None)[0],
            np.linalg.lstsq(A, dst[:, 1], rcond=None)[0],
        ]
    )
    pred = A @ M.T
    res = dst - pred
    per_point = [
        {
            "src": src_centers[i],
            "dst": dst_centers[i],
            "predicted": [float(pred[i, 0]), float(pred[i, 1])],
            "residual": [float(res[i, 0]), float(res[i, 1])],
        }
        for i in range(len(src_centers))
    ]
    return {
        "matrix": M.tolist(),
        "max_abs_residual": float(np.max(np.abs(res))),
        "rms_residual": float(math.sqrt(np.mean(res**2))),
        "per_annotation": per_point,
    }


def _match_annots(
    source: list[dict[str, Any]], target: list[dict[str, Any]]
) -> list[tuple[dict[str, Any], dict[str, Any]]]:
    used: set[int] = set()
    pairs: list[tuple[dict[str, Any], dict[str, Any]]] = []
    for s in source:
        best_i = -1
        best_score = -1.0
        for i, t in enumerate(target):
            if i in used:
                continue
            score = 0.0
            if s["Subtype"] == t["Subtype"]:
                score += 3.0
            if s.get("Contents", "") == t.get("Contents", ""):
                score += 2.0
            if s.get("Subj", "") == t.get("Subj", ""):
                score += 1.0
            if score > best_score:
                best_score = score
                best_i = i
        if best_i >= 0:
            used.add(best_i)
            pairs.append((s, target[best_i]))
    return pairs


def _ap_n_entry(entry: dict[str, Any]) -> dict[str, Any]:
    ap = entry.get("AP")
    if not isinstance(ap, dict):
        return {}
    n = ap.get("/N")
    return n if isinstance(n, dict) else {}


def _compare_pair(src: dict[str, Any], dst: dict[str, Any]) -> dict[str, Any]:
    sm = src["metrics"]
    dm = dst["metrics"]
    delta: dict[str, Any] = {
        "Rect_delta": [dst["Rect"][i] - src["Rect"][i] for i in range(4)],
        "center_delta": [dm["center_x"] - sm["center_x"], dm["center_y"] - sm["center_y"]],
        "width_delta": dm["width"] - sm["width"],
        "height_delta": dm["height"] - sm["height"],
        "aspect_ratio_delta": dm["aspect_ratio"] - sm["aspect_ratio"],
        "Rotate_delta": dst.get("Rotate", 0) - src.get("Rotate", 0),
    }
    src_ap = _ap_n_entry(src)
    dst_ap = _ap_n_entry(dst)
    if isinstance(src_ap, dict) and isinstance(dst_ap, dict):
        sb = src_ap.get("/BBox")
        db = dst_ap.get("/BBox")
        smx = src_ap.get("/Matrix")
        dmx = dst_ap.get("/Matrix")
        if sb and db:
            delta["AP_BBox_delta"] = [db[i] - sb[i] for i in range(min(len(db), len(sb)))]
        if smx and dmx:
            delta["AP_Matrix_delta"] = [dmx[i] - smx[i] for i in range(min(len(dmx), len(smx)))]
        delta["dst_has_image_xobject"] = bool(dst_ap.get("has_image"))
        delta["src_has_image_xobject"] = bool(src_ap.get("has_image"))
        delta["dst_has_text_operators"] = bool(dst_ap.get("has_text_operators"))
        delta["src_has_text_operators"] = bool(src_ap.get("has_text_operators"))
    src_keys = set((src.get("full_object") or {}).keys())
    dst_keys = set((dst.get("full_object") or {}).keys())
    delta["missing_in_dst"] = sorted(k for k in src_keys if k not in dst_keys)
    delta["added_in_dst"] = sorted(k for k in dst_keys if k not in src_keys)
    return delta


def _identify_added_annots(
    before: list[dict[str, Any]], after: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    before_keys = {
        (a["Subtype"], a.get("Contents", ""), a.get("Subj", ""), tuple(round(v, 3) for v in a["Rect"]))
        for a in before
    }
    added = []
    for a in after:
        key = (a["Subtype"], a.get("Contents", ""), a.get("Subj", ""), tuple(round(v, 3) for v in a["Rect"]))
        if key not in before_keys:
            added.append(a)
    return added


def build_report(
    *,
    qc_pdf: Path,
    manual_pdf: Path,
    stamp_pdf: Path,
    generated_pdf: Path | None,
    page_index: int = 0,
) -> dict[str, Any]:
    qc_annots = _extract_stamp_annots(qc_pdf, page_index)
    manual_annots = _extract_stamp_annots(manual_pdf, page_index)
    source_annots = _extract_stamp_annots(stamp_pdf, page_index)
    generated_annots = (
        _extract_stamp_annots(generated_pdf, page_index) if generated_pdf and generated_pdf.is_file() else []
    )

    qc_page = _page_info(qc_pdf, page_index)
    manual_page = _page_info(manual_pdf, page_index)
    source_page = _page_info(stamp_pdf, page_index)
    generated_page = _page_info(generated_pdf, page_index) if generated_pdf and generated_pdf.is_file() else None

    manual_added = _identify_added_annots(qc_annots, manual_annots)

    src_manual_pairs = _match_annots(source_annots, manual_added or manual_annots)
    affine = _fit_affine(
        [(s["metrics"]["center_x"], s["metrics"]["center_y"]) for s, _ in src_manual_pairs],
        [(d["metrics"]["center_x"], d["metrics"]["center_y"]) for _, d in src_manual_pairs],
    )

    gen_manual_pairs = _match_annots(manual_added or manual_annots, generated_annots) if generated_annots else []
    generated_vs_manual = [
        {
            "manual": m,
            "generated": g,
            "delta": _compare_pair(m, g),
        }
        for m, g in gen_manual_pairs
    ]

    subtype_counts = {
        "source": {k: sum(1 for a in source_annots if a["Subtype"] == k) for k in ("/Square", "/Stamp", "/FreeText")},
        "manual_added": {
            k: sum(1 for a in (manual_added or manual_annots) if a["Subtype"] == k)
            for k in ("/Square", "/Stamp", "/FreeText")
        },
        "generated": {
            k: sum(1 for a in generated_annots if a["Subtype"] == k) for k in ("/Square", "/Stamp", "/FreeText")
        },
    }

    bluebeam_behavior = {
        "rotates_complete_block_in_pdf_space": abs(affine.get("max_abs_residual", 999)) < 1.0,
        "candidate_affine_matrix_src_to_manual": affine.get("matrix"),
        "affine_max_residual": affine.get("max_abs_residual"),
        "swaps_block_width_height_in_pdf_space": None,
        "sets_rotate_on_freetext": all(
            a.get("Rotate") == manual_page["Rotate"]
            for a in (manual_added or manual_annots)
            if a["Subtype"] == "/FreeText"
        ),
        "freetext_rotate_values": sorted(
            {a.get("Rotate") for a in (manual_added or manual_annots) if a["Subtype"] == "/FreeText"}
        ),
        "stamp_ap_matrix_rotated": None,
        "preserves_stamp_ap_bbox": None,
        "square_ap_bbox_matches_rect": None,
        "freetext_ap_bbox_matches_rect": None,
    }

    src_union = _relative_geometry(source_annots)
    man_union = _relative_geometry(manual_added or manual_annots)
    if src_union and man_union:
        bluebeam_behavior["swaps_block_width_height_in_pdf_space"] = (
            abs(src_union["union_aspect_ratio"] - (1.0 / man_union["union_aspect_ratio"] if man_union["union_aspect_ratio"] else 0))
            < 0.05
        )

    for s, m in src_manual_pairs:
        if s["Subtype"] == "/Stamp":
            s_ap = _ap_n_entry(s)
            m_ap = _ap_n_entry(m)
            if isinstance(s_ap, dict) and isinstance(m_ap, dict):
                bluebeam_behavior["preserves_stamp_ap_bbox"] = s_ap.get("/BBox") == m_ap.get("/BBox")
                bluebeam_behavior["stamp_ap_matrix_rotated"] = s_ap.get("/Matrix") != m_ap.get("/Matrix")
        if s["Subtype"] == "/Square":
            m_ap = _ap_n_entry(m)
            if m_ap.get("/BBox"):
                bluebeam_behavior["square_ap_bbox_matches_rect"] = [
                    round(float(v), 3) for v in m_ap["/BBox"]
                ] == [round(float(v), 3) for v in m["Rect"]]
        if s["Subtype"] == "/FreeText":
            m_ap = _ap_n_entry(m)
            if isinstance(m_ap, dict) and m_ap.get("/BBox"):
                bluebeam_behavior["freetext_ap_bbox_matches_rect"] = [
                    round(float(v), 3) for v in m_ap["/BBox"]
                ] == [round(float(v), 3) for v in m["Rect"]]

    ranking = [
        "Missing FreeText /Rotate equal to page /Rotate causes vertical labels in Bluebeam",
        "Stamp /AP /Matrix not counter-rotated causes upside-down CHECK PRINT linework",
        "View-space rect mapped with transformation_matrix instead of rotation_matrix distorts PDF /Rect aspect ratio on rotated pages",
        "Recreating annotations with fitz instead of copying /AP streams loses Bluebeam appearance model",
        "Independent per-annotation rotation breaks block alignment",
    ]

    return {
        "pages": {
            "qc_before": qc_page,
            "manual_after": manual_page,
            "source_stamp": source_page,
            "generated": generated_page,
        },
        "annotation_counts": {
            "qc_before_stamp_subtypes": len(qc_annots),
            "manual_total": manual_page["annotation_count"],
            "manual_stamp_subtypes": len(manual_annots),
            "manual_added_stamp_subtypes": len(manual_added or manual_annots),
            "source_stamp_subtypes": len(source_annots),
            "generated_stamp_subtypes": len(generated_annots),
            "subtype_counts": subtype_counts,
        },
        "manual_added_annotations": manual_added or manual_annots,
        "source_annotations": source_annots,
        "generated_annotations": generated_annots,
        "source_to_manual_pairs": [
            {"source": s, "manual": m, "delta": _compare_pair(s, m)} for s, m in src_manual_pairs
        ],
        "block_analysis": {
            "source": _relative_geometry(source_annots),
            "manual": _relative_geometry(manual_added or manual_annots),
            "generated": _relative_geometry(generated_annots),
            "affine_src_to_manual": affine,
        },
        "bluebeam_behavior": bluebeam_behavior,
        "generated_vs_manual": generated_vs_manual,
        "likely_root_causes_ranked": ranking,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Compare Bluebeam manual stamp vs source/generated PDFs.")
    parser.add_argument("--qc", type=Path, default=ROOT / "test" / "050_D-02.10_d0847drn-qc.pdf")
    parser.add_argument("--manual", type=Path, default=ROOT / "test" / "050_D-02.10_d0847drn-manual_stamp.pdf")
    parser.add_argument("--stamp", type=Path, default=ROOT / "stamps" / "I-15_DR_Stamp.pdf")
    parser.add_argument("--generated", type=Path, default=ROOT / "test" / "output" / "stamp-test.pdf")
    parser.add_argument(
        "--output",
        type=Path,
        default=ROOT / "test" / "output" / "bluebeam_stamp_comparison.json",
    )
    parser.add_argument("--page-index", type=int, default=0)
    args = parser.parse_args(argv)

    generated = args.generated if args.generated.is_file() else None
    report = build_report(
        qc_pdf=args.qc,
        manual_pdf=args.manual,
        stamp_pdf=args.stamp,
        generated_pdf=generated,
        page_index=args.page_index,
    )
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, indent=2, default=str), encoding="utf-8")
    print(f"Wrote {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
