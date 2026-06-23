"""Tests for qc_review_stamp placement and field mapping."""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

OVERLAY = Path(__file__).resolve().parents[2] / "tools" / "overlay"
ROOT = Path(__file__).resolve().parents[2]
if str(OVERLAY) not in sys.path:
    sys.path.insert(0, str(OVERLAY))

import qc_review_stamp  # noqa: E402
from pdf_utils import make_minimal_pdf  # noqa: E402


def _stamp_annot_count(stamp_path: Path) -> int:
    import pikepdf

    with pikepdf.open(stamp_path) as doc:
        page = doc.pages[0]
        if "/Annots" not in page:
            return 0
        return len(page.Annots)


def _extract_stamp_subtype_annots(path: Path) -> list[dict]:
    import pikepdf

    stamp_subtypes = {"/Square", "/Stamp", "/FreeText"}
    out: list[dict] = []
    with pikepdf.open(path) as doc:
        page = doc.pages[0]
        for ref in page.Annots or []:
            subtype = str(ref.get("/Subtype", ""))
            if subtype not in stamp_subtypes:
                continue
            ap_n = None
            ap_bytes = b""
            if "/AP" in ref and ref["/AP"].get("/N") is not None:
                ap_n = ref["/AP"]["/N"]
                try:
                    ap_bytes = ap_n.read_bytes()
                except Exception:
                    ap_bytes = b""
            out.append(
                {
                    "Subtype": subtype,
                    "Rect": [float(v) for v in ref["/Rect"]],
                    "Rotate": int(ref.get("/Rotate", 0) or 0),
                    "Contents": str(ref.get("/Contents", "")),
                    "AP_BBox": [float(v) for v in ap_n["/BBox"]] if ap_n and "/BBox" in ap_n else None,
                    "AP_Matrix": [float(v) for v in ap_n["/Matrix"]] if ap_n and "/Matrix" in ap_n else None,
                    "AP_bytes": ap_bytes,
                }
            )
    return out


def _union_aspect(annots: list[dict]) -> float:
    xs0 = [a["Rect"][0] for a in annots]
    ys0 = [a["Rect"][1] for a in annots]
    xs1 = [a["Rect"][2] for a in annots]
    ys1 = [a["Rect"][3] for a in annots]
    w = max(xs1) - min(xs0)
    h = max(ys1) - min(ys0)
    return w / h if h else 0.0


def _normalized_block(annots: list[dict]) -> list[dict]:
    xs0 = [a["Rect"][0] for a in annots]
    ys0 = [a["Rect"][1] for a in annots]
    xs1 = [a["Rect"][2] for a in annots]
    ys1 = [a["Rect"][3] for a in annots]
    ux0, uy0, ux1, uy1 = min(xs0), min(ys0), max(xs1), max(ys1)
    uw = ux1 - ux0 or 1.0
    uh = uy1 - uy0 or 1.0
    rel = []
    for a in annots:
        x0, y0, x1, y1 = a["Rect"]
        rel.append(
            {
                "Subtype": a["Subtype"],
                "Contents": a["Contents"],
                "left": (x0 - ux0) / uw,
                "top": (y1 - uy0) / uh,
                "width": (x1 - x0) / uw,
                "height": (y1 - y0) / uh,
            }
        )
    return rel


def _stamp_has_vector_no_image(stamp_annot: dict) -> bool:
    import re

    data = stamp_annot.get("AP_bytes", b"").decode("latin-1", errors="replace")
    if not data:
        return False
    has_image = "/Im" in data
    has_vector = bool(
        re.search(r"(\d+\.?\d*\s+[mlc]|\bre\b|/\w+\s+Do)", data)
    )
    return has_vector and not has_image


def _make_rotated_sheet(path: Path, rotation: int) -> None:
    fitz = qc_review_stamp._get_fitz()
    doc = fitz.open()
    page = doc.new_page(width=1584, height=2448)
    page.insert_text((72, 72), f"rot{rotation}")
    page.set_rotation(rotation)
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))
    doc.close()


def test_compute_stamp_rect_overlay_top_left() -> None:
    fitz = qc_review_stamp._get_fitz()
    page = fitz.Rect(0, 0, 1584, 2448)  # 22x34 in at 72 dpi
    stamp_rect = qc_review_stamp.compute_stamp_rect_overlay_top_left(
        page, stamp_width=120, stamp_height=160, margin_inset=12
    )
    assert stamp_rect.x0 == pytest.approx(12)
    assert stamp_rect.x1 == pytest.approx(132)
    assert stamp_rect.y0 == pytest.approx(12)
    assert stamp_rect.y1 == pytest.approx(172)
    assert stamp_rect.x1 <= page.x1
    assert stamp_rect.y1 <= page.y1


def test_compute_stamp_rect_at_offset_negative_places_outside_top_left() -> None:
    fitz = qc_review_stamp._get_fitz()
    page = fitz.Rect(0, 0, 1584, 2448)
    stamp_rect = qc_review_stamp.compute_stamp_rect_at_offset(
        page, stamp_width=120, stamp_height=160, x_pt=-400, y_pt=-400
    )
    assert stamp_rect.x0 == pytest.approx(-400)
    assert stamp_rect.x1 == pytest.approx(-280)
    assert stamp_rect.y0 == pytest.approx(-400)
    assert stamp_rect.y1 == pytest.approx(-240)


def test_compute_stamp_rect_at_offset_matches_uniform_margin() -> None:
    fitz = qc_review_stamp._get_fitz()
    page = fitz.Rect(0, 0, 1584, 2448)
    by_margin = qc_review_stamp.compute_stamp_rect_overlay_top_left(
        page, stamp_width=120, stamp_height=160, margin_inset=12
    )
    by_xy = qc_review_stamp.compute_stamp_rect_at_offset(
        page, stamp_width=120, stamp_height=160, x_pt=12, y_pt=12
    )
    assert by_margin == by_xy


def test_build_peer_review_field_values_updater_matches_originator() -> None:
    fields = qc_review_stamp.build_peer_review_field_values(
        originator="designer@x.com",
        checker="reviewer@x.com",
        backchecker="checker@x.com",
        originator_date="06/01/2026",
    )
    assert fields["qc_updater_name"] == "designer@x.com"
    assert fields["qc_originator_date"] == "06/01/2026"
    assert "qc_rechecker_name" not in fields


def test_apply_ic_stamp_copies_all_markup(tmp_path: Path) -> None:
    fitz = qc_review_stamp._get_fitz()
    stamp_src = Path(__file__).resolve().parents[2] / "stamps" / "IC_Stamp.pdf"
    if not stamp_src.is_file():
        pytest.skip("IC_Stamp.pdf not in repo")

    target = tmp_path / "sheet.pdf"
    make_minimal_pdf(target, label="sheet", width=1584, height=2448)
    expected = _stamp_annot_count(stamp_src)

    result = qc_review_stamp.apply_review_stamp(
        target,
        stamp_src,
        qc_review_stamp.build_peer_review_field_values(
            originator="designer@x.com",
            checker="checker@x.com",
            backchecker="reviewer@x.com",
        ),
        stamp_height_pt=180,
    )

    assert result["copy_markup"] is True
    assert result["stamp_annotation"] is False
    assert result["stamp_markup_count"] == expected

    after = fitz.open(target)
    page = after[0]
    assert len(list(page.annots() or [])) >= 13
    assert len(list(page.widgets() or [])) >= 10
    after.close()


def test_apply_review_stamp_without_field_population(tmp_path: Path) -> None:
    fitz = qc_review_stamp._get_fitz()
    stamp_src = Path(__file__).resolve().parents[2] / "stamps" / "Peer_Review_Stamp.pdf"
    if not stamp_src.is_file():
        pytest.skip("Peer_Review_Stamp.pdf not in repo")

    target = tmp_path / "sheet.pdf"
    make_minimal_pdf(target, label="sheet", width=1584, height=2448)
    expected = _stamp_annot_count(stamp_src)

    result = qc_review_stamp.apply_review_stamp(
        target,
        stamp_src,
        qc_review_stamp.build_peer_review_field_values(
            originator="o@x.com",
            checker="c@x.com",
            backchecker="b@x.com",
        ),
        stamp_height_pt=180,
        populate_text_fields=False,
    )

    assert result["copy_markup"] is True
    assert result["field_values"] == {}
    assert result["stamp_markup_count"] == expected

    after = fitz.open(target)
    page = after[0]
    assert len(list(page.annots() or [])) >= 13
    assert len(list(page.widgets() or [])) >= 10
    after.close()


def test_apply_review_stamp_flatten_single_annotation(tmp_path: Path) -> None:
    fitz = qc_review_stamp._get_fitz()
    stamp_src = Path(__file__).resolve().parents[2] / "stamps" / "Peer_Review_Stamp.pdf"
    if not stamp_src.is_file():
        pytest.skip("Peer_Review_Stamp.pdf not in repo")

    target = tmp_path / "sheet.pdf"
    make_minimal_pdf(target, label="sheet", width=1584, height=2448)

    before = fitz.open(target)
    old_mb = before[0].mediabox
    old_cb = before[0].cropbox
    old_rot = before[0].rotation
    before.close()

    result = qc_review_stamp.apply_review_stamp(
        target,
        stamp_src,
        qc_review_stamp.build_peer_review_field_values(
            originator="o@x.com",
            checker="c@x.com",
            backchecker="b@x.com",
        ),
        stamp_height_pt=180,
        flatten_stamp_annotation=True,
    )

    assert result["stamp_annotation"] is True
    assert result["copy_markup"] is False
    assert result["field_values"]["qc_originator_name"] == "o@x.com"

    after = fitz.open(target)
    page = after[0]
    assert tuple(page.mediabox) == tuple(old_mb)
    assert tuple(page.cropbox) == tuple(old_cb)
    assert page.rotation == old_rot

    stamp_annots = [a for a in page.annots() or [] if a.type[1] == "Stamp"]
    assert len(stamp_annots) == 1
    assert "Originator: o@x.com" in (stamp_annots[0].info.get("content") or "")
    assert len(list(page.widgets() or [])) == 0
    after.close()


def test_apply_i15_stamp_on_rotated_page_visual_placement(tmp_path: Path) -> None:
    fitz = qc_review_stamp._get_fitz()
    stamp_src = Path(__file__).resolve().parents[2] / "stamps" / "I-15_DR_Stamp.pdf"
    target_src = Path(__file__).resolve().parents[2] / "test" / "powershell" / "050_D-02.10_d0847drn-qc.pdf"
    if not stamp_src.is_file() or not target_src.is_file():
        pytest.skip("I-15 stamp or rotated sample sheet not in repo")

    target = tmp_path / "rotated-sheet.pdf"
    target.write_bytes(target_src.read_bytes())

    result = qc_review_stamp.apply_review_stamp(
        target,
        stamp_src,
        {},
        stamp_height_pt=300,
        position_x_pt=0,
        position_y_pt=0,
    )

    src_bounds = qc_review_stamp._get_stamp_source_bounds(stamp_src)
    src_aspect = (src_bounds[2] - src_bounds[0]) / (src_bounds[3] - src_bounds[1])

    with fitz.open(target) as doc:
        page = doc[0]
        assert int(page.rotation or 0) == 270
        view = fitz.Rect(result["stamp_rect_view"])
        expected_visual = qc_review_stamp.compute_stamp_rect_at_offset(
            page.rect, view.width, view.height, 0, 0
        )
        visual_aspect = view.width / view.height
        assert visual_aspect == pytest.approx(src_aspect, rel=0.02)
        assert view.x0 == pytest.approx(expected_visual.x0, abs=1.0)
        assert view.y0 == pytest.approx(expected_visual.y0, abs=1.0)


def test_apply_i15_stamp_copies_markup_with_colors(tmp_path: Path) -> None:
    fitz = qc_review_stamp._get_fitz()
    stamp_src = Path(__file__).resolve().parents[2] / "stamps" / "I-15_DR_Stamp.pdf"
    if not stamp_src.is_file():
        pytest.skip("I-15_DR_Stamp.pdf not in repo")

    target = tmp_path / "sheet.pdf"
    make_minimal_pdf(target, label="sheet", width=1584, height=2448)
    expected = _stamp_annot_count(stamp_src)

    result = qc_review_stamp.apply_review_stamp(
        target,
        stamp_src,
        {},
        stamp_height_pt=180,
    )

    assert result["copy_markup"] is True
    assert result["stamp_annotation"] is False
    assert result["stamp_markup_count"] == expected

    after = fitz.open(target)
    page = after[0]
    assert len(list(page.annots() or [])) == expected
    after.close()

    annots = _extract_stamp_subtype_annots(target)
    freetext_annots = [a for a in annots if a["Subtype"] == "/FreeText"]
    assert len(freetext_annots) >= 1

    red_found = any(
        "1 0 0 rg" in a.get("AP_bytes", b"").decode("latin-1", errors="replace")
        or "1 0 0 RG" in a.get("AP_bytes", b"").decode("latin-1", errors="replace")
        or "1 0 0 sc" in a.get("AP_bytes", b"").decode("latin-1", errors="replace")
        for a in freetext_annots
    )
    assert red_found


def test_i15_bluebeam_structural_model_on_rotated_sample(tmp_path: Path) -> None:
    stamp_src = ROOT / "stamps" / "I-15_DR_Stamp.pdf"
    target_src = ROOT / "test" / "050_D-02.10_d0847drn-qc.pdf"
    manual_src = ROOT / "test" / "050_D-02.10_d0847drn-manual_stamp.pdf"
    if not stamp_src.is_file() or not target_src.is_file() or not manual_src.is_file():
        pytest.skip("I-15 stamp or sample PDFs not in repo")

    out = tmp_path / "stamped.pdf"
    target = tmp_path / "sheet.pdf"
    target.write_bytes(target_src.read_bytes())

    qc_review_stamp.apply_review_stamp(
        target,
        stamp_src,
        {},
        stamp_height_pt=300,
        position_x_pt=0,
        position_y_pt=0,
        in_place=False,
        output_path=out,
    )

    generated = _extract_stamp_subtype_annots(out)
    manual = _extract_stamp_subtype_annots(manual_src)
    assert len(generated) == 14
    assert {a["Subtype"] for a in generated} == {"/Square", "/Stamp", "/FreeText"}
    assert sum(1 for a in generated if a["Subtype"] == "/Square") == 1
    assert sum(1 for a in generated if a["Subtype"] == "/Stamp") == 1
    assert sum(1 for a in generated if a["Subtype"] == "/FreeText") == 12

    fitz = qc_review_stamp._get_fitz()
    with fitz.open(out) as doc:
        page = doc[0]
        page_rot = int(page.rotation or 0)
        assert page_rot == 270

    for annot in generated:
        if annot["Subtype"] == "/FreeText":
            assert annot["Rotate"] == page_rot
            assert annot["AP_BBox"] == pytest.approx(annot["Rect"], abs=0.05)
            assert annot["AP_Matrix"][:4] == pytest.approx([1, 0, 0, 1], abs=0.001)
        if annot["Subtype"] == "/Square":
            assert annot["Rotate"] == 0
            assert annot["AP_BBox"] == pytest.approx(annot["Rect"], abs=0.05)

    stamp = next(a for a in generated if a["Subtype"] == "/Stamp")
    manual_stamp = next(a for a in manual if a["Subtype"] == "/Stamp")
    assert stamp["AP_BBox"] == pytest.approx(manual_stamp["AP_BBox"], abs=0.01)
    assert stamp["AP_Matrix"] == pytest.approx(manual_stamp["AP_Matrix"], abs=0.01)
    assert _stamp_has_vector_no_image(stamp)

    gen_norm = _normalized_block(generated)
    man_norm = _normalized_block(manual)
    for g, m in zip(gen_norm, man_norm):
        if g["Subtype"] != m["Subtype"] or g["Contents"] != m["Contents"]:
            continue
        assert g["width"] == pytest.approx(m["width"], rel=0.02, abs=0.02)
        assert g["height"] == pytest.approx(m["height"], rel=0.02, abs=0.02)
        assert g["top"] == pytest.approx(m["top"], rel=0.02, abs=0.02)
        assert g["left"] == pytest.approx(m["left"], rel=0.02, abs=0.02)

    # Block aspect in PDF space inverts source stamp aspect on rot270 pages.
    stamp_src_annots = _extract_stamp_subtype_annots(stamp_src)
    src_aspect = _union_aspect(stamp_src_annots)
    assert _union_aspect(generated) == pytest.approx(1.0 / src_aspect, rel=0.02)


def test_i15_stamp_preserves_page_boxes_and_existing_link(tmp_path: Path) -> None:
    stamp_src = ROOT / "stamps" / "I-15_DR_Stamp.pdf"
    target_src = ROOT / "test" / "050_D-02.10_d0847drn-qc.pdf"
    if not stamp_src.is_file() or not target_src.is_file():
        pytest.skip("I-15 stamp or sample PDFs not in repo")

    import pikepdf

    out = tmp_path / "stamped.pdf"
    target = tmp_path / "sheet.pdf"
    target.write_bytes(target_src.read_bytes())

    with pikepdf.open(target) as doc:
        before_page = doc.pages[0]
        before_mb = [float(v) for v in before_page.MediaBox]
        before_cb = [float(v) for v in before_page.CropBox]
        before_rot = int(before_page.get("/Rotate", 0) or 0)
        before_link_count = sum(
            1 for ref in before_page.Annots or [] if str(ref.get("/Subtype", "")) == "/Link"
        )

    qc_review_stamp.apply_review_stamp(
        target,
        stamp_src,
        {},
        stamp_height_pt=300,
        position_x_pt=0,
        position_y_pt=0,
        in_place=False,
        output_path=out,
    )

    with pikepdf.open(out) as doc:
        after_page = doc.pages[0]
        assert [float(v) for v in after_page.MediaBox] == before_mb
        assert [float(v) for v in after_page.CropBox] == before_cb
        assert int(after_page.get("/Rotate", 0) or 0) == before_rot
        after_link_count = sum(
            1 for ref in after_page.Annots or [] if str(ref.get("/Subtype", "")) == "/Link"
        )
        assert after_link_count == before_link_count


@pytest.mark.parametrize("rotation", [0, 90, 180, 270])
def test_i15_stamp_orientation_metadata_all_page_rotations(tmp_path: Path, rotation: int) -> None:
    stamp_src = ROOT / "stamps" / "I-15_DR_Stamp.pdf"
    if not stamp_src.is_file():
        pytest.skip("I-15 stamp not in repo")

    sheet = tmp_path / f"sheet_r{rotation}.pdf"
    out = tmp_path / f"stamped_r{rotation}.pdf"
    _make_rotated_sheet(sheet, rotation)

    qc_review_stamp.apply_review_stamp(
        sheet,
        stamp_src,
        {},
        stamp_height_pt=180,
        position_x_pt=12,
        position_y_pt=12,
        in_place=False,
        output_path=out,
    )

    annots = _extract_stamp_subtype_annots(out)
    assert len(annots) == 14

    fitz = qc_review_stamp._get_fitz()
    with fitz.open(out) as doc:
        page_rot = int(doc[0].rotation or 0)
    assert page_rot == rotation

    stamp = next(a for a in annots if a["Subtype"] == "/Stamp")
    expected_ap_rot = (360 - rotation) % 360
    expected_matrix = qc_review_stamp._stamp_ap_matrix_for_rotation(
        expected_ap_rot, stamp["AP_BBox"] or stamp["Rect"]
    )
    assert stamp["AP_Matrix"] == pytest.approx(expected_matrix, abs=0.05)
    assert _stamp_has_vector_no_image(stamp)

    for annot in annots:
        if annot["Subtype"] == "/FreeText":
            assert annot["Rotate"] == (rotation if rotation else 0)
            assert annot["Contents"]
            assert annot["AP_BBox"] == pytest.approx(annot["Rect"], abs=0.05)
