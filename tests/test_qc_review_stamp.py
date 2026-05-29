"""Tests for qc_review_stamp placement and field mapping."""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

OVERLAY = Path(__file__).resolve().parents[1] / "overlay"
if str(OVERLAY) not in sys.path:
    sys.path.insert(0, str(OVERLAY))

import qc_review_stamp  # noqa: E402
from pdf_utils import make_minimal_pdf  # noqa: E402


def test_compute_stamp_rect_outside_top_left() -> None:
    fitz = qc_review_stamp._get_fitz()
    page = fitz.Rect(0, 0, 1584, 2448)  # 22x34 in at 72 dpi
    stamp_rect, expanded = qc_review_stamp.compute_stamp_rect_outside_top_left(
        page, stamp_width=120, stamp_height=160, margin_outside=12
    )
    assert stamp_rect.x1 <= page.x0 - 12
    assert stamp_rect.y0 >= page.y1 + 12
    assert expanded.x0 <= stamp_rect.x0
    assert expanded.y1 >= stamp_rect.y1


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


def test_apply_review_stamp_expands_mediabox(tmp_path: Path) -> None:
    fitz = qc_review_stamp._get_fitz()
    stamp_src = Path(__file__).resolve().parents[1] / "stamps" / "Peer_Review_Stamp.pdf"
    if not stamp_src.is_file():
        pytest.skip("Peer_Review_Stamp.pdf not in repo")

    target = tmp_path / "sheet.pdf"
    make_minimal_pdf(target, label="sheet", width=1584, height=2448)

    before = fitz.open(target)
    old_mb = before[0].mediabox
    before.close()

    qc_review_stamp.apply_review_stamp(
        target,
        stamp_src,
        qc_review_stamp.build_peer_review_field_values(
            originator="o@x.com",
            checker="c@x.com",
            backchecker="b@x.com",
        ),
        stamp_height_pt=180,
    )

    after = fitz.open(target)
    page = after[0]
    mb = page.mediabox
    assert mb.x0 < old_mb.x0
    assert mb.y1 > old_mb.y1
    names = {w.field_name for w in page.widgets() or []}
    assert "qc_originator_name" in names
    after.close()
