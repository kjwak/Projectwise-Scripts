"""Tests for QC overlay prepend pipeline."""
from __future__ import annotations

import sys
from pathlib import Path

import pikepdf
import pytest

import qc_overlay_prepend
from pdf_utils import make_layered_single_page_qc, make_minimal_pdf, make_multi_page_pdf


def _list_ocg_names_in_doc(pdf: pikepdf.Pdf) -> list[str]:
    """OCG /Name entries from the document catalog (robust vs page /XObject layout)."""
    ocprops = pdf.Root.get("/OCProperties")
    if not ocprops:
        return []
    ocgs = ocprops.get("/OCGs")
    if not ocgs:
        return []
    names: list[str] = []
    for ocg in ocgs:
        nm = ocg.get("/Name")
        if nm is not None:
            names.append(str(nm))
    return names


def test_first_run_creates_output_from_incoming(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    incoming = tmp_path / "in.pdf"
    make_minimal_pdf(incoming, label="incoming")
    qc_out = tmp_path / "qc.pdf"
    monkeypatch.setattr(
        sys,
        "argv",
        ["qc_overlay_prepend", str(incoming), str(qc_out)],
    )
    with pytest.raises(SystemExit) as exc:
        qc_overlay_prepend.main()
    assert exc.value.code == 0
    assert qc_out.exists()
    with pikepdf.open(qc_out) as pdf:
        assert len(pdf.pages) == 1


def test_extract_flat_history_full_page(tmp_path: Path) -> None:
    hist = tmp_path / "history.pdf"
    make_minimal_pdf(hist, label="flat")
    out = tmp_path / "extracted.pdf"
    assert qc_overlay_prepend.extract_page_1_current_as_old(hist, out)
    with pikepdf.open(out) as pdf:
        assert len(pdf.pages) == 1


def test_extract_layered_history_prefers_current(tmp_path: Path) -> None:
    hist = make_layered_single_page_qc(tmp_path, "layered_hist")
    out = tmp_path / "extracted.pdf"
    assert qc_overlay_prepend.extract_page_1_current_as_old(hist, out)
    with pikepdf.open(out) as pdf:
        assert len(pdf.pages) == 1
        page = pdf.pages[0]
        # Round-2 path: fitz grafts page 1 with OCG state (not necessarily /Cur wrapper)
        assert page.get("/Contents") is not None or page.get("/Resources") is not None


def test_prepend_increments_page_count(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    incoming = tmp_path / "incoming.pdf"
    make_minimal_pdf(incoming, label="inc")
    history = tmp_path / "history.pdf"
    make_multi_page_pdf(history, 2)
    out = tmp_path / "out.pdf"
    monkeypatch.setattr(
        sys,
        "argv",
        ["qc_overlay_prepend", str(incoming), str(history), "-o", str(out)],
    )
    with pytest.raises(SystemExit) as exc:
        qc_overlay_prepend.main()
    assert exc.value.code == 0
    with pikepdf.open(out) as pdf:
        assert len(pdf.pages) == 3
    with pikepdf.open(history) as h0:
        orig_boxes = [h0.pages[i].get("/MediaBox") for i in range(len(h0.pages))]
    with pikepdf.open(out) as h1:
        for i in range(2):
            assert h1.pages[i + 1].get("/MediaBox") == orig_boxes[i]


def test_prepended_page_has_three_ocgs(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    incoming = tmp_path / "incoming.pdf"
    make_minimal_pdf(incoming, label="new")
    history = tmp_path / "history.pdf"
    make_minimal_pdf(history, label="oldhist")
    out = tmp_path / "out.pdf"
    monkeypatch.setattr(
        sys,
        "argv",
        ["qc_overlay_prepend", str(incoming), str(history), "-o", str(out)],
    )
    with pytest.raises(SystemExit) as exc:
        qc_overlay_prepend.main()
    assert exc.value.code == 0
    with pikepdf.open(out) as pdf:
        names = _list_ocg_names_in_doc(pdf)
    joined = " ".join(names)
    assert "Old" in joined
    assert "New" in joined
    assert "Current" in joined
