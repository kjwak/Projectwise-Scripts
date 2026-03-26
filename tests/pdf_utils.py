"""Programmatic minimal PDFs for tests."""
from __future__ import annotations

import sys
from pathlib import Path

import pymupdf as fitz  # type: ignore

_OVERLAY = Path(__file__).resolve().parent.parent / "overlay"
if str(_OVERLAY) not in sys.path:
    sys.path.insert(0, str(_OVERLAY))


def make_minimal_pdf(
    path: Path,
    *,
    width: float = 612,
    height: float = 792,
    label: str = "",
) -> None:
    """Write a one-page PDF (vector linework)."""
    doc = fitz.open()
    page = doc.new_page(width=width, height=height)
    if label:
        page.insert_text((72, 72), label)
    else:
        page.draw_rect(fitz.Rect(0, 0, 200, 400), color=(0, 0, 0))
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))
    doc.close()


def make_multi_page_pdf(path: Path, n_pages: int) -> None:
    """Write a flat PDF with n_pages (no OCGs)."""
    doc = fitz.open()
    for i in range(n_pages):
        page = doc.new_page()
        page.insert_text((72, 72), f"Page{i + 1}")
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))
    doc.close()


def make_layered_single_page_qc(tmp_path: Path, stem: str = "layered_qc") -> Path:
    """One-page QC-style PDF with Old/New/Current OCGs (overlay_build + layerize_overlay)."""
    from overlay_build import build_overlay
    from overlay_layerize import layerize_overlay

    old_pdf = tmp_path / f"{stem}_old.pdf"
    new_pdf = tmp_path / f"{stem}_new.pdf"
    raw = tmp_path / f"{stem}_raw_overlay.pdf"
    out = tmp_path / f"{stem}.pdf"
    make_minimal_pdf(old_pdf, label="OLD")
    make_minimal_pdf(new_pdf, label="NEW")
    build_overlay(
        old_pdf,
        new_pdf,
        raw,
        pages_spec=None,
        fit=False,
        canvas="new",
        add_configs=True,
    )
    layerize_overlay(raw, out, old_alpha=0.6, new_alpha=0.6)
    try:
        raw.unlink()
    except OSError:
        pass
    return out
