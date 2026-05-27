"""Tests for overlay/qc_pdf_comments.py (isolated parser)."""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "overlay"))

from qc_pdf_comments import extract_comments  # noqa: E402
from tests.pdf_utils import make_minimal_pdf  # noqa: E402


def test_extract_empty_pdf_has_empty_or_ok_status(tmp_path: Path) -> None:
    pdf = tmp_path / "empty.pdf"
    make_minimal_pdf(pdf, label="no annots")
    result = extract_comments(pdf)
    assert result["parser_status"] in ("empty", "ok", "error")
    assert "annotations" in result
    assert result["parser_version"]


def test_extract_returns_json_serializable(tmp_path: Path) -> None:
    pdf = tmp_path / "t.pdf"
    make_minimal_pdf(pdf)
    result = extract_comments(pdf)
    text = json.dumps(result)
    parsed = json.loads(text)
    assert parsed["parser_version"] == result["parser_version"]
