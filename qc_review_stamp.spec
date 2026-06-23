# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for qc_review_stamp (sibling to qc_overlay_prepend in dist)."""
from pathlib import Path

from PyInstaller.utils.hooks import collect_all

SPECDIR = Path(SPEC).resolve().parent
OVERLAY = SPECDIR / "tools" / "overlay"
workpath = str(SPECDIR / "tools" / "overlay" / "build")
distpath = str(SPECDIR / "dist")

_extra_datas = []
_extra_binaries = []
_extra_hidden = []
for _pkg in ("pymupdf",):
    try:
        d, b, h = collect_all(_pkg)
        _extra_datas += d
        _extra_binaries += b
        _extra_hidden += h
    except Exception:
        pass

a = Analysis(
    [str(OVERLAY / "qc_review_stamp.py")],
    pathex=[str(OVERLAY)],
    binaries=_extra_binaries,
    datas=_extra_datas,
    hiddenimports=list(dict.fromkeys(["pymupdf", "fitz"] + _extra_hidden)),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["multiprocessing"],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="qc_review_stamp",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="qc_review_stamp",
)
