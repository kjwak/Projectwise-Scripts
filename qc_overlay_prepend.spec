# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for qc_overlay_prepend (portable paths; spec lives at repo root)."""
from pathlib import Path

# SPECDIR = repo root when this file is qc_overlay_prepend.spec at project root
SPECDIR = Path(SPEC).resolve().parent
OVERLAY = SPECDIR / "overlay"

a = Analysis(
    [str(OVERLAY / "qc_overlay_prepend.py")],
    pathex=[str(OVERLAY)],
    binaries=[],
    datas=[
        (str(OVERLAY / "overlay_build.py"), "."),
        (str(OVERLAY / "overlay_layerize.py"), "."),
        (str(OVERLAY / "flatten_source_layers.py"), "."),
    ],
    hiddenimports=["pymupdf", "pikepdf", "fitz"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="qc_overlay_prepend",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="qc_overlay_prepend",
)
