# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for qc_overlay_prepend (portable paths; spec lives at repo root)."""
from pathlib import Path

from PyInstaller.utils.hooks import collect_all

# SPECDIR = repo root when this file is qc_overlay_prepend.spec at project root
SPECDIR = Path(SPEC).resolve().parent
OVERLAY = SPECDIR / "overlay"

# Pull full packages (binaries + data + submodules); hiddenimports alone often misses pikepdf/pymupdf native bits.
_extra_datas = []
_extra_binaries = []
_extra_hidden = []
for _pkg in ("pikepdf", "pymupdf", "lxml", "PIL"):
    try:
        d, b, h = collect_all(_pkg)
        _extra_datas += d
        _extra_binaries += b
        _extra_hidden += h
    except Exception:
        pass

a = Analysis(
    [str(OVERLAY / "qc_overlay_prepend.py")],
    pathex=[str(OVERLAY)],
    binaries=_extra_binaries,
    datas=[
        (str(OVERLAY / "overlay_build.py"), "."),
        (str(OVERLAY / "overlay_layerize.py"), "."),
        (str(OVERLAY / "flatten_source_layers.py"), "."),
    ]
    + _extra_datas,
    hiddenimports=list(
        dict.fromkeys(
            ["pymupdf", "pikepdf", "fitz", "PIL", "lxml"]
            + _extra_hidden
        )
    ),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # App does not use multiprocessing; excluding avoids pyi_rth_multiprocessing (socket/_socket) on some hosts.
    # UPX must stay off: compressing .pyd often breaks _socket and other stdlib extensions on Windows.
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
    name="qc_overlay_prepend",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
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
    upx=False,
    upx_exclude=[],
    name="qc_overlay_prepend",
)
