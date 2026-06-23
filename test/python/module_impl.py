"""Resolve Phase 4E module implementation paths (flat shims forward to subfolders)."""
from __future__ import annotations

import re
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
MODULES_DIR = REPO_ROOT / "modules"

_SHIM_RE = re.compile(
    r"Join-Path\s+\$PSScriptRoot\s+'([^']+\\[^']+\.psm1)'",
    re.IGNORECASE,
)


def module_impl_path(name: str) -> Path:
    """Return on-disk path to module implementation (follows flat compatibility shim)."""
    shim = MODULES_DIR / name
    if not shim.exists():
        raise FileNotFoundError(f"Module not found: {shim}")
    text = shim.read_text(encoding="utf-8")
    match = _SHIM_RE.search(text)
    if match:
        impl = MODULES_DIR / match.group(1).replace("\\", "/")
        if impl.is_file():
            return impl
    return shim


def read_module_source(name: str) -> str:
    return module_impl_path(name).read_text(encoding="utf-8")
