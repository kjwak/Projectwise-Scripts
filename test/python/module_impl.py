"""Resolve Phase 4E module implementation paths (flat shims forward to subfolders)."""
from __future__ import annotations

import re
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
MODULES_DIR = REPO_ROOT / "modules"

_SHIM_RE = re.compile(
    r"Join-Path\s+\$PSScriptRoot\s+'([^']+\\[^']+\.psm1)'",
    re.IGNORECASE,
)


def module_impl_path(name: str) -> Path:
    """Return on-disk path to module implementation (follows flat compatibility shim)."""
    shim = MODULES_DIR / name
    if shim.is_file():
        text = shim.read_text(encoding="utf-8")
        match = _SHIM_RE.search(text)
        if match:
            impl = MODULES_DIR / match.group(1).replace("\\", "/")
            if impl.is_file():
                return impl
        return shim

    folder_matches = sorted(MODULES_DIR.glob(f"*/{name}"))
    if len(folder_matches) == 1:
        return folder_matches[0]
    if len(folder_matches) > 1:
        paths = ", ".join(str(p) for p in folder_matches)
        raise FileNotFoundError(f"Ambiguous module name {name}: {paths}")

    raise FileNotFoundError(f"Module not found: {shim}")


def read_module_source(name: str) -> str:
    return module_impl_path(name).read_text(encoding="utf-8")
