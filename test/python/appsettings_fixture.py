"""Read the checked-in base settings with the pipeline's JSONC parser."""

import json
import shutil
import subprocess
from functools import cache
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]


@cache
def read_base_appsettings() -> dict:
    shell = shutil.which("pwsh") or shutil.which("powershell")
    if shell is None:
        raise RuntimeError("PowerShell is required to read appsettings.json")

    script = """
    $ErrorActionPreference = 'Stop'
    $WarningPreference = 'SilentlyContinue'
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    Import-Module './modules/Core/Core.Runtime.psm1' -DisableNameChecking
    Read-QCAppSettingsFile -Path './appsettings.json' | ConvertTo-Json -Depth 100 -Compress
    """
    result = subprocess.run(
        [shell, "-NoProfile", "-NonInteractive", "-Command", script],
        cwd=REPO_ROOT,
        capture_output=True,
        encoding="utf-8-sig",
        check=True,
    )
    return json.loads(result.stdout)
