"""MCP stdio server for pw-qc-debug. Tool calls delegate to PowerShell QC.DebugMcp modules.

The PowerShell worker speaks a line-oriented JSON protocol. Every request carries a
unique ``id``; responses echo that ``id`` (and ``tool``). This prevents one-behind
desync when the MCP host cancels or overlaps tool calls — a common failure mode that
previously returned an unrelated prior payload (often sheet-identity) for
get_recent_errors / get_process_health / get_audit_scan_history.
"""

from __future__ import annotations

import json
import os
import subprocess
import threading
import time
import uuid
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP

_DIR = Path(__file__).resolve().parent
_REPO = Path(os.environ.get("PWQC_REPO_ROOT", _DIR.parent.parent))
_WORKER_PS1 = _DIR / "pw_qc_worker.ps1"
_POWERSHELL = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32/WindowsPowerShell/v1.0/powershell.exe"

# Discard at most this many non-matching stdout lines before restarting the worker.
_MAX_STALE_DISCARDS = 32

mcp = FastMCP("pw-qc-debug")

_worker_lock = threading.Lock()
_worker: subprocess.Popen[str] | None = None


def _worker_env() -> dict[str, str]:
    env = os.environ.copy()
    env.setdefault("PWQC_REPO_ROOT", str(_REPO))
    env.setdefault("PWQC_APPSETTINGS", str(_REPO / "appsettings.json"))
    for key in ("PWQC_SQL_SERVER", "PWQC_SQL_DATABASE", "PWQC_SQL_TRUST_CERT", "PWQC_SQL_DRIVER"):
        val = os.environ.get(key)
        if val:
            env[key] = val
    return env


def _powershell_args(*extra: str) -> list[str]:
    # pwps_dab requires MTA; default PowerShell is STA and PW calls can hang.
    return [
        "-MTA",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        *extra,
    ]


def _start_worker() -> subprocess.Popen[str]:
    if not _POWERSHELL.is_file():
        raise RuntimeError(f"PowerShell not found: {_POWERSHELL}")
    if not _WORKER_PS1.is_file():
        raise RuntimeError(f"Worker script not found: {_WORKER_PS1}")
    proc = subprocess.Popen(
        [
            str(_POWERSHELL),
            *_powershell_args("-File", str(_WORKER_PS1), "-Worker"),
        ],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        bufsize=1,
        env=_worker_env(),
        cwd=str(_DIR),
    )
    assert proc.stdout is not None
    line = proc.stdout.readline()
    if not line:
        err = proc.stderr.read() if proc.stderr else ""
        raise RuntimeError(f"PowerShell worker failed to start: {err}")
    resp = json.loads(line)
    if not resp.get("ok"):
        raise RuntimeError(resp.get("error") or "Worker failed to initialize")
    return proc


def _kill_worker() -> None:
    global _worker
    proc = _worker
    _worker = None
    if proc is None:
        return
    try:
        if proc.stdin:
            proc.stdin.close()
    except Exception:
        pass
    try:
        proc.kill()
    except Exception:
        pass
    try:
        proc.wait(timeout=5)
    except Exception:
        pass


def _get_worker(*, restart: bool = False) -> subprocess.Popen[str]:
    global _worker
    if restart:
        _kill_worker()
    if _worker is None or _worker.poll() is not None:
        _worker = _start_worker()
    return _worker


def _read_matched_response(
    proc: subprocess.Popen[str],
    *,
    request_id: str,
    tool: str,
) -> dict[str, Any]:
    """Read stdout until a JSON line matching request_id arrives; discard stale lines."""
    assert proc.stdout is not None
    discarded = 0
    while discarded <= _MAX_STALE_DISCARDS:
        line = proc.stdout.readline()
        if not line:
            err = ""
            try:
                if proc.stderr:
                    err = proc.stderr.read()
            except Exception:
                pass
            raise RuntimeError(f"PowerShell worker closed stdout while waiting for {tool}: {err}")
        try:
            resp = json.loads(line)
        except json.JSONDecodeError as exc:
            discarded += 1
            continue

        resp_id = resp.get("id")
        if resp_id is None:
            # Legacy worker without correlation — accept only if we have not discarded
            # anything yet (strict FIFO). Otherwise treat as stale noise.
            if discarded == 0:
                return resp
            discarded += 1
            continue

        if str(resp_id) != str(request_id):
            discarded += 1
            continue

        resp_tool = resp.get("tool")
        if resp_tool is not None and str(resp_tool) != tool:
            raise RuntimeError(
                f"Worker response tool mismatch: expected {tool!r}, got {resp_tool!r} (id={request_id})"
            )
        return resp

    raise RuntimeError(
        f"Too many stale worker responses while waiting for {tool} (id={request_id}); "
        "worker will be restarted on next call."
    )


def invoke_ps(tool: str, arguments: dict[str, Any]) -> dict[str, Any]:
    """Run a tool via the persistent PowerShell worker (request/response correlated by id)."""
    request_id = uuid.uuid4().hex
    req = json.dumps(
        {"id": request_id, "tool": tool, "arguments": arguments},
        separators=(",", ":"),
    )

    with _worker_lock:
        last_error: Exception | None = None
        for attempt in range(2):
            try:
                proc = _get_worker(restart=(attempt > 0))
                assert proc.stdin is not None
                proc.stdin.write(req + "\n")
                proc.stdin.flush()
                resp = _read_matched_response(proc, request_id=request_id, tool=tool)
                if not resp.get("ok"):
                    raise RuntimeError(resp.get("error") or "Tool call failed")
                data = resp.get("data")
                return data if isinstance(data, dict) else {"result": data}
            except Exception as exc:
                last_error = exc
                _kill_worker()
        assert last_error is not None
        raise last_error


def _prewarm_worker() -> None:
    """Start the PS worker in the background so first tool call is faster."""
    time.sleep(1.0)
    try:
        with _worker_lock:
            _get_worker()
    except Exception:
        pass


def _start_prewarm_thread() -> None:
    thread = threading.Thread(target=_prewarm_worker, name="pw-qc-prewarm", daemon=True)
    thread.start()


def _tool_args(**kwargs: Any) -> dict[str, Any]:
    return {k: v for k, v in kwargs.items() if v is not None and v != ""}


@mcp.tool()
def search_sheet(
    sheet_number: str = "",
    document_guid: str = "",
    package_id: str = "",
    document_path: str = "",
    sheet_name: str = "",
) -> dict[str, Any]:
    """Search likely sheet/package/document tables for a sheet number, document GUID, or package ID."""
    return invoke_ps("search_sheet", _tool_args(
        sheet_number=sheet_number, document_guid=document_guid, package_id=package_id,
        document_path=document_path, sheet_name=sheet_name,
    ))


@mcp.tool()
def get_sheet_identity(
    sheet_number: str = "",
    document_guid: str = "",
    package_id: str = "",
    document_path: str = "",
    sheet_name: str = "",
) -> dict[str, Any]:
    """Resolve candidate sheet package IDs, document GUIDs, names, and roles."""
    return invoke_ps("get_sheet_identity", _tool_args(
        sheet_number=sheet_number, document_guid=document_guid, package_id=package_id,
        document_path=document_path, sheet_name=sheet_name,
    ))


@mcp.tool()
def get_sheet_package_members(
    sheet_number: str = "",
    document_guid: str = "",
    package_id: str = "",
    document_path: str = "",
    sheet_name: str = "",
) -> dict[str, Any]:
    """Return package members (DGN, sheet PDF, QC PDF) and cross-table consistency hints."""
    return invoke_ps("get_sheet_package_members", _tool_args(
        sheet_number=sheet_number, document_guid=document_guid, package_id=package_id,
        document_path=document_path, sheet_name=sheet_name,
    ))


@mcp.tool()
def get_sheet_debug_timeline(
    sheet_number: str = "",
    document_guid: str = "",
    package_id: str = "",
    document_path: str = "",
    sheet_name: str = "",
    limit: int = 200,
) -> dict[str, Any]:
    """Build a combined timeline from available QC telemetry tables."""
    return invoke_ps("get_sheet_debug_timeline", _tool_args(
        sheet_number=sheet_number, document_guid=document_guid, package_id=package_id,
        document_path=document_path, sheet_name=sheet_name, limit=limit,
    ))


@mcp.tool()
def get_notification_diagnostics(
    sheet_number: str = "",
    document_guid: str = "",
    package_id: str = "",
    document_path: str = "",
    sheet_name: str = "",
    limit: int = 100,
) -> dict[str, Any]:
    """Diagnose notification queue/log outcomes for a sheet."""
    return invoke_ps("get_notification_diagnostics", _tool_args(
        sheet_number=sheet_number, document_guid=document_guid, package_id=package_id,
        document_path=document_path, sheet_name=sheet_name, limit=limit,
    ))


@mcp.tool()
def get_data_integrity_report(
    sheet_number: str = "",
    document_guid: str = "",
    package_id: str = "",
    document_path: str = "",
    sheet_name: str = "",
) -> dict[str, Any]:
    """Compare package/document identity and flag stale or inconsistent rows."""
    return invoke_ps("get_data_integrity_report", _tool_args(
        sheet_number=sheet_number, document_guid=document_guid, package_id=package_id,
        document_path=document_path, sheet_name=sheet_name,
    ))


@mcp.tool()
def get_qc_process_type_diagnostics(
    sheet_number: str = "",
    document_guid: str = "",
    package_id: str = "",
    document_path: str = "",
    sheet_name: str = "",
) -> dict[str, Any]:
    """Compare qc_process_type across lane filenames, sheet_index, lane registry, and ProjectWise."""
    return invoke_ps("get_qc_process_type_diagnostics", _tool_args(
        sheet_number=sheet_number, document_guid=document_guid, package_id=package_id,
        document_path=document_path, sheet_name=sheet_name,
    ))


@mcp.tool()
def warm_projectwise_session() -> dict[str, Any]:
    """Pre-connect ProjectWise in the MCP worker. Call before compare_projectwise_to_database to reduce timeout risk."""
    return invoke_ps("warm_projectwise_session", {})


@mcp.tool()
def compare_projectwise_to_database(
    sheet_number: str = "",
    document_guid: str = "",
    package_id: str = "",
    document_path: str = "",
    sheet_name: str = "",
) -> dict[str, Any]:
    """Read-only comparison of live ProjectWise workflow state vs QC_Pipeline telemetry."""
    return invoke_ps("compare_projectwise_to_database", _tool_args(
        sheet_number=sheet_number, document_guid=document_guid, package_id=package_id,
        document_path=document_path, sheet_name=sheet_name,
    ))


@mcp.tool()
def get_recent_errors(
    limit: int = 100,
    hours: int = 168,
    force_jsonl_fallback: bool = False,
) -> dict[str, Any]:
    """Recent warning/error automation events from automation_events (DB-first)."""
    return invoke_ps("get_recent_errors", {
        "limit": limit, "hours": hours, "force_jsonl_fallback": force_jsonl_fallback,
    })


@mcp.tool()
def get_process_health(force_jsonl_fallback: bool = False) -> dict[str, Any]:
    """Per-process automation health summary from automation_events."""
    return invoke_ps("get_process_health", {"force_jsonl_fallback": force_jsonl_fallback})


@mcp.tool()
def get_audit_scan_history(
    limit: int = 200,
    hours: int = 72,
    force_jsonl_fallback: bool = False,
) -> dict[str, Any]:
    """Watcher audit scan events from automation_events."""
    return invoke_ps("get_audit_scan_history", {
        "limit": limit, "hours": hours, "force_jsonl_fallback": force_jsonl_fallback,
    })


@mcp.tool()
def get_job_timeline(
    job_id: str = "",
    sheet_number: str = "",
    document_guid: str = "",
    package_id: str = "",
    document_path: str = "",
    sheet_name: str = "",
    limit: int = 200,
    force_jsonl_fallback: bool = False,
) -> dict[str, Any]:
    """Automation event timeline for a processing job."""
    return invoke_ps("get_job_timeline", _tool_args(
        job_id=job_id, sheet_number=sheet_number, document_guid=document_guid,
        package_id=package_id, document_path=document_path, sheet_name=sheet_name,
        limit=limit, force_jsonl_fallback=force_jsonl_fallback,
    ))


@mcp.tool()
def get_document_debug_events(
    sheet_number: str = "",
    document_guid: str = "",
    package_id: str = "",
    document_path: str = "",
    sheet_name: str = "",
    limit: int = 200,
    force_jsonl_fallback: bool = False,
) -> dict[str, Any]:
    """Automation events for a document GUID."""
    return invoke_ps("get_document_debug_events", _tool_args(
        sheet_number=sheet_number, document_guid=document_guid, package_id=package_id,
        document_path=document_path, sheet_name=sheet_name, limit=limit,
        force_jsonl_fallback=force_jsonl_fallback,
    ))


@mcp.tool()
def get_package_debug_events(
    sheet_number: str = "",
    document_guid: str = "",
    package_id: str = "",
    document_path: str = "",
    sheet_name: str = "",
    limit: int = 200,
    force_jsonl_fallback: bool = False,
) -> dict[str, Any]:
    """Automation events for a sheet package."""
    return invoke_ps("get_package_debug_events", _tool_args(
        sheet_number=sheet_number, document_guid=document_guid, package_id=package_id,
        document_path=document_path, sheet_name=sheet_name, limit=limit,
        force_jsonl_fallback=force_jsonl_fallback,
    ))


if __name__ == "__main__":
    _start_prewarm_thread()
    mcp.run()
