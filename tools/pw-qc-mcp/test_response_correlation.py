"""Regression: invoke_ps must match responses by request id, not FIFO.

Simulates the failure mode where a cancelled/overlapped MCP call leaves a stale
sheet-identity payload on the worker stdout pipe. Without id correlation,
get_recent_errors would return that unrelated payload.
"""

from __future__ import annotations

import io
import json
import sys
import threading
import types
from pathlib import Path

# Load server.py as a module without starting FastMCP stdio.
SERVER_PATH = Path(__file__).resolve().parent / "server.py"
spec = types.ModuleType("pw_qc_server_under_test")
spec.__file__ = str(SERVER_PATH)
spec.__name__ = "pw_qc_server_under_test"
code = SERVER_PATH.read_text(encoding="utf-8")
# Avoid mcp.run side effects; exec only definitions by stripping __main__.
code = code.replace('if __name__ == "__main__":\n    _start_prewarm_thread()\n    mcp.run()\n', "")
exec(compile(code, str(SERVER_PATH), "exec"), spec.__dict__)


class FakeProc:
    def __init__(self, lines: list[str]):
        self.stdin = io.StringIO()
        self.stdout = io.StringIO("".join(lines))
        self.stderr = io.StringIO("")
        self._poll = None

    def poll(self):
        return self._poll

    def kill(self):
        self._poll = 1

    def wait(self, timeout=None):
        self._poll = 1


def test_discards_stale_sheet_identity_before_errors():
    stale = {
        "ok": True,
        "id": "stale-sheet",
        "tool": "get_sheet_identity",
        "data": {
            "source_tables": ["sheet_index"],
            "data": {"lookup": {"document_guids": ["dff9b94b-154f-4840-8b0f-6cd6d0451b37"]}},
        },
    }
    # Capture the request id written by invoke_ps, then answer with matching id.
    written: dict[str, str] = {}

    class CapturingProc(FakeProc):
        def __init__(self):
            super().__init__([])
            self.stdin = _CapturingStdin(written, self)

    class _CapturingStdin:
        def __init__(self, store: dict[str, str], proc: "CapturingProc"):
            self._store = store
            self._proc = proc
            self._buf = io.StringIO()

        def write(self, s: str):
            self._buf.write(s)
            req = json.loads(s.strip())
            self._store["id"] = req["id"]
            self._store["tool"] = req["tool"]
            # Stale unrelated payload first, then the matching response.
            match = {
                "ok": True,
                "id": req["id"],
                "tool": req["tool"],
                "data": {
                    "source_tables": ["v_mcp_recent_errors"],
                    "data": {"events": [], "event_count": 0, "hours": 3, "fallback_mode": False},
                    "query_assumptions": ["Primary source is automation_events via v_mcp_recent_errors."],
                    "warnings": [],
                },
            }
            # Note: invoke_ps expects the ToolResult shape returned by PS as `data`.
            # Mirror what pw_qc_worker returns: the whole _QDM-ToolResult object.
            match["data"] = {
                "source_tables": ["v_mcp_recent_errors"],
                "query_assumptions": ["Primary source is automation_events via v_mcp_recent_errors."],
                "data": {"events": [{"code": "X"}], "event_count": 1, "hours": 3, "fallback_mode": False},
                "warnings": [],
            }
            self._proc.stdout = io.StringIO(
                json.dumps(stale) + "\n" + json.dumps(match) + "\n"
            )

        def flush(self):
            pass

        def close(self):
            pass

    proc = CapturingProc()
    spec._worker = proc  # type: ignore[attr-defined]

    result = spec.invoke_ps("get_recent_errors", {"limit": 5, "hours": 3})
    assert "source_tables" in result
    assert result["source_tables"] == ["v_mcp_recent_errors"]
    assert result["data"]["event_count"] == 1
    assert "lookup" not in result.get("data", {})
    print("test_discards_stale_sheet_identity_before_errors: PASS")


if __name__ == "__main__":
    test_discards_stale_sheet_identity_before_errors()
