"""Smoke tests for schema-aware MCP server (direct function calls)."""
from __future__ import annotations

import json
import sys

import server


def run(name: str, fn, *args, **kwargs):
    print(f"\n=== {name} ===")
    try:
        result = fn(*args, **kwargs)
        print(json.dumps(result, indent=2, default=str)[:8000])
        if len(json.dumps(result, default=str)) > 8000:
            print("... [truncated]")
        return True
    except Exception as exc:
        print(f"FAILED: {exc}")
        return False


def main() -> int:
    ok = True
    ok &= run("health_check", server.health_check)
    ok &= run("describe_table(audit_events)", server.describe_table, "audit_events")
    ok &= run("search_sheet", server.search_sheet, "080J082001ca001")
    ok &= run("get_sheet_debug_timeline", server.get_sheet_debug_timeline, "080J082001ca001", limit=20)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
