"""Schema-aware read-only MCP debug server for QC_Pipeline (SQL Server)."""
from __future__ import annotations

import os
import re
from datetime import date, datetime, time
from decimal import Decimal
from typing import Any
from uuid import UUID

import pyodbc
from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP

load_dotenv()

mcp = FastMCP("pw-qc-debug")

_TABLE_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
_COLUMNS_CACHE: dict[str, list[str]] = {}
_COLUMN_TYPES_CACHE: dict[str, dict[str, str]] = {}
_TABLES_CACHE: list[str] | None = None

_CAST_AS_NVARCHAR_TYPES = frozenset(
    {
        "datetimeoffset",
        "datetime2",
        "datetime",
        "date",
        "time",
        "smalldatetime",
        "uniqueidentifier",
    }
)

# Tables commonly searched for sheet/package identity (order matters for reporting).
_SHEET_SEARCH_TABLES: list[tuple[str, list[str]]] = [
    ("sheet_index", ["document_name", "document_guid", "folder_path", "sheet_package_id", "pw_state_name"]),
    ("sheet_packages", ["sheet_stem", "folder_path", "sheet_package_id", "dgn_guid", "sheet_pdf_guid", "qc_pdf_guid", "pw_state_name"]),
    ("sheet_documents", ["document_name", "document_guid", "sheet_package_id", "document_role", "pw_state_name"]),
    ("audit_events", ["pw_itemname", "pw_objguid", "resolved_folder", "pw_textparam"]),
    ("transition_events", ["document_name", "document_guid", "folder_path", "from_value", "to_value"]),
    ("document_state_history", ["document_name", "document_guid", "folder_path", "old_value", "new_value"]),
    ("notification_log", ["document_name", "document_guid", "subject", "folder_path"]),
    ("processing_jobs", ["source_path", "source_folder", "job_id", "dedupe_key"]),
    ("qc_workflow_events", ["document_id", "payload_json"]),
]

_TIMELINE_SOURCES: list[dict[str, Any]] = [
    {
        "source": "audit_events",
        "table": "audit_events",
        "time": "captured_at",
        "id": "id",
        "document_name": "pw_itemname",
        "document_guid": "pw_objguid",
        "sheet_package_id": None,
        "event_type": "pw_action_name",
        "state_from": None,
        "state_to": None,
        "status": "processed",
        "actor": "pw_userno",
        "detail_cols": ["pw_action", "pw_acttime", "resolved_folder", "candidate_type", "enqueued_job_id"],
        "where_cols": ["pw_itemname", "pw_objguid", "resolved_folder"],
    },
    {
        "source": "document_state_history",
        "table": "document_state_history",
        "time": "captured_at",
        "id": "id",
        "document_name": "document_name",
        "document_guid": "document_guid",
        "sheet_package_id": "sheet_package_id",
        "event_type": "event_type",
        "state_from": "old_value",
        "state_to": "new_value",
        "status": None,
        "actor": "changed_by_username",
        "detail_cols": ["field_name", "source_audit_id", "transition_group_id", "folder_path"],
        "where_cols": ["document_name", "document_guid", "folder_path"],
    },
    {
        "source": "transition_events",
        "table": "transition_events",
        "time": "detected_at",
        "id": "id",
        "document_name": "document_name",
        "document_guid": "document_guid",
        "sheet_package_id": "sheet_package_id",
        "event_type": "transition_type",
        "state_from": "from_value",
        "state_to": "to_value",
        "status": "notification_sent",
        "actor": "changed_by_username",
        "detail_cols": ["job_id", "job_type", "trigger_audit_id", "notification_id", "folder_path"],
        "where_cols": ["document_name", "document_guid", "folder_path"],
    },
    {
        "source": "qc_workflow_events",
        "table": "qc_workflow_events",
        "time": "created_utc",
        "id": "event_id",
        "document_name": None,
        "document_guid": "document_id",
        "sheet_package_id": "sheet_package_id",
        "event_type": "event_type",
        "state_from": "previous_pw_state",
        "state_to": "target_pw_state",
        "status": "decision_code",
        "actor": None,
        "detail_cols": ["transition_event_id", "payload_json", "processor_version"],
        "where_cols": ["document_id", "payload_json"],
    },
    {
        "source": "notification_log",
        "table": "notification_log",
        "time": "sent_at",
        "id": "id",
        "document_name": "document_name",
        "document_guid": "document_guid",
        "sheet_package_id": "sheet_package_id",
        "event_type": "event_type",
        "state_from": None,
        "state_to": None,
        "status": "success",
        "actor": None,
        "detail_cols": ["recipients", "subject", "dedupe_key", "provider", "error_message", "transition_id"],
        "where_cols": ["document_name", "document_guid", "subject", "folder_path"],
    },
    {
        "source": "processing_jobs",
        "table": "processing_jobs",
        "time": "created_at",
        "id": "id",
        "document_name": "source_path",
        "document_guid": None,
        "sheet_package_id": "sheet_package_id",
        "event_type": "job_type",
        "state_from": None,
        "state_to": None,
        "status": "status",
        "actor": None,
        "detail_cols": ["job_id", "error_code", "error_message", "dedupe_key", "trigger_audit_id", "source_folder"],
        "where_cols": ["source_path", "source_folder", "job_id", "dedupe_key"],
    },
]


def _connect() -> pyodbc.Connection:
    server = os.getenv("PWQC_SQL_SERVER", "localhost")
    database = os.getenv("PWQC_SQL_DATABASE")
    driver = os.getenv("PWQC_SQL_DRIVER", "ODBC Driver 13 for SQL Server")
    trust_cert = os.getenv("PWQC_SQL_TRUST_CERT", "yes")

    if not database:
        raise RuntimeError("Missing PWQC_SQL_DATABASE in .env")

    user = os.getenv("PWQC_SQL_USER")
    password = os.getenv("PWQC_SQL_PASSWORD")

    parts = [
        f"DRIVER={{{driver}}}",
        f"SERVER={server}",
        f"DATABASE={database}",
        f"TrustServerCertificate={trust_cert}",
    ]
    if user and password:
        parts.extend([f"UID={user}", f"PWD={password}"])
    else:
        parts.append("Trusted_Connection=yes")
    return pyodbc.connect(";".join(parts) + ";")


def _serialize_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    if isinstance(value, (Decimal, UUID)):
        return str(value)
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace")
    return str(value)


def execute_dicts(sql: str, params: tuple[Any, ...] | list[Any] | None = None) -> list[dict[str, Any]]:
    """Run a parameterized read-only query and return JSON-safe dict rows."""
    with _connect() as conn:
        cur = conn.cursor()
        cur.execute(sql, params or ())
        if not cur.description:
            return []
        columns = [col[0] for col in cur.description]
        return [
            {col: _serialize_value(val) for col, val in zip(columns, row)}
            for row in cur.fetchall()
        ]


def build_warning(message: str, table: str | None = None, column: str | None = None) -> dict[str, str]:
    warning: dict[str, str] = {"message": message}
    if table:
        warning["table"] = table
    if column:
        warning["column"] = column
    return warning


def safe_top_limit(limit: int, max_limit: int = 500) -> int:
    try:
        value = int(limit)
    except (TypeError, ValueError):
        value = max_limit
    return max(1, min(value, max_limit))


def table_exists(table_name: str) -> bool:
    name = table_name.strip()
    if not _TABLE_NAME_RE.match(name):
        return False
    rows = execute_dicts(
        """
        SELECT 1 AS ok
        FROM INFORMATION_SCHEMA.TABLES
        WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = ?
        """,
        (name,),
    )
    return bool(rows)


def get_column_types(table_name: str) -> dict[str, str]:
    name = table_name.strip()
    if name in _COLUMN_TYPES_CACHE:
        return dict(_COLUMN_TYPES_CACHE[name])
    if not _TABLE_NAME_RE.match(name):
        return {}
    rows = execute_dicts(
        """
        SELECT COLUMN_NAME, DATA_TYPE
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = ?
        ORDER BY ORDINAL_POSITION
        """,
        (name,),
    )
    types = {str(r["COLUMN_NAME"]): str(r["DATA_TYPE"]).lower() for r in rows}
    _COLUMN_TYPES_CACHE[name] = types
    return types


def get_columns(table_name: str) -> list[str]:
    name = table_name.strip()
    if name in _COLUMNS_CACHE:
        return list(_COLUMNS_CACHE[name])
    types = get_column_types(name)
    cols = list(types.keys())
    _COLUMNS_CACHE[name] = cols
    return cols


def sql_column_expr(table_name: str, column_name: str) -> str:
    """Return a SELECT expression that avoids unsupported ODBC types."""
    dtype = get_column_types(table_name).get(column_name, "").lower()
    col = f"[{column_name}]"
    if dtype in _CAST_AS_NVARCHAR_TYPES:
        width = "36" if dtype == "uniqueidentifier" else "50"
        return f"CAST({col} AS NVARCHAR({width})) AS {col}"
    if dtype in {"varbinary", "binary", "image", "timestamp", "rowversion"}:
        return f"CAST({col} AS NVARCHAR(MAX)) AS {col}"
    return col


def quoted_select_columns(table_name: str, columns: list[str]) -> str:
    return ", ".join(sql_column_expr(table_name, col) for col in columns)


def has_columns(table_name: str, required_columns: list[str]) -> bool:
    existing = set(get_columns(table_name))
    return all(col in existing for col in required_columns)


def select_existing_columns(table_name: str, requested_columns: list[str]) -> list[str]:
    existing = set(get_columns(table_name))
    return [col for col in requested_columns if col in existing]


def _list_all_tables() -> list[str]:
    global _TABLES_CACHE
    if _TABLES_CACHE is not None:
        return list(_TABLES_CACHE)
    rows = execute_dicts(
        """
        SELECT TABLE_NAME
        FROM INFORMATION_SCHEMA.TABLES
        WHERE TABLE_SCHEMA = 'dbo'
        ORDER BY TABLE_NAME
        """
    )
    _TABLES_CACHE = [str(r["TABLE_NAME"]) for r in rows]
    return list(_TABLES_CACHE)


def _validate_table_name(table_name: str) -> str | None:
    name = table_name.strip()
    if not _TABLE_NAME_RE.match(name):
        return None
    if not table_exists(name):
        return None
    return name


def _like_pattern(sheet_name: str) -> str:
    return f"%{sheet_name.strip()}%"


def _tool_result(
    data: Any,
    *,
    warnings: list[dict[str, str]] | None = None,
    source_tables: list[str] | None = None,
    query_assumptions: list[str] | None = None,
) -> dict[str, Any]:
    return {
        "data": data,
        "warnings": warnings or [],
        "source_tables": source_tables or [],
        "query_assumptions": query_assumptions or [],
    }


def _first_existing(table: str, candidates: list[str]) -> str | None:
    cols = set(get_columns(table))
    for name in candidates:
        if name in cols:
            return name
    return None


def _build_like_where(table: str, columns: list[str], like_param: str) -> tuple[str, list[Any]] | None:
    usable = select_existing_columns(table, columns)
    if not usable:
        return None
    clauses = [f"CAST([{col}] AS NVARCHAR(MAX)) LIKE ?" for col in usable]
    params = [like_param] * len(usable)
    return " OR ".join(clauses), params


def _collect_sheet_guids(sheet_name: str) -> tuple[set[str], set[str], list[dict[str, str]]]:
    """Return document GUIDs, package IDs, and warnings discovered for a sheet search."""
    like = _like_pattern(sheet_name)
    guids: set[str] = set()
    package_ids: set[str] = set()
    warnings: list[dict[str, str]] = []

    for table, _ in _SHEET_SEARCH_TABLES:
        if not table_exists(table):
            continue
        where = _build_like_where(table, _searchable_columns_for_table(table), like)
        if not where:
            continue
        clause, params = where
        guid_col = _first_existing(table, ["document_guid", "pw_objguid", "document_id", "dgn_guid", "sheet_pdf_guid", "qc_pdf_guid"])
        pkg_col = _first_existing(table, ["sheet_package_id"])
        select_cols = select_existing_columns(
            table,
            ["document_guid", "document_name", "sheet_package_id", "pw_objguid", "document_id", "sheet_stem", "document_role"],
        )
        if not select_cols:
            continue
        quoted = quoted_select_columns(table, select_cols)
        sql = f"SELECT TOP (200) {quoted} FROM [{table}] WHERE {clause}"
        try:
            rows = execute_dicts(sql, tuple(params))
        except pyodbc.Error as exc:
            warnings.append(build_warning(f"Search failed on {table}: {exc}", table=table))
            continue
        for row in rows:
            for key in ("document_guid", "pw_objguid", "document_id"):
                val = row.get(key)
                if val:
                    guids.add(str(val))
            for key in ("dgn_guid", "sheet_pdf_guid", "qc_pdf_guid"):
                val = row.get(key)
                if val:
                    guids.add(str(val))
            if pkg_col:
                val = row.get(pkg_col)
                if val:
                    package_ids.add(str(val))

    return guids, package_ids, warnings


def _searchable_columns_for_table(table: str) -> list[str]:
    for known_table, cols in _SHEET_SEARCH_TABLES:
        if known_table == table:
            return cols
    textish = []
    for col in get_columns(table):
        lower = col.lower()
        if any(token in lower for token in ("name", "guid", "path", "stem", "subject", "item", "folder", "json", "dedupe")):
            textish.append(col)
    return textish


def _timeline_row(
    source: str,
    row: dict[str, Any],
    mapping: dict[str, Any],
) -> dict[str, Any]:
    def pick(field: str | None) -> Any:
        if not field:
            return None
        return row.get(field)

    details_parts: list[str] = []
    for col in mapping.get("detail_cols") or []:
        val = row.get(col)
        if val is not None and str(val).strip():
            details_parts.append(f"{col}={val}")

    return {
        "event_time": pick(mapping["time"]),
        "source": source,
        "source_id": pick(mapping["id"]),
        "document_name": pick(mapping["document_name"]),
        "document_guid": pick(mapping["document_guid"]),
        "sheet_package_id": pick(mapping["sheet_package_id"]),
        "event_type": pick(mapping["event_type"]),
        "state_from": pick(mapping["state_from"]),
        "state_to": pick(mapping["state_to"]),
        "status": pick(mapping["status"]),
        "actor": pick(mapping["actor"]),
        "details": "; ".join(details_parts) if details_parts else None,
    }


# ---------------------------------------------------------------------------
# Core tools
# ---------------------------------------------------------------------------


@mcp.tool()
def health_check() -> dict[str, Any]:
    """Verify database connectivity and return basic server metadata."""
    rows = execute_dicts("SELECT DB_NAME() AS database_name, SYSDATETIME() AS server_time")
    row = rows[0] if rows else {}
    tables = _list_all_tables()
    return _tool_result(
        {
            "database_name": row.get("database_name"),
            "server_time": row.get("server_time"),
            "table_count": len(tables),
        },
        query_assumptions=["Read-only metadata query against current database."],
    )


@mcp.tool()
def list_tables(name_filter: str = "") -> dict[str, Any]:
    """List dbo tables, optionally filtered by substring (case-insensitive)."""
    tables = _list_all_tables()
    filt = name_filter.strip().lower()
    if filt:
        tables = [t for t in tables if filt in t.lower()]
    return _tool_result(
        {"tables": tables, "count": len(tables)},
        source_tables=["INFORMATION_SCHEMA.TABLES"],
        query_assumptions=["Lists base tables in schema dbo only."],
    )


@mcp.tool()
def describe_table(table_name: str) -> dict[str, Any]:
    """Describe columns for a dbo table using INFORMATION_SCHEMA."""
    name = _validate_table_name(table_name)
    if not name:
        return _tool_result(
            None,
            warnings=[build_warning(f"Table not found or invalid name: {table_name}", table=table_name)],
            query_assumptions=["Table names must match dbo identifier rules."],
        )
    rows = execute_dicts(
        """
        SELECT
            COLUMN_NAME AS column_name,
            DATA_TYPE AS data_type,
            CHARACTER_MAXIMUM_LENGTH AS char_max_length,
            IS_NULLABLE AS is_nullable
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = ?
        ORDER BY ORDINAL_POSITION
        """,
        (name,),
    )
    return _tool_result(
        {"table_name": name, "columns": rows, "column_count": len(rows)},
        source_tables=["INFORMATION_SCHEMA.COLUMNS"],
        query_assumptions=[f"Describes dbo.{name} only."],
    )


@mcp.tool()
def find_columns(column_filter: str = "") -> dict[str, Any]:
    """Find columns across dbo tables matching a substring (case-insensitive)."""
    filt = column_filter.strip().lower()
    rows = execute_dicts(
        """
        SELECT TABLE_NAME AS table_name, COLUMN_NAME AS column_name, DATA_TYPE AS data_type
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = 'dbo'
        ORDER BY TABLE_NAME, ORDINAL_POSITION
        """
    )
    if filt:
        rows = [r for r in rows if filt in str(r.get("column_name", "")).lower()]
    return _tool_result(
        {"matches": rows, "count": len(rows)},
        source_tables=["INFORMATION_SCHEMA.COLUMNS"],
        query_assumptions=["Searches column names only (not data values)."],
    )


@mcp.tool()
def sample_table(table_name: str, limit: int = 5) -> dict[str, Any]:
    """Return a small sample of rows from a dbo table (read-only)."""
    name = _validate_table_name(table_name)
    if not name:
        return _tool_result(
            None,
            warnings=[build_warning(f"Table not found or invalid name: {table_name}", table=table_name)],
        )
    top = safe_top_limit(limit, max_limit=50)
    cols = get_columns(name)
    if not cols:
        return _tool_result(
            None,
            warnings=[build_warning(f"No columns discovered for table {name}", table=name)],
            source_tables=[name],
        )
    quoted = quoted_select_columns(name, cols)
    sql = f"SELECT TOP ({top}) {quoted} FROM [{name}]"
    rows = execute_dicts(sql)
    return _tool_result(
        {"table_name": name, "rows": rows, "row_count": len(rows), "limit": top},
        source_tables=[name],
        query_assumptions=[f"Unfiltered sample TOP {top} from dbo.{name}."],
    )


# ---------------------------------------------------------------------------
# QC tools
# ---------------------------------------------------------------------------


@mcp.tool()
def search_sheet(sheet_name: str) -> dict[str, Any]:
    """Search likely sheet/package/document tables for a sheet or document name."""
    like = _like_pattern(sheet_name)
    grouped: dict[str, list[dict[str, Any]]] = {}
    warnings: list[dict[str, str]] = []
    source_tables: list[str] = []
    assumptions = [f"Case-insensitive LIKE match using pattern {like!r}."]

    for table, preferred_cols in _SHEET_SEARCH_TABLES:
        if not table_exists(table):
            warnings.append(build_warning(f"Table dbo.{table} is not present; skipped.", table=table))
            continue
        where = _build_like_where(table, preferred_cols, like)
        if not where:
            warnings.append(build_warning(f"No searchable text columns on dbo.{table}; skipped.", table=table))
            continue
        clause, params = where
        cols = get_columns(table)
        quoted = quoted_select_columns(table, cols)
        sql = f"SELECT TOP (100) {quoted} FROM [{table}] WHERE {clause} ORDER BY 1 DESC"
        try:
            rows = execute_dicts(sql, tuple(params))
        except pyodbc.Error as exc:
            warnings.append(build_warning(f"Query failed: {exc}", table=table))
            continue
        if rows:
            grouped[table] = rows
            source_tables.append(table)

    if table_exists("v_sheet_package_status"):
        vcols = get_columns("v_sheet_package_status")
        vsearch = [c for c in vcols if any(t in c.lower() for t in ("name", "stem", "package"))]
        where = _build_like_where("v_sheet_package_status", vsearch, like)
        if where:
            clause, params = where
            quoted = quoted_select_columns("v_sheet_package_status", vcols)
            sql = f"SELECT TOP (100) {quoted} FROM [v_sheet_package_status] WHERE {clause}"
            try:
                rows = execute_dicts(sql, tuple(params))
                if rows:
                    grouped["v_sheet_package_status"] = rows
                    source_tables.append("v_sheet_package_status")
            except pyodbc.Error as exc:
                warnings.append(build_warning(f"View query failed: {exc}", table="v_sheet_package_status"))

    return _tool_result(
        {"sheet_name": sheet_name, "matches_by_table": grouped, "table_count": len(grouped)},
        warnings=warnings,
        source_tables=source_tables,
        query_assumptions=assumptions,
    )


@mcp.tool()
def get_sheet_identity(sheet_name: str) -> dict[str, Any]:
    """Resolve candidate sheet package IDs, document GUIDs, names, and roles."""
    like = _like_pattern(sheet_name)
    warnings: list[dict[str, str]] = []
    source_tables: list[str] = []
    candidates: dict[str, list[dict[str, Any]]] = {
        "sheet_package_ids": [],
        "document_guids": [],
        "document_names": [],
        "sheet_stems": [],
        "roles": [],
    }

    def add_unique(bucket: str, value: Any, source: str, extra: dict[str, Any] | None = None):
        if value is None or str(value).strip() == "":
            return
        entry = {"value": str(value), "source": source}
        if extra:
            entry.update(extra)
        existing = {item["value"] for item in candidates[bucket]}
        if str(value) not in existing:
            candidates[bucket].append(entry)

    for table in ("sheet_index", "sheet_packages", "sheet_documents"):
        if not table_exists(table):
            warnings.append(build_warning(f"Missing table dbo.{table}", table=table))
            continue
        where = _build_like_where(table, _searchable_columns_for_table(table), like)
        if not where:
            continue
        clause, params = where
        cols = select_existing_columns(
            table,
            [
                "document_guid",
                "document_name",
                "sheet_package_id",
                "sheet_stem",
                "document_role",
                "folder_path",
                "pw_state_name",
                "dgn_guid",
                "sheet_pdf_guid",
                "qc_pdf_guid",
            ],
        )
        if not cols:
            continue
        quoted = quoted_select_columns(table, cols)
        rows = execute_dicts(f"SELECT TOP (200) {quoted} FROM [{table}] WHERE {clause}", tuple(params))
        source_tables.append(table)
        for row in rows:
            add_unique("document_guids", row.get("document_guid"), table, row)
            add_unique("document_guids", row.get("dgn_guid"), table, row)
            add_unique("document_guids", row.get("sheet_pdf_guid"), table, row)
            add_unique("document_guids", row.get("qc_pdf_guid"), table, row)
            add_unique("document_names", row.get("document_name"), table, row)
            add_unique("sheet_package_ids", row.get("sheet_package_id"), table, row)
            add_unique("sheet_stems", row.get("sheet_stem"), table, row)
            add_unique("roles", row.get("document_role"), table, row)

    if table_exists("v_sheet_package_status"):
        vcols = get_columns("v_sheet_package_status")
        vsearch = [c for c in vcols if any(t in c.lower() for t in ("name", "stem", "package", "guid"))]
        where = _build_like_where("v_sheet_package_status", vsearch, like)
        if where:
            clause, params = where
            quoted = quoted_select_columns("v_sheet_package_status", vcols)
            rows = execute_dicts(f"SELECT TOP (100) {quoted} FROM [v_sheet_package_status] WHERE {clause}", tuple(params))
            source_tables.append("v_sheet_package_status")
            for row in rows:
                for col in vcols:
                    val = row.get(col)
                    if val is None:
                        continue
                    lower = col.lower()
                    if "guid" in lower:
                        add_unique("document_guids", val, "v_sheet_package_status", row)
                    elif "name" in lower or "stem" in lower:
                        add_unique("document_names", val, "v_sheet_package_status", row)
                    elif "package" in lower and "id" in lower:
                        add_unique("sheet_package_ids", val, "v_sheet_package_status", row)

    pkg_values = {c["value"].lower() for c in candidates["sheet_package_ids"]}
    guid_values = {c["value"].lower() for c in candidates["document_guids"]}
    if len(pkg_values) > 1:
        warnings.append(build_warning("Multiple distinct sheet_package_id values found; identity may be inconsistent."))
    if len(guid_values) > 6:
        warnings.append(build_warning("Many document GUID candidates found; sheet may have historical/duplicate rows."))

    return _tool_result(
        {"sheet_name": sheet_name, "candidates": candidates},
        warnings=warnings,
        source_tables=source_tables,
        query_assumptions=["Aggregates identity hints from available package/document tables only."],
    )


@mcp.tool()
def get_sheet_package_members(sheet_name: str) -> dict[str, Any]:
    """Return package members (DGN, sheet PDF, QC PDF) and cross-table consistency hints."""
    identity = get_sheet_identity(sheet_name)
    warnings = list(identity.get("warnings", []))
    package_ids = [c["value"] for c in identity["data"]["candidates"]["sheet_package_ids"]]

    members: dict[str, Any] = {
        "by_role": {},
        "sheet_index_rows": [],
        "sheet_packages_rows": [],
        "sheet_documents_rows": [],
    }
    source_tables: list[str] = []

    like = _like_pattern(sheet_name)

    if table_exists("sheet_packages"):
        source_tables.append("sheet_packages")
        cols = select_existing_columns(
            "sheet_packages",
            [
                "sheet_package_id",
                "sheet_stem",
                "folder_path",
                "pw_state_name",
                "dgn_guid",
                "dgn_name",
                "sheet_pdf_guid",
                "sheet_pdf_name",
                "qc_pdf_guid",
                "qc_pdf_name",
                "designer_email",
                "reviewer_email",
                "checker_email",
            ],
        )
        if cols:
            quoted = quoted_select_columns("sheet_packages", cols)
            clauses: list[str] = []
            params: list[Any] = []
            if "sheet_stem" in cols:
                clauses.append("[sheet_stem] LIKE ?")
                params.append(like)
            if package_ids and "sheet_package_id" in cols:
                placeholders = ", ".join("?" for _ in package_ids)
                clauses.append(f"CAST([sheet_package_id] AS NVARCHAR(36)) IN ({placeholders})")
                params.extend(package_ids)
            if clauses:
                sql = f"SELECT {quoted} FROM [sheet_packages] WHERE " + " OR ".join(clauses)
                members["sheet_packages_rows"] = execute_dicts(sql, tuple(params))

    if table_exists("sheet_documents"):
        source_tables.append("sheet_documents")
        cols = select_existing_columns(
            "sheet_documents",
            ["document_guid", "document_name", "document_role", "pw_state_name", "sheet_package_id", "extension"],
        )
        if cols:
            quoted = quoted_select_columns("sheet_documents", cols)
            clauses = []
            params: list[Any] = []
            if "document_name" in cols:
                clauses.append("[document_name] LIKE ?")
                params.append(like)
            if package_ids and "sheet_package_id" in cols:
                placeholders = ", ".join("?" for _ in package_ids)
                clauses.append(f"CAST([sheet_package_id] AS NVARCHAR(36)) IN ({placeholders})")
                params.extend(package_ids)
            if clauses:
                sql = f"SELECT {quoted} FROM [sheet_documents] WHERE " + " OR ".join(clauses)
                members["sheet_documents_rows"] = execute_dicts(sql, tuple(params))
                for row in members["sheet_documents_rows"]:
                    role = str(row.get("document_role") or "unknown")
                    members["by_role"][role] = row

    if table_exists("sheet_index"):
        source_tables.append("sheet_index")
        cols = select_existing_columns(
            "sheet_index",
            [
                "document_guid",
                "document_name",
                "extension",
                "pw_state_name",
                "sheet_package_id",
                "folder_path",
                "qc_pdf_guid",
                "last_updated_at",
            ],
        )
        if cols:
            quoted = quoted_select_columns("sheet_index", cols)
            where = _build_like_where("sheet_index", ["document_name", "folder_path"], like)
            if where:
                clause, params = where
                sql = f"SELECT TOP (50) {quoted} FROM [sheet_index] WHERE {clause} ORDER BY [last_updated_at] DESC"
                if "last_updated_at" not in cols:
                    sql = f"SELECT TOP (50) {quoted} FROM [sheet_index] WHERE {clause}"
                members["sheet_index_rows"] = execute_dicts(sql, tuple(params))

    pkg = members["sheet_packages_rows"][0] if members["sheet_packages_rows"] else {}
    for role, guid_key, name_key in (
        ("dgn", "dgn_guid", "dgn_name"),
        ("sheet_pdf", "sheet_pdf_guid", "sheet_pdf_name"),
        ("qc_pdf", "qc_pdf_guid", "qc_pdf_name"),
    ):
        pkg_guid = str(pkg.get(guid_key) or "").lower()
        doc_row = members["by_role"].get(role) or {}
        doc_guid = str(doc_row.get("document_guid") or "").lower()
        if pkg_guid and doc_guid and pkg_guid != doc_guid:
            warnings.append(
                build_warning(
                    f"GUID mismatch for role {role}: sheet_packages.{guid_key}={pkg_guid} vs sheet_documents={doc_guid}",
                    table="sheet_packages",
                    column=guid_key,
                )
            )

    return _tool_result(
        {"sheet_name": sheet_name, "members": members},
        warnings=warnings,
        source_tables=source_tables,
        query_assumptions=["Compares package registry, role table, and sheet_index when present."],
    )


@mcp.tool()
def get_sheet_debug_timeline(sheet_name: str, limit: int = 200) -> dict[str, Any]:
    """Build a combined timeline from available QC telemetry tables."""
    top = safe_top_limit(limit, max_limit=500)
    like = _like_pattern(sheet_name)
    guids, package_ids, id_warnings = _collect_sheet_guids(sheet_name)
    warnings = list(id_warnings)
    events: list[dict[str, Any]] = []
    source_tables: list[str] = []
    assumptions = [
        f"Timeline LIKE pattern {like!r}.",
        f"Returns at most {top} events after merge.",
        "Skipped sources are reported in warnings.",
    ]

    for spec in _TIMELINE_SOURCES:
        table = spec["table"]
        if not table_exists(table):
            warnings.append(build_warning(f"Timeline source dbo.{table} not present; skipped.", table=table))
            continue

        time_col = spec["time"]
        if time_col not in get_columns(table):
            warnings.append(build_warning(f"Required time column {time_col} missing on dbo.{table}; skipped.", table=table, column=time_col))
            continue

        needed: list[str] = []
        for field in (
            "id",
            "time",
            "document_name",
            "document_guid",
            "sheet_package_id",
            "event_type",
            "state_from",
            "state_to",
            "status",
            "actor",
        ):
            col = spec.get(field)
            if col:
                needed.append(col)
        needed.extend(spec.get("where_cols") or [])
        needed.extend(spec.get("detail_cols") or [])
        select_cols = select_existing_columns(table, needed)
        if not select_cols:
            warnings.append(build_warning(f"No usable columns on dbo.{table}; skipped.", table=table))
            continue

        where_parts: list[str] = []
        params: list[Any] = []
        text_where = _build_like_where(table, spec.get("where_cols") or [], like)
        if text_where:
            where_parts.append(f"({text_where[0]})")
            params.extend(text_where[1])

        guid_col = spec.get("document_guid")
        if guids and guid_col and guid_col in get_columns(table):
            placeholders = ", ".join("?" for _ in guids)
            where_parts.append(f"LOWER(CAST([{guid_col}] AS NVARCHAR(36))) IN ({placeholders})")
            params.extend(g.lower() for g in guids)

        if package_ids and "sheet_package_id" in get_columns(table):
            placeholders = ", ".join("?" for _ in package_ids)
            where_parts.append(f"CAST([sheet_package_id] AS NVARCHAR(36)) IN ({placeholders})")
            params.extend(package_ids)

        if not where_parts:
            warnings.append(build_warning(f"No WHERE strategy for dbo.{table}; skipped.", table=table))
            continue

        quoted = quoted_select_columns(table, select_cols)
        sql = f"SELECT TOP ({top}) {quoted} FROM [{table}] WHERE " + " OR ".join(where_parts)
        if time_col in select_cols:
            sql += f" ORDER BY [{time_col}] DESC"
        try:
            rows = execute_dicts(sql, tuple(params))
        except pyodbc.Error as exc:
            warnings.append(build_warning(f"Timeline query failed on dbo.{table}: {exc}", table=table))
            continue

        source_tables.append(table)
        for row in rows:
            events.append(_timeline_row(spec["source"], row, spec))

    events.sort(key=lambda e: (e.get("event_time") or ""), reverse=True)
    events = events[:top]

    return _tool_result(
        {"sheet_name": sheet_name, "events": events, "event_count": len(events)},
        warnings=warnings,
        source_tables=source_tables,
        query_assumptions=assumptions,
    )


@mcp.tool()
def get_notification_diagnostics(sheet_name: str, limit: int = 100) -> dict[str, Any]:
    """Diagnose notification queue/log outcomes for a sheet."""
    top = safe_top_limit(limit, max_limit=200)
    like = _like_pattern(sheet_name)
    guids, package_ids, id_warnings = _collect_sheet_guids(sheet_name)
    warnings = list(id_warnings)
    source_tables: list[str] = []

    notifications: list[dict[str, Any]] = []
    jobs: list[dict[str, Any]] = []
    transitions: list[dict[str, Any]] = []
    recipients: list[dict[str, Any]] = []

    if table_exists("notification_log"):
        source_tables.append("notification_log")
        cols = select_existing_columns(
            "notification_log",
            [
                "id",
                "sent_at",
                "event_type",
                "document_guid",
                "document_name",
                "recipients",
                "subject",
                "success",
                "error_message",
                "dedupe_key",
                "transition_id",
                "sheet_package_id",
            ],
        )
        if cols:
            quoted = quoted_select_columns("notification_log", cols)
            where = _build_like_where("notification_log", ["document_name", "subject", "folder_path"], like)
            clauses = [where[0]] if where else []
            params = list(where[1]) if where else []
            if guids and "document_guid" in cols:
                placeholders = ", ".join("?" for _ in guids)
                clauses.append(f"LOWER(CAST([document_guid] AS NVARCHAR(36))) IN ({placeholders})")
                params.extend(g.lower() for g in guids)
            if clauses:
                sql = f"SELECT TOP ({top}) {quoted} FROM [notification_log] WHERE " + " OR ".join(clauses)
                if "sent_at" in cols:
                    sql += " ORDER BY [sent_at] DESC"
                notifications = execute_dicts(sql, tuple(params))

    if table_exists("processing_jobs"):
        source_tables.append("processing_jobs")
        cols = select_existing_columns(
            "processing_jobs",
            [
                "id",
                "job_id",
                "job_type",
                "status",
                "created_at",
                "completed_at",
                "source_path",
                "dedupe_key",
                "error_code",
                "error_message",
                "sheet_package_id",
            ],
        )
        if cols:
            quoted = quoted_select_columns("processing_jobs", cols)
            params: list[Any] = []
            text_where = _build_like_where("processing_jobs", ["source_path", "source_folder", "dedupe_key", "job_id"], like)
            sheet_filters: list[str] = []
            if text_where:
                sheet_filters.append(f"({text_where[0]})")
                params.extend(text_where[1])
            if "job_type" in cols and sheet_filters:
                sql = (
                    f"SELECT TOP ({top}) {quoted} FROM [processing_jobs] "
                    f"WHERE [job_type] = ? AND (" + " OR ".join(sheet_filters) + ")"
                )
                params = ["QC_NOTIFICATION", *params]
                if "created_at" in cols:
                    sql += " ORDER BY [created_at] DESC"
                jobs = execute_dicts(sql, tuple(params))
            elif sheet_filters:
                sql = f"SELECT TOP ({top}) {quoted} FROM [processing_jobs] WHERE " + " OR ".join(sheet_filters)
                if "created_at" in cols:
                    sql += " ORDER BY [created_at] DESC"
                jobs = execute_dicts(sql, tuple(params))
            elif "job_type" in cols:
                warnings.append(
                    build_warning(
                        "processing_jobs QC_NOTIFICATION query skipped: no sheet-specific filter columns matched.",
                        table="processing_jobs",
                    )
                )

    if table_exists("transition_events"):
        source_tables.append("transition_events")
        cols = select_existing_columns(
            "transition_events",
            [
                "id",
                "detected_at",
                "document_name",
                "document_guid",
                "from_value",
                "to_value",
                "transition_type",
                "notification_sent",
                "notification_id",
                "trigger_audit_id",
            ],
        )
        if cols:
            quoted = quoted_select_columns("transition_events", cols)
            where = _build_like_where("transition_events", ["document_name", "folder_path"], like)
            clauses = [where[0]] if where else []
            params = list(where[1]) if where else []
            if guids and "document_guid" in cols:
                placeholders = ", ".join("?" for _ in guids)
                clauses.append(f"LOWER(CAST([document_guid] AS NVARCHAR(36))) IN ({placeholders})")
                params.extend(g.lower() for g in guids)
            if clauses:
                sql = f"SELECT TOP ({top}) {quoted} FROM [transition_events] WHERE " + " OR ".join(clauses)
                if "detected_at" in cols:
                    sql += " ORDER BY [detected_at] DESC"
                transitions = execute_dicts(sql, tuple(params))

    for table in ("sheet_index", "sheet_packages"):
        if not table_exists(table):
            continue
        cols = select_existing_columns(table, ["designer_email", "reviewer_email", "checker_email", "sheet_stem", "document_name", "sheet_package_id"])
        if not cols:
            continue
        quoted = quoted_select_columns(table, cols)
        where = _build_like_where(table, ["document_name", "sheet_stem"], like)
        if not where:
            continue
        clause, params = where
        rows = execute_dicts(f"SELECT TOP (20) {quoted} FROM [{table}] WHERE {clause}", tuple(params))
        if rows:
            source_tables.append(table)
            recipients.extend({"source": table, **row} for row in rows)

    assessments: list[dict[str, Any]] = []
    for tr in transitions:
        if str(tr.get("transition_type")) != "STATE_CHANGE":
            continue
        to_state = tr.get("to_value")
        sent_flag = tr.get("notification_sent")
        matching = [
            n
            for n in notifications
            if (n.get("transition_id") == tr.get("id"))
            or (to_state and to_state in str(n.get("subject", "")))
        ]
        matching_jobs = [j for j in jobs if j.get("dedupe_key") and j.get("dedupe_key") in str(tr.get("notification_id") or "")]
        if matching and any(n.get("success") in (True, 1, "1", "True") for n in matching):
            outcome = "sent"
        elif matching and any(n.get("success") in (False, 0, "0", "False") for n in matching):
            outcome = "logged_but_failed"
        elif matching_jobs and any(str(j.get("status", "")).lower() in ("failed", "dead") for j in matching_jobs):
            outcome = "queued_but_failed"
        elif matching_jobs:
            outcome = "queued"
        elif sent_flag in (True, 1, "1", "True"):
            outcome = "transition_marked_sent"
        else:
            outcome = "not_queued"
        assessments.append(
            {
                "transition_id": tr.get("id"),
                "to_value": to_state,
                "notification_sent_flag": sent_flag,
                "outcome": outcome,
            }
        )

    return _tool_result(
        {
            "sheet_name": sheet_name,
            "notification_log": notifications,
            "qc_notification_jobs": jobs,
            "transition_rows": transitions,
            "recipient_fields": recipients,
            "transition_assessments": assessments,
        },
        warnings=warnings,
        source_tables=source_tables,
        query_assumptions=[
            "QC_NOTIFICATION jobs filtered by job_type when column exists.",
            "Outcome is heuristic based on available log/job/transition rows.",
        ],
    )


@mcp.tool()
def get_data_integrity_report(sheet_name: str) -> dict[str, Any]:
    """Compare package/document identity and flag stale or inconsistent rows."""
    members_result = get_sheet_package_members(sheet_name)
    warnings = list(members_result.get("warnings", []))
    members = members_result["data"]["members"]
    issues: list[dict[str, Any]] = []
    source_tables = list(members_result.get("source_tables", []))

    index_rows = members.get("sheet_index_rows") or []
    pkg_rows = members.get("sheet_packages_rows") or []
    doc_rows = members.get("sheet_documents_rows") or []

    if not pkg_rows:
        issues.append({"code": "missing_package", "message": "No sheet_packages row found for sheet."})
    if not doc_rows:
        issues.append({"code": "missing_sheet_documents", "message": "No sheet_documents rows found for sheet."})

    expected_roles = {"dgn", "sheet_pdf", "qc_pdf"}
    found_roles = {str(r.get("document_role") or "").lower() for r in doc_rows}
    for role in sorted(expected_roles - found_roles):
        issues.append({"code": "missing_role", "message": f"sheet_documents missing role {role}.", "role": role})

    name_to_guids: dict[str, set[str]] = {}
    for row in index_rows:
        name = str(row.get("document_name") or "").lower()
        guid = str(row.get("document_guid") or "").lower()
        if name:
            name_to_guids.setdefault(name, set()).add(guid)
    for name, guids in name_to_guids.items():
        if len(guids) > 1:
            issues.append(
                {
                    "code": "duplicate_active_rows",
                    "message": f"sheet_index has {len(guids)} GUIDs for document name {name}.",
                    "document_name": name,
                    "guids": sorted(guids),
                }
            )
            warnings.append(build_warning(f"Duplicate sheet_index rows for {name}", table="sheet_index", column="document_name"))

    pkg = pkg_rows[0] if pkg_rows else {}
    states: dict[str, str] = {}
    if pkg.get("pw_state_name"):
        states["sheet_packages"] = str(pkg["pw_state_name"])
    for row in doc_rows:
        role = str(row.get("document_role") or "unknown")
        if row.get("pw_state_name"):
            states[f"sheet_documents:{role}"] = str(row["pw_state_name"])
    unique_states = set(states.values())
    if len(unique_states) > 1:
        issues.append({"code": "inconsistent_states", "message": "Package/member states disagree.", "states": states})

    for row in index_rows:
        if not row.get("sheet_package_id"):
            issues.append(
                {
                    "code": "missing_package_link",
                    "message": "sheet_index row missing sheet_package_id.",
                    "document_guid": row.get("document_guid"),
                    "document_name": row.get("document_name"),
                }
            )

    if table_exists("v_sheet_package_status"):
        source_tables.append("v_sheet_package_status")
        vcols = get_columns("v_sheet_package_status")
        where = _build_like_where("v_sheet_package_status", [c for c in vcols if "name" in c.lower() or "stem" in c.lower()], _like_pattern(sheet_name))
        view_rows: list[dict[str, Any]] = []
        if where:
            clause, params = where
            quoted = quoted_select_columns("v_sheet_package_status", vcols)
            view_rows = execute_dicts(f"SELECT TOP (20) {quoted} FROM [v_sheet_package_status] WHERE {clause}", tuple(params))
        if not view_rows and pkg_rows:
            issues.append({"code": "view_package_gap", "message": "sheet_packages row exists but v_sheet_package_status returned no matches."})

    return _tool_result(
        {
            "sheet_name": sheet_name,
            "issues": issues,
            "issue_count": len(issues),
            "members_snapshot": {
                "sheet_packages": pkg_rows,
                "sheet_documents": doc_rows,
                "sheet_index": index_rows,
            },
        },
        warnings=warnings,
        source_tables=sorted(set(source_tables)),
        query_assumptions=["Integrity checks use only tables present in the database."],
    )


if __name__ == "__main__":
    mcp.run()
