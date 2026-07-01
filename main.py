import asyncio
import csv
import io
import json
import os
import re
import traceback
import uuid
from rules import (
    validate_sql_against_rules,
    validate_result_against_rules,
    get_intent,
    get_rulebook_entry,
    validate_summary_against_facts,
)
from datetime import date, datetime, timedelta
from typing import Dict, List, Literal, Optional
from charts import build_chart_html, reply_already_has_chart

import httpx
import anthropic
from fastapi import FastAPI, Header, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from pydantic import BaseModel
from dotenv import load_dotenv

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate,
    PageBreak, Paragraph, Spacer, Table, TableStyle,
)

# =============================================================================
# [01] Imports & Config
# =============================================================================
load_dotenv()


# =============================================================================
# [02] QUERY RESULT STORE  (ClickHouse-backed)
# =============================================================================
class QueryResult:
    def __init__(self, sql: str, columns: List[str], rows: List[dict],
                 total_rows: int, captured_at: str, filters_applied: str = ""):
        self.sql             = sql
        self.columns         = columns
        self.rows            = rows
        self.total_rows      = total_rows
        self.captured_at     = captured_at
        self.filters_applied = filters_applied


_RULE_AUDIT_LOG: List[dict] = []   # most-recent-first, capped — in-memory is fine, non-critical

def _log_rule_audit(session_id: Optional[str], sql: str, violations: List[str],
                     stage: str, user_message: str):
    entry = {
        "ts": datetime.utcnow().isoformat(),
        "session_id": session_id,
        "stage": stage,                # "pre_execute" | "post_execute" | "post_summary"
        "violations": violations,
        "sql_preview": sql[:300],
        "user_message": user_message[:200],
    }
    _RULE_AUDIT_LOG.insert(0, entry)
    del _RULE_AUDIT_LOG[200:]          # cap log size
    if violations:
        print(f"[RULE-AUDIT][{stage}] session={session_id} violations={violations}")


async def _store_result(session_key: str, result: QueryResult):
    """Writes a QueryResult into the diud_sessions ClickHouse table."""
    sql_esc   = result.sql.replace("'", "''")
    cols_json = json.dumps(result.columns).replace("'", "''")
    rows_json = json.dumps(result.rows, default=str).replace("'", "''")
    filt_esc  = result.filters_applied.replace("'", "''")
    insert_sql = f"""
        INSERT INTO kore_ai_hubspot.diud_sessions
        (session_key, sql, columns, rows, total_rows, filters_applied)
        VALUES ('{session_key}', '{sql_esc}', '{cols_json}', '{rows_json}',
                {result.total_rows}, '{filt_esc}')
    """
    await _ch_execute_raw(insert_sql)


async def _get_result(session_key: str) -> Optional[QueryResult]:
    """Reads the latest QueryResult for a session_key from ClickHouse."""
    select_sql = f"""
        SELECT sql, columns, rows, total_rows, filters_applied, captured_at
        FROM kore_ai_hubspot.diud_sessions FINAL
        WHERE session_key = '{session_key}'
        ORDER BY captured_at DESC
        LIMIT 1
    """
    payload = await _ch_execute_raw(select_sql)
    if not payload:
        return None
    row = payload[0] if isinstance(payload, list) and payload else None
    if not row:
        return None
    return QueryResult(
        sql             = row["sql"],
        columns         = json.loads(row["columns"]),
        rows            = json.loads(row["rows"]),
        total_rows      = int(row["total_rows"]),
        captured_at     = str(row["captured_at"]),
        filters_applied = row["filters_applied"],
    )


def _run_sync(coro):
    """
    Bridge for calling the async ClickHouse-backed query result store
    (_get_result / _store_result / _ch_execute_raw-based helpers) from
    synchronous code paths.

    Safe here because every caller is either a `def` (non-async) FastAPI
    route — which FastAPI runs in a worker thread with no event loop
    active — or a sync helper called from within one of those threads.

    Do NOT call this from inside an `async def` route; use `await`
    on the coroutine directly there instead.
    """
    return asyncio.run(coro)


async def _count_active_sessions() -> int:
    """Replaces the old len(_SESSION_STORE) in-memory count."""
    payload = await _ch_execute_raw(
        "SELECT count(DISTINCT session_key) AS cnt FROM kore_ai_hubspot.diud_sessions"
    )
    if not payload:
        return 0
    row = payload[0] if isinstance(payload, list) and payload else None
    return int(row["cnt"]) if row else 0


async def _cleanup_sessions(max_age_days: int = 14) -> None:
    """
    Deletes session rows older than max_age_days.
    Wrapped defensively — a cleanup failure should never block startup.
    """
    try:
        await _ch_execute_raw(f"""
            ALTER TABLE kore_ai_hubspot.diud_sessions
            DELETE WHERE captured_at < now() - INTERVAL {max_age_days} DAY
        """)
    except Exception as e:
        print(f"⚠️  Session cleanup failed (non-fatal): {e}")


# =============================================================================
# [03] FastAPI App
# =============================================================================
app = FastAPI(title="DIUD", description="Decision Intelligence Using Data", version="4.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# =============================================================================
# [04] Claude Client
# =============================================================================
_ai_client    = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
_CLAUDE_MODEL = "claude-sonnet-4-6"
ALLOWED_MODELS = {
    "sonnet": "claude-sonnet-4-6",
    "opus":   "claude-opus-4-6",
}

CHAT_MAX_TOKENS   = 6144
EXPORT_MAX_TOKENS = 6144
CLICKHOUSE_TIMEOUT_SECONDS = 60


# =============================================================================
# [05] ClickHouse Proxy — Low-Level HTTP Helpers
# =============================================================================
def _base_url() -> str:
    return (os.getenv("CLICKHOUSE_API_URL") or "").rstrip("/")

def _token() -> str:
    return os.getenv("CLICKHOUSE_API_TOKEN") or ""

def _auth_headers() -> dict:
    return {
        "Authorization": f"Bearer {_token()}",
        "Content-Type":  "application/json",
    }

FORBIDDEN_KEYWORDS = ["INSERT", "UPDATE", "DELETE", "DROP", "ALTER", "TRUNCATE", "CREATE"]


async def _ch_execute_raw(sql: str):
    """
    Internal-only proxy call — bypasses the SELECT/WITH-only guard that
    run_clickhouse_query() enforces. Used ONLY by the query result store
    above. Never expose this function to Claude or any user-facing code
    path.
    """
    base_url = _base_url()
    token    = _token()
    if not base_url or not token:
        return None
    async with httpx.AsyncClient() as client:
        try:
            r = await client.post(
                f"{base_url}/query",
                headers=_auth_headers(),
                json={"query": sql},
                timeout=CLICKHOUSE_TIMEOUT_SECONDS,
            )
            if r.status_code == 200:
                return r.json()
            print(f"   ⚠️  _ch_execute_raw → HTTP {r.status_code}: {r.text[:200]}")
            return None
        except Exception as e:
            print(f"   ⚠️  _ch_execute_raw failed: {e}")
            return None


async def _proxy_get(path: str) -> dict | list | None:
    base_url = _base_url()
    token    = _token()
    if not base_url or not token:
        return None
    async with httpx.AsyncClient() as client:
        try:
            r = await client.get(f"{base_url}{path}", headers=_auth_headers(), timeout=20)
            if r.status_code == 200:
                return r.json()
            print(f"   ⚠️  GET {path} → HTTP {r.status_code}: {r.text[:200]}")
            return None
        except Exception as e:
            print(f"   ⚠️  GET {path} → {e}")
            return None


# =============================================================================
# [06] Schema Discovery
# =============================================================================
_LIVE_SCHEMA: dict = {}
_SCHEMA_BLOCK: str = "Schema not yet loaded."


async def discover_schema() -> str:
    global _LIVE_SCHEMA, _SCHEMA_BLOCK

    print("🔎 Discovering schema from ClickHouse proxy…")
    databases_raw = await _proxy_get("/databases")
    if not databases_raw:
        msg = "⚠️  Could not fetch databases — check CLICKHOUSE_API_URL / CLICKHOUSE_API_TOKEN."
        print(msg); _SCHEMA_BLOCK = msg; return msg

    if isinstance(databases_raw, list) and databases_raw:
        if isinstance(databases_raw[0], str):
            databases = databases_raw
        elif isinstance(databases_raw[0], dict):
            databases = [d.get("name") or d.get("database") or list(d.values())[0] for d in databases_raw]
        else:
            databases = [str(d) for d in databases_raw]
    elif isinstance(databases_raw, dict):
        databases = databases_raw.get("data") or databases_raw.get("databases") or list(databases_raw.values())[0] if databases_raw else []
    else:
        databases = []

    SKIP_DBS = {"system", "information_schema", "INFORMATION_SCHEMA"}
    databases = [d for d in databases if d not in SKIP_DBS]
    print(f"   Databases found: {databases}")

    schema_lines, schema_dict = [], {}

    for db in databases:
        tables_raw = await _proxy_get(f"/tables/{db}")
        if not tables_raw:
            continue
        if isinstance(tables_raw, list) and tables_raw:
            tables = tables_raw if isinstance(tables_raw[0], str) else [t.get("name") or t.get("table") or list(t.values())[0] for t in tables_raw]
        elif isinstance(tables_raw, dict):
            tables = tables_raw.get("data") or tables_raw.get("tables") or list(tables_raw.values())[0] if tables_raw else []
        else:
            tables = []

        print(f"   {db}: tables = {tables}")
        for tbl in tables:
            schema_raw = await _proxy_get(f"/schema/{db}/{tbl}")
            if not schema_raw:
                schema_lines.append(f"\nTABLE: {db}.{tbl}\n  (schema unavailable)")
                continue
            if isinstance(schema_raw, list):
                cols = schema_raw
            elif isinstance(schema_raw, dict):
                cols = schema_raw.get("columns") or schema_raw.get("data") or schema_raw.get("schema") or [schema_raw]
            else:
                cols = []

            schema_dict[f"{db}.{tbl}"] = cols
            col_lines = []
            for col in cols:
                if isinstance(col, dict):
                    col_name    = col.get("name") or col.get("column_name") or col.get("Field") or list(col.keys())[0]
                    col_type    = col.get("type") or col.get("data_type") or col.get("Type") or ""
                    col_comment = col.get("comment") or col.get("Comment") or ""
                    col_lines.append(f"  {col_name:<35} {col_type}" + (f"  — {col_comment}" if col_comment else ""))
                else:
                    col_lines.append(f"  {col}")
            schema_lines.append(f"\nTABLE: {db}.{tbl}")
            schema_lines.extend(col_lines)

    _LIVE_SCHEMA  = schema_dict
    _SCHEMA_BLOCK = "\n".join(schema_lines) if schema_lines else "No tables found."
    print(f"✅ Schema loaded: {list(schema_dict.keys())}")
    return _SCHEMA_BLOCK


# =============================================================================
# [07] System Prompt Builder
# =============================================================================
from system_prompt_template import (
    CORE_SECTION, PATTERN_A_SECTION, PATTERN_B_SECTION, PATTERN_C_SECTION,
    MQL_SECTION, DASHBOARD_SECTION, SCHEMA_FOOTER,
)

_PATTERN_MODULES = {"A": PATTERN_A_SECTION, "B": PATTERN_B_SECTION, "C": PATTERN_C_SECTION}

def _build_system_prompt(pattern: Optional[str] = None, needs_mql: bool = False,
                          needs_dashboards: bool = False) -> str:
    compact_lines = []
    for table_key, cols in _LIVE_SCHEMA.items():
        col_parts = []
        for col in cols:
            if isinstance(col, dict):
                name = col.get("name") or col.get("column_name") or col.get("Field") or list(col.keys())[0]
                typ  = col.get("type") or col.get("data_type") or col.get("Type") or ""
                col_parts.append(f"{name}:{typ}")
            else:
                col_parts.append(str(col))
        compact_lines.append(f"{table_key}({', '.join(col_parts)})")

    schema = "\n".join(compact_lines) or "Schema not yet loaded."
    if len(schema) > 20000:
        schema = schema[:20000] + "\n[schema truncated]"

    parts = [CORE_SECTION]
    if pattern in _PATTERN_MODULES:
        parts.append(_PATTERN_MODULES[pattern])
    else:
        # unknown/ambiguous pattern — load all three rather than guess wrong
        parts += [PATTERN_A_SECTION, PATTERN_B_SECTION, PATTERN_C_SECTION]
    if needs_mql:
        parts.append(MQL_SECTION)
    if needs_dashboards:
        parts.append(DASHBOARD_SECTION)
    parts.append(SCHEMA_FOOTER)

    return "\n".join(parts).format(schema=schema)


# -----------------------------------------------------------------------------
# NARRATIVE SYSTEM PROMPT — Stage 2/3 only.
#
# Deliberately NOT the same as CORE_SECTION. CORE_SECTION is full of
# SQL-generation, tool-usage, greeting, and export-intent instructions
# that don't apply here — the narrative call has no tools and is only
# ever handed data that's already been queried and validated. Feeding it
# CORE_SECTION would risk confusing it into thinking it should ask for a
# tool, handle greetings, or emit __EXPORT_INTENT__, none of which apply
# once Stage 1 has already routed past those cases.
# -----------------------------------------------------------------------------
NARRATIVE_SYSTEM_PROMPT = """
You are DIUD's narrative writer. You turn a fixed block of query result
data into a clear, executive-grade written answer. You have NO tools and
NO access to the database — the data block given to you in the user
message is the ONLY source of truth you may use.

RULES:
1. Never invent, estimate, or recall a number not literally present in
   the data block below. If the data doesn't answer part of the
   question, say so plainly instead of guessing.
2. Lead with the direct answer to the question, then supporting detail.
3. Use clean markdown. Tables for tabular data. Bold key numbers.
4. Format dollar amounts consistently with how they appear in the data
   (e.g. keep $M scale if the data is already in millions).
5. Do not mention SQL, queries, ClickHouse, or that this came from "a
   database" — write as if you already knew the answer.
6. Do NOT write Chart.js, SVG, or any ```html chart code — charts are
   generated separately by the system.
7. Do NOT append a "Filters Applied" section — the system appends that
   automatically after your response.
"""


# =============================================================================
# [08] SQL Execution
# =============================================================================
def run_clickhouse_query(sql: str, session_id: Optional[str] = None) -> str:
    base_url = _base_url()
    token    = _token()

    if not base_url:
        return "DATABASE CONNECTION FAILED: CLICKHOUSE_API_URL is not set."
    if not token:
        return "DATABASE CONNECTION FAILED: CLICKHOUSE_API_TOKEN is not set."

    cleaned = re.sub(r'^\s*--[^\n]*\n', '', sql.strip(), flags=re.MULTILINE).strip()
    stripped = cleaned.upper()
    if not (stripped.startswith("SELECT") or stripped.startswith("WITH")):
        return "ERROR: Only SELECT/WITH queries are permitted."
    for kw in FORBIDDEN_KEYWORDS:
        if re.search(rf'\b{kw}\b', stripped):
            return f"ERROR: Forbidden keyword: {kw}"

    print(f"🔍 SQL (session={session_id}) → {sql[:200]}")

    try:
        resp = httpx.post(
            f"{base_url}/query",
            headers=_auth_headers(),
            json={"query": sql},
            timeout=60,
        )

        if resp.status_code == 401:
            return "DATABASE CONNECTION FAILED: 401 Unauthorized."
        if resp.status_code == 403:
            return "DATABASE CONNECTION FAILED: 403 Forbidden."
        if resp.status_code == 422:
            return f"ERROR: Proxy rejected query (422): {resp.text[:400]}"
        if resp.status_code == 500:
            return f"DATABASE ERROR: HTTP 500: {resp.text[:400]}"
        if resp.status_code != 200:
            return f"DATABASE ERROR: HTTP {resp.status_code} — {resp.text[:300]}"

        payload = resp.json()

        if isinstance(payload, list):
            rows = payload
        elif isinstance(payload, dict):
            rows = payload.get("data") or payload.get("rows") or payload.get("result") or payload.get("results")
            api_columns = payload.get("columns") or payload.get("meta") or payload.get("column_names")
            if rows is None:
                return json.dumps(payload, indent=2, default=str)[:3000]
        else:
            return f"Unexpected response type: {type(payload)}"

        if not rows:
            return "Query returned 0 rows."

        if isinstance(rows[0], dict):
            columns = list(rows[0].keys())
            norm_rows = rows
        else:
            if api_columns and len(api_columns) == len(rows[0]):
                columns = [c["name"] if isinstance(c, dict) else c for c in api_columns]
            else:
                columns = [f"col_{i}" for i in range(len(rows[0]))]
            norm_rows = [dict(zip(columns, r)) for r in rows]

        total_rows = len(norm_rows)

        if session_id:
            existing = _run_sync(_get_result(session_id))
            _run_sync(_store_result(f"{session_id}:latest", QueryResult(
                sql             = sql,
                columns         = columns,
                rows            = norm_rows,
                total_rows      = total_rows,
                captured_at     = datetime.utcnow().isoformat() + "Z",
                filters_applied = _extract_filters_from_sql(sql),
            )))
            if not existing or total_rows > existing.total_rows:
                _run_sync(_store_result(session_id, QueryResult(
                    sql             = sql,
                    columns         = columns,
                    rows            = norm_rows,
                    total_rows      = total_rows,
                    captured_at     = datetime.utcnow().isoformat() + "Z",
                    filters_applied = _extract_filters_from_sql(sql),
                )))
            else:
                print(f"   ⏭️  Skipping EXPORT session store update — {total_rows} row(s) won't overwrite existing {existing.total_rows} rows.")

        CHAT_DISPLAY_LIMIT = 100
        header = " | ".join(columns)
        lines  = [header, "-" * min(len(header), 140)]
        for row in norm_rows[:CHAT_DISPLAY_LIMIT]:
            lines.append(" | ".join(str(row.get(c, "")) for c in columns))

        if total_rows > CHAT_DISPLAY_LIMIT:
            lines.append(
                f"\n📊 **Showing {CHAT_DISPLAY_LIMIT} of {total_rows} rows.** "
                f"Say **\"export these deals as CSV\"** or **\"export as PDF\"** "
                f"to download all {total_rows} records."
            )

        result = "\n".join(lines)
        print(f"   ✅ {total_rows} rows returned. Session store updated.")
        return result

    except httpx.ConnectError as e:
        return f"DATABASE CONNECTION FAILED: Could not reach {base_url}. {e}"
    except httpx.TimeoutException:
        return "DATABASE CONNECTION FAILED: Query timed out after 60 seconds."
    except Exception as exc:
        traceback.print_exc()
        return f"DATABASE CONNECTION FAILED: {type(exc).__name__}: {exc}"


def _extract_filters_from_sql(sql: str) -> str:
    sql_upper = sql.upper()
    filters = []
    if "PIPELINE = 'DEFAULT'" in sql_upper:
        filters.append("Pipeline: default")

    cohort_stages = sorted(set(re.findall(r'BECAME_(\d+)_DEAL_DATE', sql_upper)), key=int)
    for stage in cohort_stages:
        filters.append(f"Cohort: {stage}% qualified deals")

    if "CLOSE_DATE" in sql_upper and "2026-04-01" in sql:
        filters.append("FY27 active pipeline")
    if "DEAL_STAGE" in sql_upper:
        m = re.search(r"deal_stage\s+IN\s*\(([^)]+)\)", sql, re.IGNORECASE)
        if m:
            filters.append(f"Stage filter: {m.group(1)[:60]}")
    if "REGION" in sql_upper:
        m = re.search(r"region\s*=\s*'([^']+)'", sql, re.IGNORECASE)
        if m:
            filters.append(f"Region: {m.group(1)}")
    if "COMPANY_PRIORITY" in sql_upper:
        filters.append("Company Priority: P1–P7")
    if "LEAD_STATUS" in sql_upper and "BAD DATA" in sql_upper:
        filters.append("Lead Status: excludes 'Bad Data'")
    if "LIFECYCLE_STAGE" in sql_upper and "MARKETINGQUALIFIEDLEAD" in sql_upper:
        filters.append("Lifecycle: MQL only")
    return "; ".join(filters) if filters else "Standard base filters applied"


# =============================================================================
# [09] Validation
# =============================================================================
def _is_numericish(v) -> bool:
    try:
        float(v)
        return True
    except (TypeError, ValueError):
        return False


def _to_float(v) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _compute_deterministic_headline(stored: Optional["QueryResult"]) -> Optional[str]:
    """For small aggregate results, compute totals/averages in Python so Claude
    never has to do the arithmetic itself when writing the summary."""
    if not stored or not stored.rows or len(stored.rows) > 20:
        return None

    numeric_cols = [
        c for c in stored.columns
        if any(_is_numericish(r.get(c)) for r in stored.rows)
    ]
    if not numeric_cols:
        return None

    parts = []
    for col in numeric_cols:
        total = sum(_to_float(r.get(col)) for r in stored.rows)
        parts.append(f"{col}_total={total:,.2f}")
        if len(stored.rows) > 1:
            avg = total / len(stored.rows)
            parts.append(f"{col}_avg={avg:,.2f}")
    return " | ".join(parts)


def _parse_rows_for_validation(query_result: str, session_id: Optional[str]) -> List[dict]:
    """
    run_clickhouse_query() returns a human-readable pipe-delimited text table,
    not JSON, so we can't parse query_result directly. Instead, pull the
    structured rows already stored in the ClickHouse-backed query result store.
    """
    if not session_id:
        return []
    stored = _run_sync(_get_result(session_id))
    if not stored:
        return []
    return stored.rows


# =============================================================================
# [10] Startup
# =============================================================================
@app.on_event("startup")
async def on_startup():
    try:
        # discover_schema() is async — must be awaited or _LIVE_SCHEMA
        # stays permanently empty.
        await discover_schema()
    except Exception as e:
        print(f"⚠️  Schema discovery failed: {e}")
    await _cleanup_sessions()
    print("🚀 DIUD v4 started — session-store export enabled.")


# =============================================================================
# [11] Claude Tool Definitions
# =============================================================================
_QUERY_TOOL = {
    "name": "query_clickhouse",
    "description": (
        "Execute a SELECT query against ClickHouse. "
        "Use for deal pipeline, AE performance, win/loss, regions, stages, MQL metrics. "
        "ALWAYS use fully-qualified table names. "
        "Relay DATABASE CONNECTION FAILED errors directly to the user."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "sql": {
                "type": "string",
                "description": (
                    "Valid ClickHouse SELECT or WITH query. "
                    "For deal LIST queries: NO LIMIT unless user asks for 'top N'. "
                    "Return all matching rows — the system displays them safely."
                ),
            }
        },
        "required": ["sql"],
    },
}

_RULEBOOK_TOOL = {
    "name": "lookup_business_rule",
    "description": (
        "Look up the exact business rule for a specific metric/topic BEFORE "
        "writing SQL for it. ALWAYS call this before generating SQL for MQL, "
        "active pipeline, cohort funnels, attainment/targets, closed won, "
        "partner targets, or dashboard-specific questions."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "topic": {
                "type": "string",
                "enum": [
                    "mql", "active_pipeline", "cohort_funnel", "attainment",
                    "closed_won", "partner_targets", "dashboard_definitions",
                ],
            }
        },
        "required": ["topic"],
    },
}


# =============================================================================
# [12] Session / Pydantic Models
# =============================================================================
class ChatMessage(BaseModel):
    role: Literal["user", "assistant"]
    content: str

class ChatRequest(BaseModel):
    message: str
    history: List[ChatMessage] = []
    session_id: Optional[str] = None
    model: str = "sonnet"

class ExportPreviewRequest(BaseModel):
    conversation: List[ChatMessage] = []
    title: str = "Pipeline Intelligence Report"
    export_type: Literal["pdf", "pptx"] = "pdf"
    detail_level: Literal["summary", "detailed"] = "detailed"
    session_id: Optional[str] = None

class ExportDownloadRequest(BaseModel):
    format: Literal["pdf", "pptx", "csv"]
    content: Optional[str] = None
    title: str = "Pipeline Intelligence Report"
    session_id: Optional[str] = None

class RetryRequest(BaseModel):
    history:    List[ChatMessage] = []
    session_id: Optional[str] = None
    model:      str = "sonnet"


def _extract_text(content_blocks) -> str:
    return "\n".join(
        b.text for b in content_blocks if hasattr(b, "text") and b.text
    ).strip()


def _normalize_messages(messages: list) -> list:
    """Shared message-flattening logic used by the Stage 1 SQL generation stage."""
    safe_messages = []
    for m in messages:
        content = m.get("content", "")
        if isinstance(content, list):
            text_parts = [
                b.get("text", "") if isinstance(b, dict) else (b.text if hasattr(b, "text") else "")
                for b in content if (isinstance(b, dict) and b.get("type") == "text")
                   or (hasattr(b, "type") and b.type == "text")
            ]
            text = "\n".join(t for t in text_parts if t).strip()
            if text:
                safe_messages.append({"role": m["role"], "content": text})
        else:
            safe_messages.append({"role": m["role"], "content": content})
    return safe_messages


def _last_user_message(messages: list) -> str:
    return next(
        (m.get("content", "") for m in reversed(messages)
         if m.get("role") == "user" and isinstance(m.get("content", ""), str)),
        ""
    )


# =============================================================================
# [13] Pipeline — Stage 1: Intent & SQL Generation
# =============================================================================

_SQL_STAGE_TOOLS      = [_QUERY_TOOL, _RULEBOOK_TOOL]
_SQL_STAGE_MAX_ROUNDS  = 6
_SQL_STAGE_MAX_TOKENS  = 2048   # SQL + rulebook text only, no narrative — small budget is deliberate


class SqlStageResult:
    """
    Structured output of Stage 1. Exactly one of query_result /
    conversational_reply is set, matching `stopped_reason`.
    """
    def __init__(self):
        self.query_result:          Optional[QueryResult] = None
        self.conversational_reply:  Optional[str] = None
        self.safe_messages:         list = []
        self.last_error:            Optional[str] = None
        self.stopped_reason:        str = ""   # "clean_result" | "conversational" | "max_rounds"
        self.intent_pattern:        Optional[str] = None


def _resolve_intent(user_message: str) -> dict:
    """
    Wraps get_intent() so a missing/malformed return value never crashes
    the request — falls back to "load everything" (safe default).

    ASSUMED SHAPE — not yet confirmed against rules.py:
        {"pattern": "A" | "B" | "C" | None,
         "needs_mql": bool, "needs_dashboards": bool}

    If rules.py's real return shape differs, this silently falls back to
    the safe default rather than crashing the request — but the
    classifier effectively won't be doing anything until this is fixed.
    """
    try:
        intent = get_intent(user_message) or {}
    except Exception as e:
        print(f"⚠️ get_intent() failed, falling back to full prompt: {e}")
        intent = {}
    return {
        "pattern":          intent.get("pattern"),
        "needs_mql":        bool(intent.get("needs_mql", False)),
        "needs_dashboards": bool(intent.get("needs_dashboards", False)),
    }


def _run_sql_generation_stage(
    messages: list,
    session_id: str,
    model: str = "sonnet",
) -> SqlStageResult:
    """
    Stage 1: Intent & SQL.

    Tools available: query_clickhouse, lookup_business_rule ONLY.
    Stops the moment one of these is true:
      - a query_clickhouse call returns a clean (non-violating,
        non-error) result           -> stopped_reason = "clean_result"
      - Claude answers without ever calling query_clickhouse
        (e.g. "what can you help with?", greetings, export requests)
                                     -> stopped_reason = "conversational"
      - _SQL_STAGE_MAX_ROUNDS exhausted -> stopped_reason = "max_rounds"

    Any prose Claude writes alongside a tool call in this stage is
    discarded on purpose — narrative is not this stage's job.

    `session_id` is required (non-optional) here: the "clean_result"
    stop condition reads back through the session store, so a caller
    without a real session_id would never be able to stop cleanly.
    The orchestrator (_run_pipeline) guarantees one is always generated.
    """
    selected_model = ALLOWED_MODELS.get(model, ALLOWED_MODELS["sonnet"])
    result = SqlStageResult()

    original_user_message = _last_user_message(messages)

    intent = _resolve_intent(original_user_message)
    result.intent_pattern = intent["pattern"]
    system_prompt = _build_system_prompt(
        pattern          = intent["pattern"],
        needs_mql        = intent["needs_mql"],
        needs_dashboards = intent["needs_dashboards"],
    )

    safe_messages = _normalize_messages(messages)

    response = _ai_client.messages.create(
        model=selected_model, system=system_prompt, messages=safe_messages,
        tools=_SQL_STAGE_TOOLS, temperature=0, max_tokens=_SQL_STAGE_MAX_TOKENS,
    )

    for round_num in range(_SQL_STAGE_MAX_ROUNDS):
        if response.stop_reason != "tool_use":
            result.stopped_reason = "conversational"
            result.conversational_reply = _extract_text(response.content)
            result.safe_messages = safe_messages
            return result

        tool_blocks = [b for b in response.content if b.type == "tool_use"]
        if not tool_blocks:
            result.stopped_reason = "conversational"
            result.conversational_reply = _extract_text(response.content)
            result.safe_messages = safe_messages
            return result

        tool_result_blocks = []
        clean_result_this_round: Optional[QueryResult] = None

        for tool_block in tool_blocks:
            if tool_block.name == "lookup_business_rule":
                tool_result_blocks.append({
                    "type": "tool_result", "tool_use_id": tool_block.id,
                    "content": get_rulebook_entry(tool_block.input.get("topic", "")),
                    "is_error": False,
                })
                continue

            sql = tool_block.input.get("sql", "")
            sql_violations = validate_sql_against_rules(sql, original_user_message)
            _log_rule_audit(session_id, sql, sql_violations, "pre_execute", original_user_message)

            if sql_violations:
                cohort_violations = [v for v in sql_violations if "cohort" in v.lower() or "8b" in v.lower()]
                cast_violations = [v for v in sql_violations if "target_table_float_cast" in v.lower() or "toFloat64OrZero" in v]
                if cohort_violations:
                    query_result_text = (
                        "RULE VIOLATION — cohort funnel SQL rejected (§8b). NOT executed.\n\n"
                        "Violations:\n" + "\n".join(f"- {v}" for v in sql_violations)
                        + "\n\n⚠️ Before rewriting, call lookup_business_rule('cohort_funnel') to get "
                        "the exact §8b template (sentinel anchor, single WITH-cohort CTE, "
                        "countDistinct(deal_id), GROUP BY deal_stage). Do NOT reuse the Pattern A "
                        "OR-chain approach or IS NOT NULL checks for this query — this is a true "
                        "cohort exclusion query, not cumulative stage counting. Then resubmit."
                    )
                elif cast_violations:
                    query_result_text = (
                        "RULE VIOLATION — query rejected, NOT executed against the database:\n"
                        + "\n".join(f"- {v}" for v in sql_violations)
                        + "\n\n⚠️ Check EVERY target-table column referenced in this query, not just "
                        "the one named above — a query often sums multiple target columns "
                        "(e.g. amount_target_20 AND deals_target_20), and each one independently "
                        "needs its own SUM(toFloat64OrZero(col)) wrapper. Rewrite the SQL to satisfy "
                        "these rules, then call the tool again."
                    )
                else:
                    query_result_text = (
                        "RULE VIOLATION — query rejected, NOT executed against the database:\n"
                        + "\n".join(f"- {v}" for v in sql_violations)
                        + "\n\nRewrite the SQL to satisfy these rules, then call the tool again."
                    )
                is_error = True
            else:
                query_result_text = run_clickhouse_query(sql, session_id=session_id)
                is_error = any(query_result_text.startswith(p) for p in [
                    "DATABASE CONNECTION FAILED", "ERROR:", "DATABASE ERROR:"
                ])
                if not is_error:
                    parsed_rows = _parse_rows_for_validation(query_result_text, session_id)
                    result_violations = validate_result_against_rules(parsed_rows, original_user_message, sql)
                    _log_rule_audit(session_id, sql, result_violations, "post_execute", original_user_message)
                    if result_violations:
                        query_result_text += (
                            "\n\n⚠️ RESULT RULE VIOLATION:\n"
                            + "\n".join(f"- {v}" for v in result_violations)
                            + "\nRe-derive using the single-CTE cohort pattern from §8b and call the tool again."
                        )
                        is_error = True
                    else:
                        clean_result_this_round = _run_sync(_get_result(f"{session_id}:latest"))

            if is_error:
                result.last_error = query_result_text

            tool_result_blocks.append({
                "type": "tool_result", "tool_use_id": tool_block.id,
                "content": query_result_text, "is_error": is_error,
            })

        safe_messages = safe_messages + [
            {"role": "assistant", "content": response.content},
            {"role": "user", "content": tool_result_blocks},
        ]

        if clean_result_this_round is not None:
            result.stopped_reason = "clean_result"
            result.query_result   = clean_result_this_round
            result.safe_messages  = safe_messages
            return result

        is_last_round = (round_num == _SQL_STAGE_MAX_ROUNDS - 1)
        response = _ai_client.messages.create(
            model=selected_model, system=system_prompt, messages=safe_messages,
            tools=[] if is_last_round else _SQL_STAGE_TOOLS,
            temperature=0, max_tokens=_SQL_STAGE_MAX_TOKENS,
        )

    result.stopped_reason = "max_rounds"
    result.safe_messages = safe_messages
    return result


# =============================================================================
# [14] Pipeline — Stage 2: Data Processing
# =============================================================================

class FactEnvelope:
    """
    Single source of truth for one query result. Chart Builder and
    Summary Generator both read ONLY from this — that's what keeps them
    synchronized (fixes "charts sometimes don't match the narrative").
    """
    def __init__(self, query_result: QueryResult, verified_headline: Optional[str]):
        self.sql              = query_result.sql
        self.columns          = query_result.columns
        self.rows             = query_result.rows
        self.total_rows       = query_result.total_rows
        self.filters_applied  = query_result.filters_applied
        self.captured_at      = query_result.captured_at
        self.verified_headline = verified_headline   # deterministic totals/avgs, or None


def _build_fact_envelope(query_result: QueryResult) -> FactEnvelope:
    """Data Normalizer — wraps Stage 1's QueryResult into the shared envelope."""
    headline = _compute_deterministic_headline(query_result)
    return FactEnvelope(query_result, headline)


def _generate_chart(envelope: FactEnvelope) -> Optional[str]:
    """Chart Builder — reads only from the Fact Envelope."""
    if not envelope.rows:
        return None
    return build_chart_html(envelope.columns, envelope.rows, envelope.filters_applied)


_SUMMARY_MAX_TOKENS = CHAT_MAX_TOKENS


def _envelope_to_facts_block(envelope: FactEnvelope, row_cap: int = 100) -> str:
    """The ONLY text the Summary Generator is allowed to pull numbers from."""
    header = " | ".join(envelope.columns)
    lines  = [header, "-" * min(len(header), 140)]
    for row in envelope.rows[:row_cap]:
        lines.append(" | ".join(str(row.get(c, "")) for c in envelope.columns))
    if envelope.total_rows > row_cap:
        lines.append(f"... ({envelope.total_rows} total rows, {row_cap} shown)")
    block = "\n".join(lines)
    if envelope.verified_headline:
        block += f"\n\n[VERIFIED TOTALS — use these exact figures, do not recompute]: {envelope.verified_headline}"
    return block


def _generate_summary(user_question: str, envelope: FactEnvelope, model: str = "sonnet") -> str:
    """
    Summary Generator — Stage 2's Claude call. NO tools. Its only input
    is the Fact Envelope; it cannot query ClickHouse, so it cannot drift
    from what Chart Builder is drawing from the same envelope.
    """
    selected_model = ALLOWED_MODELS.get(model, ALLOWED_MODELS["sonnet"])
    facts_block = _envelope_to_facts_block(envelope)

    prompt = f"""The user asked: "{user_question}"

Here is the ONLY data you have access to. Do not invent, estimate, or
recall any figures not literally present below.

{facts_block}

Write a clear, concise narrative answer using ONLY the data above.
Lead with the direct answer, then supporting context."""

    response = _ai_client.messages.create(
        model=selected_model,
        system=NARRATIVE_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": prompt}],
        tools=[], temperature=0, max_tokens=_SUMMARY_MAX_TOKENS,
    )
    reply = _extract_text(response.content)

    if response.stop_reason == "max_tokens":
        try:
            cont = _ai_client.messages.create(
                model=selected_model, system=NARRATIVE_SYSTEM_PROMPT,
                messages=[
                    {"role": "user", "content": prompt},
                    {"role": "assistant", "content": reply},
                    {"role": "user", "content": "Continue directly from where you left off. Do not repeat text."},
                ],
                tools=[], temperature=0, max_tokens=_SUMMARY_MAX_TOKENS,
            )
            reply = (reply + " " + _extract_text(cont.content)).strip()
        except Exception as e:
            print(f"⚠️ Summary continuation failed: {e}")

    return reply


# =============================================================================
# [15] Pipeline — Stage 3: Response Assembly
# =============================================================================

def _fact_bind_summary(
    user_question: str, summary_text: str, envelope: FactEnvelope,
    session_id: Optional[str], model: str = "sonnet",
) -> str:
    """
    Fact Binder — checks the Summary Generator's output against the
    Fact Envelope. Regenerates once, strictly, if numbers don't match.
    """
    violations = validate_summary_against_facts(summary_text, envelope.rows)
    if not violations:
        return summary_text

    _log_rule_audit(session_id, "summary_check", violations, "post_summary", user_question)
    print(f"⚠️ Unverified numbers in summary: {violations}")

    selected_model = ALLOWED_MODELS.get(model, ALLOWED_MODELS["sonnet"])
    facts_block = _envelope_to_facts_block(envelope)
    retry_prompt = f"""Your previous answer contained numbers that don't match the data.

DATA (ONLY SOURCE OF TRUTH):
{facts_block}

Rewrite the answer to "{user_question}" using ONLY numbers that literally
appear above. This rewrite is the ONLY version the user will see — do not
reference a previous draft or apologize."""

    try:
        retry = _ai_client.messages.create(
            model=selected_model, system=NARRATIVE_SYSTEM_PROMPT,
            messages=[{"role": "user", "content": retry_prompt}],
            tools=[], temperature=0, max_tokens=_SUMMARY_MAX_TOKENS,
        )
        return _extract_text(retry.content) or summary_text
    except Exception as e:
        print(f"⚠️ Fact-binding regeneration failed: {e}")
        return summary_text


_PATTERN_LABELS = {
    "A": "A — Cumulative Funnel",
    "B": "B — Deal Detail",
    "C": "C — Attainment",
}

def _build_filters_footer(envelope: FactEnvelope, intent_pattern: Optional[str]) -> str:
    """
    Deterministic replacement for the old LLM-written "Filters Applied"
    footer (system_prompt_template.py §14). Metadata like this should be
    generated from what actually ran, not re-derived by a model each
    time — keeps it accurate by construction instead of by prompting.
    """
    pattern_label = _PATTERN_LABELS.get(intent_pattern, "D — General/Ad-hoc")
    return (
        "\n\n---\n**Filters Applied:**\n"
        f"- Pattern used: {pattern_label}\n"
        f"- {envelope.filters_applied}\n"
        f"- Total rows: {envelope.total_rows:,}\n\n"
        "Please verify these filters match your expectation.\n"
        "---"
    )


def _build_response(summary_text: str, chart_html: Optional[str], filters_footer: str) -> str:
    """Response Builder — assembles narrative + footer + chart into the Final Response."""
    reply = summary_text.rstrip() + filters_footer
    if chart_html and not reply_already_has_chart(reply):
        reply += "\n\n" + chart_html
    return reply


# =============================================================================
# [16] Pipeline Orchestrator
# =============================================================================

def _run_pipeline(messages: list, session_id: Optional[str] = None, model: str = "sonnet") -> str:
    # Stage 1's "clean_result" stop condition depends on a real session_id
    # to read back through the store. If the caller didn't supply one,
    # generate an ephemeral one for this request only.
    session_id = session_id or f"anon-{uuid.uuid4().hex[:12]}"

    original_user_message = _last_user_message(messages)

    stage1 = _run_sql_generation_stage(messages, session_id=session_id, model=model)

    if stage1.stopped_reason == "conversational":
        return stage1.conversational_reply or "⚠️ I wasn't able to answer that. Could you rephrase the question?"

    if stage1.stopped_reason == "max_rounds" or not stage1.query_result:
        if stage1.last_error:
            return (
                "⚠️ I couldn't complete this query. The last database error was:\n\n"
                f"`{stage1.last_error[:400]}`\n\nCould you rephrase, or check **/debug/db**?"
            )
        return "⚠️ I wasn't able to finish this in time — could you try a simpler or more specific question?"

    # Stage 2
    envelope   = _build_fact_envelope(stage1.query_result)
    chart_html = _generate_chart(envelope)
    summary    = _generate_summary(original_user_message, envelope, model=model)

    # Stage 3
    bound_summary  = _fact_bind_summary(original_user_message, summary, envelope, session_id, model=model)
    filters_footer = _build_filters_footer(envelope, stage1.intent_pattern)
    return _build_response(bound_summary, chart_html, filters_footer)


# =============================================================================
# [17] Routes — Chat
# =============================================================================
@app.get("/", response_class=HTMLResponse)
def root():
    with open("chat.html", "r") as f:
        return HTMLResponse(content=f.read())

@app.get("/logo.png")
def serve_logo():
    return FileResponse("logo.png", media_type="image/png")


@app.get("/debug/db")
def debug_db():
    base_url = _base_url()
    token    = _token()
    config = {
        "CLICKHOUSE_API_URL":   base_url or "❌ NOT SET",
        "CLICKHOUSE_API_TOKEN": f"✅ set ({len(token)} chars)" if token else "❌ NOT SET",
    }
    if not base_url or not token:
        return {"status": "MISCONFIGURED", "config": config}

    tests = {}
    try:
        r = httpx.get(base_url, timeout=10)
        tests["GET /"] = {"status": r.status_code}
    except Exception as e:
        tests["GET /"] = {"error": str(e)}

    ping = run_clickhouse_query("SELECT 1 AS ping")
    query_ok = not any(ping.startswith(p) for p in ["DATABASE CONNECTION FAILED", "ERROR:", "DATABASE ERROR:"])
    tests["SELECT 1"] = {"ok": query_ok, "result": ping[:100]}

    return {
        "status":            "OK" if query_ok else "FAILED",
        "config":            config,
        "discovered_tables": list(_LIVE_SCHEMA.keys()),
        "active_sessions":   _run_sync(_count_active_sessions()),
        "tests":             tests,
    }


@app.post("/refresh-schema")
async def refresh_schema():
    schema = await discover_schema()
    return {"status": "refreshed", "tables": list(_LIVE_SCHEMA.keys())}


@app.post("/chat")
def chat(payload: ChatRequest):
    messages = [{"role": m.role, "content": m.content} for m in payload.history]
    messages.append({"role": "user", "content": payload.message})
    print(f"💬 [chat] session={payload.session_id} msg={payload.message[:80]}")

    try:
        reply = _run_pipeline(messages, session_id=payload.session_id, model=payload.model)
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=502, detail=f"Claude error: {exc}")

    stored      = _run_sync(_get_result(payload.session_id)) if payload.session_id else None
    has_dataset = stored is not None

    return {
        "reply":        reply,
        "has_dataset":  has_dataset,
        "dataset_rows": stored.total_rows if stored else 0,
        "export_intent": "__EXPORT_INTENT__" in reply,
    }


# =============================================================================
# [18] Routes — Retry
# =============================================================================
@app.post("/chat/retry")
def chat_retry(payload: RetryRequest):
    if not payload.history:
        raise HTTPException(status_code=400, detail="history must not be empty for retry.")

    clean_history = list(payload.history)
    while clean_history and clean_history[-1].role == "assistant":
        clean_history.pop()

    if not clean_history or clean_history[-1].role != "user":
        raise HTTPException(
            status_code=400,
            detail="No user message found to retry. history must end with a user turn."
        )

    last_user_msg = clean_history[-1].content
    prior_history = clean_history[:-1]

    print(f"🔄 [retry] session={payload.session_id} retrying: {last_user_msg[:80]}")

    messages = [{"role": m.role, "content": m.content} for m in prior_history]
    messages.append({"role": "user", "content": last_user_msg})

    try:
        reply = _run_pipeline(messages, session_id=payload.session_id, model=payload.model)
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=502, detail=f"Claude error on retry: {exc}")

    stored      = _run_sync(_get_result(payload.session_id)) if payload.session_id else None
    has_dataset = stored is not None

    return {
        "reply":         reply,
        "has_dataset":   has_dataset,
        "dataset_rows":  stored.total_rows if stored else 0,
        "export_intent": "__EXPORT_INTENT__" in reply,
        "retried":       True,
    }


# =============================================================================
# [19] Routes — Session Info
# =============================================================================
@app.get("/session/{session_id}/dataset-info")
def session_dataset_info(session_id: str):
    result = _run_sync(_get_result(session_id))
    if not result:
        return {"has_dataset": False}
    return {
        "has_dataset":     True,
        "total_rows":      result.total_rows,
        "columns":         result.columns,
        "captured_at":     result.captured_at,
        "filters_applied": result.filters_applied,
        "sql_preview":     result.sql[:300],
    }


# =============================================================================
# [20] Export — Preview & Download
# =============================================================================
@app.post("/export/preview")
async def export_preview(req: ExportPreviewRequest):
    if not req.conversation:
        raise HTTPException(status_code=400, detail="No conversation to export.")

    print(f"📄 [export/preview] session={req.session_id} type={req.export_type}")

    stored = await _get_result(req.session_id) if req.session_id else None

    try:
        ai_content = _generate_export_content(
            conversation   = req.conversation,
            title          = req.title,
            export_type    = req.export_type,
            detail_level   = req.detail_level,
            stored_dataset = stored,
        )
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=502, detail=f"Content generation error: {exc}")

    return {
        "content":       ai_content,
        "title":         req.title,
        "export_type":   req.export_type,
        "word_count":    len(ai_content.split()),
        "generated_at":  date.today().isoformat(),
        "total_rows":    stored.total_rows if stored else 0,
        "filters":       stored.filters_applied if stored else "",
    }


@app.post("/export/download")
async def export_download(req: ExportDownloadRequest):
    print(f"⬇️  [export/download] format={req.format} session={req.session_id}")

    if req.format == "csv":
        stored = await _get_result(req.session_id) if req.session_id else None
        if not stored:
            raise HTTPException(
                status_code=404,
                detail=(
                    "No query result found for this session. "
                    "Ask a deal-list question first, then export."
                ),
            )
        csv_bytes = _build_csv(stored)
        return StreamingResponse(
            io.BytesIO(csv_bytes),
            media_type="text/csv",
            headers={
                "Content-Disposition": f'attachment; filename="{_safe_filename(req.title)}.csv"',
                "X-Total-Rows": str(stored.total_rows),
            },
        )

    if not req.content:
        raise HTTPException(status_code=400, detail="content is required for PDF/PPTX export.")

    try:
        if req.format == "pdf":
            file_bytes = _build_pdf(req.title, req.content)
            media_type = "application/pdf"
            ext        = "pdf"
        else:
            file_bytes = _build_pptx(req.title, req.content)
            media_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            ext        = "pptx"

        return StreamingResponse(
            io.BytesIO(file_bytes),
            media_type=media_type,
            headers={"Content-Disposition": f'attachment; filename="{_safe_filename(req.title)}.{ext}"'},
        )
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"File generation error: {exc}")


# =============================================================================
# [21] Utility Functions
# =============================================================================
def _safe_filename(title: str) -> str:
    return re.sub(r'[^\w\-]', '_', title)[:60]


def _strip_md(t: str) -> str:
    t = re.sub(r'\*{1,3}(.*?)\*{1,3}', r'\1', t)
    t = re.sub(r'#{1,6}\s*', '', t)
    t = re.sub(r'^[\-\*•]\s*', '', t, flags=re.M)
    t = re.sub(r'`(.*?)`', r'\1', t)
    t = t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    return t.strip()


# =============================================================================
# [22] CSV Builder
# =============================================================================
def _build_csv(stored: QueryResult) -> bytes:
    buf = io.StringIO()
    buf.write(f"# Title: {stored.sql[:80]}\n")
    buf.write(f"# Generated: {date.today().isoformat()}\n")
    buf.write(f"# Total Records: {stored.total_rows}\n")
    buf.write(f"# Filters: {stored.filters_applied}\n")
    buf.write(f"# Captured at: {stored.captured_at}\n")
    buf.write("#\n")

    writer = csv.DictWriter(
        buf,
        fieldnames     = stored.columns,
        extrasaction   = "ignore",
        lineterminator = "\n",
    )
    writer.writeheader()
    for row in stored.rows:
        writer.writerow(row)

    return buf.getvalue().encode("utf-8")


# =============================================================================
# [23] Export Content Generation
# =============================================================================
def _render_html_to_png(html_content: str, width: int = 800) -> Optional[bytes]:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("⚠️  Playwright not installed.")
        return None

    full_html = f"""<!DOCTYPE html><html><head>
<meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ background: #ffffff; font-family: 'Inter', Arial, sans-serif; }}
</style>
</head><body>{html_content}
</body></html>"""

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                args=[
                    "--no-sandbox",
                    "--disable-setuid-sandbox",
                    "--disable-dev-shm-usage",
                    "--disable-gpu",
                    "--single-process",
                    "--no-zygote",
                ]
            )
            page = browser.new_page(viewport={"width": width, "height": 800})
            page.set_content(full_html, wait_until="networkidle", timeout=20000)
            page.wait_for_timeout(2000)
            height = page.evaluate("document.body.scrollHeight")
            page.set_viewport_size({"width": width, "height": max(height, 100)})
            png_bytes = page.screenshot(full_page=True)
            browser.close()
            return png_bytes
    except Exception as e:
        print(f"⚠️  Playwright screenshot failed: {e}")
        return None


def _extract_html_blocks(text: str):
    """
    Split a message into alternating text/html segments.
    Returns list of ("text"|"html", content) tuples.
    """
    segments = []
    last_end = 0
    for m in re.finditer(r'```html\s*\n(.*?)```', text, flags=re.DOTALL | re.IGNORECASE):
        segments.append(("text", text[last_end:m.start()]))
        segments.append(("html", m.group(1).strip()))
        last_end = m.end()
    segments.append(("text", text[last_end:]))
    return segments


def _generate_export_content(
    conversation:    List[ChatMessage],
    title:           str,
    export_type:     str,
    detail_level:    str = "detailed",
    stored_dataset:  Optional[QueryResult] = None,
) -> str:

    def _clean_text_segment(text: str) -> str:
        text = re.sub(r'```[\w]*\n.*?```', '', text, flags=re.DOTALL)
        text = re.sub(r'^\s*--\s*$', '', text, flags=re.MULTILINE)
        text = re.sub(r'^\s*`+\s*$', '', text, flags=re.MULTILINE)
        text = re.sub(r'\n{3,}', '\n\n', text)
        text = text.replace('■', '').replace('🟦', '')
        return text.strip()

    if detail_level == "detailed":
        if export_type == "pptx":
            lines = [f"# {title}", f"*Exported: {date.today().strftime('%B %d, %Y')}*", ""]
            for i, m in enumerate(conversation):
                role = "User" if m.role == "user" else "DIUD Agent"
                segments = _extract_html_blocks(m.content)
                lines.append(f"SLIDE: {role} — Turn {i + 1}")
                for seg_type, seg_content in segments:
                    if seg_type == "html":
                        lines.append("- [Interactive chart — see live DIUD interface]")
                    else:
                        cleaned = _clean_text_segment(seg_content)
                        for part in cleaned.split("\n"):
                            part = part.strip()
                            if part:
                                lines.append(f"- {part}")
            return "\n".join(lines)
        else:
            pdf_segments = []
            for i, m in enumerate(conversation):
                role = "User" if m.role == "user" else "DIUD Agent"
                pdf_segments.append({"type": "turn_header", "turn": i + 1, "role": role})
                for seg_type, seg_content in _extract_html_blocks(m.content):
                    if seg_type == "html":
                        pdf_segments.append({"type": "html_chart", "html": seg_content})
                    else:
                        cleaned = _clean_text_segment(seg_content)
                        if cleaned:
                            pdf_segments.append({"type": "text", "content": cleaned})
            return "__VERBATIM_SEGMENTS__" + json.dumps(pdf_segments)

    selected_model = ALLOWED_MODELS["sonnet"]
    conv_text = "\n\n".join(
        f"{'USER' if m.role == 'user' else 'DIUD AGENT'}: {m.content}"
        for m in conversation
    )
    format_hint = (
        "Format as a PowerPoint: use SLIDE: <title> for each slide, then bullet points."
        if export_type == "pptx"
        else "Format as a professional PDF report: ## section headers, narrative prose, tables."
    )
    dataset_hint = ""
    if stored_dataset:
        dataset_hint = (
            f"\n\nDATASET CONTEXT: The query returned {stored_dataset.total_rows} records "
            f"with columns: {', '.join(stored_dataset.columns[:12])}. "
            f"Filters: {stored_dataset.filters_applied}."
        )
    prompt = f"""You are preparing a professional {export_type.upper()} summary report.

CONVERSATION:
{conv_text}
{dataset_hint}

TASK: Create a concise executive summary titled "{title}"

{format_hint}

REQUIREMENTS:
- Executive summary at the start with key numbers only
- Logical sections: summary, key metrics, insights, recommendations
- Bold key numbers; clean professional tone
- Today: {date.today().strftime('%B %d, %Y')}
- Generate the COMPLETE document — do not truncate

Generate the summary report now:"""

    response = _ai_client.messages.create(
        model=selected_model,
        system="You are a professional document formatter. Follow the instructions exactly. Never add content not requested.",
        messages=[{"role": "user", "content": prompt}],
        temperature=0,
        max_tokens=EXPORT_MAX_TOKENS,
    )
    ai_text = _extract_text(response.content)

    if stored_dataset and stored_dataset.total_rows > 0:
        table_md  = _rows_to_markdown_table(stored_dataset)
        meta_line = (
            f"**Total records:** {stored_dataset.total_rows:,} | "
            f"**Filters:** {stored_dataset.filters_applied} | "
            f"**Exported:** {date.today().strftime('%B %d, %Y')}"
        )
        full_section = f"\n\n## Data Export ({stored_dataset.total_rows:,} records)\n\n{meta_line}\n\n{table_md}"
        if "[DEAL_TABLE_PLACEHOLDER]" in ai_text:
            ai_text = ai_text.replace("[DEAL_TABLE_PLACEHOLDER]", full_section)
        else:
            ai_text = ai_text.rstrip() + full_section

    return ai_text


def _rows_to_markdown_table(stored: QueryResult) -> str:
    if not stored.rows:
        return "_No data._"

    cols   = stored.columns
    header = "| " + " | ".join(cols) + " |"
    sep    = "| " + " | ".join("---" for _ in cols) + " |"
    lines  = [header, sep]

    for row in stored.rows:
        cells = [str(row.get(c, "")).replace("|", "\\|") for c in cols]
        lines.append("| " + " | ".join(cells) + " |")

    return "\n".join(lines)


# =============================================================================
# [24] PDF Builder
# =============================================================================
_C_NAVY  = colors.HexColor("#0D1B3E")
_C_BLUE  = colors.HexColor("#1565C0")
_C_WHITE = colors.white
_C_BG    = colors.HexColor("#F7F9FC")
_C_TXT   = colors.HexColor("#1E293B")
_C_DIM   = colors.HexColor("#94A3B8")

_SECTION_COLORS = {
    "executive": colors.HexColor("#0D1B3E"),
    "pipeline":  colors.HexColor("#1565C0"),
    "metric":    colors.HexColor("#004D40"),
    "deal":      colors.HexColor("#1565C0"),
    "regional":  colors.HexColor("#BF360C"),
    "win":       colors.HexColor("#B71C1C"),
    "loss":      colors.HexColor("#B71C1C"),
    "recommend": colors.HexColor("#1B5E20"),
    "summary":   colors.HexColor("#0D1B3E"),
    "analysis":  colors.HexColor("#1565C0"),
    "overview":  colors.HexColor("#004D40"),
    "insight":   colors.HexColor("#1B5E20"),
}

PW, PH = A4
_ML = _MR = 0.6 * inch
_MT = 0.45 * inch
_MB = 0.40 * inch
_HDR_H = 44
_FTR_H = 20
_CW    = PW - _ML - _MR


def _pdf_styles():
    return {
        "Cover_Title": ParagraphStyle("Cover_Title", fontSize=26, leading=32,
            textColor=_C_WHITE, fontName="Helvetica-Bold", spaceAfter=8),
        "Cover_Sub": ParagraphStyle("Cover_Sub", fontSize=13, leading=18,
            textColor=colors.HexColor("#B0BEC5"), fontName="Helvetica"),
        "Section_H": ParagraphStyle("Section_H", fontSize=11, leading=15,
            textColor=_C_WHITE, fontName="Helvetica-Bold"),
        "Body": ParagraphStyle("Body", fontSize=9, leading=14, textColor=_C_TXT,
            fontName="Helvetica", spaceAfter=4),
        "Bullet": ParagraphStyle("Bullet", fontSize=9, leading=14, textColor=_C_TXT,
            fontName="Helvetica", leftIndent=12, firstLineIndent=-8, spaceAfter=3),
        "H2": ParagraphStyle("H2", fontSize=11, leading=15, textColor=_C_NAVY,
            fontName="Helvetica-Bold", spaceBefore=10, spaceAfter=4),
        "H3": ParagraphStyle("H3", fontSize=9, leading=13, textColor=_C_BLUE,
            fontName="Helvetica-Bold", spaceBefore=6, spaceAfter=2),
        "TH": ParagraphStyle("TH", fontSize=7, leading=9, textColor=_C_WHITE,
            fontName="Helvetica-Bold"),
        "TD": ParagraphStyle("TD", fontSize=7, leading=9, textColor=_C_TXT,
            fontName="Helvetica"),
    }


def _parse_sections(text: str):
    parts = re.split(r'^##\s+', text, flags=re.MULTILINE)
    return [
        (lines[0].strip(), lines[1].strip() if len(lines) > 1 else "")
        for part in parts if part.strip()
        for lines in [part.strip().split("\n", 1)]
    ]


def _build_pdf(title: str, report_text: str) -> bytes:
    from reportlab.platypus import Image as RLImage
    buf    = io.BytesIO()
    styles = _pdf_styles()

    def _on_page(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(_C_NAVY)
        canvas.rect(0, PH - _HDR_H - _MT, PW, _HDR_H + _MT, fill=1, stroke=0)
        canvas.setFillColor(_C_WHITE)
        canvas.setFont("Helvetica-Bold", 10)
        canvas.drawString(_ML, PH - _MT - 28, title)
        canvas.setFillColor(_C_BG)
        canvas.rect(0, 0, PW, _FTR_H + _MB, fill=1, stroke=0)
        canvas.setFillColor(_C_DIM)
        canvas.setFont("Helvetica", 7)
        canvas.drawCentredString(
            PW / 2, _MB + 5,
            f"DIUD Report  |  AI-Generated  |  CONFIDENTIAL  |  {date.today().strftime('%B %Y')}"
        )
        canvas.drawRightString(PW - _MR, _MB + 5, f"Page {canvas.getPageNumber()}")
        canvas.restoreState()

    frame    = Frame(_ML, _MB + _FTR_H, _CW, PH - _HDR_H - _MT - _MB - _FTR_H, id="main")
    template = PageTemplate(id="main", frames=[frame], onPage=_on_page)
    doc = BaseDocTemplate(buf, pagesize=A4, leftMargin=_ML, rightMargin=_MR,
                          topMargin=_MT + _HDR_H, bottomMargin=_MB + _FTR_H)
    doc.addPageTemplates([template])

    cover_story = [
        Spacer(1, 1.0 * inch),
        Paragraph(title, styles["Cover_Title"]),
        Paragraph(f"Generated {date.today().strftime('%B %d, %Y')}", styles["Cover_Sub"]),
        PageBreak(),
    ]

    if report_text.startswith("__VERBATIM_SEGMENTS__"):
        segments = json.loads(report_text[len("__VERBATIM_SEGMENTS__"):])
        story = cover_story[:]

        for seg in segments:
            if seg["type"] == "turn_header":
                role_label = "User" if seg["role"] == "User" else "DIUD Agent"
                bar_color  = _C_BLUE if seg["role"] == "User" else _C_NAVY
                story.append(Spacer(1, 8))
                story.append(Table(
                    [[Paragraph(f"Turn {seg['turn']} — {role_label}", styles["Section_H"])]],
                    colWidths=[_CW],
                    style=TableStyle([
                        ("BACKGROUND",    (0,0),(-1,-1), bar_color),
                        ("TOPPADDING",    (0,0),(-1,-1), 6),
                        ("BOTTOMPADDING", (0,0),(-1,-1), 6),
                        ("LEFTPADDING",   (0,0),(-1,-1), 10),
                    ])
                ))
                story.append(Spacer(1, 6))

            elif seg["type"] == "html_chart":
                png_bytes = _render_html_to_png(seg["html"], width=760)
                if png_bytes:
                    img = RLImage(io.BytesIO(png_bytes), width=_CW, height=_CW * 0.55)
                    img.hAlign = "CENTER"
                    story.append(img)
                    story.append(Spacer(1, 10))
                else:
                    story.append(Paragraph(
                        "[Chart could not be rendered — view in live DIUD interface]",
                        styles["Body"]
                    ))

            elif seg["type"] == "text":
                lines = seg["content"].split("\n")
                i = 0
                while i < len(lines):
                    line = lines[i].strip()
                    if line.startswith("|") and i + 1 < len(lines) and "---" in lines[i + 1]:
                        table_lines = []
                        while i < len(lines) and lines[i].strip().startswith("|"):
                            table_lines.append(lines[i].strip())
                            i += 1
                        story.append(_md_table_to_rl(table_lines, styles))
                        story.append(Spacer(1, 6))
                        continue
                    if not line:
                        story.append(Spacer(1, 3))
                    elif line.startswith("### "):
                        story.append(Paragraph(_strip_md(line), styles["H3"]))
                    elif line.startswith("## "):
                        story.append(Paragraph(_strip_md(line), styles["H2"]))
                    elif line.startswith(("- ", "* ", "• ")):
                        story.append(Paragraph("• " + _strip_md(line[2:]), styles["Bullet"]))
                    else:
                        story.append(Paragraph(_strip_md(line), styles["Body"]))
                    i += 1

        doc.build(story)
        return buf.getvalue()

    sections = _parse_sections(report_text)
    story = cover_story[:]

    if not sections:
        clean_text = re.sub(r'```html.*?```', '[Chart omitted]', report_text, flags=re.DOTALL | re.IGNORECASE)
        clean_text = re.sub(r'```.*?```', '', clean_text, flags=re.DOTALL)
        clean_text = re.sub(r'^\s*--\s*$', '', clean_text, flags=re.MULTILINE)
        for line in clean_text.split("\n"):
            line = line.strip()
            if line:
                story.append(Paragraph(_strip_md(line), styles["Body"]))
    else:
        for sec_title, sec_body in sections:
            color_key = next((k for k in _SECTION_COLORS if k in sec_title.lower()), None)
            bar_color = _SECTION_COLORS.get(color_key, _C_BLUE)
            story.append(Table(
                [[Paragraph(sec_title.upper(), styles["Section_H"])]],
                colWidths=[_CW],
                style=TableStyle([
                    ("BACKGROUND",    (0, 0), (-1, -1), bar_color),
                    ("TOPPADDING",    (0, 0), (-1, -1), 8),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                    ("LEFTPADDING",   (0, 0), (-1, -1), 10),
                ])
            ))
            story.append(Spacer(1, 6))

            lines = sec_body.split("\n")
            i = 0
            while i < len(lines):
                line = lines[i].strip()
                if line.startswith("|") and i + 1 < len(lines) and "---" in lines[i + 1]:
                    table_lines = []
                    while i < len(lines) and lines[i].strip().startswith("|"):
                        table_lines.append(lines[i].strip())
                        i += 1
                    story.append(_md_table_to_rl(table_lines, styles))
                    story.append(Spacer(1, 6))
                    continue
                if not line:
                    story.append(Spacer(1, 3))
                elif line.startswith("### "):
                    story.append(Paragraph(_strip_md(line), styles["H3"]))
                elif line.startswith("## "):
                    story.append(Paragraph(_strip_md(line), styles["H2"]))
                elif line.startswith(("- ", "* ", "• ")):
                    story.append(Paragraph("• " + _strip_md(line[2:]), styles["Bullet"]))
                else:
                    story.append(Paragraph(_strip_md(line), styles["Body"]))
                i += 1

            story.extend([Spacer(1, 12), PageBreak()])

    doc.build(story)
    return buf.getvalue()


def _md_table_to_rl(table_lines: list, styles: dict):
    data = []
    for idx, line in enumerate(table_lines):
        if "---" in line:
            continue
        cells = [c.strip().replace("\\|", "|") for c in line.strip("|").split("|")]
        if idx == 0:
            row = [Paragraph(_strip_md(c), styles["TH"]) for c in cells]
        else:
            row = [Paragraph(_strip_md(c), styles["TD"]) for c in cells]
        data.append(row)

    if not data:
        return Spacer(1, 1)

    num_cols = max(len(r) for r in data)
    col_w    = _CW / max(num_cols, 1)

    tbl = Table(data, colWidths=[col_w] * num_cols, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0),  _C_NAVY),
        ("TEXTCOLOR",     (0, 0), (-1, 0),  _C_WHITE),
        ("FONTNAME",      (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, -1), 7),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
        ("GRID",          (0, 0), (-1, -1), 0.3, colors.HexColor("#E2E8F0")),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING",   (0, 0), (-1, -1), 4),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return tbl


# =============================================================================
# [25] PPTX Builder
# =============================================================================
_C_NAVY_P  = RGBColor(0x0D, 0x1B, 0x3E)
_C_DNAV_P  = RGBColor(0x0A, 0x11, 0x28)
_C_BLUE_P  = RGBColor(0x1E, 0x88, 0xE5)
_C_WHITE_P = RGBColor(0xFF, 0xFF, 0xFF)
_C_LTBG_P  = RGBColor(0xF5, 0xF7, 0xFA)
_C_TXT_P   = RGBColor(0x1A, 0x1A, 0x2E)
_C_DIM_P   = RGBColor(0x88, 0x99, 0xAA)

_SLIDE_ACCENT = {
    "overview":  RGBColor(0x1E, 0x88, 0xE5),
    "pipeline":  RGBColor(0x00, 0x89, 0x7B),
    "deal":      RGBColor(0x1E, 0x88, 0xE5),
    "metric":    RGBColor(0x2E, 0x7D, 0x32),
    "regional":  RGBColor(0xBF, 0x36, 0x0C),
    "win":       RGBColor(0x2E, 0x7D, 0x32),
    "loss":      RGBColor(0xC6, 0x28, 0x28),
    "recommend": RGBColor(0x1B, 0x5E, 0x20),
    "summary":   RGBColor(0x1E, 0x88, 0xE5),
    "analysis":  RGBColor(0x00, 0x89, 0x7B),
    "insight":   RGBColor(0x1B, 0x5E, 0x20),
    "executive": RGBColor(0x0D, 0x1B, 0x3E),
}


def _pptx_bg(slide, color):
    f = slide.background.fill; f.solid(); f.fore_color.rgb = color

def _pptx_rect(slide, l, t, w, h, color):
    shp = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    shp.fill.solid(); shp.fill.fore_color.rgb = color; shp.line.fill.background()
    return shp

def _pptx_txt(slide, text, l, t, w, h, bold=False, size=18, color=None, align=PP_ALIGN.LEFT):
    txb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    txb.word_wrap = True
    tf = txb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = align
    run = p.add_run(); run.text = text
    run.font.size = Pt(size); run.font.bold = bold
    run.font.color.rgb = color or _C_TXT_P
    return txb

def _parse_slides(text: str):
    slides, cur_title, cur_bullets = [], None, []
    for line in text.split("\n"):
        line = line.rstrip()
        if line.startswith("SLIDE:"):
            if cur_title is not None:
                slides.append((cur_title, cur_bullets))
            cur_title, cur_bullets = line[6:].strip(), []
        elif line.startswith("- ") and cur_title:
            cur_bullets.append(line[2:].strip())
    if cur_title is not None:
        slides.append((cur_title, cur_bullets))
    return slides

def _build_pptx(title: str, slide_text: str) -> bytes:
    slides_data = _parse_slides(slide_text) or [(title, [slide_text[:400]])]
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    footer_text = f"DIUD  |  AI-Generated  |  CONFIDENTIAL  |  {date.today().strftime('%B %Y')}"
    blank = prs.slide_layouts[6]

    def _footer(s):
        _pptx_rect(s, 0, 7.1, 13.33, 0.4, _C_DNAV_P)
        _pptx_txt(s, footer_text, 0.3, 7.12, 12, 0.35, size=7, color=_C_DIM_P, align=PP_ALIGN.CENTER)

    def _accent(t):
        for k, c in _SLIDE_ACCENT.items():
            if k in t: return c
        return _C_BLUE_P

    cover = prs.slides.add_slide(blank)
    _pptx_bg(cover, _C_NAVY_P)
    _pptx_rect(cover, 0, 3.2, 13.33, 0.06, _C_BLUE_P)
    _pptx_txt(cover, title, 0.8, 1.6, 11.5, 1.4, bold=True, size=34, color=_C_WHITE_P)
    _pptx_txt(cover, "Deals Intelligence Report", 0.8, 3.0, 8, 0.6,
              size=15, color=RGBColor(0xB0, 0xBE, 0xC5))
    _pptx_txt(cover, f"Generated: {date.today().strftime('%B %d, %Y')}", 0.8, 3.6, 6, 0.45,
              size=12, color=RGBColor(0x78, 0x90, 0x9C))

    for i, (s_title, bullets) in enumerate(slides_data):
        slide = prs.slides.add_slide(blank)
        ac = _accent(s_title.lower())
        _pptx_bg(slide, _C_LTBG_P)
        _pptx_rect(slide, 0, 0, 13.33, 0.9, ac)
        _pptx_txt(slide, s_title.upper(), 0.35, 0.1, 12.5, 0.7, bold=True, size=18, color=_C_WHITE_P)
        _pptx_txt(slide, str(i + 1), 12.5, 0.12, 0.6, 0.6, size=11, color=_C_WHITE_P, align=PP_ALIGN.RIGHT)
        _pptx_rect(slide, 0.3, 1.0, 12.73, 5.9, _C_WHITE_P)
        if bullets:
            txb = slide.shapes.add_textbox(Inches(0.5), Inches(1.1), Inches(12.3), Inches(5.6))
            txb.word_wrap = True
            tf = txb.text_frame; tf.word_wrap = True
            for j, bullet in enumerate(bullets[:12]):
                p = tf.add_paragraph() if j > 0 else tf.paragraphs[0]
                p.space_before = Pt(4)
                dot = p.add_run(); dot.text = "●  "; dot.font.size = Pt(8); dot.font.color.rgb = ac
                run = p.add_run(); run.text = bullet; run.font.size = Pt(12); run.font.color.rgb = _C_TXT_P
        else:
            _pptx_txt(slide, "No data available.", 0.5, 1.2, 12, 0.5, size=11, color=_C_DIM_P)
        _footer(slide)

    buf = io.BytesIO(); prs.save(buf)
    return buf.getvalue()
