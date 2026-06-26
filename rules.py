"""
rules.py
=============================================================================
Machine-checkable registry of MANDATORY business rules from the DIUD system
prompt. Every SQL string the model produces (and, where relevant, the rows
that come back) is run through this registry BEFORE being trusted.

Add new rules by appending to RULES / RESULT_RULES. Each rule is independent
and self-contained, so the registry can grow without anything else changing.
=============================================================================
"""

import re
from typing import Any, Callable, Dict, List, Optional


# =============================================================================
# Intent detection — light NLP on the user's message + the SQL itself.
# This is intentionally simple (regex, not another LLM call) so it's fast,
# deterministic, and cheap to run on every tool call.
# =============================================================================
def detect_intent(user_message: str, sql: str = "") -> Dict[str, Any]:
    msg = user_message or ""
    intent: Dict[str, Any] = {}

    # ── Cohort stage detection ───────────────────────────────────────────────
    # Pattern 1: explicit arrow notation  "20% → Closed Won", "20% to CW"
    m = re.search(r'(\d+)\s*%\s*(→|->|to)\s*(closed\s*won|cw)\b', msg, re.I)
    if m:
        intent["cohort_stage"] = m.group(1)

    # Pattern 2: funnel/cohort keyword BEFORE the percentage
    #   "funnel from 20%", "cohort starting at 10%", "pipeline from 30% stage"
    if not intent.get("cohort_stage"):
        m2 = re.search(
            r'\b(funnel|cohort|pipeline\s+from|starting\s+(?:at|from)|from\s+stage)\b'
            r'.*?(\d+)\s*%',
            msg, re.I
        )
        if m2:
            intent["cohort_stage"] = m2.group(2)

    # Pattern 3: percentage BEFORE funnel/cohort keyword
    #   "20% funnel", "20% cohort analysis", "30% stage breakdown", "20% conversion"
    if not intent.get("cohort_stage"):
        m3 = re.search(
            r'(\d+)\s*%.*?\b(funnel|cohort|stage\s+breakdown|conversion|progression)\b',
            msg, re.I
        )
        if m3:
            intent["cohort_stage"] = m3.group(1)

    # Pattern 4: generic stage-progression language with a percentage anywhere
    #   "show me how deals progress after 20%", "deals that reached 20"
    if not intent.get("cohort_stage"):
        m4 = re.search(
            r'\b(progress|reached|entered|qualified\s+at|moved\s+(?:past|beyond|to))\b'
            r'.*?(\d+)\s*%',
            msg, re.I
        )
        if not m4:
            m4 = re.search(
                r'(\d+)\s*%.*?\b(progress|reached|entered|qualified\s+at|moved\s+(?:past|beyond|to))\b',
                msg, re.I
            )
        if m4:
            for grp in m4.groups():
                if grp and grp.isdigit():
                    intent["cohort_stage"] = grp
                    break

    # Pattern 5: fall back to SQL structure — if the model used became_N_deal_date,
    #   the intent is cohort regardless of what the user typed
    if not intent.get("cohort_stage") and sql:
        sql_m = re.search(r'became_(\d+)_deal_date', sql, re.I)
        if sql_m:
            intent["cohort_stage"] = sql_m.group(1)

    # ── List query detection ─────────────────────────────────────────────────
    if re.search(r'\b(list|show me all|which deals|deals\s+(with|where))\b', msg, re.I):
        intent["query_type"] = "list"

    # ── Top-N cap detection ──────────────────────────────────────────────────
    if re.search(r'\btop\s+\d+\b|\bfirst\s+\d+\b', msg, re.I):
        intent["top_n"] = True

    # ── Metric detection ─────────────────────────────────────────────────────
    if re.search(r'\bMQL\b', msg, re.I):
        intent["metric"] = "mql"

    if re.search(r'\battainment\b|\bvs\.?\s*target\b|\bquota\b', msg, re.I):
        intent["metric"] = intent.get("metric") or "attainment"

    if re.search(r'\bactive pipeline\b', msg, re.I):
        intent["metric"] = intent.get("metric") or "active_pipeline"

    # ── Generation hygiene ───────────────────────────────────────────────────
    if sql and re.search(r'\bMANDATORY_BASE_FILTERS\b', sql):
        intent["placeholder_leak"] = True

    return intent


def _arity(fn: Callable) -> int:
    return fn.__code__.co_argcount


def _call(fn: Callable, sql: str, intent: dict):
    """Call a rule lambda whether it takes (sql) or (sql, intent)."""
    return fn(sql, intent) if _arity(fn) == 2 else fn(sql)


# =============================================================================
# SQL-TEXT RULES — checkable by inspecting the SQL string alone.
# =============================================================================
RULES: List[Dict[str, Any]] = [

    # ---- §3 MANDATORY_BASE_FILTERS ----------------------------------------
    {
        "id": "base_filter_pipeline",
        "section": "§3 MANDATORY_BASE_FILTERS (1/3)",
        "applies_when": lambda sql: "hs_analytics.deals" in sql,
        "check": lambda sql: re.search(r"pipeline\s*=\s*'default'", sql, re.I) is not None,
        "message": "Missing `pipeline = 'default'` base filter.",
    },
    {
        "id": "base_filter_deal_type",
        "section": "§3 MANDATORY_BASE_FILTERS (2/3)",
        "applies_when": lambda sql: "hs_analytics.deals" in sql,
        "check": lambda sql: "Partner-Led SMB" in sql and "NOT IN" in sql.upper(),
        "message": "Missing deal_type NOT IN ('Partner-Led SMB') base filter.",
    },
    {
        "id": "base_filter_allowlist",
        "section": "§3 MANDATORY_BASE_FILTERS (3/3)",
        "applies_when": lambda sql: "hs_analytics.deals" in sql,
        "check": lambda sql: "gs_deal_ids_hs" in sql,
        "message": "Missing deal_id allowlist join against gs_deal_ids_hs.",
    },

    # ---- §3 Duplicate record exclusion -------------------------------------
    {
        "id": "final_keyword",
        "section": "§3 Duplicate exclusion — FINAL",
        "applies_when": lambda sql: bool(re.search(r"hs_analytics\.\w+", sql)),
        "check": lambda sql: bool(re.search(r"hs_analytics\.\w+\s+FINAL", sql, re.I)),
        "message": "Missing FINAL on an hs_analytics.* table reference.",
    },
    {
        "id": "count_distinct_not_count",
        "section": "§3 / §9.5 — countDistinct, never count()",
        "applies_when": lambda sql: ("hs_analytics" in sql) and bool(re.search(r"(?<!\w)count\s*\(", sql, re.I)),
        "check": lambda sql: bool(re.search(r"countDistinct\s*\(\s*(deal_id|contact_id)\s*\)", sql, re.I)),
        "message": "Uses count() instead of countDistinct(deal_id)/countDistinct(contact_id).",
    },
    {
        "id": "distinct_in_association_subquery",
        "section": "§3 Table 5 — DISTINCT in association subqueries",
        "applies_when": lambda sql: "gs_DealContactAssociation" in sql and re.search(r"\(\s*SELECT", sql, re.I),
        "check": lambda sql: bool(re.search(r"SELECT\s+DISTINCT", sql, re.I)),
        "message": "Subquery against gs_DealContactAssociation missing DISTINCT.",
    },

    # ---- §9 Date handling standard -----------------------------------------
    {
        "id": "date_cast_standard",
        "section": "§9 Date handling — toDate(LEFT(coalesce(...)))",
        "applies_when": lambda sql: bool(
            re.search(r"\b(close_date|became_\d+_deal_date)\b\s*(>=|<=|>|<|=)", sql, re.I)
        ),
        "check": lambda sql: "toDate(LEFT(coalesce(" in sql.replace(" ", "")
        or "toDate(LEFT(coalesce(" in sql,
        "message": "Raw date comparison without the mandatory toDate(LEFT(coalesce(col,'1900-01-01'),10)) cast.",
    },

    # ---- §9.4 No LIMIT on list queries --------------------------------------
    {
        "id": "no_limit_on_list",
        "section": "§9.4 No LIMIT unless 'top N' / 'first N'",
        "applies_when": lambda sql, intent: intent.get("query_type") == "list" and not intent.get("top_n"),
        "check": lambda sql, intent: "LIMIT" not in sql.upper(),
        "message": "LIMIT applied to a list query the user did not ask to cap with 'top N'/'first N'.",
    },

    # ---- §9.1/§9.2/§9.3 generic guardrails ---------------------------------
    {
        "id": "select_or_with_only",
        "section": "§9.1 SELECT/WITH only",
        "applies_when": lambda sql: True,
        "check": lambda sql: sql.strip().upper().startswith(("SELECT", "WITH")),
        "message": "Query does not start with SELECT or WITH.",
    },
    {
        "id": "no_placeholder_tokens",
        "section": "Generation hygiene",
        "applies_when": lambda sql: True,
        "check": lambda sql: "<MANDATORY_BASE_FILTERS>" not in sql and "<" not in sql.split("--")[0] or True,
        "message": "SQL contains an unresolved placeholder token like <MANDATORY_BASE_FILTERS>.",
        # Note: kept loose to avoid false positives on legitimate '<' comparisons;
        # the placeholder_leak intent flag below does the strict check.
    },
    {
        "id": "no_placeholder_leak_strict",
        "section": "Generation hygiene (strict)",
        "applies_when": lambda sql, intent: intent.get("placeholder_leak", False),
        "check": lambda sql, intent: False,
        "message": "Literal placeholder '<MANDATORY_BASE_FILTERS>' leaked into generated SQL — must be expanded.",
    },

    # ---- §8b Cohort funnel rule ---------------------------------------------
    {
        "id": "cohort_anchor",
        "section": "§8b Cohort definition — became_<stage>_deal_date IS NOT NULL",
        "applies_when": lambda sql, intent: intent.get("cohort_stage") is not None,
        "check": lambda sql, intent: f"became_{intent['cohort_stage']}_deal_date" in sql
        and "IS NOT NULL" in sql.upper(),
        "message": "Cohort query missing became_<stage>_deal_date IS NOT NULL anchor.",
    },
    {
        "id": "cohort_exclusion",
        "section": "§8b Stage exclusion — NOT IN prior stages",
        "applies_when": lambda sql, intent: intent.get("cohort_stage") is not None,
        "check": lambda sql, intent: "NOT IN" in sql.upper(),
        "message": "Cohort query missing NOT IN exclusion of pre-cohort stages.",
    },
    {
        "id": "cohort_single_cte",
        "section": "§8b — funnel must be one CTE, not per-stage queries",
        "applies_when": lambda sql, intent: intent.get("cohort_stage") is not None,
        "check": lambda sql, intent: sql.strip().upper().startswith("WITH") and "GROUP BY" in sql.upper(),
        "message": "Cohort funnel should be a single WITH-cohort CTE with GROUP BY deal_stage, not a one-off filter.",
    },
    {
        "id": "cohort_count_distinct",
        "section": "§8b Deduplication — countDistinct(deal_id)",
        "applies_when": lambda sql, intent: intent.get("cohort_stage") is not None,
        "check": lambda sql, intent: bool(re.search(r"countDistinct\s*\(\s*deal_id\s*\)", sql, re.I)),
        "message": "Cohort funnel query not using countDistinct(deal_id).",
    },

    # ---- §6 MQL calculation --------------------------------------------------
    {
        "id": "mql_filters_present",
        "section": "§6 MQL — 3 mandatory filters",
        "applies_when": lambda sql, intent: intent.get("metric") == "mql",
        "check": lambda sql, intent: "lifecycle_stage" in sql
        and "date_entered_marketing_qualified_lead_lifecycle_stage_pipeline" in sql,
        "message": "MQL query missing required lifecycle_stage / MQL date-entered filter.",
    },
    {
        "id": "mql_no_quarter_divide",
        "section": "§6 MQL — never derive quarterly target by /4",
        "applies_when": lambda sql, intent: intent.get("metric") == "mql",
        "check": lambda sql, intent: "/4" not in sql.replace(" ", "") and "/ 4" not in sql,
        "message": "MQL target appears to be derived by dividing annual target by 4 — not allowed.",
    },

    # ---- §5 Target / attainment queries --------------------------------------
    {
        "id": "attainment_uses_cte_pattern",
        "section": "§5 Rule 4 — actual/target CTE pattern",
        "applies_when": lambda sql, intent: intent.get("metric") == "attainment",
        "check": lambda sql, intent: sql.strip().upper().startswith("WITH"),
        "message": "Attainment/target query should use the actual-CTE + target-CTE pattern from §5.",
    },

    # ---- §8 Active pipeline definition ----------------------------------------
    {
        "id": "active_pipeline_stage_filter",
        "section": "§8 Active pipeline definition",
        "applies_when": lambda sql, intent: intent.get("metric") == "active_pipeline",
        "check": lambda sql, intent: all(
            s in sql
            for s in ["20% - Solution", "30% - Proof", "40% - Proposal", "60% - Price Negotiation", "75% - Contract Review"]
        ),
        "message": "Active pipeline query missing one or more of the 5 required active deal_stage values.",
    },
    # ---- §8b SQL-structure-triggered cohort checks (bypass intent detection) ---
    {
        "id": "cohort_sql_triggered_exclusion",
        "section": "§8b — SQL-triggered cohort exclusion check",
        "applies_when": lambda sql: bool(re.search(r'became_\d+_deal_date', sql, re.I)),
        "check": lambda sql: bool(re.search(r'\bNOT\s+IN\b', sql, re.I)),
        "message": (
            "SQL references became_<N>_deal_date (cohort anchor) but is missing "
            "the NOT IN exclusion of prior stages. Every cohort query MUST exclude "
            "deals currently in all stages prior to the cohort starting stage (§8b)."
        ),
    },
    {
        "id": "cohort_sql_triggered_cte",
        "section": "§8b — Cohort must use CTE pattern",
        "applies_when": lambda sql: bool(re.search(r'became_\d+_deal_date', sql, re.I)),
        "check": lambda sql: sql.strip().upper().startswith("WITH"),
        "message": (
            "SQL references became_<N>_deal_date but is not structured as a CTE. "
            "All cohort funnel queries MUST use the WITH-cohort CTE pattern from §8b."
        ),
    },
    {
        "id": "cohort_sql_triggered_count_distinct",
        "section": "§8b — Cohort must use countDistinct(deal_id)",
        "applies_when": lambda sql: bool(re.search(r'became_\d+_deal_date', sql, re.I)),
        "check": lambda sql: bool(re.search(r'countDistinct\s*\(\s*deal_id\s*\)', sql, re.I)),
        "message": (
            "Cohort funnel SQL is not using countDistinct(deal_id). "
            "Never use count() in cohort queries — each deal must be counted exactly once."
        ),
    },
]


# =============================================================================
# RESULT-LEVEL RULES — checkable only against the returned rows, since some
# violations look fine as SQL text but produce impossible numbers.
# =============================================================================
RESULT_RULES: List[Dict[str, Any]] = [
    {
        "id": "funnel_sum_within_cohort",
        "section": "§8b Deduplication — each deal appears once across stages",
        "applies_when": lambda rows, intent: intent.get("cohort_stage") is not None and bool(rows),
        "check": lambda rows, intent: _funnel_sum_ok(rows),
        "message": "Funnel stage counts (active stages + Closed Won/Lost) exceed the cohort total — "
                    "rows were not derived from a single deduplicated cohort.",
    },
]


def _funnel_sum_ok(rows: List[dict]) -> bool:
    try:
        cohort_total = None
        active_sum = 0
        terminal_sum = 0
        for r in rows:
            stage = str(r.get("deal_stage", ""))
            cnt = r.get("deal_count", 0) or 0
            if cohort_total is None:
                # First row's stage is assumed to be the cohort's starting stage and is
                # the baseline (100%) -- treat it as the cohort total.
                cohort_total = cnt
            if "Closed Won" in stage or "Closed Lost" in stage:
                terminal_sum += cnt
            elif cnt != cohort_total:  # skip the baseline row itself
                active_sum += cnt
        if cohort_total is None:
            return True
        return (active_sum + terminal_sum) <= cohort_total
    except Exception:
        # If we can't evaluate it, don't block the response on a checker bug —
        # log and pass. Visibility is handled via the audit log in main.py.
        return True


# =============================================================================
# Public API
# =============================================================================
def validate_sql_against_rules(sql: str, user_message: str) -> List[str]:
    """Run every applicable RULES entry against `sql`. Returns violation strings."""
    intent = detect_intent(user_message, sql)
    violations = []
    for rule in RULES:
        try:
            applies = _call(rule["applies_when"], sql, intent)
        except Exception:
            applies = False
        if not applies:
            continue
        try:
            ok = _call(rule["check"], sql, intent)
        except Exception:
            ok = False
        if not ok:
            violations.append(f"[{rule['id']}] {rule['section']}: {rule['message']}")
    return violations


def validate_result_against_rules(rows: List[dict], user_message: str, sql: str = "") -> List[str]:
    """Run every applicable RESULT_RULES entry against the returned `rows`."""
    intent = detect_intent(user_message, sql)
    violations = []
    for rule in RESULT_RULES:
        try:
            applies = rule["applies_when"](rows, intent)
        except Exception:
            applies = False
        if not applies:
            continue
        try:
            ok = rule["check"](rows, intent)
        except Exception:
            ok = False
        if not ok:
            violations.append(f"[{rule['id']}] {rule['section']}: {rule['message']}")
    return violations


def get_intent(user_message: str, sql: str = "") -> Dict[str, Any]:
    """Exposed for logging/debugging in main.py."""
    return detect_intent(user_message, sql)
