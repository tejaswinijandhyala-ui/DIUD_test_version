import re
import json
from typing import Dict, List, Optional, Tuple

PALETTE = [
    "#4E79A7", "#F28E2B", "#E15759", "#76B7B2", "#59A14F",
    "#EDC948", "#B07AA1", "#FF9DA7", "#9C755F", "#BAB0AC",
]

# Canonical funnel stage order — used to sort detected stage rows correctly
_STAGE_ORDER = [
    "5% - IQM Held", "10% - Discovery", "20% - Solution", "30% - Proof",
    "40% - Proposal", "60% - Price Negotiation", "75% - Contract Review",
    "90% - Deal Desk Review", "Closed Won",
]


def _stage_rank(stage: str) -> int:
    for i, s in enumerate(_STAGE_ORDER):
        if s.split(" - ")[0].strip() in str(stage):
            return i
    return 99


def _fmt_money(v: float) -> str:
    if abs(v) >= 1_000_000:
        return f"${v/1_000_000:.1f}M"
    if abs(v) >= 1_000:
        return f"${v/1_000:.0f}K"
    return f"${v:.0f}"


def _to_float(v) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _is_numericish(v) -> bool:
    try:
        float(v)
        return True
    except (TypeError, ValueError):
        return False


# =============================================================================
# DETECTION — figure out what kind of chart (if any) fits this result
# =============================================================================

def detect_chart_type(columns: List[str], rows: List[dict]) -> Optional[str]:
    """
    Returns one of: 'funnel', 'attainment', 'donut', 'bar_h', 'bar_v', None.
    None means: don't chart this (e.g. raw deal-list / single-value result).
    """
    if not rows or not columns:
        return None

    cols_lower = {c.lower() for c in columns}

    # Pure deal-list (Pattern B) — has deal_id / deal_name as a row-identity
    # column and many distinct rows → it's a list, not an aggregate. Skip.
    if "deal_id" in cols_lower and "deal_name" in cols_lower:
        return None

    # Pattern A — funnel: has a stage column + a count column (deals).
    # NOTE: this app's own SQL templates (§6, §8b) alias the stage column
    # as "deal_stage", not bare "stage" — match on substring, not exact
    # membership, or every real Pattern A result silently misses this
    # branch and falls through to a generic donut/bar chart instead.
    stage_col = next((c for c in columns if "stage" in c.lower()), None)
    if stage_col and any(c.lower() in ("deals", "deal_count") for c in columns):
        return "funnel"

    # Pattern C — attainment: actual vs target columns present
    has_actual = any("actual" in c.lower() for c in columns)
    has_target = any("target" in c.lower() for c in columns)
    if has_actual and has_target:
        return "attainment"

    # Generic aggregate: one categorical column + one numeric column
    numeric_cols = [c for c in columns if any(_is_numericish(r.get(c)) for r in rows)]
    categorical_cols = [c for c in columns if c not in numeric_cols]

    if len(numeric_cols) >= 1 and len(categorical_cols) >= 1 and len(rows) >= 2:
        if len(rows) <= 6:
            return "donut"
        return "bar_h"

    # Single aggregate row (e.g. one total) — not chart-worthy on its own
    return None


# =============================================================================
# DATA SHAPING
# =============================================================================

def _aggregate_funnel_rows(rows: List[dict]) -> List[Tuple[str, int, float]]:
    """Collapse possibly-multiple rows per stage (different region/source
    breakdowns) into one (stage, total_deals, total_amount) tuple per stage."""
    if not rows:
        return []
    # Real SQL templates alias this column "deal_stage", not bare "stage" —
    # detect the actual key present instead of assuming "stage" literally.
    stage_key = next((k for k in rows[0].keys() if "stage" in k.lower()), "stage")

    agg: Dict[str, List[float]] = {}
    for r in rows:
        stage = str(r.get(stage_key, "")).strip()
        if not stage:
            continue
        deals = _to_float(r.get("deals") or r.get("deal_count") or 0)
        amount = _to_float(r.get("amount") or r.get("pipeline_m") or 0)
        if stage not in agg:
            agg[stage] = [0.0, 0.0]
        agg[stage][0] += deals
        agg[stage][1] += amount

    out = [(stage, vals[0], vals[1]) for stage, vals in agg.items()]
    out.sort(key=lambda x: _stage_rank(x[0]))
    return out


def _aggregate_attainment_rows(rows: List[dict], columns: List[str]) -> Tuple[List[str], List[float], List[float]]:
    label_col = next(
        (c for c in columns if c.lower() in ("region", "deal_source_rollup", "create_quarter", "quarter")),
        columns[0],
    )
    actual_col = next((c for c in columns if "actual" in c.lower()), None)
    target_col = next((c for c in columns if "target" in c.lower() and "attainment" not in c.lower()), None)

    agg: Dict[str, List[float]] = {}
    for r in rows:
        label = str(r.get(label_col, "")).strip() or "Unknown"
        a = _to_float(r.get(actual_col)) if actual_col else 0
        t = _to_float(r.get(target_col)) if target_col else 0
        if label not in agg:
            agg[label] = [0.0, 0.0]
        agg[label][0] += a
        agg[label][1] += t

    labels = list(agg.keys())
    actuals = [agg[l][0] for l in labels]
    targets = [agg[l][1] for l in labels]
    return labels, actuals, targets


_NA_LIKE = {"n/a", "na", "unknown", "none", "null", "not specified", "not set", ""}


def _aggregate_generic_rows(
    rows: List[dict],
    columns: List[str],
    label_col_override: Optional[str] = None,
    value_col_override: Optional[str] = None,
    extra_exclude: Optional[List[str]] = None,
) -> Tuple[List[str], List[float], str, Optional[Tuple[str, float]]]:
    numeric_cols = [c for c in columns if any(_is_numericish(r.get(c)) for r in rows)]
    categorical_cols = [c for c in columns if c not in numeric_cols]

    # label_col_override / value_col_override come from Claude's chart spec
    # (choose_chart_spec tool) and are ONLY accepted here if they're real
    # column names present in this exact result — validated by the caller
    # before this function is ever reached. Falls back to the same
    # auto-detection as before when no spec was given.
    label_col = label_col_override if label_col_override in columns else (
        categorical_cols[0] if categorical_cols else columns[0]
    )
    value_col = value_col_override if value_col_override in columns else (
        numeric_cols[0] if numeric_cols else columns[-1]
    )

    exclude_set = set(_NA_LIKE)
    if extra_exclude:
        exclude_set |= {str(v).strip().lower() for v in extra_exclude}

    agg: Dict[str, float] = {}
    excluded_label: Optional[str] = None
    excluded_total = 0.0
    for r in rows:
        raw_label = str(r.get(label_col, "")).strip()
        # Unlogged/blank categories (e.g. "N/A" competitor on a lost deal),
        # or anything Claude explicitly asked to exclude via the chart spec,
        # are real data — but letting one dominant bucket set the scale
        # crushes every named category into an unreadable sliver. Pull it
        # out and let the caller footnote it instead.
        if raw_label.lower() in exclude_set:
            excluded_label = raw_label or "Unknown"
            excluded_total += _to_float(r.get(value_col))
            continue
        agg[raw_label] = agg.get(raw_label, 0.0) + _to_float(r.get(value_col))

    items = sorted(agg.items(), key=lambda kv: kv[1], reverse=True)[:12]
    labels = [k for k, _ in items]
    values = [v for _, v in items]
    excluded = (excluded_label, excluded_total) if excluded_total > 0 else None
    return labels, values, value_col, excluded


# =============================================================================
# HTML BUILDERS
# =============================================================================

def _wrap_html(body: str, title: str, subtitle: str) -> str:
    return f"""```html
<!DOCTYPE html>
<html><head>
<meta charset="UTF-8">
<style>
  * {{ box-sizing: border-box; margin:0; padding:0; font-family: 'Inter', system-ui, sans-serif; }}
  body {{ background:#fff; padding:16px; }}
  .chart-card {{ background:#fff; border:1px solid #E2E8F0; border-radius:12px; padding:18px; }}
  .chart-title {{ font-size:14px; font-weight:700; color:#0D1B3E; margin-bottom:2px; }}
  .chart-subtitle {{ font-size:11px; color:#94A3B8; margin-bottom:14px; }}
  .legend {{ display:flex; flex-wrap:wrap; gap:10px; margin-top:12px; }}
  .legend-item {{ display:flex; align-items:center; gap:6px; font-size:11px; color:#334155; }}
  .legend-swatch {{ width:10px; height:10px; border-radius:3px; flex-shrink:0; }}
</style>
</head>
<body>
<div class="chart-card">
  <div class="chart-title">{title}</div>
  <div class="chart-subtitle">{subtitle}</div>
  {body}
</div>
</body></html>
```"""


def build_funnel_html(
    stored_rows: List[dict], subtitle: str,
    title_override: Optional[str] = None, insight_note: Optional[str] = None,
) -> Optional[str]:
    stages = _aggregate_funnel_rows(stored_rows)
    if not stages:
        return None

    max_deals = max(s[1] for s in stages) or 1
    rows_html = []
    for i, (stage, deals, amount) in enumerate(stages):
        pct_width = max((deals / max_deals) * 88, 18)
        color = PALETTE[i % len(PALETTE)]
        conv = f"{(deals / stages[0][1] * 100):.0f}%" if stages[0][1] else "—"
        rows_html.append(f"""
        <div style="display:flex;align-items:center;gap:12px;margin-bottom:10px;">
          <div style="width:160px;font-size:11px;color:#334155;text-align:right;flex-shrink:0;">{stage}</div>
          <div style="flex:1;background:#F1F5F9;border-radius:6px;height:28px;position:relative;overflow:hidden;">
            <div style="width:{pct_width:.1f}%;height:100%;background:{color};border-radius:6px;
                        display:flex;align-items:center;justify-content:flex-end;padding-right:8px;">
              <span style="font-size:11px;color:white;font-weight:700;white-space:nowrap;">
                {int(deals):,} deals
              </span>
            </div>
          </div>
          <div style="width:110px;font-size:11px;color:#475569;flex-shrink:0;">
            {_fmt_money(amount * (1_000_000 if amount < 10000 else 1))} · {conv}
          </div>
        </div>""")

    note_html = (
        f'<div style="margin-top:8px;font-size:11.5px;color:#0D1B3E;font-weight:600;">{insight_note}</div>'
        if insight_note else ""
    )
    body = "<div>" + "".join(rows_html) + "</div>" + note_html
    return _wrap_html(body, title_override or "Pipeline Funnel", subtitle)


def build_attainment_html(
    stored_rows: List[dict], columns: List[str], subtitle: str,
    title_override: Optional[str] = None, insight_note: Optional[str] = None,
) -> Optional[str]:
    labels, actuals, targets = _aggregate_attainment_rows(stored_rows, columns)
    if not labels:
        return None

    max_val = max(max(actuals, default=0), max(targets, default=0)) or 1
    rows_html = []
    for i, label in enumerate(labels):
        a, t = actuals[i], targets[i]
        a_pct = max((a / max_val) * 100, 1)
        t_pct = max((t / max_val) * 100, 1)
        attain = f"{(a / t * 100):.0f}%" if t else "—"
        rows_html.append(f"""
        <div style="margin-bottom:14px;">
          <div style="display:flex;justify-content:space-between;font-size:11px;color:#334155;margin-bottom:4px;">
            <span style="font-weight:600;">{label}</span>
            <span style="color:#1565C0;font-weight:700;">{attain}</span>
          </div>
          <div style="background:#F1F5F9;border-radius:5px;height:14px;margin-bottom:3px;position:relative;">
            <div style="width:{a_pct:.1f}%;height:100%;background:#1E88E5;border-radius:5px;"></div>
          </div>
          <div style="background:#F8FAFC;border:1px dashed #CBD5E1;border-radius:5px;height:8px;width:{t_pct:.1f}%;"></div>
          <div style="font-size:10px;color:#94A3B8;margin-top:2px;">
            Actual: {_fmt_money(a)} &nbsp;|&nbsp; Target: {_fmt_money(t)}
          </div>
        </div>""")

    legend = """
    <div class="legend">
      <div class="legend-item"><span class="legend-swatch" style="background:#1E88E5;"></span>Actual</div>
      <div class="legend-item"><span class="legend-swatch" style="background:#CBD5E1;"></span>Target</div>
    </div>"""
    note_html = (
        f'<div style="margin-top:8px;font-size:11.5px;color:#0D1B3E;font-weight:600;">{insight_note}</div>'
        if insight_note else ""
    )
    body = "<div>" + "".join(rows_html) + "</div>" + legend + note_html
    return _wrap_html(body, title_override or "Actual vs Target", subtitle)


def build_bar_html(
    stored_rows: List[dict], columns: List[str], horizontal: bool, subtitle: str,
    label_col: Optional[str] = None, value_col: Optional[str] = None,
    exclude_values: Optional[List[str]] = None, title_override: Optional[str] = None,
    insight_note: Optional[str] = None,
) -> Optional[str]:
    labels, values, value_col_used, excluded = _aggregate_generic_rows(
        stored_rows, columns, label_col, value_col, exclude_values
    )
    if not labels:
        return None

    max_val = max(values, default=0) or 1
    rows_html = []
    for i, (label, val) in enumerate(zip(labels, values)):
        pct = max((val / max_val) * 100, 2)
        color = PALETTE[i % len(PALETTE)]
        display_val = _fmt_money(val) if val >= 1000 else f"{val:,.1f}"
        # Value label sits in its own fixed-width column OUTSIDE the bar
        # track, not inside it — so it never gets clipped, no matter how
        # small the bar is relative to the largest value in the set.
        rows_html.append(f"""
        <div style="display:flex;align-items:center;gap:10px;margin-bottom:9px;">
          <div style="width:140px;font-size:11px;color:#334155;text-align:right;flex-shrink:0;
                      overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="{label}">{label}</div>
          <div style="flex:1;background:#F1F5F9;border-radius:5px;height:22px;position:relative;">
            <div style="width:{pct:.1f}%;height:100%;background:{color};border-radius:5px;"></div>
          </div>
          <div style="width:52px;font-size:11px;color:#334155;font-weight:700;
                      flex-shrink:0;text-align:left;">{display_val}</div>
        </div>""")

    footnote = ""
    if excluded:
        excl_label, excl_total = excluded
        excl_display = _fmt_money(excl_total) if excl_total >= 1000 else f"{excl_total:,.0f}"
        footnote = (
            f'<div style="margin-top:10px;padding-top:10px;border-top:1px solid #E2E8F0;'
            f'font-size:11px;color:#94A3B8;">+ {excl_display} rows with {value_col_used.replace("_"," ")} '
            f'recorded as "{excl_label}" (excluded above so real categories stay readable)</div>'
        )
    if insight_note:
        footnote += (
            f'<div style="margin-top:8px;font-size:11.5px;color:#0D1B3E;font-weight:600;">'
            f'{insight_note}</div>'
        )

    body = "<div>" + "".join(rows_html) + "</div>" + footnote
    title = title_override or value_col_used.replace("_", " ").title()
    return _wrap_html(body, title, subtitle)


def build_donut_html(
    stored_rows: List[dict], columns: List[str], subtitle: str,
    label_col: Optional[str] = None, value_col: Optional[str] = None,
    exclude_values: Optional[List[str]] = None, title_override: Optional[str] = None,
    insight_note: Optional[str] = None,
) -> Optional[str]:
    labels, values, value_col_used, excluded = _aggregate_generic_rows(
        stored_rows, columns, label_col, value_col, exclude_values
    )
    if not labels:
        return None
    total = sum(values) or 1

    legend_items = []
    for i, (label, val) in enumerate(zip(labels, values)):
        color = PALETTE[i % len(PALETTE)]
        pct = val / total * 100
        legend_items.append(f"""
        <div class="legend-item">
          <span class="legend-swatch" style="background:{color};"></span>
          {label}: {_fmt_money(val) if val >= 1000 else f'{val:,.1f}'} ({pct:.0f}%)
        </div>""")

    # Conic-gradient donut, pure CSS — no JS dependency, always renders
    gradient_parts = []
    cum = 0.0
    for i, val in enumerate(values):
        color = PALETTE[i % len(PALETTE)]
        start = cum / total * 360
        cum += val
        end = cum / total * 360
        gradient_parts.append(f"{color} {start:.1f}deg {end:.1f}deg")
    gradient = ", ".join(gradient_parts)

    body = f"""
    <div style="display:flex;align-items:center;gap:24px;flex-wrap:wrap;">
      <div style="width:140px;height:140px;border-radius:50%;
                  background:conic-gradient({gradient});
                  position:relative;flex-shrink:0;">
        <div style="position:absolute;inset:22px;background:#fff;border-radius:50%;
                    display:flex;align-items:center;justify-content:center;
                    font-size:11px;color:#64748B;text-align:center;">
          {_fmt_money(total) if total >= 1000 else f'{total:,.0f}'}
        </div>
      </div>
      <div class="legend" style="flex-direction:column;gap:8px;">{"".join(legend_items)}</div>
    </div>"""

    if excluded:
        excl_label, excl_total = excluded
        excl_display = _fmt_money(excl_total) if excl_total >= 1000 else f"{excl_total:,.0f}"
        body += (
            f'<div style="margin-top:10px;padding-top:10px;border-top:1px solid #E2E8F0;'
            f'font-size:11px;color:#94A3B8;">+ {excl_display} rows with {value_col_used.replace("_"," ")} '
            f'recorded as "{excl_label}" (excluded above so real categories stay readable)</div>'
        )
    if insight_note:
        body += (
            f'<div style="margin-top:8px;font-size:11.5px;color:#0D1B3E;font-weight:600;">'
            f'{insight_note}</div>'
        )

    title = title_override or value_col_used.replace("_", " ").title()
    return _wrap_html(body, title, subtitle)


# =============================================================================
# PUBLIC ENTRY POINT
# =============================================================================

_ALLOWED_CHART_TYPES = {"funnel", "attainment", "donut", "bar_h"}


def build_chart_html(
    columns: List[str],
    rows: List[dict],
    filters_applied: str = "",
    spec: Optional[dict] = None,
) -> Optional[str]:
    """
    Deterministically build an embeddable ```html chart block from stored
    query results. Returns None if the result isn't chart-worthy (e.g. a
    raw deal list or a single scalar value) — callers should NOT force a
    chart in that case.

    `spec` is an OPTIONAL, pre-validated dict from Claude's choose_chart_spec
    tool call (validated by the caller in main.py against this exact
    result's real column names before it ever reaches here — this function
    does not trust it further). Claude may only choose WHICH chart type and
    WHICH columns to use, and give a title/insight line — every number,
    percentage, and pixel width below is still computed here in Python from
    the real rows, never by Claude. If spec is None or doesn't name a valid
    chart_type, this falls back to the original deterministic
    detect_chart_type() heuristic, unchanged.
    """
    chart_type = None
    label_col = value_col = title_override = insight_note = None
    exclude_values = None

    if spec and spec.get("chart_type") in _ALLOWED_CHART_TYPES:
        chart_type = spec["chart_type"]
        label_col = spec.get("label_column")
        value_col = spec.get("value_column")
        exclude_values = spec.get("exclude_values")
        title_override = spec.get("title")
        insight_note = spec.get("insight_note")

    if chart_type is None:
        chart_type = detect_chart_type(columns, rows)
        if chart_type is None:
            return None

    subtitle = filters_applied[:140] if filters_applied else f"{len(rows)} rows"

    if chart_type == "funnel":
        return build_funnel_html(rows, subtitle, title_override, insight_note)
    if chart_type == "attainment":
        return build_attainment_html(rows, columns, subtitle, title_override, insight_note)
    if chart_type == "donut":
        return build_donut_html(rows, columns, subtitle, label_col, value_col, exclude_values, title_override, insight_note)
    if chart_type == "bar_h":
        return build_bar_html(rows, columns, True, subtitle, label_col, value_col, exclude_values, title_override, insight_note)

    return None


def reply_already_has_chart(reply: str) -> bool:
    """Detect if Claude already included an ```html chart block, so we don't
    double-inject."""
    return bool(re.search(r'```html', reply, re.IGNORECASE))
