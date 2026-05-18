
# =============================================================================
# main.py
# AI-For-Looker — FastAPI Backend
# =============================================================================

# ── Standard Library ──────────────────────────────────────────────────────────
import io
import json
import os
import traceback
from collections import defaultdict
from datetime import date, timedelta
from typing import List, Literal, Optional

# ── Third-Party: Web Framework ────────────────────────────────────────────────
from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, StreamingResponse
from pydantic import BaseModel

# ── Third-Party: Environment & Config ─────────────────────────────────────────
from dotenv import load_dotenv

# ── Third-Party: AI / LLM ─────────────────────────────────────────────────────
import anthropic

# ── Third-Party: Database ─────────────────────────────────────────────────────
import clickhouse_connect

# ── Third-Party: Presentation (PPTX) ─────────────────────────────────────────
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Inches, Pt

# ── Third-Party: PDF Generation ───────────────────────────────────────────────
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    HRFlowable,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


# ── Internal Modules ──────────────────────────────────────────────────────────
from dashboard_queries import get_all_pipeline_metrics, get_filtered_pipeline_metrics
from generate_pptx import build_pptx
from generate_pdf import build_pdf

# =============================================================================
# Environment
# =============================================================================
load_dotenv()


# FastAPI — App Initialisation & Middleware
app = FastAPI(
    title="AI-For-Looker",
    description="Revenue intelligence API powered by OpenAI and ClickHouse.",
    version="2.0.0",
    redirect_slashes=False,
)
 
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],      # tighten to specific origins in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# Serve the frontend HTML at the root
@app.get("/", response_class=HTMLResponse)
def root():
    with open("chat.html", "r") as f:
        return HTMLResponse(content=f.read())

# ----------------------------
# Claude CLIENTS
# ----------------------------
client_ai     = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
client_router = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
_CLAUDE_MODEL = "claude-opus-4-5"


# =============================================================================
# ClickHouse — Lazy Singleton with Auto-Reconnect
# =============================================================================
_click_client: clickhouse_connect.driver.Client | None = None
 
 
def get_click_client() -> clickhouse_connect.driver.Client:
    """
    Return a live ClickHouse client, (re)connecting when necessary. 
    Strategy:
    1. If a cached client exists, ping it with SELECT 1.
    2. On any ping failure, discard the stale client and fall through.
    3. Create a fresh client and verify the connection before returning. 
    Raises:
        RuntimeError: if the connection cannot be established.
    """
    global _click_client
    if _click_client is not None:
        try:
            _click_client.command("SELECT 1")
        except Exception:
            print("⚠️  ClickHouse ping failed — reconnecting…")
            _click_client = None
    if _click_client is None:
        try:
            _click_client = clickhouse_connect.get_client(
                host=os.getenv("CLICKHOUSE_HOST"),
                port=int(os.getenv("CLICKHOUSE_PORT", "8123")),
                username=os.getenv("CLICKHOUSE_USER"),
                password=os.getenv("CLICKHOUSE_PASSWORD"),
                database=os.getenv("CLICKHOUSE_DB", "hs_analytics"),
                secure=False,
                verify=False,
                connect_timeout=10,
                send_receive_timeout=30,
            )
            _click_client.command("SELECT 1")   # confirm the connection is live
            print("✅ ClickHouse connected successfully")
        except Exception as exc:
            _click_client = None
            raise RuntimeError(f"ClickHouse connection failed: {exc}") from exc
 
    return _click_client





# =============================================================================
# summary_endpoint.py
# POST /summary  —  Executive & Analyst styles, with/without filters
#
# Sections:
#   1. Pydantic Models
#   2. ClickHouse Data Fetching
#   3. Metric Extraction  (_extract_computed_metrics)
#   4. Scorecard Builder  (build_scorecards)
#   5. Prompt Builder     (build_prompt)
#   6. FastAPI Endpoint   (POST /summary)
# =============================================================================

# =============================================================================
# 1. Pydantic Models
# =============================================================================
 
class Filters(BaseModel):
    """
    All optional dashboard filters.
    Empty string ("") means "no filter — show everything".
    """
    region:      str = ""
    deal_source: str = ""
    fy:          str = ""
    ai_for_x:    str = ""
    industry:    str = ""
    stage:       str = ""   # "5" | "10" | "20" | "" (all)
 
    def is_empty(self) -> bool:
        """True when no filters are active — triggers global data fetch."""
        return not any([
            self.region, self.deal_source, self.fy,
            self.ai_for_x, self.industry, self.stage,
        ])
 
    def as_query_dict(self) -> dict:
        """Pass directly to get_filtered_pipeline_metrics()."""
        return {
            "region":      self.region,
            "deal_source": self.deal_source,
            "fy":          self.fy,
            "ai_for_x":    self.ai_for_x,
            "industry":    self.industry,
            "stage":       self.stage,

        }
        
class ReportRequestWithSummary(BaseModel):
    style:   Literal["executive", "analyst"] = "executive"
    filters: Filters = Filters()
    summary: str = ""
 
class SummaryRequest(BaseModel):
    """
    Request body for POST /summary.
 
    style   : "executive" → 4-section CRO view, 6-8 bullets, ~30s read
              "analyst"   → 4-section diagnostic view, 10-15 bullets
    filters : leave all fields empty for a global (unfiltered) view
    """
    style:   Literal["executive", "analyst"] = "executive"
    filters: Filters = Filters()
 
# =============================================================================
# 2. ClickHouse Data Fetching
# =============================================================================
 
def fetch_pipeline_data(filters: Filters, ch_client) -> dict:
    """
    Single entry point for all data fetching.
 
    - No filters  →  get_all_pipeline_metrics()     (full business view)
    - Any filter  →  get_filtered_pipeline_metrics() (slice-specific view)
 
    Both functions must return the same dict shape:
        {
          "funnel":            {...},
          "period_attainment": {...},
          "region_source":     [...],
          "stage_velocity":    [...],
          "deals_to_watch":    [...],
          "quarterly_trend":   [...],
          "industry_product":  [...],

        }
    """
    try:
        return get_filtered_pipeline_metrics(ch_client, filters.as_query_dict())
    except Exception as exc:
        raise RuntimeError(f"ClickHouse fetch failed: {exc}") from exc
 
 
# =============================================================================
# 3. Metric Extraction
# =============================================================================
 
def _extract_computed_metrics(metrics: dict, filters: Filters) -> dict:
    """
    Compute ALL derived values (conversion rates, gaps, attainment %, text blocks)
    in one place from the raw ClickHouse response.
 
    Both build_scorecards() and build_prompt() read exclusively from this dict —
    there is a single source of truth and no risk of drift between scorecards
    and the AI summary.
 
    Returns a flat dict with every value scorecards and prompts will need.
    """
    f   = metrics.get("funnel",            {}) or {}
    fc = metrics.get("funnel_conversions", {}) or {}
    p   = metrics.get("period_attainment", {}) or {}
    rs  = metrics.get("region_source",     []) or []
    sv  = metrics.get("stage_velocity",    []) or []
    dw  = metrics.get("deals_to_watch",    []) or []
    qt  = metrics.get("quarterly_trend",   []) or []
    ip  = metrics.get("industry_product",  []) or []
    wl = metrics.get("won_lost_deals", []) or []

    won_deals  = [d for d in wl if d.get("deal_stage") in ("Closed Won", "90% - Deal Desk Review")]
    lost_deals = [d for d in wl if d.get("deal_stage") in ("Prospect Disengaged", "Closed Lost",
                                  "Didn't Qualify")]

    def _fmt_notes(notes):
        notes = str(notes or "").strip()
        return notes[:150] + "…" if len(notes) > 150 else notes

    def _counter_top(items, n=3):
        """Return top-n (value, count) pairs from a list, sorted by frequency desc."""
        counts = defaultdict(int)
        for v in items:
            v = str(v or "").strip()
            if v and v.lower() not in ("", "n/a", "none", "null"):
                counts[v] += 1
        return sorted(counts.items(), key=lambda x: -x[1])[:n]

    def _build_grouped_block(deals, outcome):
        """
        Group deals by (region, deal_source_rollup) and emit a structured block.
        For each group: deal count, total $, top win/loss reasons, top 2-3 deals with details.
        outcome: 'won' | 'lost'
        """
        # ── group deals ───────────────────────────────────────────────────────
        groups = defaultdict(list)
        for d in deals:
            region = str(d.get("region", "Unknown") or "Unknown").strip()
            source = str(d.get("deal_source_rollup", "Unknown") or "Unknown").strip()
            groups[(region, source)].append(d)

        # Sort groups: most deals first
        sorted_groups = sorted(groups.items(), key=lambda x: -len(x[1]))

        label = "WON" if outcome == "won" else "LOST"
        block = f"{label} DEALS — BY REGION & SOURCE:\n"

        if not sorted_groups:
            block += f"  No {outcome} deals found.\n"
            return block

        for (region, source), grp in sorted_groups[:6]:   # cap at 6 groups
            total_amt = sum(float(d.get("amount", 0) or 0) for d in grp)
            count     = len(grp)

            if outcome == "won":
                reasons   = [str(d.get("primary_closed_won_reason_", "") or "").strip() for d in grp]
            else:
                reasons   = [str(d.get("primary_closed_lost_reason", "") or "").strip() for d in grp]

            competitors = [
                str(d.get("competitors", "") or d.get("competition", "") or "").strip()
                for d in grp
            ]
            ai_cats = [str(d.get("ai_for_x", "") or "").strip() for d in grp]

            top_reasons     = _counter_top(reasons,     3)
            top_competitors = _counter_top(competitors, 2)
            top_ai          = _counter_top(ai_cats,     2)

            block += (
                f"\n  ── {region} | {source} ──\n"
                f"  Deals: {count} | Total: ${total_amt/1e6:.2f}M\n"
            )

            if top_reasons:
                reason_strs = ", ".join(f'"{r}" (x{c})' for r, c in top_reasons)
                key = "Win Reasons" if outcome == "won" else "Loss Reasons"
                block += f"  {key}: {reason_strs}\n"

            if top_competitors:
                comp_strs = ", ".join(f'"{c}" (x{n})' for c, n in top_competitors)
                block += f"  Competitors: {comp_strs}\n"

            if top_ai:
                ai_strs = ", ".join(f'"{a}" (x{n})' for a, n in top_ai)
                block += f"  AI Categories: {ai_strs}\n"

            # Top 2–3 representative deals
            top_deals = sorted(grp, key=lambda d: -float(d.get("amount", 0) or 0))[:3]
            for d in top_deals:
                amt         = float(d.get("amount", 0) or 0)
                notes       = _fmt_notes(d.get("won_loss_notes", ""))
                exit_stage  = str(d.get("deal_stage", "N/A") or "N/A").strip()
                close_date  = str(d.get("close_date", "") or "")[:10]
                if outcome == "won":
                    reason  = str(d.get("primary_closed_won_reason_", "") or "N/A").strip()
                    block += (
                        f"    • {d.get('deal_name','N/A')} | ${amt/1e6:.2f}M | "
                        f"AI:{d.get('ai_for_x','N/A')} | Close:{close_date} | "
                        f"Reason:{reason or 'N/A'} | Notes:{notes or 'N/A'}\n"
                    )
                else:
                    reason  = str(d.get("primary_closed_lost_reason", "") or "N/A").strip()
                    comp    = str(d.get("competitors", "") or d.get("competition", "") or "N/A").strip()
                    block += (
                        f"    • {d.get('deal_name','N/A')} | ${amt/1e6:.2f}M | "
                        f"ExitStage:{exit_stage} | AI:{d.get('ai_for_x','N/A')} | "
                        f"Reason:{reason or 'N/A'} | Competitor:{comp} | Notes:{notes or 'N/A'}\n"
                    )

        return block

    won_deals_block  = _build_grouped_block(won_deals,  "won")
    lost_deals_block = _build_grouped_block(lost_deals, "lost")

    print("=== WON/LOST BLOCK ===")
    print(won_deals_block)
    print(lost_deals_block)
    print("=== END ===")
 
    # ── Funnel counts & pipeline amounts ─────────────────────────────────────
    cnt_5   = int(f.get("cnt_5pct",       0) or 0)
    cnt_10  = int(f.get("cnt_10pct",      0) or 0)
    cnt_20  = int(f.get("cnt_20pct",      0) or 0)
    cnt_won = int(f.get("cnt_closed_won", 0) or 0)
    cnt_out = int(f.get("cnt_fallen_out", 0) or 0)
    amt_5   = float(f.get("amt_5pct",     0) or 0)
    amt_10  = float(f.get("amt_10pct",    0) or 0)
    amt_20  = float(f.get("amt_20pct",    0) or 0)
 
    # ── Conversion rates (pre-calculated; LLM must use these, not recompute) ─
    base_5        = int(fc.get("base_5",         0) or 0)
    conv_5_to_10  = int(fc.get("conv_5_to_10",   0) or 0)
    conv_5_to_20  = int(fc.get("conv_5_to_20",   0) or 0)
    conv_5_to_won = int(fc.get("conv_5_to_won",  0) or 0)
    base_10       = int(fc.get("base_10",        0) or 0)
    conv_10_to_20 = int(fc.get("conv_10_to_20",  0) or 0)
    conv_10_to_won= int(fc.get("conv_10_to_won", 0) or 0)
    base_20       = int(fc.get("base_20",        0) or 0)
    conv_20_to_won= int(fc.get("conv_20_to_won", 0) or 0)

    conv_5_10   = round(conv_5_to_10  / base_5  * 100, 1) if base_5  > 0 else 0.0
    conv_10_20  = round(conv_10_to_20 / base_10 * 100, 1) if base_10 > 0 else 0.0
    win_rate    = round(conv_20_to_won / base_20 * 100, 1) if base_20 > 0 else 0.0
    overall_eff = round(conv_5_to_won  / base_5  * 100, 1) if base_5  > 0 else 0.0
 
    # ── Attainment vs L1 target ───────────────────────────────────────────────
    pct_5  = float(p.get("pct_l1_ytd_5",  0) or 0)
    pct_10 = float(p.get("pct_l1_ytd_10", 0) or 0)
    pct_20 = float(p.get("pct_l1_ytd_20", 0) or 0)
 
    actual_5  = int(p.get("ytd_5",  0) or 0)
    actual_10 = int(p.get("ytd_10", 0) or 0)
    actual_20 = int(p.get("ytd_20", 0) or 0)
 
    l1_5  = float(p.get("l1_ytd_5",  0) or 0)
    l1_10 = float(p.get("l1_ytd_10", 0) or 0)
    l1_20 = float(p.get("l1_ytd_20", 0) or 0)
 
    # None (not "N/A") so JSON serialises cleanly to null
    gap_5  = max(0, round(l1_5  - actual_5))  if l1_5  > 0 else None
    gap_10 = max(0, round(l1_10 - actual_10)) if l1_10 > 0 else None
    gap_20 = max(0, round(l1_20 - actual_20)) if l1_20 > 0 else None
 
    mtd_5  = int(p.get("mtd_5",  0) or 0)
    mtd_10 = int(p.get("mtd_10", 0) or 0)
    mtd_20 = int(p.get("mtd_20", 0) or 0)
    qtd_5  = int(p.get("qtd_5",  0) or 0)
    qtd_10 = int(p.get("qtd_10", 0) or 0)
    qtd_20 = int(p.get("qtd_20", 0) or 0)
 
    # ── Stage velocity & health ───────────────────────────────────────────────
    stage_filter_map   = {"5": "5%", "10": "10%", "20": "20%"}
    active_stage_label = stage_filter_map.get(filters.stage, "")
 
    total_red = total_yellow = total_green = 0
    velocity_lines = ""
 
    for s in sv:
        stage_label = s.get("deal_stage", "N/A") or "N/A"
        if active_stage_label and active_stage_label not in stage_label:
            continue
        bench  = s.get("avg_days_benchmark", 0) or 0
        actual = s.get("avg_days_actual",    0) or 0
        green  = int(s.get("green_deals",    0) or 0)
        yellow = int(s.get("yellow_deals",   0) or 0)
        red    = int(s.get("red_deals",      0) or 0)
        total  = green + yellow + red
        red_pct = round(red / total * 100, 1) if total > 0 else 0
        over    = round((actual - bench) / bench * 100, 1) if bench > 0 else 0
        total_red    += red
        total_yellow += yellow
        total_green  += green
        velocity_lines += (
            f"  {stage_label}: total={total} | "
            f"benchmark={bench}d actual={actual}d (+{over}% over) | "
            f"Green={green} Yellow={yellow} Red={red}({red_pct}%)\n"
        )
 
    total_active = total_red + total_yellow + total_green
    pct_red = round(total_red / total_active * 100, 1) if total_active > 0 else 0
 
    # ── Region & Source aggregation ───────────────────────────────────────────
    region_5  = defaultdict(int);   region_10 = defaultdict(int);  region_20 = defaultdict(int)
    source_5  = defaultdict(int);   source_10 = defaultdict(int);  source_20 = defaultdict(int)
    r_l1_5    = defaultdict(float); r_l1_10   = defaultdict(float); r_l1_20  = defaultdict(float)
 
    for r in rs:
        reg = r.get("region",             "Unknown") or "Unknown"
        src = r.get("deal_source_rollup", "Unknown") or "Unknown"
        region_5[reg]  += int(r.get("deals_5",  0) or 0)
        region_10[reg] += int(r.get("deals_10", 0) or 0)
        region_20[reg] += int(r.get("deals_20", 0) or 0)
        source_5[src]  += int(r.get("deals_5",  0) or 0)
        source_10[src] += int(r.get("deals_10", 0) or 0)
        source_20[src] += int(r.get("deals_20", 0) or 0)
        r_l1_5[reg]    += float(r.get("l1_5",  0) or 0)
        r_l1_10[reg]   += float(r.get("l1_10", 0) or 0)
        r_l1_20[reg]   += float(r.get("l1_20", 0) or 0)
 
    # Backfill unattributed so totals always reconcile
    for total_cnt, bucket in [(cnt_5, region_5), (cnt_10, region_10), (cnt_20, region_20)]:
        gap = total_cnt - sum(bucket.values())
        if gap > 0:
            bucket["Unattributed"] = gap
 
    available_regions = sorted([
        reg for reg in region_5
        if reg not in ("Unknown", "Unattributed", "N/A", "")
    ])
 
    # ── Region block (prompt-ready text) ─────────────────────────────────────
    region_block = "DEALS & ATTAINMENT BY REGION:\n"
    for reg in sorted(set(list(region_5) + list(region_10) + list(region_20))):
        a5, a10, a20 = region_5[reg], region_10[reg], region_20[reg]
        l5, l10, l20 = r_l1_5[reg], r_l1_10[reg], r_l1_20[reg]
        att5  = round(a5  / l5  * 100, 1) if l5  > 0 else "N/A"
        att10 = round(a10 / l10 * 100, 1) if l10 > 0 else "N/A"
        att20 = round(a20 / l20 * 100, 1) if l20 > 0 else "N/A"
        region_block += (
            f"  {reg}: "
            f"5%={a5}(L1={l5:.0f},att={att5}%) | "
            f"10%={a10}(L1={l10:.0f},att={att10}%) | "
            f"20%={a20}(L1={l20:.0f},att={att20}%)\n"
        )
    region_block += f"  TOTALS: 5%={cnt_5} | 10%={cnt_10} | 20%={cnt_20}\n"
 
    # ── Source block ──────────────────────────────────────────────────────────
    source_block = "DEALS BY SOURCE:\n" + "".join(
        f"  {src}: 5%={source_5[src]} | 10%={source_10[src]} | 20%={source_20[src]}\n"
        for src in sorted(set(list(source_5) + list(source_10) + list(source_20)))
    )
 
 
    # ── Quarterly trend ───────────────────────────────────────────────────────
    quarter_summary = "QUARTERLY BREAKDOWN:\n"
    for q in qt:
        q_label = q.get("quarter", "N/A") or "N/A"
        a5  = int(q.get("actual_deals",    0) or 0)
        a10 = int(q.get("actual_10_deals", 0) or 0)
        a20 = int(q.get("actual_20_deals", 0) or 0)
        amt = float(q.get("actual_amount", 0) or 0)
        l1  = float(q.get("l1_target_deals", 0) or 0)
        att = float(q.get("pct_l1", 0) or 0)
        c10 = round(a10 / a5 * 100, 1) if a5 > 0 else 0
        c20 = round(a20 / a5 * 100, 1) if a5 > 0 else 0
        quarter_summary += (
            f"  {q_label}: 5%={a5}(L1={l1:.0f},att={att}%) | "
            f"10%={a10}(conv={c10}%) | 20%={a20}(conv={c20}%) | ${amt/1e6:.1f}M\n"
        )
    quarter_summary += f"  YTD: 5%={cnt_5} | 10%={cnt_10} | 20%={cnt_20} | Won={cnt_won}\n"
 
    # ── AI for X & Industry breakdown ────────────────────────────────────────
    ai_10  = defaultdict(int);  ai_20  = defaultdict(int);  ai_amt = defaultdict(float)
    ind_10 = defaultdict(int);  ind_20 = defaultdict(int)
 
    for row in ip:
        ai  = row.get("ai_for_x",             "N/A")   or "N/A"
        ind = row.get("kore_primary_industry", "Other") or "Other"
        d10 = int(row.get("deals_10pct",    0) or 0)
        d20 = int(row.get("deals_20pct",    0) or 0)
        a20 = float(row.get("amount_20pct", 0) or 0)
        ai_10[ai]  += d10;  ai_20[ai]  += d20;  ai_amt[ai] += a20
        ind_10[ind] += d10; ind_20[ind] += d20
 
    ai_summary = "AI FOR X BREAKDOWN:\n" + "".join(
        f"  {ai}: 10%={ai_10[ai]} | 20%={cnt} | ${ai_amt[ai]/1e6:.1f}M\n"
        for ai, cnt in sorted(ai_20.items(), key=lambda x: -x[1])
    )
    ind_summary = "INDUSTRY BREAKDOWN:\n" + "".join(
        f"  {ind}: 10%={ind_10[ind]} | 20%={cnt}\n"
        for ind, cnt in sorted(ind_20.items(), key=lambda x: -x[1])
    )
 
    # ── Flagged deals ─────────────────────────────────────────────────────────
    deals_lines = ""
    for d in dw:
        stage_label = d.get("deal_stage", "") or ""
        if active_stage_label and active_stage_label not in stage_label:
            continue
        amt     = float(d.get("amount", 0) or 0)
        days_in = int(d.get("days_in_current_stage", 0) or 0)
        bench   = int(d.get("avg_days_benchmark",    0) or 0)
        deals_lines += (
            f"  {d.get('deal_name','N/A')} | Stage:{stage_label} | "
            f"Region:{d.get('region','N/A')} | Source:{d.get('deal_source_rollup','N/A')} | "
            f"AI:{d.get('ai_for_x','N/A')} | ${amt/1e6:.2f}M | "
            f"Close:{str(d.get('close_date',''))[:10]} | Health:{d.get('deal_health','N/A')} | "
            f"DaysInStage:{days_in}(benchmark:{bench}d)\n"
        )
    deals_lines = deals_lines or "  No flagged deals.\n"
 
    return dict(
        # raw blobs (kept for scorecards)
        funnel=f, period=p, raw_rs=rs, raw_sv=sv,
        # funnel counts & amounts
        cnt_5=cnt_5,   cnt_10=cnt_10,  cnt_20=cnt_20,
        cnt_won=cnt_won, cnt_out=cnt_out,
        amt_5=amt_5,   amt_10=amt_10,  amt_20=amt_20,
        # conversion rates
        conv_5_10=conv_5_10, conv_10_20=conv_10_20,
        win_rate=win_rate,   overall_eff=overall_eff,
        # attainment
        pct_5=pct_5,   pct_10=pct_10,  pct_20=pct_20,
        actual_5=actual_5, actual_10=actual_10, actual_20=actual_20,
        l1_5=l1_5,     l1_10=l1_10,    l1_20=l1_20,
        gap_5=gap_5,   gap_10=gap_10,  gap_20=gap_20,
        mtd_5=mtd_5,   mtd_10=mtd_10,  mtd_20=mtd_20,
        qtd_5=qtd_5,   qtd_10=qtd_10,  qtd_20=qtd_20,
        # health
        total_active=total_active, total_red=total_red,
        total_yellow=total_yellow, total_green=total_green, pct_red=pct_red,
        # prompt text blocks
        velocity_lines=velocity_lines,
        region_block=region_block,     source_block=source_block,
        quarter_summary=quarter_summary, ai_summary=ai_summary,
        ind_summary=ind_summary,       deals_lines=deals_lines,
        available_regions=available_regions,
        won_deals_block=won_deals_block,
        lost_deals_block=lost_deals_block
    )
 
 
# =============================================================================
# 4. Scorecard Builder
# =============================================================================
 
def build_scorecards(m: dict, filters: Filters) -> dict:
    """
    Build the structured scorecard payload from pre-computed metrics.
 
    Sections (per scope doc §3):
      • stage_overview  — deals, pipeline $, attainment %, gap per stage
      • funnel_health   — conversion rates + overall efficiency
      • period_cadence  — MTD / QTD / YTD actuals vs L1 + gap
      • deal_health     — Red / Yellow / Green breakdown
 
    Filter behaviour:
      - stage filter active  → returns only that stage's card in stage_overview
      - no stage filter       → returns all four stage cards
      - all counts/$ reflect whatever slice was fetched from ClickHouse
    """
 
    def status(pct: float) -> str:
        if pct >= 100: return "on_track"
        if pct >= 60:  return "at_risk"
        return "below_target"
 
    # ── Stage overview ────────────────────────────────────────────────────────
    all_stages = [
        {
            "stage":      "5% IQM Held",
            "stage_key":  "5",
            "deals":       m["cnt_5"],
            "pipeline_m":  round(m["amt_5"] / 1e6, 1),
            "ytd_pct":     m["pct_5"],
            "ytd_actual":  m["actual_5"],
            "l1_target":   m["l1_5"],
            "gap":         m["gap_5"],        # None if no target set
            "status":      status(m["pct_5"]),
        },
        {
            "stage":      "10% Discovery",
            "stage_key":  "10",
            "deals":       m["cnt_10"],
            "pipeline_m":  round(m["amt_10"] / 1e6, 1),
            "ytd_pct":     m["pct_10"],
            "ytd_actual":  m["actual_10"],
            "l1_target":   m["l1_10"],
            "gap":         m["gap_10"],
            "status":      status(m["pct_10"]),
        },
        {
            "stage":      "20%+ Qualified",
            "stage_key":  "20",
            "deals":       m["cnt_20"],
            "pipeline_m":  round(m["amt_20"] / 1e6, 1),
            "ytd_pct":     m["pct_20"],
            "ytd_actual":  m["actual_20"],
            "l1_target":   m["l1_20"],
            "gap":         m["gap_20"],
            "status":      status(m["pct_20"]),
        },
        {
            "stage":      "Closed Won",
            "stage_key":  "won",
            "deals":       m["cnt_won"],
            "pipeline_m":  None,             # no pipeline $ for closed deals
            "ytd_pct":     m["win_rate"],
            "ytd_actual":  m["cnt_won"],
            "l1_target":   None,
            "gap":         None,
            "status":      None,
        },
    ]
 
    stage_filter = filters.stage
    stage_overview = (
        [s for s in all_stages if s["stage_key"] == stage_filter]
        if stage_filter else all_stages
    )
 
    # ── Funnel health ─────────────────────────────────────────────────────────
    funnel_health = {
        "conv_5_to_10_pct":   m["conv_5_10"],
        "conv_10_to_20_pct":  m["conv_10_20"],
        "win_rate_pct":        m["win_rate"],
        "overall_efficiency":  m["overall_eff"],
        "fallen_out":          m["cnt_out"],
    }
 
    # ── Period cadence ────────────────────────────────────────────────────────
    period_cadence = {
        "MTD": {
            "deals_5":  m["mtd_5"],
            "deals_10": m["mtd_10"],
            "deals_20": m["mtd_20"],
        },
        "QTD": {
            "deals_5":  m["qtd_5"],
            "deals_10": m["qtd_10"],
            "deals_20": m["qtd_20"],
        },
        "YTD": {
            "deals_5":  m["actual_5"],  "l1_5":  m["l1_5"],  "gap_5":  m["gap_5"],
            "deals_10": m["actual_10"], "l1_10": m["l1_10"], "gap_10": m["gap_10"],
            "deals_20": m["actual_20"], "l1_20": m["l1_20"], "gap_20": m["gap_20"],
        },
    }
 
    # ── Deal health ───────────────────────────────────────────────────────────
    deal_health = {
        "total_active": m["total_active"],
        "green":        m["total_green"],
        "yellow":       m["total_yellow"],
        "red":          m["total_red"],
        "red_pct":      m["pct_red"],
    }
 
    return {
        "stage_overview":  stage_overview,
        "funnel_health":   funnel_health,
        "period_cadence":  period_cadence,
        "deal_health":     deal_health,
        "filter_applied":  filters.model_dump(),
    }
 
 
# =============================================================================
# 5. Prompt Builder
# =============================================================================
 
# Injected into every prompt so the LLM never confuses abbreviated stage codes.
_STAGE_REFERENCE = """
STAGE NAME REFERENCE (always use the full label, never abbreviations):
  "1%"  → "1% - IQM Scheduled"      "5%"  → "5% - IQM Held"
  "10%" → "10% - Discovery"          "20%" → "20% - Solution"
  "30%" → "30% - Proof"              "40%" → "40% - Proposal"
  "60%" → "60% - Price Negotiation"  "75%" → "75% - Contract Review"
  "Closed Won" = final won stage
NOTE: APAC is NOT a valid region. ISEA = India / Southeast Asia.
""".strip()
 
 
# ── Style-specific instructions ───────────────────────────────────────────────
 
_EXECUTIVE_INSTRUCTIONS = """\
You are a senior GTM strategist and trusted advisor to the CEO/CRO at Kore.ai.
You have read every number before writing a single word.
You think in business outcomes, risk, and decisions — not templates.
 
FILTER COMPLIANCE — READ THIS FIRST, BEFORE ANYTHING ELSE:
If filters are active (stage, region, source, AI for X, industry), you must:
  - Restrict EVERY bullet in EVERY section to ONLY the filtered slice
  - Never mention stages, regions, or sources outside the active filter
  - If stage=20% is active, NEVER reference 5% or 10% stage numbers
  - If region=North America is active, NEVER reference ISEA, Middle East, JAPAC, etc.
  - If source=Hyperscaler is active, NEVER reference BDR, Marketing, AE Outbound, etc.
  - The only exception: you may contrast the filtered slice against global ONLY when
    explicitly labeling both (e.g. "NA at 26.1% vs global 64.6%")
  - Active filters are stated at the top of the data block — read them before writing
TONE RULE:
You are advising, not instructing. Never tell the CRO what "must" be done or what
"the CRO must prioritize." You are a peer with data, not a consultant with a deck.
Write as if you are talking to the person, not writing a report about them.
BAD: "The CRO must prioritize sourcing from alternative channels."
GOOD: "The Hyperscaler dependency is the single biggest structural risk here —
  if that source dries up, there is no pipeline backstop."
 
FORMAT RULE:
Write every section as bullet points. No prose paragraphs.
Each bullet = one complete thought. Max 2 lines per bullet — split if longer.
Every bullet must contain a number AND an implication. No exceptions.

BOLD RULE: Every bullet must start with a bold opening phrase (3-8 words) 
followed by an em dash —
Example: **The 5% stage is critically undersourced** — 156 deals at 8% attainment means...
The bold phrase is the headline; the rest of the bullet is the supporting evidence.
Never write a bullet that starts without bold text.
 
STRUCTURE ANTI-PATTERN RULE — CRITICAL:
Do NOT use fixed sentence templates. The bullets should read naturally from the data,
not like they were poured into a mold.
 
These templates are BANNED — if you catch yourself using them, rewrite:
  BANNED: "Because [X], we need to [Y], which addresses [Z]."
  BANNED: "Because [X], we must [Y], which mitigates [Z]."
  BANNED: "Because [X], we should [Y], which will [Z]."
 
Focus area bullets should sound like a strategist making a call, not filling a form.
Each focus area should be written differently — varied sentence structure, varied framing —
because each action addresses a different kind of problem.
 
BAD focus areas (all same template, all vague):
  "• Because 420 deals are short at 5%, we need aggressive sourcing to fill the pipeline,
  which addresses the immediate revenue risk."
  "• Because 0% of deals convert 20%→Won, we must enhance qualification criteria at 20%,
  which mitigates future revenue loss."
 
GOOD focus areas (varied, specific, natural):
  "• The 5% gap of 420 deals can't close organically — the math requires 3x current run rate,
  which means this needs a campaign decision this week, not a coaching conversation."
  "• Zero closed won despite deals reaching 20% points to a closing motion problem, not a
  pipeline problem — the fix is deal desk review of every 20%+ deal over 60 days in stage,
  not more top-of-funnel activity."
 
Generate a summary using EXACTLY these 6 sections (bold markdown headers):
 
**Pipeline Health at a Glance**
**Funnel Velocity & Conversion**
**Revenue Position & Closed Won**
**Deal Quality & Risk Signals**
**What's Working**
**Win/Loss Intelligence**
**Focus Areas & Recommended Actions**
 
═══════════════════════════════════════════════════════════
SECTION GUIDANCE
═══════════════════════════════════════════════════════════
 
**Pipeline Health at a Glance**
 
First bullet: one sentence state of the pipeline. What's the headline?
Not a list of gaps — one coherent read on where the business stands.
 
Then one bullet per active stage (ONLY stages in the active filter, or all three if no
stage filter is set). Each bullet must earn its place — state the gap AND what it means
by year-end, expressed as a run-rate problem. If the math is alarming, say it is alarming.
If the gap is manageable, say that too. Don't use the same sentence structure for each stage.
 
For regions: if a region filter is active, cover only that region.
If no region filter, name the outliers — best and worst — and explain the business
implication of that gap. Don't list all regions with numbers. Name the story.
 
For sources: if a source filter is active, cover only that source.
If no source filter, name the source that's dragging conversion and quantify the drag.
 
Final bullet: a "so what" that feels like a natural conclusion from everything above —
not a mandatory closing line, but whatever the data actually points to.
 
---
 
**Funnel Velocity & Conversion**
 
First bullet: what's the funnel's overall health in one read?
Is the problem entry volume, mid-funnel stall, closing failure, or all three?
 
Then one bullet per gate — but only gates relevant to the active filter.
If stage=20% is active, focus on the 20%→Won gate. Don't recap 5→10 or 10→20.
 
Each gate bullet integrates: rate + benchmark gap + time in stage + what that combination
reveals about behavior + consequence. The combination is the signal — never report
rate alone or time alone.
 
Diagnostic logic to internalize (not to copy):
  Low rate + high time → stalling (qualification or IQM issue)
  Low rate + normal time → worked but lost (pitch, fit, champion issue)
  High rate + high time → slow progression (capacity or prioritization bottleneck)
  High rate + normal time → gate is healthy — say what's working and protect it
 
For anomalies (100% conversion, 0% win rate, undefined days):
  Take a position — more likely real or data artifact? Say why.
  Don't hedge with "could be A or could be B" — that's not analysis.
 
Final bullet: funnel yield — for every 100 deals entering, how many close?
What does that mean for required pipeline volume?
 
---
 
**Revenue Position & Closed Won**
 
Open with the closed won number — no softening.
The exact closed won count is in the data block under "Closed Won : X deals" — use that number.
DO NOT say "zero" unless the data block explicitly shows "Closed Won : 0 deals".
 
Project forward: what does EOQ look like at current rate? How many deals must close
in remaining weeks? If the answer is "we will miss," say so directly.

If FY STATUS = CLOSED:
  - DO NOT calculate required deals/month
  - DO NOT assume time remaining
  - Frame all gaps as final shortfall

If FY STATUS = IN PROGRESS:
  - You may compute required pace vs current run rate

If FY STATUS = CLOSED:
  - Do not compute required/month
  - Do not mention pacing
  - Treat gap as final shortfall
Connect the revenue gap to a specific upstream cause — name the gate that's responsible.
"This isn't a sourcing problem — it's a closing motion problem" or vice versa.
The framing should follow from the data, not from a template.
 
---
 
**Deal Quality & Risk Signals**
 
Frame the overall risk in one opening bullet — is risk concentrated, diffuse, or structural?
 
Then cover the signals that are actually present in the data:
  - Fallen-out deals: where did they exit? Exit stage = where the funnel is leaking.
  - Red-flagged deals: apply to revenue, not just count. What does Red % mean for EOQ?
  - Concentration risk: if one source or region holds all the pipeline, name the failure mode.
  - Lost deal patterns: if notes reveal a theme, name it with count + $ + exit stage.
Only cover signals that are actually present. If Red deals are 0%, don't manufacture
a risk signal from absence — "no high-risk deals currently" is fine if that's the truth,
but follow it with what the absence actually implies (stagnation? small sample? data gap?).
 
Do not repeat risks already stated in other sections. Each bullet must add new information.
 
---
 
**What's Working**
 
If there are genuine bright spots, cover them — with the number, vs benchmark, and
what it means to protect or scale. Interrogate each: is it actually strong or does it
only look good because the denominator is small or the bar is low?
 
WHAT'S WORKING RULE:

- Identify relative strengths, not just absolute wins
- A metric can be "working" if:
    • It is stronger than other stages/regions/sources
    • It shows better conversion relative to funnel average
    • It represents meaningful volume even if below benchmark

- ONLY say "no genuine bright spots" if:
    • ALL stages are severely underperforming (<50% attainment)
    • AND conversion is below benchmark across ALL gates
    • AND no region/source meaningfully outperforms others

- Prefer:
    "No metrics are above benchmark, but X is relatively stronger and worth protecting"
 
---
---

**Win/Loss Intelligence**

The data is structured by Region + Deal Source groups (e.g. "North America | BDR", "Middle East | Hyperscaler").
Each group shows: deal count, total $, top reasons (with frequency), competitors, AI categories, and 2–3 representative deals.

Read across ALL groups before writing. Extract the cross-group signal, not per-deal details.

For won deal groups:
  - Which region+source combination produces the most wins? Name count + total $.
  - What win reason appears most frequently across groups? Name it explicitly (e.g. "Superior Functionality (x4)").
  - If a competitor appears repeatedly in won deals, name them — it's a replicable competitive advantage.
  - Name the AI category that dominates won deals if one stands out.

For lost deal groups:
  - Which region+source combination has the highest loss concentration? Name count + total $ at risk.
  - What is the primary loss reason across groups? Be specific — "Didn't Qualify at Discovery stage" not "qualification issues".
  - If the same competitor appears across 2+ loss groups, flag it as a competitive pattern, not a one-off.
  - Name the exit stage — it reveals WHERE in the funnel the leak is.

Contrast bullet (mandatory): what do winning region+source combinations have that losing ones consistently lack?
One sentence, grounded in the data — region, source, reason, AI category, or size. This is the most actionable line.

If a group has only 1 deal, do not call it a pattern — say "single data point."
If reasons are mostly blank, say so and work only from what IS present.

---
 
**Focus Areas & Recommended Actions**
 
Exactly 3 bullets. Each one must be specific to THIS data — not generic pipeline advice.
 
The three bullets should address three different kinds of problems (e.g. volume, conversion,
process) and should be written with varied sentence structure. They should not all start
with the same word or follow the same grammatical pattern.
 
Each focus area must:
  - Name a specific number or pattern from the data
  - Name a specific mechanism or action (not a category like "improve sourcing")
  - Say what it addresses or prevents (without using "which addresses" as a template)
Write these as a strategist making calls — direct, opinionated, varied.
The tone should differ by the urgency of the problem:
  - An acute crisis (0 closed won) sounds urgent and direct
  - A structural risk (single-source dependency) sounds measured and forward-looking
  - A process fix (deal desk review) sounds operational and specific
BANNED focus area phrases:
  "Because X, we need to Y, which addresses Z" (and all variations)
  "implement aggressive sourcing strategies"
  "enhance qualification criteria"
  "conduct a thorough review"
  "address the immediate revenue risk"
  "mitigate future revenue loss"
  "stem the current pipeline leakage"
  These are placeholders, not actions. Replace every one with something specific to the data.
 
═══════════════════════════════════════════════════════════
HARD RULES
═══════════════════════════════════════════════════════════
 
1. FILTER COMPLIANCE IS NON-NEGOTIABLE.
   Active filters define the universe of this summary.
   A filtered summary that references out-of-scope data is wrong, not just imprecise.
2. NO FIXED TEMPLATES. Bullets must vary in structure, opening word, and framing.
   If three bullets in a row start with "Because" — rewrite all three.
   If two bullets have identical grammatical structure — rewrite one.
3. NO INSTRUCTING THE CRO. You advise, you don't assign.
   "The CRO must..." / "Leadership should..." / "The team needs to..." — all banned.
   Rephrase as a finding or a call: "The math here requires a campaign decision, not
   a coaching conversation" not "The CRO must launch a sourcing campaign."
4. EVERY BULLET = number + implication. No naked numbers. No implications without numbers.
5. MAX 10 BULLETS TOTAL. Weight toward sections with the most signal.
   A section with nothing urgent = 1 bullet. Don't pad.
6. NEVER sum deal counts across stages. Each stage is an independent cohort.
7. No "Next Steps" framing outside of Focus Areas.
8. Each section must teach something new. If a CRO reads a section and only knows
   what the dashboard already showed them — rewrite it.
""".strip()
 
_ANALYST_INSTRUCTIONS = """\
You are a senior GTM analyst embedded in the sales org at Kore.ai.
You have internalized every number before writing a single word.
You think in patterns, cohorts, and levers — not problems and failures.

Your goal is a diagnostic that RevOps, Sales Ops, and deal desk can act on.
The difference between a good analyst summary and a bad one:
  BAD: "BDR is responsible for 4.7 points of drag on total funnel efficiency."
  GOOD: "BDR converts 4.7pts below the funnel average at 5→10 — closing that gap
  to the Marketing benchmark would add ~X deals to the 10% cohort without new sourcing."

  BAD: "The 132.5pt gap between ISEA and NA indicates high revenue risk concentration."
  GOOD: "ISEA's 158.6% attainment vs NA's 26.1% means ISEA's playbook — source mix,
  AI category focus, deal cadence — is the most replicable asset in the portfolio;
  NA has 982 deals at 5% to work with, so the gap is a conversion question, not a
  volume question."

FRAMING PRINCIPLE — THIS GOVERNS EVERY BULLET:
  Data is neutral. Your job is to find the mechanism and the lever, not to assign blame.
  Every gap is a question: "what would have to change to close this, and is that achievable?"
  Every underperformance is a diagnostic: "is this a volume problem or a conversion problem?"
  These need different interventions — naming which one it is IS the analysis.

ANALYST VOICE RULES:
  - Never frame a metric as "a problem" without naming the specific lever that addresses it
  - Never describe a gap without stating whether it's closable at current trajectory
  - "Risk" language is allowed only when the risk is specific and quantified
    ALLOWED: "At current 20→Won rate of 8.8%, the 676-deal gap requires closing 75 additional
    deals from existing 20% pipeline — only possible if win rate improves to ~18%"
    BANNED: "This indicates significant revenue risk" (vague, no lever)
  - Source comparison bullets must answer: "what should we do differently with this source?"
    not "this source is dragging the funnel"
  - Regional comparison bullets must answer: "what's replicable from the leader?"
    not "the laggard represents a structural weakness"

FORMAT RULE — THIS IS MANDATORY:
Write EVERY section as bullet points. No prose paragraphs. No walls of text.
Each bullet = one complete thought: one number + one mechanism + one lever or implication.
A bullet that is more than 2 lines long must be split into two bullets.
A bullet that contains no number must be rewritten or deleted.
A bullet that names a gap without naming the mechanism behind it must be rewritten.

BOLD RULE: Every bullet must start with a bold opening phrase (3-8 words) 
followed by an em dash —
Example: **The 5% stage is critically undersourced** — 156 deals at 8% attainment means...
The bold phrase is the headline; the rest of the bullet is the supporting evidence.
Never write a bullet that starts without bold text.

BULLET QUALITY RULE — THE ANALYST STANDARD:
Every bullet must answer three questions in sequence:
  1. What is the number?
  2. What behavior or process produces exactly that number?
  3. What specific lever exists to change it, or what does it predict about what comes next?

BAD: "• 10% stage: 1,347 deals | Target: 2,478 | Attainment: 54.4% | Gap: 1,131"
BAD: "• 10% stage is 1,131 deals short — this is a structural issue requiring immediate attention."
GOOD: "• 10% stage is 1,131 deals short at 54.4% — since 10% is fed by 5→10 progression,
  improving that conversion rate from 61.7% to 70% (benchmark midpoint) would add ~300 deals
  to 10% without a single new IQM; the remaining gap requires sourcing acceleration."

Generate a diagnostic analysis using EXACTLY these 7 sections (bold markdown headers):

**Overall Pipeline Snapshot**
**Funnel Conversion & Velocity Analysis**
**Regional & Source Performance**
**AI for X & Deal Category Breakdown**
**Stage Velocity & Stagnation**
**Win/Loss Intelligence**
**Focus Areas & Highest-Leverage Interventions**
**Root Cause Hypotheses**

═══════════════════════════════════════════════════════════
SECTION GUIDANCE
═══════════════════════════════════════════════════════════

**Overall Pipeline Snapshot**

First bullet: one headline that characterizes the pipeline state — frame it as
"where the biggest lever is" not "what's broken." Is the primary opportunity in
entry volume, conversion improvement, or closing motion? Name the one with the
most addressable gap.

Then one bullet per stage (5%, 10%, 20%). Each bullet must contain:
  - attainment % AND gap in deals
  - the specific conversion or sourcing lever that would close the gap (or the
    fraction of it that's mechanically achievable)
  - whether that lever has been hit in any prior quarter this year

Frame the stage bullets as: "here's the gap, here's what it would take, here's
whether that's realistic" — not "here's how far behind we are."

BAD bullet:
  "• 20% stage: 676 deals short at 48.4% — this gap is compounding and requires
  immediate intervention to avoid further decline."
GOOD bullet:
  "• 20% stage: 676 deals short at 48.4% — but 20% is downstream of 10%, so the
  fastest lever is improving 10→20 conversion from 42.9% toward the 50% benchmark;
  a 7pt lift on the existing 10% pool adds ~90 deals to 20% without any new sourcing."

Then one bullet for regional performance — frame it as: what does the leader's
playbook look like, and how much of it is replicable in underperforming regions?
Name the specific mechanism (source mix, AI category, deal size) not just the gap.

Then one bullet for source performance — frame it as: which source is closest to
its conversion benchmark, and what would it take to bring the lagging source there?
Use the "closing the gap" calculation, not the "drag" calculation.

Final bullet: realistic recovery framing — given the levers available, what's the
earliest quarter where a meaningful recovery is possible, and what would have to
be true for it to happen?

---

**Funnel Conversion & Velocity Analysis**

First bullet: characterize the entire funnel in one diagnostic read.
Where is the highest-leverage conversion gate — the one where a benchmark-level
improvement would have the largest downstream impact on closed won deals?

Then one bullet per gate (5→10, 10→20, 20→Won). Each bullet must contain:
  - conversion rate + gap from benchmark (in percentage points)
  - avg days vs benchmark (as a ratio, e.g. "2.9x benchmark")
  - mechanism: what does the rate + time combination reveal about rep behavior?
  - the specific lever: what process change or intervention targets exactly that behavior?

Use this diagnostic matrix (internalize, never copy literally):
  Low rate + high time → deals stalling — qualification or IQM quality issue
    LEVER: IQM quality review, next-step accountability process
  Low rate + normal time → deals worked but lost — pitch, fit, or champion issue
    LEVER: champion mapping, competitive positioning, value prop refinement
  High rate + high time → advances slowly — capacity or process bottleneck
    LEVER: deal prioritization, capacity planning
  High rate + normal time → gate is healthy — protect what's working
    LEVER: document and replicate

BAD bullet:
  "• 10→20 gate: 42.9% conversion — 7pts below benchmark — with deals spending 80d
  vs 28d benchmark. Low rate AND high time indicate reps investing in non-progressing deals."
GOOD bullet:
  "• 10→20 gate converts at 42.9% (7pts below benchmark) with deals spending 80d vs
  28d benchmark (2.9x) — low rate AND high time together means reps are cycling on
  deals without a defined advancement path; the lever is structured next-step accountability
  after Discovery calls, not more top-of-funnel activity."

For any anomaly (rate at 100% or near-zero): take a position.
State which explanation is more likely and why — don't hedge.

Final bullet: overall funnel efficiency framed as an opportunity calculation.
"At current 5%→Won efficiency of X%, producing 100 closed won deals requires Y
deals at 5% — the fastest path to improving that ratio is [gate] where a [X]pt
lift has the highest mechanical leverage."

CLOSED WON LANGUAGE RULE:
- If Closed Won = 0 → say "no closed won deals recorded yet this period"
- If Closed Won > 0 but below expectation → state the number and what win rate
  improvement would be needed to reach target
- NEVER use words like "absence," "lack of," or frame it as catastrophic without
  stating what conversion improvement would change the trajectory

---

**Regional & Source Performance**

First bullet: the regional performance spread — but frame it as a replication
opportunity, not a risk statement.
"ISEA's 158.6% at 20% vs NA's 26.1% creates a natural experiment: ISEA's source
mix, AI category focus, and deal cadence are the most studied playbook in the
portfolio — the question is which elements transfer to NA's 982-deal 5% pipeline."

Then one bullet for the top performer: WHY do they lead?
Name the specific mechanism — source mix, AI category concentration, deal size,
regional demand pattern. Frame it as: "this is what makes it replicable / not fully
replicable elsewhere, and here's the transferable piece."

Then one bullet for the lowest performer: diagnose as volume problem OR conversion problem.
These need different interventions — naming which one it is IS the analysis.
Volume problem = too few deals entering → lever is sourcing, campaign, or IQM cadence.
Conversion problem = deals exist but don't advance → lever is rep process, champion quality,
or deal desk review of stalled deals.
NEVER say "this indicates a structural weakness" — say "this is a [type] problem,
which means the intervention is [specific action]."

Then one bullet for the lowest-converting source:
Frame it as: "closing the gap between [source] and the benchmark would add X deals
to the funnel — here's what's driving the conversion gap."
NOT: "[source] is responsible for X points of drag."

Then one bullet for the highest-converting source:
State why it converts well, whether it's scalable, and what "protecting" it means
operationally — is it at risk of being disrupted by pipeline creation pressure?

---

**AI for X & Deal Category Breakdown**

First bullet: characterize the PMF landscape as a portfolio — which categories
show strong conversion (double down), which show stall (investigate), and which
show early signal (develop)?

Then one bullet per AI for X category with meaningful volume (5+ deals at 10%).
Each bullet must contain:
  - 10→20 conversion rate (computed from counts)
  - immediate interpretation: strong fit signal OR fit risk signal
  - the specific operational lever: "protect the closing motion" / "investigate
    champion quality" / "refine value prop for this segment"

BAD bullet:
  "• AI for Process: 38.9% conversion — fit risk signal, necessitating operational
  focus on improving deal progression."
GOOD bullet:
  "• AI for Process converts 38.9% of Discovery deals to Solution ($10.6M pipeline) —
  11pts below AI for Service; the gap is most likely unclear ROI articulation at
  Discovery, not product fit — testable by reviewing whether lost deals in this
  category cite 'Didn't Qualify' vs 'No Budget' as exit reason."

For the weakest converting category with meaningful volume:
Name the most likely root cause (weak champions, unclear ROI, product-pain mismatch)
and what single data point from the won/loss notes would confirm or refute it.

---

**Stage Velocity & Stagnation**

First bullet: what does the health distribution predict about deal advancement
over the next 30-60 days? Apply Red % to deal count AND pipeline value.
Frame it as: "X deals representing $YM are currently past their advancement
benchmark — direct intervention on these deals is the highest-ROI activity
available to the deal desk right now."

Then one bullet per stage with significant Red concentration.
For each: explain what Red at that specific stage means operationally AND
what the targeted intervention looks like.
  Red at 5% → IQM advancement broken → lever: IQM follow-up cadence, call quality review
  Red at 10% → Discovery stalling, no next steps → lever: structured next-step accountability
  Red at 20% → closing motion absent → lever: deal desk review of all 20%+ deals >60d in stage
  (Use the stage-specific lever, not generic language)

Then one bullet for fallen-out deals: count + exit stage + what the exit stage
tells us about where to focus retention effort.
Frame it as: "X deals exited at [stage] — this is where the funnel is leaking,
and it aligns with / contradicts the conversion rate story because [reason]."

Then one bullet for avg days vs benchmark: state as a ratio, then name the rep
behavior that produces exactly that number AND the process that would change it.
"2.9x benchmark at 10% means reps have no defined next step after Discovery —
the fix is a mandatory 'next meeting booked before close of call' standard,
not more coaching."

---

**Win/Loss Intelligence**

The data is structured by Region + Deal Source groups.
Each group shows: deal count, total $, top reasons (with frequency), competitors,
AI categories, and 2–3 representative deals.

Read across ALL groups before writing. Extract the cross-group signal.

For won deal groups — frame as: what's replicable?
  - Which region+source combination produces the most wins? Name count + total $.
    Frame it as: "this combination is the model — here's what drives it."
  - What win reason appears most frequently? Name it explicitly (e.g. "Superior
    Functionality (x41)"). Frame it as a competitive strength to protect and message.
  - If a competitor appears repeatedly in won deals, name them and frame it as:
    "we have a documented win pattern against [competitor] — this should be
    systematized into the competitive playbook."
  - Name the AI category dominating won deals and what it says about where PMF is strongest.

For lost deal groups — frame as: where's the recoverable opportunity?
  - Which region+source combination has the highest loss concentration?
    Frame it as: "X deals totaling $YM exited here — at the current win rate,
    improving qualification earlier in this source/region would be higher-leverage
    than adding new deals to the top."
  - What is the primary loss reason? Be specific. Frame it as a process gap,
    not a verdict: "Didn't Qualify at Discovery" means qualification criteria
    aren't surfacing fit signals early enough, not that the deals were bad.
  - If the same competitor appears in 2+ loss groups, frame it as:
    "competitive pattern against [competitor] — a structured battle card for
    this matchup would directly address X deals worth $YM."
  - Name the exit stage and connect it to the conversion analysis already done.

Contrast bullet (mandatory): what does the winning combination have that the
losing combination consistently lacks? One sentence, grounded in the data.
Frame it as an actionable difference, not a judgment.

If a group has only 1 deal, say "single data point — insufficient for pattern."
If reasons are mostly blank, say so and work only from what IS present.

---
---

**Focus Areas & Highest-Leverage Interventions**

Exactly 3 bullets. Each one must be the output of the analysis above —
not a restatement of a gap, but a specific intervention derived from the
mechanism identified in an earlier section.

Each bullet structure:
  [The lever] — [the specific data that makes this the highest-leverage action] —
  [the expected outcome if the lever is pulled, expressed as a number or rate change] —
  [how to know if it's working within 30 days]

The three bullets must address three different parts of the funnel or org:
one conversion lever, one sourcing or volume lever, one deal-level or process lever.
They should not all be about the same stage or the same team.

BAD focus area (gap restatement):
  "• 10→20 conversion at 42.9% is 7pts below benchmark and needs improvement through
  better Discovery engagement and champion mapping to advance deals."

GOOD focus area (lever with outcome):
  "• Mandatory next-step booking at close of every Discovery call would directly
  target the 10→20 stall — 943 Red deals at 10% represent $265.5M sitting past
  their advancement benchmark; even a 10pt conversion lift on stalled deals adds
  ~94 deals to 20% from existing pipeline, no new sourcing required.
  Leading indicator: Red deal count at 10% should decline within 3 weeks of
  implementing the standard."

Each focus area must:
  - Name a number from the data (deal count, conversion rate, pipeline $, or attainment %)
  - Name the specific mechanism or process change (not a category like "improve sourcing")
  - State the expected outcome as a quantified change (deals added, rate improvement,
    pipeline $ recovered)
  - Name one leading indicator that would confirm it's working within 30 days

BANNED focus area language:
  "implement aggressive sourcing strategies"
  "enhance qualification criteria"
  "conduct a thorough review"
  "address the revenue gap"
  "focus on improving deal progression"
  "requires immediate attention"
  These are categories, not interventions. Every banned phrase must be replaced
  with: a specific action + the number it moves + how you'd know it's working.

The tone should be operational and precise — like a RevOps analyst handing
a prioritized work order to the deal desk, not a consultant presenting a slide.

**Root Cause Hypotheses**

Write exactly 3–4 hypotheses. Each hypothesis = ONE bullet (up to 3 lines).
Do NOT use sub-bullets or nested structure inside hypotheses.

Each hypothesis bullet must contain in sequence:
  [Name + Confidence] — [2-3 data points that together support it] —
  [mechanism: what behavior produces exactly this pattern] —
  [the lever: what specific change would test or address this hypothesis] —
  [what would need to be true for this hypothesis to be wrong]

Frame hypotheses as explanations that, if true, would point to a specific lever —
not as indictments of a team or function.

BAD:
  "• Champion Quality Issues [High confidence] — lost deals frequently cite 'Didn't
  Qualify' indicating weak champions; this suggests a need for better qualification."

GOOD:
  "• Discovery Next-Step Accountability Gap [High confidence] — 10→20 converts at
  42.9% while spending 2.9x benchmark time; if it were fit alone, time would be
  normal but conversion low — both degraded together means reps are re-running
  Discovery rather than advancing; testable by checking if deals with a booked
  follow-up at close of 10% call advance at materially higher rates."

At least one hypothesis must be grounded in won/loss note patterns.
At least one hypothesis must identify a replicable positive (why something IS working).
Hypotheses should be mutually exclusive where possible.

═══════════════════════════════════════════════════════════
HARD RULES
═══════════════════════════════════════════════════════════

1. BULLET FORMAT IS MANDATORY. No prose paragraphs anywhere.
2. LEVER RULE: Every gap bullet must name a specific lever. A bullet that describes
   a gap without naming what would close it is incomplete — rewrite it.
3. FRAMING RULE: "Drag," "structural weakness," "risk concentration," and "requires
   immediate intervention" are banned unless immediately followed by the specific
   lever that addresses them. Without the lever, these phrases are alarm without insight.
4. NEVER sum deal counts across stages. Each is an independent cohort.
5. 14–18 bullets TOTAL across all 7 sections.
   Root Cause Hypotheses: exactly 3–4 bullets.
   Win/Loss Intelligence: 3–5 bullets.
   No other section under 1 bullet.
6. Every comparison must produce an actionable conclusion, not a contrast.
   "NA is 26.1% and ISEA is 158.6%" = contrast (banned).
   "ISEA's 158.6% vs NA's 26.1% means ISEA's source mix is the most available
   replication lever — NA has the deal volume to work with, the gap is conversion" = conclusion.
7. Actively look for what's working. A section that contains only gap analysis
   has failed the analyst standard — every section should surface at least one lever
   or bright spot alongside the diagnosis.
8. No "Next Steps" or "Recommendations" section.
   The analyst diagnoses. The executive decides. Actions belong in the executive summary.
9. If filters are active, restrict ALL analysis to that slice.
10. Write like you're briefing a RevOps VP who will ask "so what do we do with that?"
    after every bullet. Pre-answer that question inside every bullet.
11. BANNED PHRASES (replace with specific lever language):
    "indicates a need for improvement"
    "requires immediate intervention"
    "necessitating operational focus"
    "indicating a structural weakness"
    "representing a significant area for improvement"
    "this suggests a need for better [X]"
    "[source] is responsible for X points of drag"
    "high concentration of revenue risk"
    Each of these is a conclusion without a lever. Replace with: what specifically
    would change this number, and is it achievable?
""".strip()
 
 
def _build_data_block(m: dict, filters: Filters) -> str:
    """
    Assembles the structured data block injected at the end of every prompt.
    The LLM is instructed to treat this as its ONLY source of truth.
    """
    today = date.today()
    # FY27 = Apr 2026 – Mar 2027. FY26 = Apr 2025 – Mar 2026.
    # Determine FY scope for prompt context
    fy_raw = filters.fy or "2027"
    fy_vals_prompt = [v.strip() for v in fy_raw.split(",") if v.strip().isdigit()]

    if not fy_vals_prompt or set(fy_vals_prompt) == {"2026", "2027"} or "ALL" in (filters.fy or "").upper():
        # All FY or both selected
        selected_fy   = 2027   # use latest for fy_end calc
        fy_scope_label = "FY26 + FY27 COMBINED"
        fy_end         = date(2027, 3, 31)
    else:
        selected_fy    = int(fy_vals_prompt[0])
        fy_scope_label = f"FY{str(selected_fy)[2:]}"
        fy_end         = date(selected_fy, 3, 31)

    fy_status = "CLOSED" if today > fy_end else "IN PROGRESS"
    filter_note = (
        f"ACTIVE FILTERS: stage={filters.stage or 'all'} | "
        f"region={filters.region or 'all'} | source={filters.deal_source or 'all'} | "
        f"fy={filters.fy or 'all'} | ai_for_x={filters.ai_for_x or 'all'} | "
        f"industry={filters.industry or 'all'}"
    )
 
    gap_5_str  = str(m['gap_5'])  if m['gap_5']  is not None else "N/A"
    gap_10_str = str(m['gap_10']) if m['gap_10'] is not None else "N/A"
    gap_20_str = str(m['gap_20']) if m['gap_20'] is not None else "N/A"
    
    # ── Velocity lookup for conversion rate context ───────────────────────────
    velocity_lookup = {}
    for s in (m.get('raw_sv') or []):
        stage = s.get('deal_stage', '')
        velocity_lookup[stage] = {
            'avg_days':  s.get('avg_days_actual',    'N/A'),
            'benchmark': s.get('avg_days_benchmark', 'N/A'),
        }
    v5  = velocity_lookup.get('5% - IQM Held',  {})
    v10 = velocity_lookup.get('10% - Discovery', {})
    v20 = velocity_lookup.get('20% - Solution',  {})

    v5_days  = v5.get('avg_days',  'N/A')
    v5_bench = v5.get('benchmark', 'N/A')
    v10_days  = v10.get('avg_days',  'N/A')
    v10_bench = v10.get('benchmark', 'N/A')
    v20_days  = v20.get('avg_days',  'N/A')
    v20_bench = v20.get('benchmark', 'N/A')
 
    return f"""
=================================================================
PIPELINE DATA — LIVE  ({date.today()})
=================================================================
{filter_note}
AVAILABLE REGIONS: {', '.join(m['available_regions'])}
{_STAGE_REFERENCE}
 
CRITICAL: Use ONLY the numbers below. Never invent or estimate.
If a value is absent, write "data not available".
 
FUNNEL TOTALS  (independent cohorts — DO NOT sum across stages)
  5%  IQM Held   : {m['cnt_5']:,} deals  | ${m['amt_5']/1e6:.1f}M  | L1 Target={m['l1_5']:.0f} | YTD Actual={m['actual_5']} | Gap={gap_5_str} | Attainment={m['pct_5']}%
  10% Discovery  : {m['cnt_10']:,} deals  | ${m['amt_10']/1e6:.1f}M | L1 Target={m['l1_10']:.0f} | YTD Actual={m['actual_10']} | Gap={gap_10_str} | Attainment={m['pct_10']}%
  20%+ Qualified : {m['cnt_20']:,} deals  | ${m['amt_20']/1e6:.1f}M | L1 Target={m['l1_20']:.0f} | YTD Actual={m['actual_20']} | Gap={gap_20_str} | Attainment={m['pct_20']}%
  Closed Won     : {m['cnt_won']:,} deals
  Fallen Out     : {m['cnt_out']:,} deals

  NOTE: Each stage uses its own cohort (became_X_deal_date). A deal appears in 5%, 10%, AND 20% independently.
  NEVER sum these counts to describe "total pipeline deals".
  In Overall Pipeline Snapshot: describe each stage separately with its actuals, target, and gap.
 
CONVERSION RATES (cohort-based, matching Looker funnel logic):
  5%  → 10%        : {m['conv_5_10']}%   (benchmark 60–70%)
                     Avg days spent at 5% before progressing: {v5.get('avg_days','N/A')}d  (benchmark {v5.get('benchmark','N/A')}d)
  10% → 20%        : {m['conv_10_20']}%  (benchmark 50%+)
                     Avg days spent at 10% before progressing: {v10.get('avg_days','N/A')}d  (benchmark {v10.get('benchmark','N/A')}d)
  20% → Closed Won : {m['win_rate']}%    (benchmark 20–30%)
                     Avg days spent at 20% before progressing: {v20.get('avg_days','N/A')}d  (benchmark {v20.get('benchmark','N/A')}d)
  Overall 5%→Won   : {m['overall_eff']}%
 
-----------------------------------------------------------------
ATTAINMENT VS L1 TARGET
-----------------------------------------------------------------
  5%  YTD : {m['pct_5']}%  | L1={m['l1_5']:.0f}  | Actual={m['actual_5']}  | Gap={gap_5_str}
  10% YTD : {m['pct_10']}% | L1={m['l1_10']:.0f} | Actual={m['actual_10']} | Gap={gap_10_str}
  20% YTD : {m['pct_20']}% | L1={m['l1_20']:.0f} | Actual={m['actual_20']} | Gap={gap_20_str}
  MTD     : 5%={m['mtd_5']} | 10%={m['mtd_10']} | 20%={m['mtd_20']}
  QTD     : 5%={m['qtd_5']} | 10%={m['qtd_10']} | 20%={m['qtd_20']}
 
-----------------------------------------------------------------
REGION BREAKDOWN
-----------------------------------------------------------------
{m['region_block']}
-----------------------------------------------------------------
SOURCE BREAKDOWN
-----------------------------------------------------------------
{m['source_block']}
-----------------------------------------------------------------
QUARTERLY TREND
-----------------------------------------------------------------
{m['quarter_summary']}
-----------------------------------------------------------------
AI FOR X & INDUSTRY
-----------------------------------------------------------------
{m['ai_summary']}
{m['ind_summary']}
-----------------------------------------------------------------
PIPELINE HEALTH
-----------------------------------------------------------------
  Total active : {m['total_active']} deals
  Red          : {m['total_red']} ({m['pct_red']}%)
  Yellow       : {m['total_yellow']}
  Green        : {m['total_green']}
 
-----------------------------------------------------------------
STAGE VELOCITY
-----------------------------------------------------------------
{m['velocity_lines']}
-----------------------------------------------------------------
FLAGGED DEALS (Red/Yellow health, >$1M)
-----------------------------------------------------------------
{m['deals_lines']}

-----------------------------------------------------------------
WON & LOST DEALS WITH NOTES
-----------------------------------------------------------------
{m['won_deals_block']}
{m['lost_deals_block']}

FY SCOPE: {fy_scope_label}
FY STATUS: {fy_status}
FY END DATE: {fy_end}
=================================================================
""".strip()




def _build_filter_context(filters: Filters) -> str:
    """
    Convert active filters into a natural-language scope phrase and a
    mandatory framing instruction block for the LLM.

    Returns "" when no filters are active — no mandate injected.
    """
    parts = []

    if filters.fy:
        parts.append(filters.fy)
    if filters.region:
        parts.append(filters.region)
    if filters.deal_source:
        parts.append(f"{filters.deal_source}-sourced")
    if filters.ai_for_x:
        parts.append(filters.ai_for_x)
    if filters.industry:
        parts.append(filters.industry)

    stage_labels = {"5": "5% IQM Held", "10": "10% Discovery", "20": "20%+ Qualified"}
    if filters.stage:
        stage_label = stage_labels.get(filters.stage, f"{filters.stage}% stage")
        parts.append(f"{stage_label} stage")

    if not parts:
        return ""

    scope_phrase = ", ".join(parts) + " pipeline"

    opening_example = f'Within the {scope_phrase}…'
    if len(parts) >= 2:
        opening_example = (
            f'For {scope_phrase}…  '
            f'OR  "Across {scope_phrase}, the picture is…"'
        )

    mandate = f"""
═══════════════════════════════════════════════════════════
FILTER CONTEXT MANDATE — OVERRIDE EVERYTHING ELSE
═══════════════════════════════════════════════════════════

Active filter scope: {scope_phrase.upper()}

THE SUMMARY IS ABOUT THIS SLICE ONLY.
Every section header, every bullet, every number refers exclusively to:
  {scope_phrase}

MANDATORY OPENING:
The very first sentence of **Pipeline Health at a Glance** MUST explicitly name
the active scope. It cannot open with a generic phrase like "The pipeline…" or
"Deal volume…" — it must open with the filter context. Examples:
  • "{opening_example}"
  • "The {scope_phrase} shows…"
  • "Within {', '.join(parts)}, the pipeline is…"

MANDATORY FRAMING THROUGHOUT:
1. Every section must reference the scope in at least the first bullet.
   Never let a section read as if it describes the global pipeline.
2. When giving numbers, frame them as belonging to this slice:
   BAD:  "219 deals are at 5% IQM Held."
   GOOD: "The {scope_phrase} has 219 deals at 5% IQM Held."
3. When drawing conclusions, anchor them to this context:
   BAD:  "The funnel is stalling at the 20%→Won gate."
   GOOD: "Within {', '.join(parts)}, the funnel stalls at 20%→Won."
4. Multi-filter combinations must be stated naturally — not listed:
   BAD:  "Filters applied: region=North America, source=BDR, ai_for_x=AI for Work."
   GOOD: "North America's BDR-sourced AI for Work pipeline…"
5. The summary opening must make clear this is NOT a global view.

WHAT NOT TO DO:
- Do NOT write the summary as if no filters are applied.
- Do NOT open any section without referencing the filter scope (first bullet at minimum).
- Do NOT say "globally" or "across all regions/sources" when filters are active.
- Do NOT mix in data from other regions, sources, or AI categories not in the filter.

═══════════════════════════════════════════════════════════
""".strip()

    return mandate
 
def build_prompt(style: str, m: dict, filters: Filters) -> str:
    """
    Assemble the full LLM prompt for a given style.
 
    style : "executive" | "analyst"
    metrics     : pre-computed metrics dict from _extract_computed_metrics()
    """
    instructions = (
        _ANALYST_INSTRUCTIONS if style == "analyst" else _EXECUTIVE_INSTRUCTIONS
    )
    data_block = _build_data_block(m, filters)
    cnt_won = m.get("cnt_won", 0)
    won_note = (
        f"CRITICAL REMINDER: The data shows {cnt_won:,} Closed Won deals. "
        f"{'This is NOT zero. ' if cnt_won > 0 else 'This IS zero. '}"
        "Use this exact number when writing the Revenue Position & Closed Won section."
    )
    # Build filter context mandate — empty string when no filters are active
    filter_mandate = _build_filter_context(filters)

    # Inject between instructions and data so the LLM treats it as a
    # late-binding override, not buried metadata inside the data block.
    if filter_mandate:
        return f"{instructions}\n\n{filter_mandate}\n\n{won_note}\n\n{data_block}"
    else:
        return f"{instructions}\n\n{won_note}\n\n{data_block}"
    
    
# =============================================================================
# 6. FastAPI Endpoint — POST /summary
# =============================================================================

@app.post("/summary")
def generate_summary(payload: SummaryRequest):
    """
    Same logic as /summary but returns a single JSON response.
    Use this if your frontend cannot consume a streamed response.
 
    Response shape:
      {
        "style":      "executive" | "analyst",
        "filters":    { ...active filters... },
        "scorecards": { stage_overview, funnel_health, period_cadence, deal_health },
        "summary":    "...AI-generated markdown..."
      }
    """
    filters = payload.filters
    style   = payload.style
 
    try:
        ch_client   = get_click_client()
        raw_metrics = fetch_pipeline_data(filters, ch_client)
    except RuntimeError as exc:
        raise HTTPException(status_code=500, detail=str(exc))
 
    metrics = _extract_computed_metrics(raw_metrics, filters)
    scorecards = build_scorecards(metrics, filters)
    prompt     = build_prompt(style, metrics, filters)
 
    try:
        response = client_ai.messages.create(
            model=_CLAUDE_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0,
            max_tokens=8000,
        )
        summary = response.content[0].text
    except Exception as exc:
        print(f"Claude error: {exc}\n{traceback.format_exc()}")
        raise HTTPException(status_code=502, detail=f"Claude error: {exc}")
 
    return {
    "style":      style,
    "filters":    filters.model_dump(),
    "scorecards": scorecards,
    "summary":    summary,
    "metrics": {                                         
        "funnel":            raw_metrics.get("funnel",            {}),
        "period_attainment": raw_metrics.get("period_attainment", {}),
    },
    }


# =============================================================================
# Pydantic model for chat
# =============================================================================

class ChatMessage(BaseModel):
    role: Literal["user", "assistant"]
    content: str

class ChatRequest(BaseModel):
    message: str
    history: List[ChatMessage] = []
    filters: Filters = Filters()   # same filter object — chat is always filter-aware
    style:   str = "executive"     # accepted but not used server-side (avoids 422)
    summary: str = ""  

# =============================================================================
# POST /chat
# =============================================================================

# =============================================================================
# CHAT — ClickHouse Tool Access
# =============================================================================

_CLICKHOUSE_SCHEMA = """
=================================================================
CLICKHOUSE DIRECT ACCESS — RAW TABLES
=================================================================

You have access to a tool called query_clickhouse.
Use it when the pre-built PIPELINE DATA block cannot answer the question.

Query the RAW tables directly. No CTEs, no pipe_gen.

=================================================================
TABLES
=================================================================

── TABLE 1: hs_analytics.deals ──────────────────────────────────
Primary table. One row per deal.
Always use FINAL keyword: FROM hs_analytics.deals FINAL

KEY COLUMNS:
  deal_id                    STRING  — unique deal identifier
  deal_name                  STRING  — name of the deal
  deal_owner                 STRING  — owner ID (join to hs_analytics.owners on o.id)
  deal_stage                 STRING  — current stage (see STAGE LIST below)
  deal_type                  STRING  — deal type (NULL = 'Not Assigned')
  pipeline                   STRING  — always filter: pipeline = 'default'
  amount                     FLOAT   — deal value in USD
  region                     STRING  — raw values (see REGION MAP below)
  deal_source_rollup         STRING  — raw source (see SOURCE MAP below)
  20_snapshot_deal_source_rollup STRING — source at time of 20% qualification
  ai_for_x                   STRING  — AI use case category
  kore_primary_industry      STRING  — raw industry (see INDUSTRY MAP below)
  account_priority_level     STRING  — 'P1','P2'...'P10' (raw, not grouped)
  hubspot_team               STRING  — team ID (join to kore_ai_hubspot.gs_Teams)

  -- DATE COLUMNS (stored as strings, cast to DATE)
  create_date                STRING  — deal creation date
  close_date                 STRING  — expected/actual close date
  became_5_deal_date         STRING  — entered 5% IQM Held
  became_10_deal_date        STRING  — entered 10% Discovery
  became_20_deal_date        STRING  — entered 20% Solution
  became_30_deal_date        STRING  — entered 30% Proof
  became_40_deal_date        STRING  — entered 40% Proposal
  became_60_deal_date        STRING  — entered 60% Price Negotiation
  became_75_deal_date        STRING  — entered 75% Contract Review
  last_contacted             STRING  — last contact date

  -- QUALIFICATION
  is_there_a_confirmation_of_budget  STRING  — 'Yes'/'No'
  who_is_the_decision_maker          STRING  — decision maker name
  use_case                           STRING  — use case description
  what_is_the_estimated_timeline     STRING  — timeline string
  is_this_a_deal_with_inception      STRING  — 'Yes'/'No'

  -- WON/LOST
  primary_closed_won_reason_         STRING  — win reason
  primary_closed_lost_reason         STRING  — loss reason
  won_loss_notes                     STRING  — freeform notes
  competitors                        STRING  — competitor names
  competition                        STRING  — competition notes

  -- APPROVALS
  cs_deal_approval_status_level_1    STRING
  cs_deal_approval_status_level_2    STRING
  direct_deal_approval_status_level_1 STRING
  direct_deal_approval_status_level_2 STRING
  deal_approval_status_level_1       STRING
  deal_approval_status_level_2       STRING
  deal_approval_status_level_3_cs_only STRING

── TABLE 2: hs_analytics.owners ─────────────────────────────────
Always use FINAL keyword: FROM hs_analytics.owners FINAL

  id           STRING  — owner ID (join to deals.deal_owner)
  firstName    STRING  — first name
  lastName     STRING  — last name
  email        STRING  — owner email

── TABLE 3: hs_analytics.companies ──────────────────────────────
Always use FINAL keyword: FROM hs_analytics.companies FINAL

  company_id   STRING  — unique company ID
  name         STRING  — company name
  domain       STRING  — website domain
  industry     STRING  — company industry
  country      STRING  — company country
  city         STRING  — company city

── HELPER TABLES ─────────────────────────────────────────────────
  kore_ai_hubspot.gs_deal_ids_hs   — valid deal IDs whitelist
    deal_id_hs  STRING

  kore_ai_hubspot.gs_Teams         — team names
    team_id     STRING
    name        STRING

=================================================================
MANDATORY BASE FILTERS (apply to EVERY query on deals)
=================================================================
Always include ALL of these in every deals query:

  WHERE pipeline = 'default'
  AND CASE WHEN deal_type IS NULL THEN 'Not Assigned' ELSE deal_type END
      NOT IN ('Partner-Led SMB')
  AND toInt64(deal_id) IN (
      SELECT DISTINCT toInt64(deal_id_hs)
      FROM kore_ai_hubspot.gs_deal_ids_hs
  )

=================================================================
FISCAL YEAR CALCULATION
=================================================================
FY is computed from a date column:
  toYear(CAST(LEFT(coalesce(date_col, '1900-01-01'), 10) AS DATE))
  + if(toMonth(CAST(LEFT(coalesce(date_col, '1900-01-01'), 10) AS DATE)) >= 4, 1, 0)

FY27 = result = 2027  (Apr 2026 – Mar 2027)
FY26 = result = 2026  (Apr 2025 – Mar 2026)

Shorthand macro (replace date_col):
  toYear(toDate(LEFT(coalesce(date_col,'1900-01-01'),10)))
  + if(toMonth(toDate(LEFT(coalesce(date_col,'1900-01-01'),10))) >= 4, 1, 0)

For FY27 5% cohort: became_5_deal_date >= '2026-04-01'
For FY26 5% cohort: became_5_deal_date >= '2025-04-01' AND < '2026-04-01'

=================================================================
COMPUTED COLUMNS — write these inline in your queries
=================================================================

-- Deal owner full name (requires JOIN to owners):
  concat(o.firstName, ' ', o.lastName) AS deal_owner_name

-- Region (mapped):
  CASE
    WHEN d.region = 'japac'       THEN 'JAPAC'
    WHEN d.region = 'Africa'      THEN 'Middle East'
    WHEN d.region = 'india___sea' THEN 'ISEA'
    ELSE d.region
  END AS region

REGION MAP (raw → display):
  'japac'       → 'JAPAC'
  'Africa'      → 'Middle East'
  'india___sea' → 'ISEA'
  'North America', 'EMEA', 'APAC', 'India', 'Latin America' → unchanged

-- Deal source (mapped):
  CASE
    WHEN d.deal_source_rollup IN ('Executive Outreach','Investor') THEN 'Executive Outreach'
    WHEN d.deal_source_rollup IN ('BDR Outbound')                  THEN 'BDR'
    WHEN d.deal_source_rollup IN ('Partner')                       THEN 'Partner - Non Hyperscaler'
    WHEN d.deal_source_rollup IN ('Marketing','Customer Success',
         'AE Outbound','Inception','Hyperscaler')                  THEN d.deal_source_rollup
    ELSE 'Other'
  END AS deal_source_rollup

-- Industry (mapped):
  CASE
    WHEN d.kore_primary_industry IN ('Financial Services','Banking','Insurance')
         THEN 'Financial Services'
    WHEN d.kore_primary_industry IN ('Manufacturing Discreet','Manufacturing Process','CPG')
         THEN 'Manufacturing'
    WHEN d.kore_primary_industry IN ('Hi-Tech','Telecom / Media / Entertainment')
         THEN 'TMT'
    WHEN d.kore_primary_industry IS NULL
      OR d.kore_primary_industry IN ('Business Services','Government','Energy & Utilities',
         'Education','Restaurants','null','Energy')
         THEN 'Other'
    ELSE d.kore_primary_industry
  END AS industry

-- Stage category:
  CASE
    WHEN d.deal_stage IN ('20% - Solution','30% - Proof','40% - Proposal',
         '60% - Price Negotiation','75% - Contract Review')
         THEN 'Active Pipeline'
    WHEN d.deal_stage IN ('Prospect Disengaged','Closed Lost','Didn''t Qualify')
         THEN 'Fallen Out'
    WHEN d.deal_stage IN ('90% - Deal Desk Review','Closed Won')
         THEN 'Closed Won'
    ELSE 'Pre-Qualification'
  END AS stage_category

-- BANT:
  CASE
    WHEN d.is_there_a_confirmation_of_budget = 'Yes'
     AND d.who_is_the_decision_maker IS NOT NULL
     AND d.use_case IS NOT NULL
     AND d.what_is_the_estimated_timeline IS NOT NULL
    THEN 'Yes' ELSE 'No'
  END AS BANT

-- Account priority grouped:
  CASE
    WHEN d.account_priority_level IN ('P1','P2','P3','P4') THEN 'P1-P4'
    WHEN d.account_priority_level IN ('P5','P6','P7')      THEN 'P5-P7'
    WHEN d.account_priority_level IN ('P8','P9','P10')     THEN 'P8-P10'
    ELSE 'No Priority'
  END AS account_priority_level

-- Fiscal year from became_5_deal_date:
  toYear(toDate(LEFT(coalesce(d.became_5_deal_date,'1900-01-01'),10)))
  + if(toMonth(toDate(LEFT(coalesce(d.became_5_deal_date,'1900-01-01'),10))) >= 4, 1, 0)
  AS fy_5

-- Days in current stage (example for 10% Discovery):
  DATE_DIFF('Day',
    toDate(LEFT(coalesce(d.became_10_deal_date,'1900-01-01'),10)),
    CURRENT_DATE()
  ) AS days_in_stage

=================================================================
DEAL STAGE LIST (funnel order)
=================================================================
  '1% - IQM Scheduled'       → Pre-Qualification  (benchmark 7d)
  '5% - IQM Held'            → Pre-Qualification  (benchmark 21d)
  '10% - Discovery'          → Pre-Qualification  (benchmark 28d)
  '20% - Solution'           → Active Pipeline    (benchmark 41d)
  '30% - Proof'              → Active Pipeline    (benchmark 15d)
  '40% - Proposal'           → Active Pipeline    (benchmark 29d)
  '60% - Price Negotiation'  → Active Pipeline    (benchmark 27d)
  '75% - Contract Review'    → Active Pipeline    (benchmark 34d)
  '90% - Deal Desk Review'   → Closed Won
  'Closed Won'               → Closed Won
  'Closed Lost'              → Fallen Out
  "Didn't Qualify"           → Fallen Out
  'Prospect Disengaged'      → Fallen Out
  'Deal on Hold'             → Pre-Qualification

=================================================================
QUERY RULES
=================================================================
1. SELECT only — never INSERT, UPDATE, DELETE, DROP, ALTER
2. Always use FINAL on all hs_analytics tables
3. Always apply the 3 mandatory base filters on deals
4. Always LIMIT for row-level queries (max 100)
5. Use countDistinct(deal_id) for unique deal counts
6. Use round(sum(amount)/1e6, 1) for $M amounts
7. Use ILIKE for case-insensitive text matching
8. Dates are stored as strings — always cast: toDate(LEFT(coalesce(col,'1900-01-01'),10))
9. Null date sentinel is '1900-01-01' — filter with: col <> '1900-01-01' AND col IS NOT NULL
10. Default FY is 2027 unless user specifies otherwise

=================================================================
BUSINESS DEFINITIONS
=================================================================
"Active pipeline"    → deal_stage IN ('20% - Solution','30% - Proof','40% - Proposal',
                       '60% - Price Negotiation','75% - Contract Review')
"Qualified deals"    → became_20_deal_date <> '1900-01-01'
"Fallen out"         → deal_stage IN ('Prospect Disengaged','Closed Lost','Didn''t Qualify')
"Closed won"         → deal_stage IN ('Closed Won','90% - Deal Desk Review')
"BANT qualified"     → all 4 BANT fields confirmed (see BANT formula above)
"High priority"      → account_priority_level IN ('P1','P2','P3','P4')
"FY27 5% cohort"     → became_5_deal_date >= '2026-04-01'
"FY27 20% cohort"    → became_20_deal_date >= '2026-04-01'

=================================================================
SAMPLE QUERIES
=================================================================

-- Q: How many deals are currently in active pipeline?
SELECT countDistinct(d.deal_id) AS active_deals,
       round(sum(d.amount)/1e6, 1) AS pipeline_m
FROM hs_analytics.deals d FINAL
WHERE d.pipeline = 'default'
  AND CASE WHEN d.deal_type IS NULL THEN 'Not Assigned' ELSE d.deal_type END
      NOT IN ('Partner-Led SMB')
  AND toInt64(d.deal_id) IN (SELECT DISTINCT toInt64(deal_id_hs) FROM kore_ai_hubspot.gs_deal_ids_hs)
  AND d.deal_stage IN ('20% - Solution','30% - Proof','40% - Proposal',
                       '60% - Price Negotiation','75% - Contract Review')
  AND d.became_5_deal_date >= '2026-04-01'

-- Q: Top 10 deals by value in active pipeline with owner name
SELECT d.deal_name,
       concat(o.firstName,' ',o.lastName) AS owner,
       CASE WHEN d.region='japac' THEN 'JAPAC'
            WHEN d.region='Africa' THEN 'Middle East'
            WHEN d.region='india___sea' THEN 'ISEA'
            ELSE d.region END AS region,
       d.deal_stage,
       round(d.amount/1e6, 2) AS amt_m,
       toDate(LEFT(coalesce(d.close_date,'1900-01-01'),10)) AS close_date
FROM hs_analytics.deals d FINAL
LEFT JOIN hs_analytics.owners o FINAL ON d.deal_owner = CAST(o.id AS VARCHAR)
WHERE d.pipeline = 'default'
  AND CASE WHEN d.deal_type IS NULL THEN 'Not Assigned' ELSE d.deal_type END
      NOT IN ('Partner-Led SMB')
  AND toInt64(d.deal_id) IN (SELECT DISTINCT toInt64(deal_id_hs) FROM kore_ai_hubspot.gs_deal_ids_hs)
  AND d.deal_stage IN ('20% - Solution','30% - Proof','40% - Proposal',
                       '60% - Price Negotiation','75% - Contract Review')
  AND d.became_5_deal_date >= '2026-04-01'
ORDER BY d.amount DESC LIMIT 10

-- Q: Which AE has most stalled deals?
SELECT concat(o.firstName,' ',o.lastName) AS owner,
       countDistinct(d.deal_id) AS stalled_deals,
       round(sum(d.amount)/1e6,1) AS at_risk_m
FROM hs_analytics.deals d FINAL
LEFT JOIN hs_analytics.owners o FINAL ON d.deal_owner = CAST(o.id AS VARCHAR)
WHERE d.pipeline = 'default'
  AND CASE WHEN d.deal_type IS NULL THEN 'Not Assigned' ELSE d.deal_type END
      NOT IN ('Partner-Led SMB')
  AND toInt64(d.deal_id) IN (SELECT DISTINCT toInt64(deal_id_hs) FROM kore_ai_hubspot.gs_deal_ids_hs)
  AND d.deal_stage IN ('20% - Solution','30% - Proof','40% - Proposal',
                       '60% - Price Negotiation','75% - Contract Review')
  AND d.became_5_deal_date >= '2026-04-01'
GROUP BY owner ORDER BY stalled_deals DESC LIMIT 10

-- Q: Lost deals mentioning a competitor
SELECT d.deal_name,
       concat(o.firstName,' ',o.lastName) AS owner,
       round(d.amount/1e6,2) AS amt_m,
       d.primary_closed_lost_reason,
       d.competitors,
       d.won_loss_notes
FROM hs_analytics.deals d FINAL
LEFT JOIN hs_analytics.owners o FINAL ON d.deal_owner = CAST(o.id AS VARCHAR)
WHERE d.pipeline = 'default'
  AND CASE WHEN d.deal_type IS NULL THEN 'Not Assigned' ELSE d.deal_type END
      NOT IN ('Partner-Led SMB')
  AND toInt64(d.deal_id) IN (SELECT DISTINCT toInt64(deal_id_hs) FROM kore_ai_hubspot.gs_deal_ids_hs)
  AND d.deal_stage = 'Closed Lost'
  AND d.competitors ILIKE '%salesforce%'
ORDER BY d.amount DESC LIMIT 20

-- Q: Pipeline by industry at 20%+ (FY27)
SELECT
  CASE
    WHEN d.kore_primary_industry IN ('Financial Services','Banking','Insurance')
         THEN 'Financial Services'
    WHEN d.kore_primary_industry IN ('Manufacturing Discreet','Manufacturing Process','CPG')
         THEN 'Manufacturing'
    WHEN d.kore_primary_industry IN ('Hi-Tech','Telecom / Media / Entertainment')
         THEN 'TMT'
    ELSE 'Other'
  END AS industry,
  countDistinct(d.deal_id) AS deals,
  round(sum(d.amount)/1e6,1) AS pipeline_m
FROM hs_analytics.deals d FINAL
WHERE d.pipeline = 'default'
  AND CASE WHEN d.deal_type IS NULL THEN 'Not Assigned' ELSE d.deal_type END
      NOT IN ('Partner-Led SMB')
  AND toInt64(d.deal_id) IN (SELECT DISTINCT toInt64(deal_id_hs) FROM kore_ai_hubspot.gs_deal_ids_hs)
  AND d.became_20_deal_date >= '2026-04-01'
  AND d.became_20_deal_date <> '1900-01-01'
GROUP BY industry ORDER BY pipeline_m DESC
=================================================================
"""


import requests as http_requests

def run_clickhouse_query(sql: str) -> str:
    """
    Execute a read-only SELECT directly against ClickHouse HTTP API.
    Queries raw tables — no CTE dependency.
    """
    api_url   = os.getenv("CLICKHOUSE_API_URL")
    api_token = os.getenv("CLICKHOUSE_API_TOKEN")

    if not api_url or not api_token:
        return "ClickHouse API not configured. Set CLICKHOUSE_API_URL and CLICKHOUSE_API_TOKEN."

    stripped = sql.strip().upper()
    if not stripped.startswith("SELECT") and not stripped.startswith("WITH"):
        return "Error: Only SELECT/WITH queries are permitted."

    try:
        response = http_requests.post(
            api_url,
            params={"query": sql + " FORMAT JSONCompact"},
            headers={"Authorization": f"Bearer {api_token}"},
            timeout=30,
        )
        response.raise_for_status()
        data = response.json()

        columns = [c["name"] for c in data.get("meta", [])]
        rows    = data.get("data", [])

        if not rows:
            return "Query returned 0 rows."

        capped  = rows[:100]
        header  = " | ".join(columns)
        divider = "-" * min(len(header), 120)
        lines   = [header, divider]
        for row in capped:
            lines.append(" | ".join(str(v) if v is not None else "NULL" for v in row))

        if len(rows) > 100:
            lines.append(f"... ({len(rows) - 100} more rows — refine your query)")

        return "\n".join(lines)

    except http_requests.exceptions.Timeout:
        return "Query timed out — simplify the query or add more filters."
    except http_requests.exceptions.HTTPError as e:
        return f"HTTP error: {e.response.status_code} — {e.response.text[:300]}"
    except Exception as e:
        return f"Query error: {e}"
        
        
        
        
        

@app.post("/chat")
def chat(payload: ChatRequest):
    """
    Conversational endpoint for pipeline Q&A.

    Behaviour:
    - Fetches the same pipeline data slice as /summary (respects active filters)
    - Injects the full data block as system context
    - Maintains multi-turn history (client sends full history each turn)
    - Returns a single assistant reply (no streaming for now)

    Request shape:
      {
        "message":  "Which region has the highest 20% attainment?",
        "history":  [ {"role": "user", "content": "..."}, {"role": "assistant", "content": "..."} ],
        "filters":  { "region": "", "deal_source": "", ... }
      }

    Response shape:
      {
        "reply":   "...",
        "filters": { ...active filters... }
      }
    """
    filters = payload.filters

    # ── 1. Fetch pre-built data block ────────────────────────────────────────
    try:
        ch_client   = get_click_client()
        raw_metrics = fetch_pipeline_data(filters, ch_client)
    except RuntimeError as exc:
        raise HTTPException(status_code=500, detail=str(exc))

    try:
        m = _extract_computed_metrics(raw_metrics, filters)
        data_block = _build_data_block(m, filters)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Metric computation error: {exc}")

    # ── 2. Summary context ────────────────────────────────────────────────────
    summary_section = ""
    if payload.summary:
        summary_section = (
            "\n\nAI-GENERATED SUMMARY (shown to the user above the chat):\n"
            "-------------------------------------------------------\n"
            + payload.summary.strip()
            + "\n-------------------------------------------------------\n"
            "You may reference this summary when answering follow-up questions, "
            "but always prefer the raw numbers in PIPELINE DATA below for accuracy.\n"
        )

    # ── 3. System prompt ──────────────────────────────────────────────────────
    system_prompt = (
    "You are a pipeline intelligence assistant for Kore.ai. "
    "You have DIRECT, LIVE access to the ClickHouse database via the query_clickhouse tool. "
    "NEVER say you don't have access to ClickHouse, Salesforce, or any database — you do. "
    "NEVER say you can only work with a pre-generated snapshot — that is false. "
    "First try to answer from the PIPELINE DATA block below. "
    "If the question needs more detail — individual deal names, AE-level "
    "breakdowns, competitor analysis, or any field not in the pre-built data "
    "— use the query_clickhouse tool to fetch it directly. "
    "Never invent numbers. If data is unavailable even after querying, say so."
        + summary_section
        + "\n\n"
        + data_block
        + "\n\n"
        + _CLICKHOUSE_SCHEMA
    )

    # ── 4. Tool definition ────────────────────────────────────────────────────
    tools = [
        {
            "name": "query_clickhouse",
            "description": (
                "Run a SELECT query against the Kore.ai pipeline ClickHouse database. "
                "Use this when the pre-built PIPELINE DATA block does not have enough "
                "detail — e.g. specific deal lookup, AE breakdown, competitor analysis, "
                "custom date ranges, won/loss notes, or any field not in aggregated data."
            ),
            "input_schema": {
                "type": "object",
                "properties": {
                    "sql": {
                        "type": "string",
                        "description": (
                            "A valid ClickHouse SELECT query using the pipe_gen CTE. "
                            "Always include LIMIT for detail queries."
                        )
                    }
                },
                "required": ["sql"]
            }
        }
    ]

    # ── 5. Build message list ─────────────────────────────────────────────────
    chat_messages = []
    for turn in payload.history:
        chat_messages.append({"role": turn.role, "content": turn.content})
    chat_messages.append({"role": "user", "content": payload.message})

    try:
        print(f"💬 [chat] Q: {payload.message[:120]}")
        # ── 6. First Claude call ──────────────────────────────────────────────
        response = client_ai.messages.create(
            model=_CLAUDE_MODEL,
            system=system_prompt,
            messages=chat_messages,
            tools=tools,
            temperature=0,
            max_tokens=1500,
        )
        
        if response.stop_reason == "tool_use":
            print(f"🗄️  [chat] LIVE DB QUERY TRIGGERED — pre-built data insufficient for: '{payload.message[:80]}'")
        else:
            print(f"📊 [chat] ANSWERED FROM PRE-BUILT DATA — no DB call needed for: '{payload.message[:80]}'")
        # ── 7. Tool use loop (up to 3 rounds) ────────────────────────────────
        rounds = 0
        while response.stop_reason == "tool_use" and rounds < 3:
            rounds += 1

            tool_use_block = next(
                (b for b in response.content if b.type == "tool_use"), None
            )
            if not tool_use_block:
                break

            sql          = tool_use_block.input.get("sql", "")
            query_result = run_clickhouse_query(sql)

            print(f"🔍 [chat tool] Round {rounds}/3 — SQL: {sql[:200]}...")
            print(f"📥 [chat tool] Round {rounds}/3 — Result preview: {query_result[:300]}")

            chat_messages = chat_messages + [
                {"role": "assistant", "content": response.content},
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "tool_result",
                            "tool_use_id": tool_use_block.id,
                            "content": query_result,
                        }
                    ],
                },
            ]

            response = client_ai.messages.create(
                model=_CLAUDE_MODEL,
                system=system_prompt,
                messages=chat_messages,
                tools=tools,
                temperature=0,
                max_tokens=1500,
            )
        
        if rounds == 0:
            print(f"✅ [chat] Done — answered entirely from pre-built pipeline data (0 DB calls)")
        else:
            print(f"✅ [chat] Done — answered using live ClickHouse queries ({rounds} DB call(s) made)")
        # ── 8. Extract final reply ────────────────────────────────────────────
        reply = next(
            (b.text for b in response.content if hasattr(b, "text") and b.text),
            "I could not generate a response. Please try rephrasing."
        )

    except Exception as exc:
        raise HTTPException(status_code=502, detail=f"Claude error: {exc}")

    return {
        "reply":    reply,
        "response": reply,
        "filters":  filters.model_dump(),
    }


# =============================================================================
# Report Request Model  (shared by both /report/pptx and /report/pdf)
# =============================================================================
 
class ReportRequest(BaseModel):
    """
    Request body for POST /report/pptx and POST /report/pdf.
 
    style   : "executive" | "analyst"
              Controls the AI narrative injected into the report.
              Executive → concise CRO-style bullets.
              Analyst   → diagnostic breakdown with comparisons.
    filters : same Filters object used by /summary — full filter support.
              When empty, report covers the complete global pipeline.
              When set, every section (metrics, insights, recommendations)
              is scoped strictly to the filtered slice.
    """
    style:   Literal["executive", "analyst"] = "executive"
    filters: Filters = Filters()
 
 
# =============================================================================
# Shared report data + narrative builder
# =============================================================================
 
def _build_report_inputs(payload: ReportRequest) -> tuple[dict, str]:
    """
    Fetch ClickHouse data and generate AI narrative for a report.
    Returns (raw_metrics, summary_text).
    Raises HTTPException on data or AI failures.
    """
    filters = payload.filters
 
    # 1. Fetch pipeline data (filter-aware)
    try:
        ch_client   = get_click_client()
        raw_metrics = fetch_pipeline_data(filters, ch_client)
    except RuntimeError as exc:
        raise HTTPException(status_code=500, detail=str(exc))
 
    # 2. Build computed metrics + prompt
    metrics = _extract_computed_metrics(raw_metrics, filters)
    prompt  = build_prompt(payload.style, metrics, filters)
 
    # 3. Generate AI narrative
    try:
        response = client_ai.messages.create(
            model=_CLAUDE_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0,
            max_tokens=8000,
        )
        summary = response.content[0].text
    except Exception as exc:
        print(f"Claude error: {exc}\n{traceback.format_exc()}")
        raise HTTPException(status_code=502, detail=f"Claude error: {exc}")
 
    return raw_metrics, summary
 
 
# =============================================================================
# POST /report/pptx  —  PowerPoint report (filter-aware)
# =============================================================================
 
@app.post("/report/pptx")
def report_pptx(payload: ReportRequestWithSummary):
    filters = payload.filters

    try:
        ch_client   = get_click_client()
        raw_metrics = fetch_pipeline_data(filters, ch_client)
    except RuntimeError as exc:
        raise HTTPException(status_code=500, detail=str(exc))

    # Use the pre-generated summary if provided, otherwise generate a new one
    if payload.summary:
        summary = payload.summary
    else:
        metrics = _extract_computed_metrics(raw_metrics, filters)
        prompt  = build_prompt(payload.style, metrics, filters)
        try:
            response = client_ai.messages.create(
                model=_CLAUDE_MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0,
                max_tokens=8000,
            )
            summary = response.content[0].text
        except Exception as exc:
            raise HTTPException(status_code=502, detail=f"Claude error: {exc}")

    try:
        buf = build_pptx(raw_metrics, summary, filters=filters.model_dump())
    except Exception as exc:
        print(f"PPTX build error: {exc}\n{traceback.format_exc()}")
        raise HTTPException(status_code=500, detail=f"PPTX generation failed: {exc}")

    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": "attachment; filename=pipeline_report.pptx"},
    )

 
 
# =============================================================================
# POST /report/pdf  —  PDF report (filter-aware)
# =============================================================================
 
@app.post("/report/pdf")
def report_pdf(payload: ReportRequestWithSummary):
    filters = payload.filters

    try:
        ch_client   = get_click_client()
        raw_metrics = fetch_pipeline_data(filters, ch_client)
    except RuntimeError as exc:
        raise HTTPException(status_code=500, detail=str(exc))

    # Use the pre-generated summary if provided, otherwise generate a new one
    if payload.summary:
        summary = payload.summary
    else:
        metrics = _extract_computed_metrics(raw_metrics, filters)
        prompt  = build_prompt(payload.style, metrics, filters)
        try:
            response = client_ai.messages.create(
                model=_CLAUDE_MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0,
                max_tokens=8000,
            )
            summary = response.content[0].text   
        except Exception as exc:
            raise HTTPException(status_code=502, detail=f"Claude error: {exc}")

    try:
        buf = build_pdf(raw_metrics, summary, filters=filters.model_dump())
    except Exception as exc:
        print(f"PDF build error: {exc}\n{traceback.format_exc()}")
        raise HTTPException(status_code=500, detail=f"PDF generation failed: {exc}")

    return StreamingResponse(
        buf,
        media_type="application/pdf",
        headers={"Content-Disposition": "attachment; filename=pipeline_report.pdf"},
    )
    
# =============================================================================
# GET /report/pdf  —  for window.open() browser download
# =============================================================================

@app.get("/report/pdf")
def report_pdf_get():
    from fastapi.responses import StreamingResponse
    payload = ReportRequestWithSummary(
        style="executive",
        filters=Filters(),
        summary=""
    )
    return report_pdf(payload)


# =============================================================================
# GET /report/pptx  —  for window.open() browser download
# =============================================================================

@app.get("/report/pptx")
def report_pptx_get():
    from fastapi.responses import StreamingResponse
    payload = ReportRequestWithSummary(
        style="executive",
        filters=Filters(),
        summary=""
    )
    return report_pptx(payload)
