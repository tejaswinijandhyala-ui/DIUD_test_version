# =============================================================================
# generate_pdf.py  —  v4
#
# Fixes vs v3:
#   • No blank first page — cover starts immediately, no leading PageBreak
#   • Section title rendered INSIDE the navy bar using canvas (not frame)
#   • Margins tightened — no excessive gaps
#   • AI narrative is the main event — ALL bullets, full text, properly styled
#   • KPI / attainment cells use nested Tables (not list) for ReportLab compat
#   • Every content page: compact table on top, full AI diagnostic below
#   • Each section pulls the correct AI bullets (not repeated Revenue Position)
#   • Recommendations use data-driven text with real numbers
# =============================================================================

import io
import re
from collections import defaultdict

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate, Flowable,
    HRFlowable, NextPageTemplate, PageBreak,
    Paragraph, Spacer, Table, TableStyle, KeepTogether,
)

# =============================================================================
# Palette
# =============================================================================
C_NAVY      = colors.HexColor("#0D1B3E")
C_DNAV      = colors.HexColor("#0A1128")
C_BLUE      = colors.HexColor("#1565C0")
C_TEAL      = colors.HexColor("#00695C")
C_AMBER     = colors.HexColor("#F59E0B")
C_AMBER_D   = colors.HexColor("#E65100")
C_GREEN     = colors.HexColor("#2E7D32")
C_RED       = colors.HexColor("#B71C1C")
C_WHITE     = colors.white
C_BG        = colors.HexColor("#F7F9FC")
C_ROW_A     = colors.white
C_ROW_B     = colors.HexColor("#EEF4FF")
C_BORDER    = colors.HexColor("#CBD5E1")
C_TXT       = colors.HexColor("#1E293B")
C_TXT_MID   = colors.HexColor("#475569")
C_TXT_DIM   = colors.HexColor("#94A3B8")
C_GOLD      = colors.HexColor("#F59E0B")
C_DIVIDER   = colors.HexColor("#E2E8F0")

ACCENT = {
    "cover":      C_NAVY,
    "funnel":     colors.HexColor("#1565C0"),
    "quarterly":  colors.HexColor("#004D40"),
    "region":     colors.HexColor("#BF360C"),
    "source":     colors.HexColor("#4A148C"),
    "conversion": colors.HexColor("#880E4F"),
    "period":     colors.HexColor("#006064"),
    "deals":      colors.HexColor("#B71C1C"),
    "actions":    colors.HexColor("#1B5E20"),
}
ACTION_COLS = [
    colors.HexColor("#B71C1C"),
    colors.HexColor("#E65100"),
    colors.HexColor("#F59E0B"),
    colors.HexColor("#1565C0"),
    colors.HexColor("#1B5E20"),
]

# =============================================================================
# Page geometry — portrait A4
# =============================================================================
PW, PH   = A4          # 595 × 842 pts
ML = MR  = 0.60 * inch
MT       = 0.45 * inch
MB       = 0.40 * inch
HDR_H    = 44          # painted by canvas; title drawn inside it
FTR_H    = 20
CW       = PW - ML - MR   # usable content width ≈ 474 pts

# Frame sits BELOW the header bar
FRAME_Y  = MB + FTR_H
FRAME_H  = PH - HDR_H - MT - MB - FTR_H
# Cover frame occupies the full page (navy bg, no header bar)
COVER_FRAME_Y = MB
COVER_FRAME_H = PH - MT - MB


# =============================================================================
# Text helpers
# =============================================================================

def _strip(t: str) -> str:
    if not t: return ""
    t = re.sub(r'\*{1,3}(.*?)\*{1,3}', r'\1', t)
    t = re.sub(r'#{1,6}\s*', '', t)
    t = re.sub(r'^[\-\*•]\s*', '', t, flags=re.M)
    t = re.sub(r'`(.*?)`', r'\1', t)
    return t.strip()


def _parse(summary: str) -> dict:
    out, cur = {}, None
    for line in summary.split("\n"):
        s = line.strip()
        if not s: continue
        m = re.match(r'^\*\*(.+?)\*\*:?\s*$', s)
        if m and len(s) < 120:
            cur = m.group(1).strip(); out[cur] = []
        elif cur:
            c = _strip(s)
            if c: out[cur].append(c)
    return out

def _render_full_summary(story, S, summary_text):
    lines = summary_text.split("\n")

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Section headers
        if line.startswith("**") and line.endswith("**"):
            header = line.replace("**", "")
            story.append(Spacer(1, 10))
            story.append(Paragraph(
                f"<b>{header}</b>",
                ParagraphStyle(
                    "hdr",
                    fontSize=12,
                    leading=14,
                    textColor=colors.white,
                    spaceAfter=6
                )
            ))

        # Bullets
        elif line.startswith("-") or line.startswith("•"):
            content = line[1:].strip()
            story.append(Paragraph(f"• {content}", S["bullet"]))

        # Plain text
        else:
            story.append(Paragraph(line, S["bullet"]))
            
            
def _bullets(sections: dict, *keys) -> list:
    for k in keys:
        for sk, v in sections.items():
            if k.lower() in sk.lower() and v:
                return v
    return []


def _att_color(v) -> colors.Color:
    try: v = float(v or 0)
    except: v = 0
    if v >= 100: return C_GREEN
    if v >= 60:  return C_AMBER_D
    return C_RED


def _status(p: float) -> str:
    if p >= 100: return "On Track"
    if p >= 60:  return "At Risk"
    return "Below Target"


# =============================================================================
# Styles
# =============================================================================

def _styles():
    S = {}

    # ── Cover ─────────────────────────────────────────────────────────────────
    S["cov_title"] = ParagraphStyle("cov_title",
        fontSize=28, leading=34, fontName="Helvetica-Bold",
        textColor=C_WHITE, spaceAfter=4)
    S["cov_sub"] = ParagraphStyle("cov_sub",
        fontSize=12, leading=16, fontName="Helvetica-Bold",
        textColor=C_GOLD, spaceAfter=4)
    S["cov_meta"] = ParagraphStyle("cov_meta",
        fontSize=8.5, leading=12, fontName="Helvetica",
        textColor=colors.HexColor("#93C5FD"), spaceAfter=14)
    S["kpi_lbl"] = ParagraphStyle("kpi_lbl",
        fontSize=7, leading=9, fontName="Helvetica",
        textColor=C_TXT_DIM, alignment=TA_CENTER, spaceAfter=2)
    S["kpi_val"] = ParagraphStyle("kpi_val",
        fontSize=22, leading=26, fontName="Helvetica-Bold",
        textColor=C_WHITE, alignment=TA_CENTER, spaceAfter=1)
    S["kpi_sub"] = ParagraphStyle("kpi_sub",
        fontSize=8, leading=10, fontName="Helvetica",
        textColor=C_GOLD, alignment=TA_CENTER)
    S["att_lbl"] = ParagraphStyle("att_lbl",
        fontSize=7.5, leading=9, fontName="Helvetica",
        textColor=C_TXT_DIM, alignment=TA_CENTER, spaceAfter=2)

    # ── Section page ──────────────────────────────────────────────────────────
    # (title/subtitle drawn directly on canvas, not as flowable)

    # ── Tables ────────────────────────────────────────────────────────────────
    S["th"] = ParagraphStyle("th",
        fontSize=7, leading=9, fontName="Helvetica-Bold",
        textColor=C_WHITE, alignment=TA_CENTER)
    S["th_l"] = ParagraphStyle("th_l",
        fontSize=7, leading=9, fontName="Helvetica-Bold",
        textColor=C_WHITE, alignment=TA_LEFT)
    S["td"] = ParagraphStyle("td",
        fontSize=8, leading=10, fontName="Helvetica",
        textColor=C_TXT, alignment=TA_CENTER)
    S["td_l"] = ParagraphStyle("td_l",
        fontSize=8, leading=10, fontName="Helvetica",
        textColor=C_TXT, alignment=TA_LEFT)
    S["td_r"] = ParagraphStyle("td_r",
        fontSize=8, leading=10, fontName="Helvetica",
        textColor=C_TXT, alignment=TA_RIGHT)

    # ── Narrative ─────────────────────────────────────────────────────────────
    S["insight_hdr"] = ParagraphStyle("insight_hdr",
        fontSize=8, leading=10, fontName="Helvetica-Bold",
        textColor=C_WHITE, alignment=TA_LEFT)
    S["bullet"] = ParagraphStyle(
    "bullet",
    fontSize=10,
    leading=16,
    fontName="Helvetica",
    textColor=colors.white,
    leftIndent=14,
    firstLineIndent=-14,
    spaceBefore=2,
    spaceAfter=6,
    alignment=TA_JUSTIFY
)
    S["bullet_lbl"] = ParagraphStyle("bullet_lbl",
        fontSize=9.5, leading=15, fontName="Helvetica-Bold",
        textColor=C_TXT, leftIndent=14, firstLineIndent=-14,
        spaceBefore=1, spaceAfter=5)
    S["sub_hdr"] = ParagraphStyle("sub_hdr",
        fontSize=8.5, leading=11, fontName="Helvetica-Bold",
        textColor=C_TXT_MID, spaceBefore=8, spaceAfter=3)

    # ── Actions ───────────────────────────────────────────────────────────────
    S["act_num"] = ParagraphStyle("act_num",
        fontSize=22, leading=26, fontName="Helvetica-Bold",
        textColor=C_WHITE, alignment=TA_CENTER)
    S["act_title"] = ParagraphStyle("act_title",
        fontSize=10, leading=13, fontName="Helvetica-Bold",
        textColor=C_WHITE, spaceAfter=3)
    S["act_body"] = ParagraphStyle("act_body",
        fontSize=8.5, leading=13, fontName="Helvetica",
        textColor=colors.HexColor("#DBEAFE"))

    # ── Misc ──────────────────────────────────────────────────────────────────
    S["caption"] = ParagraphStyle("caption",
        fontSize=7, leading=9, fontName="Helvetica",
        textColor=C_TXT_DIM, alignment=TA_CENTER, spaceAfter=3)
    S["warn"] = ParagraphStyle("warn",
        fontSize=8, leading=11, fontName="Helvetica-Bold",
        textColor=C_RED, spaceAfter=4)
    S["cov_bullet"] = ParagraphStyle("cov_bullet",
        fontSize=9, leading=14, fontName="Helvetica",
        textColor=C_WHITE, leftIndent=12, firstLineIndent=-12,
        spaceAfter=4)
    S["cov_blbl"] = ParagraphStyle("cov_blbl",
        fontSize=8, leading=10, fontName="Helvetica-Bold",
        textColor=C_TXT_DIM, spaceAfter=5)

    return S


# =============================================================================
# Custom flowables
# =============================================================================

class AttBar(Flowable):
    def __init__(self, pct, w=1.6*inch, h=10):
        super().__init__()
        self.pct = float(pct or 0)
        self.w = w; self.h = h
    def wrap(self, *a): return self.w + 38, self.h + 2
    def draw(self):
        c = self.canv; c.saveState()
        fw = self.w * min(self.pct, 100) / 100
        c.setFillColor(colors.HexColor("#E2E8F0"))
        c.rect(0, 0, self.w, self.h, fill=1, stroke=0)
        c.setFillColor(C_AMBER)
        c.rect(0, 0, fw, self.h, fill=1, stroke=0)
        c.setFillColor(_att_color(self.pct))
        c.setFont("Helvetica-Bold", 7.5)
        c.drawString(self.w + 4, 1, f"{self.pct:.1f}%")
        c.restoreState()


class InsightBand(Flowable):
    """Coloured label band above the narrative block."""
    def __init__(self, label: str, accent: colors.Color, w: float):
        super().__init__()
        self.label = label.upper()
        self.accent = accent
        self.w = w
    def wrap(self, *a): return self.w, 20
    def draw(self):
        c = self.canv; c.saveState()
        c.setFillColor(self.accent)
        c.rect(0, 0, self.w, 20, fill=1, stroke=0)
        c.setFillColor(C_WHITE)
        c.setFont("Helvetica-Bold", 7.5)
        c.drawString(8, 6, self.label)
        c.restoreState()


class FunnelDiagram(Flowable):
    """
    Draws a 4-stage trapezoid funnel with counts, amounts, and conversion rates.
    stages: list of (label, count, amount_str, conv_rate_str)
    conv_rate_str is the rate FROM this stage to next (empty for last stage).
    """
    def __init__(self, stages, w=None, accent=None):
        super().__init__()
        self.stages  = stages          # [(label, count, amt, conv_pct_str), ...]
        self.w       = w or (CW * 0.56)
        self.accent  = accent or colors.HexColor("#1565C0")
        # height = n stages * stage_h + (n-1) * arrow_h
        self._stage_h = 38
        self._arrow_h = 22
        n = len(stages)
        self._h = n * self._stage_h + (n - 1) * self._arrow_h

    def wrap(self, *a):
        return self.w, self._h

    def draw(self):
        c = self.canv; c.saveState()
        n        = len(self.stages)
        top_w    = self.w
        bot_w    = self.w * 0.42
        sh       = self._stage_h
        ah       = self._arrow_h
        total_h  = self._h

        # colour ramp: darker as funnel narrows
        base_r, base_g, base_b = (
            self.accent.red, self.accent.green, self.accent.blue
        )

        for i, (label, count, amt, conv) in enumerate(self.stages):
            # trapezoid width at this row (linearly narrows)
            frac_top = 1.0 - i       / (n - 1) if n > 1 else 1.0
            frac_bot = 1.0 - (i + 1) / (n - 1) if n > 1 else 0.42
            w_top = top_w * max(frac_top, 0.42)
            w_bot = top_w * max(frac_bot, 0.42)
            x_top = (top_w - w_top) / 2
            x_bot = (top_w - w_bot) / 2
            y_top = total_h - i * (sh + ah) - sh

            # shade: lighten slightly for upper stages
            shade = 1.0 - i * 0.12
            fill_col = colors.Color(
                min(base_r * shade + (1 - shade) * 0.55, 1),
                min(base_g * shade + (1 - shade) * 0.65, 1),
                min(base_b * shade + (1 - shade) * 0.85, 1),
            )

            # trapezoid path
            c.setFillColor(fill_col)
            c.setStrokeColor(C_WHITE)
            c.setLineWidth(1.2)
            p = c.beginPath()
            p.moveTo(x_top, y_top + sh)
            p.lineTo(x_top + w_top, y_top + sh)
            p.lineTo(x_bot + w_bot, y_top)
            p.lineTo(x_bot, y_top)
            p.close()
            c.drawPath(p, fill=1, stroke=1)

            # label (left of centre)
            c.setFillColor(C_WHITE)
            c.setFont("Helvetica-Bold", 8)
            cx = top_w / 2
            cy = y_top + sh / 2 + 5
            c.drawCentredString(cx, cy, label)
            c.setFont("Helvetica", 7)
            c.drawCentredString(cx, cy - 11,
                f"{count:,} deals" + (f"  ·  {amt}" if amt else ""))

            # conversion arrow + rate between stages
            if conv and i < n - 1:
                arr_y_top = y_top - ah
                arr_mid   = arr_y_top + ah / 2
                arr_cx    = top_w / 2
                # arrow shaft
                c.setStrokeColor(colors.HexColor("#94A3B8"))
                c.setLineWidth(1.0)
                c.line(arr_cx, y_top, arr_cx, arr_y_top + 6)
                # arrowhead
                c.setFillColor(colors.HexColor("#94A3B8"))
                p2 = c.beginPath()
                p2.moveTo(arr_cx - 5, arr_y_top + 8)
                p2.lineTo(arr_cx + 5, arr_y_top + 8)
                p2.lineTo(arr_cx,     arr_y_top + 1)
                p2.close()
                c.drawPath(p2, fill=1, stroke=0)
                # conversion rate label
                conv_num = float(conv.replace("%", "")) if conv else 0
                conv_col = C_GREEN if conv_num >= 50 else (
                    C_AMBER_D if conv_num >= 30 else C_RED)
                c.setFillColor(conv_col)
                c.setFont("Helvetica-Bold", 7.5)
                c.drawCentredString(arr_cx + 30, arr_mid - 3, f"{conv} conv.")

        c.restoreState()


# =============================================================================
# Page templates
# =============================================================================

def _on_cover(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(C_NAVY)
    canvas.rect(0, 0, PW, PH, fill=1, stroke=0)
    # Left amber accent bar
    canvas.setFillColor(C_AMBER)
    canvas.rect(0, 0, 5, PH, fill=1, stroke=0)
    # Bottom strip
    canvas.setFillColor(C_DNAV)
    canvas.rect(0, 0, PW, 22, fill=1, stroke=0)
    canvas.setFillColor(C_TXT_DIM)
    canvas.setFont("Helvetica", 6.5)
    canvas.drawCentredString(PW/2, 7,
        "CONFIDENTIAL  ·  Prepared for Executive Review  ·  FY2026")
    canvas.restoreState()


def _on_content(canvas, doc):
    canvas.saveState()
    # Page background
    canvas.setFillColor(C_BG)
    canvas.rect(0, 0, PW, PH, fill=1, stroke=0)
    # Header bar
    accent = getattr(doc, "_accent", C_NAVY)
    canvas.setFillColor(accent)
    canvas.rect(0, PH - HDR_H, PW, HDR_H, fill=1, stroke=0)
    # Page title & subtitle drawn directly into the bar
    title    = getattr(doc, "_pg_title",    "")
    subtitle = getattr(doc, "_pg_subtitle", "")
    canvas.setFillColor(C_WHITE)
    canvas.setFont("Helvetica-Bold", 13)
    canvas.drawString(ML, PH - HDR_H + 26, title)
    if subtitle:
        canvas.setFillColor(colors.HexColor("#93C5FD"))
        canvas.setFont("Helvetica", 8)
        canvas.drawString(ML, PH - HDR_H + 10, subtitle)
    # Footer
    canvas.setFillColor(C_DNAV)
    canvas.rect(0, 0, PW, FTR_H, fill=1, stroke=0)
    canvas.setFillColor(C_TXT_DIM)
    canvas.setFont("Helvetica", 6.5)
    canvas.drawCentredString(PW/2, 6,
        f"Pipeline Report  ·  FY2026  ·  CONFIDENTIAL  ·  Page {doc.page}")
    canvas.restoreState()


def _build_doc(buf):
    doc = BaseDocTemplate(buf, pagesize=A4,
        leftMargin=ML, rightMargin=MR,
        topMargin=MT, bottomMargin=MB)

    cover_frame = Frame(ML, COVER_FRAME_Y, CW, COVER_FRAME_H, id="cover")
    content_frame = Frame(ML, FRAME_Y, CW, FRAME_H, id="content")

    doc.addPageTemplates([
        PageTemplate(id="Cover",   frames=[cover_frame],   onPage=_on_cover),
        PageTemplate(id="Content", frames=[content_frame], onPage=_on_content),
    ])
    return doc


# =============================================================================
# Table helpers
# =============================================================================

_BASE_TBL = TableStyle([
    ("BACKGROUND",    (0,0), (-1,0),  C_NAVY),
    ("TEXTCOLOR",     (0,0), (-1,0),  C_WHITE),
    ("FONTNAME",      (0,0), (-1,0),  "Helvetica-Bold"),
    ("FONTSIZE",      (0,0), (-1,0),  7),
    ("ALIGN",         (0,0), (-1,-1), "CENTER"),
    ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
    ("FONTNAME",      (0,1), (-1,-1), "Helvetica"),
    ("FONTSIZE",      (0,1), (-1,-1), 8),
    ("ROWBACKGROUNDS",(0,1), (-1,-1), [C_ROW_A, C_ROW_B]),
    ("GRID",          (0,0), (-1,-1), 0.3, C_BORDER),
    ("TOPPADDING",    (0,0), (-1,-1), 4),
    ("BOTTOMPADDING", (0,0), (-1,-1), 4),
    ("LEFTPADDING",   (0,0), (-1,-1), 5),
    ("RIGHTPADDING",  (0,0), (-1,-1), 5),
])


def _tbl(data, cw, extra=None):
    t = Table(data, colWidths=cw, repeatRows=1)
    ts = TableStyle(_BASE_TBL.getCommands())
    for cmd in (extra or []):
        ts.add(*cmd)
    t.setStyle(ts)
    return t


# =============================================================================
# Narrative block — THE main content of each page
# =============================================================================

def _narrative(story, S, bullets: list, accent: colors.Color,
               label: str = "KEY INSIGHTS"):
    """Render ALL bullets as proper flowing Paragraphs — no truncation."""
    if not bullets:
        return
    story.append(Spacer(1, 10))
    story.append(InsightBand(label, accent, CW))

    for b in bullets:
        clean = _strip(b)
        if not clean:
            continue
        # Bold the lead phrase (before first em-dash or colon) for scanability
        if " — " in clean[:90]:
            head, tail = clean.split(" — ", 1)
            txt = f"<b>{head}</b> — {tail}"
        elif ": " in clean[:70]:
            head, tail = clean.split(": ", 1)
            txt = f"<b>{head}:</b> {tail}"
        else:
            txt = clean
        story.append(Paragraph(f"• {txt}", S["bullet"]))

    story.append(Spacer(1, 6))


# =============================================================================
# Helpers
# =============================================================================

def _page(story, doc, accent, title, subtitle=""):
    """Emit a PageBreak and set canvas metadata for the next page."""
    story.append(PageBreak())
    doc._accent       = accent
    doc._pg_title     = title
    doc._pg_subtitle  = subtitle


def _att_val_para(pct, S):
    col = _att_color(pct)
    return Paragraph(f"<font color='#{col.hexval()[2:]}'><b>{pct:.1f}%</b></font>",
                     ParagraphStyle("av", fontSize=14, leading=17,
                                    alignment=TA_CENTER, fontName="Helvetica-Bold",
                                    textColor=col))


# =============================================================================
# Section visual helpers
# =============================================================================

def _attainment_table_section(story, S, p, CW_ref):
    """
    Render a detailed attainment table:
    Rows = Stage (5%, 10%, 20%+)
    Cols = MTD Actual | MTD L1 | QTD Actual | QTD L1 | YTD Actual | YTD L1 | YTD % | Status
    """
    ytd_5   = float(p.get("pct_l1_ytd_5",  0) or 0)
    ytd_10  = float(p.get("pct_l1_ytd_10", 0) or 0)
    ytd_20  = float(p.get("pct_l1_ytd_20", 0) or 0)
    l1_5    = int(p.get("l1_ytd_5",  0) or 0)
    l1_10   = int(p.get("l1_ytd_10", 0) or 0)
    l1_20   = int(p.get("l1_ytd_20", 0) or 0)
    mtd_5   = int(p.get("mtd_5",  0) or 0)
    mtd_10  = int(p.get("mtd_10", 0) or 0)
    mtd_20  = int(p.get("mtd_20", 0) or 0)
    qtd_5   = int(p.get("qtd_5",  0) or 0)
    qtd_10  = int(p.get("qtd_10", 0) or 0)
    qtd_20  = int(p.get("qtd_20", 0) or 0)
    ytd_5a  = int(p.get("ytd_5",  0) or 0)
    ytd_10a = int(p.get("ytd_10", 0) or 0)
    ytd_20a = int(p.get("ytd_20", 0) or 0)
    l1_mtd5  = int(p.get("l1_mtd_5",  0) or 0)
    l1_mtd10 = int(p.get("l1_mtd_10", 0) or 0)
    l1_mtd20 = int(p.get("l1_mtd_20", 0) or 0)
    l1_qtd5  = int(p.get("l1_qtd_5",  0) or 0)
    l1_qtd10 = int(p.get("l1_qtd_10", 0) or 0)
    l1_qtd20 = int(p.get("l1_qtd_20", 0) or 0)

    def status_p(pct, style="td"):
        col = _att_color(pct)
        label = _status(pct)
        return Paragraph(
            f"<font color='#{col.hexval()[2:]}'><b>{label}</b></font>",
            S[style])

    def pct_p(pct, style="td"):
        col = _att_color(pct)
        return Paragraph(
            f"<font color='#{col.hexval()[2:]}'><b>{pct:.1f}%</b></font>",
            S[style])

    th = S["th"]; td = S["td"]; td_l = S["td_l"]

    header = [
        Paragraph("Stage",        th),
        Paragraph("MTD",          th),
        Paragraph("MTD L1",       th),
        Paragraph("QTD",          th),
        Paragraph("QTD L1",       th),
        Paragraph("YTD",          th),
        Paragraph("YTD L1",       th),
        Paragraph("YTD %",        th),
        Paragraph("Status",       th),
    ]

    rows = [
        header,
        [
            Paragraph("5% IQM Held",    td_l),
            Paragraph(str(mtd_5),        td),
            Paragraph(str(l1_mtd5),      td),
            Paragraph(str(qtd_5),        td),
            Paragraph(str(l1_qtd5),      td),
            Paragraph(str(ytd_5a),       td),
            Paragraph(str(l1_5),         td),
            pct_p(ytd_5),
            status_p(ytd_5),
        ],
        [
            Paragraph("10% Discovery",  td_l),
            Paragraph(str(mtd_10),       td),
            Paragraph(str(l1_mtd10),     td),
            Paragraph(str(qtd_10),       td),
            Paragraph(str(l1_qtd10),     td),
            Paragraph(str(ytd_10a),      td),
            Paragraph(str(l1_10),        td),
            pct_p(ytd_10),
            status_p(ytd_10),
        ],
        [
            Paragraph("20%+ Qualified", td_l),
            Paragraph(str(mtd_20),       td),
            Paragraph(str(l1_mtd20),     td),
            Paragraph(str(qtd_20),       td),
            Paragraph(str(l1_qtd20),     td),
            Paragraph(str(ytd_20a),      td),
            Paragraph(str(l1_20),        td),
            pct_p(ytd_20),
            status_p(ytd_20),
        ],
    ]

    cw_ref = CW_ref
    col_widths = [
        cw_ref * 0.20,   # Stage
        cw_ref * 0.08,   # MTD
        cw_ref * 0.08,   # MTD L1
        cw_ref * 0.08,   # QTD
        cw_ref * 0.08,   # QTD L1
        cw_ref * 0.08,   # YTD
        cw_ref * 0.08,   # YTD L1
        cw_ref * 0.12,   # YTD %
        cw_ref * 0.20,   # Status
    ]

    tbl = _tbl(rows, col_widths, extra=[
        ("ALIGN",  (1, 0), (-1, -1), "CENTER"),
        ("ALIGN",  (0, 0), (0, -1),  "LEFT"),
        # colour the status col bg based on attainment
        ("BACKGROUND", (8, 1), (8, 1), _att_color(ytd_5)),
        ("BACKGROUND", (8, 2), (8, 2), _att_color(ytd_10)),
        ("BACKGROUND", (8, 3), (8, 3), _att_color(ytd_20)),
        ("TEXTCOLOR",  (8, 1), (8, -1), C_WHITE),
        ("FONTNAME",   (8, 1), (8, -1), "Helvetica-Bold"),
    ])
    story.append(tbl)


def _funnel_section(story, S, f, CW_ref):
    """
    Render a FunnelDiagram flowable + side stats table side-by-side.
    """
    cnt_5   = int(f.get("cnt_5pct",       0) or 0)
    cnt_10  = int(f.get("cnt_10pct",      0) or 0)
    cnt_20  = int(f.get("cnt_20pct",      0) or 0)
    cnt_won = int(f.get("cnt_closed_won", 0) or 0)
    cnt_out = int(f.get("cnt_fallen_out", 0) or 0)
    amt_5   = float(f.get("amt_5pct",  0) or 0)
    amt_10  = float(f.get("amt_10pct", 0) or 0)
    amt_20  = float(f.get("amt_20pct", 0) or 0)

    c_5_10  = f"{round(cnt_10  / cnt_5   * 100, 1)}%" if cnt_5  > 0 else "—"
    c_10_20 = f"{round(cnt_20  / cnt_10  * 100, 1)}%" if cnt_10 > 0 else "—"
    c_20_won= f"{round(cnt_won / cnt_20  * 100, 1)}%" if cnt_20 > 0 else "—"

    stages = [
        ("5% IQM Held",    cnt_5,   f"${amt_5/1e6:.1f}M",  c_5_10),
        ("10% Discovery",  cnt_10,  f"${amt_10/1e6:.1f}M", c_10_20),
        ("20%+ Qualified", cnt_20,  f"${amt_20/1e6:.1f}M", c_20_won),
        ("Closed Won",     cnt_won, "",                     ""),
    ]

    funnel_w = CW_ref * 0.54
    funnel   = FunnelDiagram(stages, w=funnel_w,
                             accent=colors.HexColor("#1565C0"))

    # Side stats table
    th = S["th"]; td = S["td"]; td_l = S["td_l"]
    side_rows = [
        [Paragraph("Metric",      th), Paragraph("Value",  th)],
        [Paragraph("5%→10% Conv",  td_l), Paragraph(c_5_10,  td)],
        [Paragraph("10%→20% Conv", td_l), Paragraph(c_10_20, td)],
        [Paragraph("20%→Won",      td_l), Paragraph(c_20_won,td)],
        [Paragraph("Fallen Out",   td_l), Paragraph(f"{cnt_out:,}", td)],
    ]
    side_w = CW_ref * 0.38
    side_cw = [side_w * 0.58, side_w * 0.42]
    side_tbl = _tbl(side_rows, side_cw, extra=[
        ("ALIGN",  (1, 0), (-1, -1), "CENTER"),
        ("ALIGN",  (0, 0), (0, -1),  "LEFT"),
    ])

    gap   = CW_ref * 0.06
    combo = Table(
        [[funnel, side_tbl]],
        colWidths=[funnel_w + gap * 0.3, side_w],
    )
    combo.setStyle(TableStyle([
        ("VALIGN",       (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING",   (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 0),
        ("LEFTPADDING",  (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(combo)


# =============================================================================
# MAIN
# =============================================================================

def build_pdf(raw_metrics: dict, summary: str,
              filters: dict | None = None) -> io.BytesIO:

    filters  = filters or {}
    sections = _parse(summary)
    S        = _styles()

    # ── Unpack ────────────────────────────────────────────────────────────────
    f  = raw_metrics.get("funnel",            {}) or {}
    p  = raw_metrics.get("period_attainment", {}) or {}
    rs = raw_metrics.get("region_source",     []) or []
    sv = raw_metrics.get("stage_velocity",    []) or []
    dw = raw_metrics.get("deals_to_watch",    []) or []
    qt = raw_metrics.get("quarterly_trend",   []) or []
    ip = raw_metrics.get("industry_product",  []) or []

    cnt_1   = int(f.get("cnt_1pct",       0) or 0)
    cnt_5   = int(f.get("cnt_5pct",       0) or 0)
    cnt_10  = int(f.get("cnt_10pct",      0) or 0)
    cnt_20  = int(f.get("cnt_20pct",      0) or 0)
    cnt_won = int(f.get("cnt_closed_won", 0) or 0)
    cnt_out = int(f.get("cnt_fallen_out", 0) or 0)
    amt_5   = float(f.get("amt_5pct",     0) or 0)
    amt_10  = float(f.get("amt_10pct",    0) or 0)
    amt_20  = float(f.get("amt_20pct",    0) or 0)

    ytd_5  = float(p.get("pct_l1_ytd_5",  0) or 0)
    ytd_10 = float(p.get("pct_l1_ytd_10", 0) or 0)
    ytd_20 = float(p.get("pct_l1_ytd_20", 0) or 0)
    l1_5   = float(p.get("l1_ytd_5",  0) or 0)
    l1_10  = float(p.get("l1_ytd_10", 0) or 0)
    l1_20  = float(p.get("l1_ytd_20", 0) or 0)
    mtd_5  = int(p.get("mtd_5",  0) or 0)
    mtd_10 = int(p.get("mtd_10", 0) or 0)
    mtd_20 = int(p.get("mtd_20", 0) or 0)
    qtd_5  = int(p.get("qtd_5",  0) or 0)
    qtd_10 = int(p.get("qtd_10", 0) or 0)
    qtd_20 = int(p.get("qtd_20", 0) or 0)
    l2_5   = float(p.get("pct_l2_ytd_5",  0) or 0)
    l2_10  = float(p.get("pct_l2_ytd_10", 0) or 0)
    l2_20  = float(p.get("pct_l2_ytd_20", 0) or 0)

    win_rate   = round(cnt_won / cnt_20  * 100, 1) if cnt_20  > 0 else 0.0
    conv_5_10  = round(cnt_10  / cnt_5   * 100, 1) if cnt_5   > 0 else 0.0
    conv_10_20 = round(cnt_20  / cnt_10  * 100, 1) if cnt_10  > 0 else 0.0

    # Region / source aggregation
    reg_5  = defaultdict(int);   reg_10  = defaultdict(int);  reg_20  = defaultdict(int)
    reg_l5 = defaultdict(float); reg_l10 = defaultdict(float);reg_l20 = defaultdict(float)
    reg_amt= defaultdict(float)
    src_5  = defaultdict(int);   src_10  = defaultdict(int);  src_20  = defaultdict(int)
    src_l20= defaultdict(float)

    for r in rs:
        reg = r.get("region",             "Unknown") or "Unknown"
        src = r.get("deal_source_rollup", "Unknown") or "Unknown"
        reg_5[reg]   += int(r.get("deals_5",   0) or 0)
        reg_10[reg]  += int(r.get("deals_10",  0) or 0)
        reg_20[reg]  += int(r.get("deals_20",  0) or 0)
        reg_l5[reg]  += float(r.get("l1_5",    0) or 0)
        reg_l10[reg] += float(r.get("l1_10",   0) or 0)
        reg_l20[reg] += float(r.get("l1_20",   0) or 0)
        reg_amt[reg] += float(r.get("amount_20",0)or 0)
        src_5[src]   += int(r.get("deals_5",   0) or 0)
        src_10[src]  += int(r.get("deals_10",  0) or 0)
        src_20[src]  += int(r.get("deals_20",  0) or 0)
        src_l20[src] += float(r.get("l1_20",   0) or 0)

    def _ra20(reg):
        l = reg_l20[reg]
        return round(reg_20[reg] / l * 100, 1) if l > 0 else 0.0

    sorted_regs = sorted(
        [r for r in reg_5 if r.lower() not in ("unknown","unattributed")],
        key=_ra20, reverse=True
    )
    sorted_srcs = sorted(src_5, key=lambda s: -src_5[s])

    fp         = [f"{k.replace('_',' ').title()}: {v}" for k,v in filters.items() if v]
    filter_str = "  ·  ".join(fp) if fp else "Full Pipeline — No Filters"

    # ── AI bucket shorthand ───────────────────────────────────────────────────
    def B(*keys): return _bullets(sections, *keys)

    # =========================================================================
    # BUILD
    # =========================================================================
    buf   = io.BytesIO()
    doc   = _build_doc(buf)
    story = []

    # ── Set initial accent (used by _on_cover — doesn't matter) ──────────────
    doc._accent      = C_NAVY
    doc._pg_title    = ""
    doc._pg_subtitle = ""

    # =========================================================================
    # PAGE 1 — COVER  (no leading PageBreak — cover IS the first page)
    # =========================================================================
    story.append(NextPageTemplate("Cover"))
    # No PageBreak here — the doc starts on the Cover template directly

    story.append(Spacer(1, 0.5 * inch))
    story.append(Paragraph("Pipeline Performance Report", S["cov_title"]))
    story.append(Paragraph("FY2026 Executive Briefing", S["cov_sub"]))
    story.append(Paragraph(filter_str, S["cov_meta"]))
    story.append(HRFlowable(width=CW, thickness=2, color=C_AMBER, spaceAfter=14))

    # ── KPI cards ─────────────────────────────────────────────────────────────
    kw = (CW - 3*6) / 4
    def _kpi(lbl, val, sub=""):
        inner = [[Paragraph(lbl, S["kpi_lbl"])],
                 [Paragraph(val, S["kpi_val"])]]
        if sub: inner.append([Paragraph(sub, S["kpi_sub"])])
        t = Table(inner, colWidths=[kw - 12])
        t.setStyle(TableStyle([
            ("ALIGN",        (0,0),(-1,-1),"CENTER"),
            ("VALIGN",       (0,0),(-1,-1),"MIDDLE"),
            ("TOPPADDING",   (0,0),(-1,-1), 4),
            ("BOTTOMPADDING",(0,0),(-1,-1), 4),
            ("LEFTPADDING",  (0,0),(-1,-1), 0),
            ("RIGHTPADDING", (0,0),(-1,-1), 0),
        ]))
        return t

    kpi_row = [[
    _kpi("5% IQM Held",    f"{cnt_5:,}",  f"${amt_5/1e6:.1f}M"),
    _kpi("10% Discovery",  f"{cnt_10:,}", f"${amt_10/1e6:.1f}M"),
    _kpi("20%+ Qualified", f"{cnt_20:,}", f"${amt_20/1e6:.1f}M"),
    _kpi("Closed Won",     f"{cnt_won:,}", "YTD"),
    ]]
    kpi_tbl = Table(kpi_row, colWidths=[kw]*4, rowHeights=[0.95*inch])
    kpi_tbl.setStyle(TableStyle([
        ("BACKGROUND",   (0,0),(-1,-1), C_DNAV),
        ("ALIGN",        (0,0),(-1,-1),"CENTER"),
        ("VALIGN",       (0,0),(-1,-1),"MIDDLE"),
        ("TOPPADDING",   (0,0),(-1,-1), 0),
        ("BOTTOMPADDING",(0,0),(-1,-1), 0),
        ("LEFTPADDING",  (0,0),(-1,-1), 0),
        ("RIGHTPADDING", (0,0),(-1,-1), 6),
        ("LINEAFTER",    (0,0),(2,0),   1, colors.HexColor("#1E3A5F")),
    ]))
    story.append(kpi_tbl)
    story.append(Spacer(1, 10))

    # ── Attainment row ────────────────────────────────────────────────────────
    def _att(lbl, pct):
        col = _att_color(pct)
        inner = [
            [Paragraph(lbl, S["att_lbl"])],
            [Paragraph(f"<b>{pct:.1f}%</b>",
                       ParagraphStyle("av", fontSize=15, leading=18,
                                      fontName="Helvetica-Bold",
                                      textColor=col, alignment=TA_CENTER))],
        ]
        t = Table(inner, colWidths=[CW/4 - 10])
        t.setStyle(TableStyle([
            ("ALIGN",        (0,0),(-1,-1),"CENTER"),
            ("VALIGN",       (0,0),(-1,-1),"MIDDLE"),
            ("TOPPADDING",   (0,0),(-1,-1), 4),
            ("BOTTOMPADDING",(0,0),(-1,-1), 4),
            ("LEFTPADDING",  (0,0),(-1,-1), 0),
            ("RIGHTPADDING", (0,0),(-1,-1), 0),
        ]))
        return t

    att_row = [[
        _att("5% YTD Attainment",  ytd_5),
        _att("10% YTD Attainment", ytd_10),
        _att("20% YTD Attainment", ytd_20),
        _att("Win Rate 20%→Won",   win_rate),
    ]]
    att_tbl = Table(att_row, colWidths=[CW/4]*4, rowHeights=[0.72*inch])
    att_tbl.setStyle(TableStyle([
        ("BACKGROUND",   (0,0),(-1,-1), C_NAVY),
        ("ALIGN",        (0,0),(-1,-1),"CENTER"),
        ("VALIGN",       (0,0),(-1,-1),"MIDDLE"),
        ("TOPPADDING",   (0,0),(-1,-1), 0),
        ("BOTTOMPADDING",(0,0),(-1,-1), 0),
        ("LEFTPADDING",  (0,0),(-1,-1), 0),
        ("RIGHTPADDING", (0,0),(-1,-1), 0),
        ("LINEAFTER",    (0,0),(2,0),   1, colors.HexColor("#1E3A5F")),
    ]))
    story.append(att_tbl)
    story.append(Spacer(1, 14))

    # ── Cover summary bullets ─────────────────────────────────────────────────
    story.append(HRFlowable(width=CW, thickness=0.5,
                            color=C_TXT_DIM, spaceAfter=8))

    story.append(Paragraph("EXECUTIVE SUMMARY", S["cov_blbl"]))

    _render_full_summary(story, S, summary)

    # ── Build ─────────────────────────────────────────────────────────────────
    doc.build(story)
    buf.seek(0)
    return buf
