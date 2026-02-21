"""
PDF report generation using reportlab.
Produces a professional multi-section working paper with cover page,
diagonal watermark, and per-page headers/footers.
"""

import io
import logging
from datetime import date

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate, NextPageTemplate,
    Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, KeepTogether,
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

from utils.formatters import format_currency, format_percent, format_ratio, format_days

logger = logging.getLogger(__name__)

# ── Colour palette ──────────────────────────────────────────────────────────
NAVY     = colors.HexColor("#1B2A4A")
NAVY_MID = colors.HexColor("#2C3E6B")
GREY     = colors.HexColor("#6B7280")
LIGHT_GREY = colors.HexColor("#F3F4F6")
MID_GREY   = colors.HexColor("#E5E7EB")
WHITE    = colors.white
RED_COL  = colors.HexColor("#DC2626")
GREEN    = colors.HexColor("#16A34A")
AMBER    = colors.HexColor("#D97706")
GREEN_BG = colors.HexColor("#DCFCE7")
AMBER_BG = colors.HexColor("#FEF3C7")
RED_BG   = colors.HexColor("#FEE2E2")

STATUS_COLORS = {
    "green": (GREEN, GREEN_BG),
    "amber": (AMBER, AMBER_BG),
    "red":   (RED_COL, RED_BG),
    "grey":  (GREY, LIGHT_GREY),
}


# ── Styles ──────────────────────────────────────────────────────────────────

def _get_styles():
    styles = getSampleStyleSheet()

    styles.add(ParagraphStyle("CoverFirm",   parent=styles["Normal"],
        fontSize=26, textColor=WHITE, fontName="Helvetica-Bold", spaceAfter=6))
    styles.add(ParagraphStyle("CoverTitle",  parent=styles["Normal"],
        fontSize=15, textColor=colors.HexColor("#CBD5E1"), fontName="Helvetica", spaceAfter=4))
    styles.add(ParagraphStyle("CoverBanner", parent=styles["Normal"],
        fontSize=9, textColor=WHITE, fontName="Helvetica-Bold", alignment=TA_CENTER))
    styles.add(ParagraphStyle("CoverLabel",  parent=styles["Normal"],
        fontSize=10, textColor=NAVY, fontName="Helvetica-Bold"))
    styles.add(ParagraphStyle("CoverValue",  parent=styles["Normal"],
        fontSize=10, textColor=colors.black, fontName="Helvetica"))
    styles.add(ParagraphStyle("SectionHeading", parent=styles["Normal"],
        fontSize=13, textColor=NAVY, fontName="Helvetica-Bold",
        spaceBefore=14, spaceAfter=6))
    styles.add(ParagraphStyle("SubHeading",  parent=styles["Normal"],
        fontSize=11, textColor=NAVY_MID, fontName="Helvetica-Bold",
        spaceBefore=8, spaceAfter=4))
    styles.add(ParagraphStyle("MetricLabel", parent=styles["Normal"],
        fontSize=9, textColor=colors.black, fontName="Helvetica"))
    styles.add(ParagraphStyle("Body",        parent=styles["Normal"],
        fontSize=10, textColor=colors.black, fontName="Helvetica", leading=14, spaceAfter=5))
    styles.add(ParagraphStyle("BodySmall",   parent=styles["Normal"],
        fontSize=8, textColor=GREY, fontName="Helvetica", leading=11, spaceAfter=3))
    styles.add(ParagraphStyle("RedFlag",     parent=styles["Normal"],
        fontSize=10, textColor=RED_COL, fontName="Helvetica", leading=14, spaceAfter=4))
    styles.add(ParagraphStyle("Footer",      parent=styles["Normal"],
        fontSize=7, textColor=GREY, fontName="Helvetica", alignment=TA_CENTER))
    styles.add(ParagraphStyle("BulletItem",  parent=styles["Normal"],
        fontSize=10, textColor=colors.black, fontName="Helvetica",
        leading=14, spaceAfter=3, leftIndent=12))

    return styles


# ── Canvas callbacks ─────────────────────────────────────────────────────────

def _draw_watermark(canvas, w, h, alpha_grey="#DDDDDD"):
    """Draw diagonal 'INTERNAL USE ONLY' watermark across the page."""
    canvas.saveState()
    canvas.setFont("Helvetica-Bold", 55)
    canvas.setFillColor(colors.HexColor(alpha_grey))
    canvas.translate(w / 2, h / 2)
    canvas.rotate(40)
    canvas.drawCentredString(0, 0, "INTERNAL USE ONLY")
    canvas.restoreState()


def _make_cover_cb(session_info: dict, firm_name: str, analysis_date: str):
    """Return a canvas callback that draws the cover-page chrome."""
    client_name = session_info.get("client_name", "Client")

    def _draw(canvas, doc):
        canvas.saveState()
        w, h = A4

        # Navy top band
        canvas.setFillColor(NAVY)
        canvas.rect(0, h - 5 * cm, w, 5 * cm, fill=True, stroke=False)

        # Red 'INTERNAL USE ONLY' strip
        canvas.setFillColor(RED_COL)
        canvas.rect(0, h - 5.7 * cm, w, 0.7 * cm, fill=True, stroke=False)
        canvas.setFillColor(WHITE)
        canvas.setFont("Helvetica-Bold", 9)
        canvas.drawCentredString(w / 2, h - 5.4 * cm, "INTERNAL USE ONLY — NOT FOR DISTRIBUTION")

        # Watermark (lighter on cover)
        _draw_watermark(canvas, w, h, alpha_grey="#EBEBEB")

        # Bottom footer
        canvas.setStrokeColor(NAVY)
        canvas.setLineWidth(0.5)
        canvas.line(1.5 * cm, 1.4 * cm, w - 1.5 * cm, 1.4 * cm)
        canvas.setFillColor(GREY)
        canvas.setFont("Helvetica", 7)
        canvas.drawString(1.5 * cm, 0.8 * cm, client_name)
        canvas.drawCentredString(w / 2, 0.8 * cm, analysis_date)
        canvas.drawRightString(w - 1.5 * cm, 0.8 * cm, "Prepared by FinSight")

        canvas.restoreState()

    return _draw


def _make_page_cb(client_name: str, firm_name: str, analysis_date: str):
    """Return a canvas callback for all non-cover pages."""

    def _draw(canvas, doc):
        canvas.saveState()
        w, h = A4

        # Navy header bar
        canvas.setFillColor(NAVY)
        canvas.rect(0, h - 1.5 * cm, w, 1.5 * cm, fill=True, stroke=False)

        canvas.setFillColor(WHITE)
        canvas.setFont("Helvetica-Bold", 8)
        canvas.drawString(1.5 * cm, h - 1.0 * cm, "INTERNAL USE ONLY")

        canvas.setFont("Helvetica", 9)
        canvas.drawCentredString(w / 2, h - 1.0 * cm, f"FinSight Financial Analysis — {firm_name}")

        canvas.setFont("Helvetica", 9)
        canvas.drawRightString(w - 1.5 * cm, h - 1.0 * cm, f"Page {doc.page}")

        # Diagonal watermark (very light on inner pages)
        _draw_watermark(canvas, w, h, alpha_grey="#EEEEEE")

        # Footer
        canvas.setStrokeColor(NAVY)
        canvas.setLineWidth(0.5)
        canvas.line(1.5 * cm, 1.4 * cm, w - 1.5 * cm, 1.4 * cm)
        canvas.setFillColor(GREY)
        canvas.setFont("Helvetica", 7)
        canvas.drawString(1.5 * cm, 0.8 * cm, client_name)
        canvas.drawCentredString(w / 2, 0.8 * cm, analysis_date)
        canvas.drawRightString(w - 1.5 * cm, 0.8 * cm, "Prepared by FinSight")

        canvas.restoreState()

    return _draw


# ── Story helpers ────────────────────────────────────────────────────────────

def _status_dot(status: str) -> str:
    colour_map = {
        "green": "#16A34A", "amber": "#D97706", "red": "#DC2626", "grey": "#9CA3AF"
    }
    colour = colour_map.get(status, "#9CA3AF")
    return f'<font color="{colour}">●</font>'


def _status_label(status: str) -> str:
    return {"green": "Good", "amber": "Review", "red": "Concern", "grey": "N/A"}.get(status, "N/A")


def _metric_table(metrics: dict, category: str, styles, period_labels: list) -> Table | None:
    """Build a styled metrics table for one category."""
    label_cur = period_labels[0] if period_labels else "Current"
    label_pri = period_labels[1] if len(period_labels) > 1 else "Prior"

    header = ["Metric", "Status", label_cur, label_pri, "Trend"]
    rows = [header]

    for m in metrics.values():
        if m.category != category:
            continue
        dot = _status_dot(m.status)
        rows.append([
            Paragraph(m.label, styles["MetricLabel"]),
            Paragraph(dot, styles["MetricLabel"]),
            m.current_fmt,
            m.prior_fmt,
            m.trend,
        ])

    if len(rows) == 1:
        return None

    col_widths = [7 * cm, 1.5 * cm, 3 * cm, 3 * cm, 1.5 * cm]
    t = Table(rows, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0),  NAVY),
        ("TEXTCOLOR",     (0, 0), (-1, 0),  WHITE),
        ("FONTNAME",      (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, 0),  9),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [WHITE, LIGHT_GREY]),
        ("FONTSIZE",      (0, 1), (-1, -1), 9),
        ("GRID",          (0, 0), (-1, -1), 0.3, GREY),
        ("ALIGN",         (1, 0), (-1, -1), "CENTER"),
        ("ALIGN",         (0, 0), (0, -1),  "LEFT"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return t


def _snapshot_table(metrics: dict, styles, period_labels: list) -> Table:
    """Build a compact 6-metric snapshot table for the exec summary."""
    label_cur = period_labels[0] if period_labels else "Current"
    label_pri = period_labels[1] if len(period_labels) > 1 else "Prior"

    # Pick key metrics to spotlight
    spotlight = [
        "gross_profit_margin", "net_profit_margin", "current_ratio",
        "ebitda_margin", "debtor_days", "debt_to_equity",
    ]
    rows = [["Metric", "Status", label_cur, label_pri, "Trend"]]
    for key in spotlight:
        m = metrics.get(key)
        if m is None:
            continue
        dot = _status_dot(m.status)
        rows.append([
            Paragraph(m.label, styles["MetricLabel"]),
            Paragraph(dot, styles["MetricLabel"]),
            m.current_fmt,
            m.prior_fmt,
            m.trend,
        ])

    col_widths = [6.5 * cm, 1.5 * cm, 3 * cm, 3 * cm, 2 * cm]
    t = Table(rows, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0),  NAVY),
        ("TEXTCOLOR",     (0, 0), (-1, 0),  WHITE),
        ("FONTNAME",      (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, 0),  9),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [WHITE, LIGHT_GREY]),
        ("FONTSIZE",      (0, 1), (-1, -1), 9),
        ("GRID",          (0, 0), (-1, -1), 0.3, GREY),
        ("ALIGN",         (1, 0), (-1, -1), "CENTER"),
        ("ALIGN",         (0, 0), (0, -1),  "LEFT"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return t


# ── Main generator ───────────────────────────────────────────────────────────

def generate_pdf_report(
    analysis_result,
    session_info: dict,
    commentary: str,
    firm_name: str = "Your Firm Name",
    chart_images: dict = None,
) -> bytes:
    """
    Generate a polished PDF working paper and return as bytes.

    Sections:
      Page 1  — Cover (client details, internal use banner)
      Page 2  — Executive Summary (health indicator, snapshot metrics, red flags)
      Page 3+ — Detailed category sections (Profitability, Liquidity, Efficiency,
                 Leverage, Growth)
      Next    — ATO Benchmark Comparison
      Next    — AI Commentary (if provided)

    Every page has:
      • Header: "INTERNAL USE ONLY" | firm name | page number
      • Footer: client name | analysis date | "Prepared by FinSight"
      • Diagonal grey watermark
    """
    buffer = io.BytesIO()
    w, h = A4

    client_name  = session_info.get("client_name", "Client")
    abn          = session_info.get("abn", "")
    industry     = session_info.get("industry", "")
    fy_end       = session_info.get("financial_year_end", "")
    currency     = session_info.get("currency", "AUD")
    analysis_date = date.today().strftime("%d %B %Y")
    period_labels = analysis_result.period_labels

    styles = _get_styles()

    # ── Page templates ──────────────────────────────────────────────────────
    cover_cb = _make_cover_cb(session_info, firm_name, analysis_date)
    page_cb  = _make_page_cb(client_name, firm_name, analysis_date)

    # Cover frame: sits below the navy band (h - 5.7cm) with bottom margin
    cover_frame = Frame(
        1.5 * cm, 2 * cm, w - 3 * cm, h - 8 * cm,
        leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
    )
    # Content frame: below header bar (1.5cm) with bottom margin (2cm)
    content_frame = Frame(
        1.5 * cm, 2 * cm, w - 3 * cm, h - 3.8 * cm,
        leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
    )

    cover_template   = PageTemplate(id="cover",   frames=[cover_frame],   onPage=cover_cb)
    content_template = PageTemplate(id="content", frames=[content_frame], onPage=page_cb)

    doc = BaseDocTemplate(
        buffer,
        pagesize=A4,
        pageTemplates=[cover_template, content_template],
    )

    story = []

    # ── PAGE 1: Cover ───────────────────────────────────────────────────────
    # Story items rendered in the white space below the navy band
    story.append(Spacer(1, 0.5 * cm))
    story.append(Paragraph(firm_name, styles["CoverFirm"]))
    story.append(Paragraph("Financial Analysis Report", styles["CoverTitle"]))
    story.append(Spacer(1, 0.3 * cm))

    # Client info table on cover
    info_rows = [
        ["Client", client_name],
        ["ABN", abn or "Not provided"],
        ["Industry", industry],
        ["Financial Year", fy_end],
        ["Currency", currency],
        ["Date of Analysis", analysis_date],
    ]
    cover_tbl = Table(info_rows, colWidths=[4 * cm, 12 * cm])
    cover_tbl.setStyle(TableStyle([
        ("FONTNAME",       (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTSIZE",       (0, 0), (-1, -1), 10),
        ("TEXTCOLOR",      (0, 0), (0, -1), NAVY),
        ("TEXTCOLOR",      (1, 0), (1, -1), colors.black),
        ("ROWBACKGROUNDS", (0, 0), (-1, -1), [WHITE, LIGHT_GREY]),
        ("TOPPADDING",     (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING",  (0, 0), (-1, -1), 5),
        ("GRID",           (0, 0), (-1, -1), 0.3, MID_GREY),
    ]))
    story.append(cover_tbl)

    # Switch to content template for page 2+
    story.append(NextPageTemplate("content"))
    story.append(PageBreak())

    # ── PAGE 2: Executive Summary ────────────────────────────────────────────
    story.append(Paragraph("Executive Summary", styles["SectionHeading"]))
    story.append(HRFlowable(width="100%", thickness=1.5, color=NAVY))
    story.append(Spacer(1, 0.3 * cm))

    # Health indicator: count green / amber / red
    g_count = sum(1 for m in analysis_result.metrics.values() if m.status == "green")
    a_count = sum(1 for m in analysis_result.metrics.values() if m.status == "amber")
    r_count = sum(1 for m in analysis_result.metrics.values() if m.status == "red")
    total = len(analysis_result.metrics)

    health_data = [
        ["Overall Metrics Health", f"{g_count} Good", f"{a_count} Review", f"{r_count} Concern"],
    ]
    health_tbl = Table(health_data, colWidths=[6 * cm, 3 * cm, 3 * cm, 4 * cm])
    health_tbl.setStyle(TableStyle([
        ("FONTNAME",      (0, 0), (-1, -1), "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, -1), 10),
        ("TEXTCOLOR",     (0, 0), (0, -1),  NAVY),
        ("TEXTCOLOR",     (1, 0), (1, -1),  GREEN),
        ("TEXTCOLOR",     (2, 0), (2, -1),  AMBER),
        ("TEXTCOLOR",     (3, 0), (3, -1),  RED_COL),
        ("BACKGROUND",    (0, 0), (-1, -1), LIGHT_GREY),
        ("TOPPADDING",    (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING",   (0, 0), (-1, -1), 10),
        ("ALIGN",         (1, 0), (-1, -1), "CENTER"),
    ]))
    story.append(health_tbl)
    story.append(Spacer(1, 0.4 * cm))

    # Snapshot metrics table
    story.append(Paragraph("Key Metrics Snapshot", styles["SubHeading"]))
    snap = _snapshot_table(analysis_result.metrics, styles, period_labels)
    story.append(snap)
    story.append(Spacer(1, 0.4 * cm))

    # Red flags on exec summary
    if analysis_result.red_flags:
        story.append(Paragraph("Red Flags Identified", styles["SubHeading"]))
        for flag in analysis_result.red_flags:
            story.append(Paragraph(f"⚠ {flag}", styles["RedFlag"]))
    else:
        story.append(Paragraph("No red flags detected.", styles["Body"]))

    story.append(PageBreak())

    # ── DETAILED SECTIONS ────────────────────────────────────────────────────
    categories = [
        ("profitability", "Profitability"),
        ("liquidity",     "Liquidity & Working Capital"),
        ("efficiency",    "Operational Efficiency"),
        ("leverage",      "Leverage & Solvency"),
        ("growth",        "Growth Metrics"),
    ]

    for cat_key, cat_label in categories:
        cat_metrics = {k: v for k, v in analysis_result.metrics.items() if v.category == cat_key}
        if not cat_metrics:
            continue

        story.append(Paragraph(cat_label, styles["SectionHeading"]))
        story.append(HRFlowable(width="100%", thickness=1, color=NAVY))
        story.append(Spacer(1, 0.2 * cm))

        tbl = _metric_table(analysis_result.metrics, cat_key, styles, period_labels)
        if tbl:
            story.append(KeepTogether([tbl]))

        # Metric notes
        for m in cat_metrics.values():
            if m.notes:
                story.append(Paragraph(f"⚠ {m.notes}", styles["RedFlag"]))

        story.append(Spacer(1, 0.5 * cm))

    # ── ATO BENCHMARKS ───────────────────────────────────────────────────────
    if analysis_result.benchmark_comparisons:
        story.append(PageBreak())
        story.append(Paragraph("ATO Small Business Benchmark Comparison", styles["SectionHeading"]))
        story.append(HRFlowable(width="100%", thickness=1, color=NAVY))
        story.append(Paragraph(
            f"Industry: <b>{industry}</b>. Figures expressed as % of turnover. "
            "ATO benchmarks are updated annually and may lag by one financial year.",
            styles["BodySmall"],
        ))
        story.append(Spacer(1, 0.3 * cm))

        bm_header = ["Expense Category", "Client %", "ATO Low", "ATO High", "Status"]
        bm_rows = [bm_header]
        for comp in analysis_result.benchmark_comparisons.values():
            actual = comp.get("actual_pct")
            low    = comp.get("benchmark_low")
            high   = comp.get("benchmark_high")
            in_range = (
                actual is not None and low is not None and high is not None
                and low <= actual <= high
            )
            status = "green" if in_range else ("amber" if actual is not None else "grey")
            bm_rows.append([
                comp["label"],
                f"{actual:.1f}%" if actual is not None else "N/A",
                f"{low:.0f}%"    if low    is not None else "N/A",
                f"{high:.0f}%"   if high   is not None else "N/A",
                Paragraph(_status_dot(status), styles["MetricLabel"]),
            ])

        bm_table = Table(bm_rows, colWidths=[6 * cm, 3 * cm, 3 * cm, 3 * cm, 2 * cm], repeatRows=1)
        bm_table.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, 0),  NAVY),
            ("TEXTCOLOR",     (0, 0), (-1, 0),  WHITE),
            ("FONTNAME",      (0, 0), (-1, 0),  "Helvetica-Bold"),
            ("FONTSIZE",      (0, 0), (-1, -1), 9),
            ("ROWBACKGROUNDS",(0, 1), (-1, -1), [WHITE, LIGHT_GREY]),
            ("GRID",          (0, 0), (-1, -1), 0.3, GREY),
            ("ALIGN",         (1, 0), (-1, -1), "CENTER"),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING",    (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(bm_table)
        story.append(Spacer(1, 0.3 * cm))
        story.append(Paragraph(
            "Source: ATO Small Business Benchmarks (ato.gov.au). "
            "Benchmark ranges indicate the typical range for businesses in this industry sector.",
            styles["BodySmall"],
        ))

    # ── COMMENTARY ───────────────────────────────────────────────────────────
    if commentary:
        story.append(PageBreak())
        story.append(Paragraph("Financial Commentary", styles["SectionHeading"]))
        story.append(HRFlowable(width="100%", thickness=1, color=NAVY))
        story.append(Paragraph(
            "The following commentary has been generated by a local AI model to assist in "
            "client meeting preparation. <b>Review and edit as required before use.</b> "
            "This commentary does not constitute professional advice.",
            styles["BodySmall"],
        ))
        story.append(Spacer(1, 0.3 * cm))

        for line in commentary.split("\n"):
            line = line.strip()
            if not line:
                story.append(Spacer(1, 0.15 * cm))
            elif line.startswith("## "):
                story.append(Paragraph(line[3:], styles["SubHeading"]))
            elif line.startswith("### "):
                story.append(Paragraph(line[4:], styles["SubHeading"]))
            elif line.startswith("- ") or line.startswith("* "):
                story.append(Paragraph(f"• {line[2:]}", styles["BulletItem"]))
            else:
                story.append(Paragraph(line, styles["Body"]))

    # ── APPENDIX NOTE ────────────────────────────────────────────────────────
    story.append(PageBreak())
    story.append(Paragraph("Appendix — Notes & Assumptions", styles["SectionHeading"]))
    story.append(HRFlowable(width="100%", thickness=1, color=NAVY))
    story.append(Spacer(1, 0.2 * cm))
    notes = [
        "All figures are extracted from the provided financial statements and have not been independently verified.",
        "Metric thresholds are indicative benchmarks. Professional judgement should be applied.",
        "ATO benchmark data is sourced from the ATO Small Business Benchmarks and may lag by one financial year.",
        "AI-generated commentary is based solely on the data provided and should be reviewed before client delivery.",
        f"Analysis generated: {analysis_date}",
        f"Reporting currency: {currency}",
    ]
    for note in notes:
        story.append(Paragraph(f"• {note}", styles["BulletItem"]))

    # ── Build ────────────────────────────────────────────────────────────────
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()
