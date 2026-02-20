"""
PDF report generation using reportlab.
Produces a professional financial analysis report.
"""

import io
import os
import logging
from datetime import date
from typing import Optional

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, KeepTogether
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

logger = logging.getLogger(__name__)

# ── Colour palette ─────────────────────────────────────────────────────────
NAVY = colors.HexColor("#1B2A4A")
GREY = colors.HexColor("#6B7280")
LIGHT_GREY = colors.HexColor("#F3F4F6")
WHITE = colors.white
GREEN = colors.HexColor("#16A34A")
AMBER = colors.HexColor("#D97706")
RED_COL = colors.HexColor("#DC2626")
GREEN_BG = colors.HexColor("#DCFCE7")
AMBER_BG = colors.HexColor("#FEF3C7")
RED_BG = colors.HexColor("#FEE2E2")
GREY_BG = colors.HexColor("#F9FAFB")

STATUS_COLORS = {
    "green": (GREEN, GREEN_BG),
    "amber": (AMBER, AMBER_BG),
    "red": (RED_COL, RED_BG),
    "grey": (GREY, GREY_BG),
}
STATUS_LABELS = {
    "green": "●",
    "amber": "●",
    "red": "●",
    "grey": "○",
}


def _get_styles():
    styles = getSampleStyleSheet()

    styles.add(ParagraphStyle(
        "FirmHeader",
        parent=styles["Normal"],
        fontSize=20,
        textColor=NAVY,
        fontName="Helvetica-Bold",
        spaceAfter=2,
    ))
    styles.add(ParagraphStyle(
        "SubHeader",
        parent=styles["Normal"],
        fontSize=12,
        textColor=GREY,
        fontName="Helvetica",
        spaceAfter=8,
    ))
    styles.add(ParagraphStyle(
        "SectionHeading",
        parent=styles["Normal"],
        fontSize=13,
        textColor=NAVY,
        fontName="Helvetica-Bold",
        spaceBefore=14,
        spaceAfter=6,
        borderPad=4,
    ))
    styles.add(ParagraphStyle(
        "MetricLabel",
        parent=styles["Normal"],
        fontSize=9,
        textColor=colors.black,
        fontName="Helvetica",
    ))
    styles.add(ParagraphStyle(
        "CommentaryBody",
        parent=styles["Normal"],
        fontSize=10,
        textColor=colors.black,
        fontName="Helvetica",
        leading=14,
        spaceAfter=6,
    ))
    styles.add(ParagraphStyle(
        "Footer",
        parent=styles["Normal"],
        fontSize=8,
        textColor=GREY,
        fontName="Helvetica",
        alignment=TA_CENTER,
    ))
    styles.add(ParagraphStyle(
        "RedFlag",
        parent=styles["Normal"],
        fontSize=10,
        textColor=RED_COL,
        fontName="Helvetica",
        leading=14,
        spaceAfter=4,
    ))

    return styles


def _header_footer(canvas, doc):
    """Draw page header and footer."""
    canvas.saveState()
    w, h = A4

    # Top bar
    canvas.setFillColor(NAVY)
    canvas.rect(0, h - 1.5 * cm, w, 1.5 * cm, fill=True, stroke=False)

    # Firm name in header
    canvas.setFillColor(WHITE)
    canvas.setFont("Helvetica-Bold", 11)
    canvas.drawString(1.5 * cm, h - 1 * cm, "FinSight Financial Analysis")

    # Page number
    canvas.setFont("Helvetica", 9)
    canvas.drawRightString(w - 1.5 * cm, h - 1 * cm, f"Page {doc.page}")

    # Footer
    canvas.setFillColor(GREY)
    canvas.setFont("Helvetica", 7)
    footer_text = "Prepared for internal use only — not for distribution without review"
    canvas.drawCentredString(w / 2, 0.7 * cm, footer_text)

    # Bottom line
    canvas.setStrokeColor(NAVY)
    canvas.setLineWidth(0.5)
    canvas.line(1.5 * cm, 1.2 * cm, w - 1.5 * cm, 1.2 * cm)

    canvas.restoreState()


def _status_dot(status: str) -> str:
    colour_map = {"green": "#16A34A", "amber": "#D97706", "red": "#DC2626", "grey": "#9CA3AF"}
    colour = colour_map.get(status, "#9CA3AF")
    return f'<font color="{colour}">●</font>'


def _metric_table(metrics: dict, category: str, styles, period_labels: list) -> Table:
    """Build a table of metrics for a given category."""
    label_cur = period_labels[0] if period_labels else "Current"
    label_pri = period_labels[1] if len(period_labels) > 1 else "Prior"

    header = ["Metric", "Status", label_cur, label_pri, "Trend"]
    rows = [header]

    for key, m in metrics.items():
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
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 9),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, LIGHT_GREY]),
        ("FONTSIZE", (0, 1), (-1, -1), 9),
        ("GRID", (0, 0), (-1, -1), 0.3, GREY),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("ALIGN", (0, 0), (0, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return t


def generate_pdf_report(
    analysis_result,
    session_info: dict,
    commentary: str,
    firm_name: str = "Your Firm Name",
    chart_images: dict = None,
) -> bytes:
    """
    Generate a PDF report and return as bytes.

    analysis_result: AnalysisResult object from calculator
    session_info: dict with client_name, abn, industry, financial_year_end, currency
    commentary: AI-generated (possibly edited) commentary text
    firm_name: Accounting firm name for header
    chart_images: dict of section -> PIL Image or bytes
    """
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        topMargin=2.5 * cm,
        bottomMargin=2 * cm,
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
    )
    styles = _get_styles()
    story = []

    client_name = session_info.get("client_name", "Client")
    abn = session_info.get("abn", "")
    industry = session_info.get("industry", "")
    fy_end = session_info.get("financial_year_end", "")
    currency = session_info.get("currency", "AUD")
    analysis_date = date.today().strftime("%d %B %Y")
    period_labels = analysis_result.period_labels

    # ── Cover section ────────────────────────────────────────────────────────
    story.append(Spacer(1, 0.5 * cm))
    story.append(Paragraph(firm_name, styles["FirmHeader"]))
    story.append(Paragraph("Financial Analysis Report", styles["SubHeader"]))
    story.append(HRFlowable(width="100%", thickness=2, color=NAVY))
    story.append(Spacer(1, 0.3 * cm))

    # Client info table
    info_data = [
        ["Client:", client_name, "Date of Analysis:", analysis_date],
        ["ABN:", abn or "Not provided", "Financial Year:", fy_end],
        ["Industry:", industry, "Currency:", currency],
    ]
    info_table = Table(info_data, colWidths=[3 * cm, 8 * cm, 3.5 * cm, 4 * cm])
    info_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTNAME", (2, 0), (2, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("TEXTCOLOR", (0, 0), (0, -1), NAVY),
        ("TEXTCOLOR", (2, 0), (2, -1), NAVY),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 0.5 * cm))

    # ── Red Flags ─────────────────────────────────────────────────────────
    if analysis_result.red_flags:
        story.append(Paragraph("Red Flags", styles["SectionHeading"]))
        for flag in analysis_result.red_flags:
            story.append(Paragraph(flag, styles["RedFlag"]))
        story.append(Spacer(1, 0.3 * cm))

    # ── Metrics sections ─────────────────────────────────────────────────
    categories = [
        ("liquidity", "Liquidity Metrics"),
        ("profitability", "Profitability Metrics"),
        ("efficiency", "Operational Efficiency"),
        ("leverage", "Leverage & Solvency"),
        ("growth", "Growth Metrics"),
    ]

    for cat_key, cat_label in categories:
        cat_metrics = {k: v for k, v in analysis_result.metrics.items() if v.category == cat_key}
        if not cat_metrics:
            continue
        story.append(Paragraph(cat_label, styles["SectionHeading"]))
        tbl = _metric_table(analysis_result.metrics, cat_key, styles, period_labels)
        if tbl:
            story.append(tbl)

        # Notes for metrics in this category
        for k, m in cat_metrics.items():
            if m.notes:
                story.append(Paragraph(m.notes, styles["RedFlag"]))
        story.append(Spacer(1, 0.2 * cm))

    # ── ATO Benchmarks ────────────────────────────────────────────────────
    if analysis_result.benchmark_comparisons:
        story.append(PageBreak())
        story.append(Paragraph("ATO Small Business Benchmark Comparison", styles["SectionHeading"]))
        story.append(Paragraph(
            f"Industry: {industry}. Benchmarks expressed as % of turnover. "
            "Note: ATO benchmarks are updated annually and may lag by one year.",
            styles["CommentaryBody"]
        ))

        bm_header = ["Expense Category", "Client %", "ATO Low", "ATO High", "Status"]
        bm_rows = [bm_header]
        for key, comp in analysis_result.benchmark_comparisons.items():
            actual = comp.get("actual_pct")
            low = comp.get("benchmark_low")
            high = comp.get("benchmark_high")
            in_range = (actual is not None and low is not None and high is not None
                        and low <= actual <= high)
            status_dot = _status_dot("green" if in_range else ("amber" if actual else "grey"))
            bm_rows.append([
                comp["label"],
                f"{actual:.1f}%" if actual is not None else "N/A",
                f"{low:.0f}%" if low is not None else "N/A",
                f"{high:.0f}%" if high is not None else "N/A",
                Paragraph(status_dot, styles["MetricLabel"]),
            ])

        bm_table = Table(bm_rows, colWidths=[6 * cm, 3 * cm, 3 * cm, 3 * cm, 2 * cm])
        bm_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), NAVY),
            ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, LIGHT_GREY]),
            ("GRID", (0, 0), (-1, -1), 0.3, GREY),
            ("ALIGN", (1, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(bm_table)
        story.append(Spacer(1, 0.3 * cm))

    # ── Commentary ────────────────────────────────────────────────────────
    if commentary:
        story.append(PageBreak())
        story.append(Paragraph("Financial Commentary", styles["SectionHeading"]))
        story.append(Paragraph(
            "The following commentary has been generated to assist in client meeting preparation. "
            "Review and edit as required before use.",
            styles["CommentaryBody"]
        ))
        story.append(Spacer(1, 0.2 * cm))

        # Split commentary by lines and render
        for line in commentary.split("\n"):
            line = line.strip()
            if not line:
                story.append(Spacer(1, 0.1 * cm))
            elif line.startswith("### "):
                story.append(Paragraph(line[4:], styles["SectionHeading"]))
            elif line.startswith("## "):
                story.append(Paragraph(line[3:], styles["SectionHeading"]))
            elif line.startswith("- ") or line.startswith("* "):
                story.append(Paragraph(f"• {line[2:]}", styles["CommentaryBody"]))
            else:
                story.append(Paragraph(line, styles["CommentaryBody"]))

    # Build PDF
    doc.build(story, onFirstPage=_header_footer, onLaterPages=_header_footer)
    buffer.seek(0)
    return buffer.getvalue()
