"""
Word document report generation using python-docx.
Produces an editable .docx report for personalisation before client delivery.
"""

import io
import logging
from datetime import date

from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

logger = logging.getLogger(__name__)

NAVY_RGB = RGBColor(0x1B, 0x2A, 0x4A)
GREY_RGB = RGBColor(0x6B, 0x72, 0x80)
GREEN_RGB = RGBColor(0x16, 0xA3, 0x4A)
AMBER_RGB = RGBColor(0xD9, 0x77, 0x06)
RED_RGB = RGBColor(0xDC, 0x26, 0x26)

STATUS_RGB = {
    "green": GREEN_RGB,
    "amber": AMBER_RGB,
    "red": RED_RGB,
    "grey": GREY_RGB,
}
STATUS_TEXT = {
    "green": "Good",
    "amber": "Review",
    "red": "Concern",
    "grey": "N/A",
}


def _set_cell_bg(cell, hex_color: str):
    """Set table cell background colour."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color.lstrip("#"))
    tcPr.append(shd)


def _add_heading(doc: Document, text: str, level: int = 1):
    p = doc.add_heading(text, level=level)
    run = p.runs[0] if p.runs else p.add_run(text)
    run.font.color.rgb = NAVY_RGB
    return p


def _add_metric_table(doc: Document, metrics: dict, category: str, period_labels: list):
    """Add a table of metrics for a given category."""
    label_cur = period_labels[0] if period_labels else "Current"
    label_pri = period_labels[1] if len(period_labels) > 1 else "Prior"

    cat_metrics = [(k, v) for k, v in metrics.items() if v.category == category]
    if not cat_metrics:
        return

    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.LEFT

    # Column widths
    col_widths = [Cm(7), Cm(3), Cm(3), Cm(3), Cm(2.5)]
    for i, width in enumerate(col_widths):
        for cell in table.columns[i].cells:
            cell.width = width

    # Header row
    hdr = table.rows[0].cells
    headers = ["Metric", label_cur, label_pri, "Trend", "Status"]
    for i, (cell, text) in enumerate(zip(hdr, headers)):
        cell.text = text
        _set_cell_bg(cell, "1B2A4A")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.runs[0] if p.runs else p.add_run(text)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        run.font.size = Pt(9)

    # Data rows
    status_fill = {
        "green": "DCFCE7",
        "amber": "FEF3C7",
        "red": "FEE2E2",
        "grey": "F3F4F6",
    }

    for key, m in cat_metrics:
        row = table.add_row().cells
        row[0].text = m.label
        row[1].text = m.current_fmt
        row[2].text = m.prior_fmt
        row[3].text = m.trend
        row[4].text = STATUS_TEXT.get(m.status, "N/A")

        # Colour the status cell
        _set_cell_bg(row[4], status_fill.get(m.status, "F3F4F6"))

        for cell in row:
            for para in cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in para.runs:
                    run.font.size = Pt(9)

        row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT


def generate_word_report(
    analysis_result,
    session_info: dict,
    commentary: str,
    firm_name: str = "Your Firm Name",
) -> bytes:
    """
    Generate a Word document and return as bytes.
    """
    doc = Document()

    # Page margins
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)

    client_name = session_info.get("client_name", "Client")
    abn = session_info.get("abn", "")
    industry = session_info.get("industry", "")
    fy_end = session_info.get("financial_year_end", "")
    currency = session_info.get("currency", "AUD")
    period_labels = analysis_result.period_labels
    label_cur = period_labels[0] if period_labels else "Current"
    label_pri = period_labels[1] if len(period_labels) > 1 else "Prior"

    # ── Cover / Title ────────────────────────────────────────────────────
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = title.add_run(firm_name)
    run.font.size = Pt(20)
    run.font.bold = True
    run.font.color.rgb = NAVY_RGB

    subtitle = doc.add_paragraph()
    run = subtitle.add_run("Financial Analysis Report")
    run.font.size = Pt(13)
    run.font.color.rgb = GREY_RGB

    doc.add_paragraph()  # Spacer

    # Client info table
    info_table = doc.add_table(rows=3, cols=4)
    info_table.style = "Table Grid"
    info_data = [
        ["Client", client_name, "Date of Analysis", str(date.today())],
        ["ABN", abn or "Not provided", "Financial Year", fy_end],
        ["Industry", industry, "Currency", currency],
    ]
    for r_idx, row_data in enumerate(info_data):
        cells = info_table.rows[r_idx].cells
        for c_idx, text in enumerate(row_data):
            cells[c_idx].text = text
            if c_idx % 2 == 0:  # Label cells
                for run in cells[c_idx].paragraphs[0].runs:
                    run.font.bold = True
                    run.font.color.rgb = NAVY_RGB

    doc.add_paragraph()

    # ── Red Flags ────────────────────────────────────────────────────────
    if analysis_result.red_flags:
        _add_heading(doc, "Red Flags", level=2)
        for flag in analysis_result.red_flags:
            p = doc.add_paragraph()
            run = p.add_run(flag)
            run.font.color.rgb = RED_RGB
            run.font.size = Pt(10)

        doc.add_paragraph()

    # ── Metrics Sections ─────────────────────────────────────────────────
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
        _add_heading(doc, cat_label, level=2)
        _add_metric_table(doc, analysis_result.metrics, cat_key, period_labels)

        # Notes
        for key, m in cat_metrics.items():
            if m.notes:
                p = doc.add_paragraph()
                run = p.add_run(m.notes)
                run.font.color.rgb = RED_RGB
                run.font.size = Pt(9)

        doc.add_paragraph()

    # ── ATO Benchmarks ────────────────────────────────────────────────────
    if analysis_result.benchmark_comparisons:
        doc.add_page_break()
        _add_heading(doc, "ATO Small Business Benchmark Comparison", level=2)

        p = doc.add_paragraph(
            f"Industry: {industry}. Benchmarks expressed as % of turnover. "
            "Note: ATO benchmarks are updated annually and may lag by one year."
        )
        p.runs[0].font.size = Pt(9)

        bm_table = doc.add_table(rows=1, cols=5)
        bm_table.style = "Table Grid"

        hdr = bm_table.rows[0].cells
        bm_headers = ["Expense Category", "Client %", "ATO Low", "ATO High", "Status"]
        for i, (cell, text) in enumerate(zip(hdr, bm_headers)):
            cell.text = text
            _set_cell_bg(cell, "1B2A4A")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.runs[0] if p.runs else p.add_run(text)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            run.font.size = Pt(9)

        for key, comp in analysis_result.benchmark_comparisons.items():
            actual = comp.get("actual_pct")
            low = comp.get("benchmark_low")
            high = comp.get("benchmark_high")
            in_range = actual is not None and low is not None and high is not None and low <= actual <= high
            bm_status = "green" if in_range else "amber"

            row = bm_table.add_row().cells
            row[0].text = comp["label"]
            row[1].text = f"{actual:.1f}%" if actual is not None else "N/A"
            row[2].text = f"{low:.0f}%" if low is not None else "N/A"
            row[3].text = f"{high:.0f}%" if high is not None else "N/A"
            row[4].text = "Within range" if in_range else "Outside range"
            _set_cell_bg(row[4], "DCFCE7" if in_range else "FEF3C7")

            for cell in row:
                for para in cell.paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in para.runs:
                        run.font.size = Pt(9)
            row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

        doc.add_paragraph()

    # ── Commentary ────────────────────────────────────────────────────────
    if commentary:
        doc.add_page_break()
        _add_heading(doc, "Financial Commentary", level=1)

        note_p = doc.add_paragraph(
            "The following commentary has been generated to assist in client meeting preparation. "
            "Review and edit as required before use."
        )
        note_p.runs[0].font.size = Pt(9)
        note_p.runs[0].font.color.rgb = GREY_RGB
        doc.add_paragraph()

        for line in commentary.split("\n"):
            line_stripped = line.strip()
            if not line_stripped:
                doc.add_paragraph()
            elif line_stripped.startswith("### "):
                _add_heading(doc, line_stripped[4:], level=3)
            elif line_stripped.startswith("## "):
                _add_heading(doc, line_stripped[3:], level=2)
            elif line_stripped.startswith("- ") or line_stripped.startswith("* "):
                p = doc.add_paragraph(line_stripped[2:], style="List Bullet")
                for run in p.runs:
                    run.font.size = Pt(10)
            else:
                p = doc.add_paragraph(line_stripped)
                for run in p.runs:
                    run.font.size = Pt(10)

    # ── Footer note ──────────────────────────────────────────────────────
    doc.add_paragraph()
    footer_p = doc.add_paragraph(
        "Prepared for internal use only — not for distribution without review"
    )
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in footer_p.runs:
        run.font.size = Pt(8)
        run.font.color.rgb = GREY_RGB
        run.font.italic = True

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()
