"""
Excel report generation using xlsxwriter.
Produces a multi-tab Excel workbook with colour-coded metrics.
"""

import io
import logging
from datetime import date

import xlsxwriter

logger = logging.getLogger(__name__)

# Traffic light colours (RGB tuples for xlsxwriter hex)
STATUS_FILL = {
    "green": "#DCFCE7",
    "amber": "#FEF3C7",
    "red": "#FEE2E2",
    "grey": "#F3F4F6",
}
STATUS_FONT = {
    "green": "#16A34A",
    "amber": "#D97706",
    "red": "#DC2626",
    "grey": "#6B7280",
}
NAVY = "#1B2A4A"
WHITE = "#FFFFFF"
LIGHT_GREY = "#F3F4F6"


def _fmt_currency(v) -> str:
    if v is None:
        return "N/A"
    return f"${v:,.0f}"


def _fmt_pct(v) -> str:
    if v is None:
        return "N/A"
    return f"{v:.1f}%"


def _fmt_ratio(v) -> str:
    if v is None:
        return "N/A"
    return f"{v:.2f}x"


def _fmt_days(v) -> str:
    if v is None:
        return "N/A"
    return f"{v:.0f} days"


def _format_metric_value(m) -> str:
    fmt = m.format_type
    val = m.current
    if fmt == "percentage":
        return _fmt_pct(val)
    elif fmt == "ratio":
        return _fmt_ratio(val)
    elif fmt == "currency":
        return _fmt_currency(val)
    elif fmt == "days":
        return _fmt_days(val)
    return str(val) if val is not None else "N/A"


def generate_excel_report(
    analysis_result,
    session_info: dict,
    commentary: str,
) -> bytes:
    """
    Generate an Excel workbook and return as bytes.
    Tabs: Overview | Raw Data | Metrics | Benchmarks | Commentary
    """
    buffer = io.BytesIO()
    wb = xlsxwriter.Workbook(buffer, {"in_memory": True})

    client_name = session_info.get("client_name", "Client")
    industry = session_info.get("industry", "")
    fy_end = session_info.get("financial_year_end", "")
    period_labels = analysis_result.period_labels
    label_cur = period_labels[0] if period_labels else "Current"
    label_pri = period_labels[1] if len(period_labels) > 1 else "Prior"

    # ── Shared formats ─────────────────────────────────────────────────────
    hdr_fmt = wb.add_format({
        "bold": True, "font_color": WHITE, "bg_color": NAVY,
        "border": 1, "align": "center", "valign": "vcenter",
        "text_wrap": True,
    })
    title_fmt = wb.add_format({
        "bold": True, "font_size": 14, "font_color": NAVY,
    })
    sub_fmt = wb.add_format({
        "font_size": 10, "font_color": "#6B7280",
    })
    section_fmt = wb.add_format({
        "bold": True, "font_size": 11, "font_color": NAVY,
        "bg_color": LIGHT_GREY, "border": 1,
    })
    normal_fmt = wb.add_format({"border": 1, "valign": "vcenter"})
    currency_fmt = wb.add_format({
        "num_format": '$#,##0', "border": 1, "valign": "vcenter",
    })
    pct_fmt = wb.add_format({
        "num_format": '0.0%', "border": 1, "valign": "vcenter",
    })

    def status_fmt(status):
        return wb.add_format({
            "bold": True,
            "font_color": STATUS_FONT.get(status, "#6B7280"),
            "bg_color": STATUS_FILL.get(status, "#F3F4F6"),
            "border": 1, "align": "center", "valign": "vcenter",
        })

    # ── Sheet 1: Overview ─────────────────────────────────────────────────
    ws = wb.add_worksheet("Overview")
    ws.set_column("A:A", 28)
    ws.set_column("B:D", 18)
    ws.set_column("E:E", 14)

    row = 0
    ws.write(row, 0, f"FinSight Financial Analysis — {client_name}", title_fmt)
    row += 1
    ws.write(row, 0, f"Industry: {industry} | Period: {fy_end} | Generated: {date.today()}", sub_fmt)
    row += 2

    ws.write(row, 0, "Key Metrics Summary", section_fmt)
    ws.write(row, 1, label_cur, section_fmt)
    ws.write(row, 2, label_pri, section_fmt)
    ws.write(row, 3, "Trend", section_fmt)
    ws.write(row, 4, "Status", section_fmt)
    row += 1

    categories = [
        ("liquidity", "LIQUIDITY"),
        ("profitability", "PROFITABILITY"),
        ("efficiency", "EFFICIENCY"),
        ("leverage", "LEVERAGE & SOLVENCY"),
        ("growth", "GROWTH"),
    ]

    for cat_key, cat_label in categories:
        ws.merge_range(row, 0, row, 4, cat_label, section_fmt)
        row += 1
        for key, m in analysis_result.metrics.items():
            if m.category != cat_key:
                continue
            ws.write(row, 0, m.label, normal_fmt)
            ws.write(row, 1, m.current_fmt, normal_fmt)
            ws.write(row, 2, m.prior_fmt, normal_fmt)
            ws.write(row, 3, m.trend, normal_fmt)
            sf = status_fmt(m.status)
            status_text = {"green": "✅ Good", "amber": "🟡 Review", "red": "🔴 Concern", "grey": "N/A"}.get(m.status, "N/A")
            ws.write(row, 4, status_text, sf)
            row += 1

    # Red Flags
    if analysis_result.red_flags:
        row += 1
        ws.write(row, 0, "RED FLAGS", wb.add_format({"bold": True, "font_color": "#DC2626", "font_size": 11}))
        row += 1
        flag_fmt = wb.add_format({"font_color": "#DC2626", "text_wrap": True})
        ws.set_row(row, 14)
        for flag in analysis_result.red_flags:
            ws.merge_range(row, 0, row, 4, flag, flag_fmt)
            row += 1

    # ── Sheet 2: Raw Data ──────────────────────────────────────────────────
    ws2 = wb.add_worksheet("Raw Data")
    ws2.set_column("A:A", 30)
    ws2.set_column("B:D", 20)

    ws2.write(0, 0, "Account", hdr_fmt)
    ws2.write(0, 1, label_cur, hdr_fmt)
    ws2.write(0, 2, label_pri, hdr_fmt)

    raw_data = analysis_result.raw_data
    cur = raw_data.get("current", {}) or {}
    prior = raw_data.get("prior", {}) or {}

    pl_fields = [
        ("revenue", "Revenue"),
        ("cogs", "Cost of Goods Sold"),
        ("gross_profit", "Gross Profit"),
        ("operating_expenses", "Operating Expenses"),
        ("ebit", "EBIT"),
        ("depreciation", "Depreciation & Amortisation"),
        ("ebitda", "EBITDA"),
        ("interest_expense", "Interest Expense"),
        ("tax_expense", "Tax Expense"),
        ("net_profit", "Net Profit"),
    ]
    bs_fields = [
        ("cash", "Cash & Bank"),
        ("accounts_receivable", "Accounts Receivable"),
        ("inventory", "Inventory"),
        ("current_assets", "Total Current Assets"),
        ("non_current_assets", "Total Non-Current Assets"),
        ("total_assets", "Total Assets"),
        ("accounts_payable", "Accounts Payable"),
        ("current_liabilities", "Total Current Liabilities"),
        ("non_current_liabilities", "Total Non-Current Liabilities"),
        ("total_liabilities", "Total Liabilities"),
        ("equity", "Total Equity"),
        ("total_debt", "Total Debt"),
    ]
    cf_fields = [
        ("operating_cash_flow", "Operating Cash Flow"),
        ("investing_cash_flow", "Investing Cash Flow"),
        ("financing_cash_flow", "Financing Cash Flow"),
    ]

    data_row = 1

    def write_section(label, fields):
        nonlocal data_row
        ws2.merge_range(data_row, 0, data_row, 2, label, section_fmt)
        data_row += 1
        for key, field_label in fields:
            c_val = cur.get(key)
            p_val = prior.get(key)
            ws2.write(data_row, 0, field_label, normal_fmt)
            if c_val is not None:
                ws2.write_number(data_row, 1, c_val, currency_fmt)
            else:
                ws2.write(data_row, 1, "N/A", normal_fmt)
            if p_val is not None:
                ws2.write_number(data_row, 2, p_val, currency_fmt)
            else:
                ws2.write(data_row, 2, "N/A", normal_fmt)
            data_row += 1

    write_section("PROFIT & LOSS", pl_fields)
    write_section("BALANCE SHEET", bs_fields)
    write_section("CASH FLOW", cf_fields)

    # ── Sheet 3: Metrics Detail ────────────────────────────────────────────
    ws3 = wb.add_worksheet("Metrics")
    ws3.set_column("A:A", 30)
    ws3.set_column("B:G", 16)

    ws3.write(0, 0, "Metric", hdr_fmt)
    ws3.write(0, 1, label_cur, hdr_fmt)
    ws3.write(0, 2, label_pri, hdr_fmt)
    ws3.write(0, 3, "Trend", hdr_fmt)
    ws3.write(0, 4, "Status", hdr_fmt)
    ws3.write(0, 5, "Benchmark Low", hdr_fmt)
    ws3.write(0, 6, "Benchmark High", hdr_fmt)

    m_row = 1
    for cat_key, cat_label in categories:
        ws3.merge_range(m_row, 0, m_row, 6, cat_label, section_fmt)
        m_row += 1
        for key, m in analysis_result.metrics.items():
            if m.category != cat_key:
                continue
            ws3.write(m_row, 0, m.label, normal_fmt)
            ws3.write(m_row, 1, m.current_fmt, status_fmt(m.status))
            ws3.write(m_row, 2, m.prior_fmt, normal_fmt)
            ws3.write(m_row, 3, m.trend, normal_fmt)
            status_text = {"green": "Good", "amber": "Review", "red": "Concern", "grey": "N/A"}.get(m.status, "N/A")
            ws3.write(m_row, 4, status_text, status_fmt(m.status))
            ws3.write(m_row, 5, f"{m.benchmark_low:.1f}%" if m.benchmark_low is not None else "N/A", normal_fmt)
            ws3.write(m_row, 6, f"{m.benchmark_high:.1f}%" if m.benchmark_high is not None else "N/A", normal_fmt)
            m_row += 1

    # ── Sheet 4: ATO Benchmarks ───────────────────────────────────────────
    ws4 = wb.add_worksheet("Benchmarks")
    ws4.set_column("A:A", 30)
    ws4.set_column("B:E", 18)

    ws4.write(0, 0, "ATO Benchmark Comparison", title_fmt)
    ws4.write(1, 0, f"Industry: {industry}", sub_fmt)

    ws4.write(3, 0, "Expense Category", hdr_fmt)
    ws4.write(3, 1, "Client % of Turnover", hdr_fmt)
    ws4.write(3, 2, "ATO Benchmark Low", hdr_fmt)
    ws4.write(3, 3, "ATO Benchmark High", hdr_fmt)
    ws4.write(3, 4, "Status", hdr_fmt)

    b_row = 4
    for key, comp in analysis_result.benchmark_comparisons.items():
        actual = comp.get("actual_pct")
        low = comp.get("benchmark_low")
        high = comp.get("benchmark_high")
        in_range = actual is not None and low is not None and high is not None and low <= actual <= high
        bm_status = "green" if in_range else ("amber" if actual else "grey")
        ws4.write(b_row, 0, comp["label"], normal_fmt)
        ws4.write(b_row, 1, f"{actual:.1f}%" if actual is not None else "N/A", status_fmt(bm_status))
        ws4.write(b_row, 2, f"{low:.0f}%" if low is not None else "N/A", normal_fmt)
        ws4.write(b_row, 3, f"{high:.0f}%" if high is not None else "N/A", normal_fmt)
        status_text = "Within range" if in_range else "Outside range"
        ws4.write(b_row, 4, status_text, status_fmt(bm_status))
        b_row += 1

    ws4.write(b_row + 1, 0,
              "Note: ATO benchmarks are updated annually and may lag by one year.",
              sub_fmt)

    # ── Sheet 5: Commentary ────────────────────────────────────────────────
    ws5 = wb.add_worksheet("Commentary")
    ws5.set_column("A:A", 100)

    ws5.write(0, 0, "Financial Commentary", title_fmt)
    ws5.write(1, 0, f"Client: {client_name} | {fy_end} | Generated: {date.today()}", sub_fmt)

    c_row = 3
    commentary_fmt = wb.add_format({"text_wrap": True, "valign": "top"})
    heading_fmt = wb.add_format({"bold": True, "font_color": NAVY, "font_size": 12})

    for line in (commentary or "").split("\n"):
        line = line.strip()
        if not line:
            c_row += 1
        elif line.startswith("### ") or line.startswith("## "):
            ws5.write(c_row, 0, line.lstrip("# "), heading_fmt)
            c_row += 1
        else:
            ws5.set_row(c_row, None, None, {"level": 0})
            ws5.write(c_row, 0, line, commentary_fmt)
            c_row += 1

    wb.close()
    buffer.seek(0)
    return buffer.getvalue()
