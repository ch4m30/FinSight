"""
Excel report generation using xlsxwriter.
Produces a 7-tab professional workbook with internal use banners,
conditional formatting, and structured financial data.

Sheets:
  1. Cover               — client details & internal use notice
  2. Executive Summary   — health snapshot with conditional formatting
  3. Detailed Metrics    — all metrics across all categories
  4. P&L Data            — profit & loss line items
  5. Balance Sheet Data  — balance sheet line items
  6. Charts              — revenue/GP/NP comparison chart
  7. Commentary          — AI-generated commentary (if available)
"""

import io
import logging
from datetime import date

import xlsxwriter

from utils.formatters import format_currency, format_percent, format_ratio, format_days

logger = logging.getLogger(__name__)

# Colour palette
NAVY       = "#1B2A4A"
NAVY_MID   = "#2C3E6B"
RED_ALERT  = "#DC2626"
WHITE      = "#FFFFFF"
LIGHT_GREY = "#F3F4F6"
MID_GREY   = "#E5E7EB"
DARK_GREY  = "#6B7280"

STATUS_FILL = {
    "green": "#DCFCE7",
    "amber": "#FEF3C7",
    "red":   "#FEE2E2",
    "grey":  "#F3F4F6",
}
STATUS_FONT = {
    "green": "#16A34A",
    "amber": "#D97706",
    "red":   "#DC2626",
    "grey":  "#6B7280",
}
STATUS_TEXT = {
    "green": "Good",
    "amber": "Review",
    "red":   "Concern",
    "grey":  "N/A",
}

# Fields for data sheets
PL_FIELDS = [
    ("revenue",             "Revenue"),
    ("cogs",                "Cost of Goods Sold"),
    ("gross_profit",        "Gross Profit"),
    ("operating_expenses",  "Operating Expenses"),
    ("ebit",                "EBIT"),
    ("depreciation",        "Depreciation & Amortisation"),
    ("ebitda",              "EBITDA"),
    ("interest_expense",    "Interest Expense"),
    ("tax_expense",         "Tax Expense"),
    ("net_profit",          "Net Profit"),
]
BS_FIELDS = [
    ("cash",                    "Cash & Bank"),
    ("accounts_receivable",     "Accounts Receivable"),
    ("inventory",               "Inventory"),
    ("current_assets",          "Total Current Assets"),
    ("non_current_assets",      "Total Non-Current Assets"),
    ("total_assets",            "Total Assets"),
    ("accounts_payable",        "Accounts Payable"),
    ("current_liabilities",     "Total Current Liabilities"),
    ("non_current_liabilities", "Total Non-Current Liabilities"),
    ("total_liabilities",       "Total Liabilities"),
    ("equity",                  "Total Equity"),
    ("total_debt",              "Total Debt"),
]
CF_FIELDS = [
    ("operating_cash_flow",  "Operating Cash Flow"),
    ("investing_cash_flow",  "Investing Cash Flow"),
    ("financing_cash_flow",  "Financing Cash Flow"),
]

CATEGORIES = [
    ("profitability", "PROFITABILITY"),
    ("liquidity",     "LIQUIDITY"),
    ("efficiency",    "OPERATIONAL EFFICIENCY"),
    ("leverage",      "LEVERAGE & SOLVENCY"),
    ("growth",        "GROWTH"),
]


def generate_excel_report(
    analysis_result,
    session_info: dict,
    commentary: str,
) -> bytes:
    """
    Generate a polished 7-tab Excel workbook and return as bytes.
    """
    buffer = io.BytesIO()
    wb = xlsxwriter.Workbook(buffer, {"in_memory": True})

    # ── Metadata ──────────────────────────────────────────────────────────────
    client_name = session_info.get("client_name", "Client")
    abn         = session_info.get("abn", "")
    industry    = session_info.get("industry", "")
    fy_end      = session_info.get("financial_year_end", "")
    currency    = session_info.get("currency", "AUD")
    today_str   = date.today().strftime("%d %B %Y")
    period_labels = analysis_result.period_labels
    label_cur   = period_labels[0] if period_labels else "Current"
    label_pri   = period_labels[1] if len(period_labels) > 1 else "Prior"

    # ── Shared formats ────────────────────────────────────────────────────────
    def _f(props):
        return wb.add_format(props)

    # Banner (internal use)
    banner_fmt = _f({"bold": True, "font_size": 10, "font_color": WHITE,
                     "bg_color": RED_ALERT, "align": "center", "valign": "vcenter",
                     "border": 0})
    # Titles
    title_lg_fmt = _f({"bold": True, "font_size": 18, "font_color": NAVY})
    title_sm_fmt = _f({"bold": True, "font_size": 13, "font_color": NAVY})
    sub_fmt      = _f({"font_size": 10, "font_color": DARK_GREY})
    # Table headers
    hdr_fmt      = _f({"bold": True, "font_color": WHITE, "bg_color": NAVY,
                       "border": 1, "align": "center", "valign": "vcenter",
                       "text_wrap": True})
    # Section separators
    sect_fmt     = _f({"bold": True, "font_size": 10, "font_color": WHITE,
                       "bg_color": NAVY_MID, "border": 1})
    # Data cells
    normal_fmt   = _f({"border": 1, "valign": "vcenter"})
    currency_fmt = _f({"num_format": "$#,##0", "border": 1, "valign": "vcenter"})
    pct_fmt      = _f({"num_format": "0.0%",   "border": 1, "valign": "vcenter"})
    label_fmt    = _f({"bold": True, "font_color": NAVY, "border": 1, "bg_color": LIGHT_GREY})
    value_fmt    = _f({"border": 1, "bg_color": WHITE})
    # Commentary
    para_fmt     = _f({"text_wrap": True, "valign": "top", "font_size": 10})
    head_fmt     = _f({"bold": True, "font_color": NAVY, "font_size": 12})

    def status_fmt(status):
        return wb.add_format({
            "bold": True,
            "font_color": STATUS_FONT.get(status, DARK_GREY),
            "bg_color":   STATUS_FILL.get(status, LIGHT_GREY),
            "border": 1, "align": "center", "valign": "vcenter",
        })

    def _banner(ws, row, ncols, text="INTERNAL USE ONLY — NOT FOR DISTRIBUTION"):
        """Write a full-width red internal use banner row."""
        if ncols > 1:
            ws.merge_range(row, 0, row, ncols - 1, text, banner_fmt)
        else:
            ws.write(row, 0, text, banner_fmt)
        ws.set_row(row, 18)

    # ── SHEET 1: Cover ────────────────────────────────────────────────────────
    ws1 = wb.add_worksheet("Cover")
    ws1.set_column("A:A", 28)
    ws1.set_column("B:B", 40)
    ws1.set_tab_color(NAVY)

    _banner(ws1, 0, 2)

    ws1.merge_range(2, 0, 2, 1, "FinSight Financial Analysis", title_lg_fmt)
    ws1.set_row(2, 32)
    ws1.merge_range(3, 0, 3, 1, "Internal Working Paper", sub_fmt)

    ws1.set_row(5, 14)  # spacer

    info_rows = [
        ("Client / Business Name", client_name),
        ("ABN",                    abn or "Not provided"),
        ("Industry",               industry),
        ("Financial Year End",     fy_end),
        ("Reporting Currency",     currency),
        ("Date of Analysis",       today_str),
        ("Current Period",         label_cur),
        ("Prior Period",           label_pri),
        ("Prepared by",            "FinSight — Internal Use Only"),
    ]
    for i, (lbl, val) in enumerate(info_rows, start=6):
        ws1.write(i, 0, lbl, label_fmt)
        ws1.write(i, 1, val, value_fmt)
        ws1.set_row(i, 20)

    last_info_row = 6 + len(info_rows) + 1
    ws1.merge_range(
        last_info_row, 0, last_info_row, 1,
        "This document is prepared for internal accountant use only. "
        "Do not distribute without appropriate review and sign-off.",
        _f({"text_wrap": True, "font_color": DARK_GREY, "font_size": 9,
            "border": 1, "bg_color": LIGHT_GREY, "valign": "top"}),
    )
    ws1.set_row(last_info_row, 40)

    # ── SHEET 2: Executive Summary ────────────────────────────────────────────
    ws2 = wb.add_worksheet("Executive Summary")
    ws2.set_column("A:A", 28)
    ws2.set_column("B:E", 18)
    ws2.set_tab_color(NAVY)

    _banner(ws2, 0, 5)
    ws2.write(2, 0, f"Executive Summary — {client_name}", title_sm_fmt)
    ws2.write(3, 0, f"Period: {fy_end}  |  Industry: {industry}  |  Generated: {today_str}", sub_fmt)
    ws2.set_row(5, 8)  # spacer

    # Health counts
    g_count = sum(1 for m in analysis_result.metrics.values() if m.status == "green")
    a_count = sum(1 for m in analysis_result.metrics.values() if m.status == "amber")
    r_count = sum(1 for m in analysis_result.metrics.values() if m.status == "red")

    ws2.write(6, 0, "Overall Health", hdr_fmt)
    ws2.write(6, 1, f"{g_count} Good",    _f({"bold": True, "font_color": "#16A34A", "bg_color": "#DCFCE7", "border": 1, "align": "center"}))
    ws2.write(6, 2, f"{a_count} Review",  _f({"bold": True, "font_color": "#D97706", "bg_color": "#FEF3C7", "border": 1, "align": "center"}))
    ws2.write(6, 3, f"{r_count} Concern", _f({"bold": True, "font_color": "#DC2626", "bg_color": "#FEE2E2", "border": 1, "align": "center"}))
    ws2.write(6, 4, "",                   normal_fmt)
    ws2.set_row(6, 22)
    ws2.set_row(7, 8)  # spacer

    # Snapshot metrics header
    ws2.write(8, 0, "Key Metric",     hdr_fmt)
    ws2.write(8, 1, label_cur,        hdr_fmt)
    ws2.write(8, 2, label_pri,        hdr_fmt)
    ws2.write(8, 3, "Trend",          hdr_fmt)
    ws2.write(8, 4, "Status",         hdr_fmt)
    ws2.set_row(8, 18)

    spotlight_keys = [
        "gross_profit_margin", "net_profit_margin", "ebitda_margin",
        "current_ratio", "debtor_days", "debt_to_equity",
    ]
    snap_row = 9
    for key in spotlight_keys:
        m = analysis_result.metrics.get(key)
        if m is None:
            continue
        ws2.write(snap_row, 0, m.label,                   normal_fmt)
        ws2.write(snap_row, 1, m.current_fmt,             status_fmt(m.status))
        ws2.write(snap_row, 2, m.prior_fmt,               normal_fmt)
        ws2.write(snap_row, 3, m.trend,                   normal_fmt)
        ws2.write(snap_row, 4, STATUS_TEXT[m.status],     status_fmt(m.status))
        snap_row += 1

    # Red flags
    rf_start = snap_row + 2
    ws2.write(rf_start, 0, "Red Flags", _f({"bold": True, "font_color": RED_ALERT, "font_size": 11}))
    rf_start += 1
    if analysis_result.red_flags:
        flag_fmt = _f({"font_color": RED_ALERT, "text_wrap": True, "border": 1, "bg_color": "#FEE2E2"})
        for flag in analysis_result.red_flags:
            ws2.merge_range(rf_start, 0, rf_start, 4, f"⚠ {flag}", flag_fmt)
            ws2.set_row(rf_start, 18)
            rf_start += 1
    else:
        ws2.merge_range(rf_start, 0, rf_start, 4,
                        "No red flags detected.", normal_fmt)

    # ── SHEET 3: Detailed Metrics ─────────────────────────────────────────────
    ws3 = wb.add_worksheet("Detailed Metrics")
    ws3.set_column("A:A", 30)
    ws3.set_column("B:G", 16)
    ws3.set_tab_color(NAVY_MID)

    _banner(ws3, 0, 7)
    ws3.write(2, 0, "Detailed Metric Analysis", title_sm_fmt)
    ws3.set_row(3, 8)  # spacer

    ws3.write(4, 0, "Metric",          hdr_fmt)
    ws3.write(4, 1, label_cur,         hdr_fmt)
    ws3.write(4, 2, label_pri,         hdr_fmt)
    ws3.write(4, 3, "Trend",           hdr_fmt)
    ws3.write(4, 4, "Status",          hdr_fmt)
    ws3.write(4, 5, "Benchmark Low",   hdr_fmt)
    ws3.write(4, 6, "Benchmark High",  hdr_fmt)
    ws3.set_row(4, 18)

    m_row = 5
    for cat_key, cat_label in CATEGORIES:
        ws3.merge_range(m_row, 0, m_row, 6, cat_label, sect_fmt)
        ws3.set_row(m_row, 16)
        m_row += 1
        for key, m in analysis_result.metrics.items():
            if m.category != cat_key:
                continue
            ws3.write(m_row, 0, m.label,          normal_fmt)
            ws3.write(m_row, 1, m.current_fmt,    status_fmt(m.status))
            ws3.write(m_row, 2, m.prior_fmt,      normal_fmt)
            ws3.write(m_row, 3, m.trend,          normal_fmt)
            ws3.write(m_row, 4, STATUS_TEXT[m.status], status_fmt(m.status))
            ws3.write(m_row, 5,
                      f"{m.benchmark_low:.1f}%" if m.benchmark_low is not None else "N/A",
                      normal_fmt)
            ws3.write(m_row, 6,
                      f"{m.benchmark_high:.1f}%" if m.benchmark_high is not None else "N/A",
                      normal_fmt)
            m_row += 1

    # Benchmark comparisons below metrics
    if analysis_result.benchmark_comparisons:
        m_row += 1
        ws3.merge_range(m_row, 0, m_row, 6, "ATO BENCHMARK COMPARISONS", sect_fmt)
        m_row += 1
        ws3.write(m_row, 0, "Expense Category", hdr_fmt)
        ws3.write(m_row, 1, "Client %",         hdr_fmt)
        ws3.write(m_row, 2, "ATO Low",          hdr_fmt)
        ws3.write(m_row, 3, "ATO High",         hdr_fmt)
        ws3.write(m_row, 4, "Status",           hdr_fmt)
        ws3.merge_range(m_row, 5, m_row, 6, "", hdr_fmt)
        m_row += 1
        for comp in analysis_result.benchmark_comparisons.values():
            actual   = comp.get("actual_pct")
            low      = comp.get("benchmark_low")
            high     = comp.get("benchmark_high")
            in_range = actual is not None and low is not None and high is not None and low <= actual <= high
            bm_st    = "green" if in_range else ("amber" if actual is not None else "grey")
            ws3.write(m_row, 0, comp["label"],                                         normal_fmt)
            ws3.write(m_row, 1, f"{actual:.1f}%" if actual is not None else "N/A",    status_fmt(bm_st))
            ws3.write(m_row, 2, f"{low:.0f}%" if low is not None else "N/A",          normal_fmt)
            ws3.write(m_row, 3, f"{high:.0f}%" if high is not None else "N/A",        normal_fmt)
            ws3.write(m_row, 4, "Within range" if in_range else "Outside range",       status_fmt(bm_st))
            ws3.merge_range(m_row, 5, m_row, 6, "", normal_fmt)
            m_row += 1

    # ── SHEET 4: P&L Data ─────────────────────────────────────────────────────
    ws4 = wb.add_worksheet("P&L Data")
    ws4.set_column("A:A", 32)
    ws4.set_column("B:D", 20)
    ws4.set_tab_color(NAVY_MID)

    _banner(ws4, 0, 3)
    ws4.write(2, 0, "Profit & Loss Statement", title_sm_fmt)

    cur_data  = analysis_result.raw_data.get("current") or {}
    prior_data = analysis_result.raw_data.get("prior")  or {}

    ws4.write(4, 0, "Line Item",  hdr_fmt)
    ws4.write(4, 1, label_cur,    hdr_fmt)
    ws4.write(4, 2, label_pri,    hdr_fmt)
    ws4.set_row(4, 18)

    pl_row = 5
    ws4.merge_range(pl_row, 0, pl_row, 2, "PROFIT & LOSS", sect_fmt)
    pl_row += 1
    for key, field_label in PL_FIELDS:
        c_val = cur_data.get(key)
        p_val = prior_data.get(key)
        ws4.write(pl_row, 0, field_label, normal_fmt)
        if c_val is not None:
            ws4.write_number(pl_row, 1, c_val, currency_fmt)
        else:
            ws4.write(pl_row, 1, "N/A", normal_fmt)
        if p_val is not None:
            ws4.write_number(pl_row, 2, p_val, currency_fmt)
        else:
            ws4.write(pl_row, 2, "N/A", normal_fmt)
        pl_row += 1

    # Cash flow appended below P&L
    pl_row += 1
    ws4.merge_range(pl_row, 0, pl_row, 2, "CASH FLOW", sect_fmt)
    pl_row += 1
    for key, field_label in CF_FIELDS:
        c_val = cur_data.get(key)
        p_val = prior_data.get(key)
        ws4.write(pl_row, 0, field_label, normal_fmt)
        if c_val is not None:
            ws4.write_number(pl_row, 1, c_val, currency_fmt)
        else:
            ws4.write(pl_row, 1, "N/A", normal_fmt)
        if p_val is not None:
            ws4.write_number(pl_row, 2, p_val, currency_fmt)
        else:
            ws4.write(pl_row, 2, "N/A", normal_fmt)
        pl_row += 1

    # ── SHEET 5: Balance Sheet Data ───────────────────────────────────────────
    ws5 = wb.add_worksheet("Balance Sheet Data")
    ws5.set_column("A:A", 32)
    ws5.set_column("B:C", 20)
    ws5.set_tab_color(NAVY_MID)

    _banner(ws5, 0, 3)
    ws5.write(2, 0, "Balance Sheet", title_sm_fmt)

    ws5.write(4, 0, "Line Item",  hdr_fmt)
    ws5.write(4, 1, label_cur,    hdr_fmt)
    ws5.write(4, 2, label_pri,    hdr_fmt)
    ws5.set_row(4, 18)

    bs_row = 5
    ws5.merge_range(bs_row, 0, bs_row, 2, "BALANCE SHEET", sect_fmt)
    bs_row += 1
    for key, field_label in BS_FIELDS:
        c_val = cur_data.get(key)
        p_val = prior_data.get(key)
        ws5.write(bs_row, 0, field_label, normal_fmt)
        if c_val is not None:
            ws5.write_number(bs_row, 1, c_val, currency_fmt)
        else:
            ws5.write(bs_row, 1, "N/A", normal_fmt)
        if p_val is not None:
            ws5.write_number(bs_row, 2, p_val, currency_fmt)
        else:
            ws5.write(bs_row, 2, "N/A", normal_fmt)
        bs_row += 1

    # ── SHEET 6: Charts ───────────────────────────────────────────────────────
    ws6 = wb.add_worksheet("Charts")
    ws6.set_column("A:A", 30)
    ws6.set_column("B:D", 18)
    ws6.set_tab_color(NAVY_MID)

    _banner(ws6, 0, 4)
    ws6.write(2, 0, "Financial Charts", title_sm_fmt)
    ws6.write(3, 0, f"Period comparison: {label_pri} vs {label_cur}", sub_fmt)
    ws6.set_row(5, 8)

    # Data table for chart
    chart_data_row = 6
    ws6.write(chart_data_row, 0, "Item",          hdr_fmt)
    ws6.write(chart_data_row, 1, label_pri,        hdr_fmt)
    ws6.write(chart_data_row, 2, label_cur,        hdr_fmt)
    ws6.write(chart_data_row, 3, "Change %",       hdr_fmt)
    ws6.set_row(chart_data_row, 18)

    chart_items = [
        ("Revenue",       "revenue"),
        ("Gross Profit",  "gross_profit"),
        ("Net Profit",    "net_profit"),
        ("EBITDA",        "ebitda"),
        ("Operating CF",  "operating_cash_flow"),
    ]
    for i, (lbl, key) in enumerate(chart_items):
        r = chart_data_row + 1 + i
        c_val = cur_data.get(key)
        p_val = prior_data.get(key)
        ws6.write(r, 0, lbl, normal_fmt)
        if p_val is not None:
            ws6.write_number(r, 1, p_val, currency_fmt)
        else:
            ws6.write(r, 1, "N/A", normal_fmt)
        if c_val is not None:
            ws6.write_number(r, 2, c_val, currency_fmt)
        else:
            ws6.write(r, 2, "N/A", normal_fmt)
        if c_val is not None and p_val is not None and p_val != 0:
            chg = (c_val - p_val) / abs(p_val) * 100
            ws6.write(r, 3, f"{chg:+.1f}%",
                      _f({"font_color": "#16A34A" if chg >= 0 else "#DC2626",
                          "border": 1, "align": "center"}))
        else:
            ws6.write(r, 3, "N/A", normal_fmt)

    # Bar chart
    chart_end_row = chart_data_row + len(chart_items)
    chart = wb.add_chart({"type": "bar"})
    chart.add_series({
        "name":       label_pri,
        "categories": ["Charts", chart_data_row + 1, 0, chart_end_row, 0],
        "values":     ["Charts", chart_data_row + 1, 1, chart_end_row, 1],
        "fill":       {"color": "#93C5FD"},
    })
    chart.add_series({
        "name":       label_cur,
        "categories": ["Charts", chart_data_row + 1, 0, chart_end_row, 0],
        "values":     ["Charts", chart_data_row + 1, 2, chart_end_row, 2],
        "fill":       {"color": NAVY},
    })
    chart.set_title({"name": f"Revenue & Profit Comparison ({label_pri} vs {label_cur})"})
    chart.set_x_axis({"name": "Amount ($)"})
    chart.set_y_axis({"name": "Item"})
    chart.set_size({"width": 540, "height": 300})
    ws6.insert_chart(chart_data_row + len(chart_items) + 2, 0, chart)

    # ── SHEET 7: Commentary ───────────────────────────────────────────────────
    ws7 = wb.add_worksheet("Commentary")
    ws7.set_column("A:A", 110)
    ws7.set_tab_color(DARK_GREY)

    _banner(ws7, 0, 1)
    ws7.write(2, 0, "Financial Commentary", title_sm_fmt)
    ws7.write(3, 0, f"Client: {client_name}  |  Period: {fy_end}  |  Generated: {today_str}", sub_fmt)
    ws7.write(4, 0,
              "Review and edit this commentary before client delivery. "
              "Generated by local AI — not professional advice.",
              _f({"font_color": DARK_GREY, "font_size": 9, "italic": True}))
    ws7.set_row(5, 8)

    c_row = 6
    for line in (commentary or "").split("\n"):
        line = line.strip()
        if not line:
            c_row += 1
        elif line.startswith("## ") or line.startswith("### "):
            ws7.write(c_row, 0, line.lstrip("# "), head_fmt)
            ws7.set_row(c_row, 18)
            c_row += 1
        else:
            ws7.write(c_row, 0, line, para_fmt)
            # Estimate row height for wrapped text
            ws7.set_row(c_row, max(15, len(line) // 10 * 15))
            c_row += 1

    wb.close()
    buffer.seek(0)
    return buffer.getvalue()
