"""
FinSight — SME Financial Analysis Tool
Main Streamlit application.
"""

import os
import sys
import logging
from pathlib import Path
from datetime import date

import streamlit as st
import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
from dotenv import load_dotenv

# ── Path setup ─────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR))
load_dotenv(BASE_DIR / ".env")

from parser.xero_parser import (
    parse_xero_pl, parse_xero_balance_sheet, parse_xero_cashflow,
    merge_financial_data, get_demo_data
)
from parser.pdf_parser import (
    parse_pdf, get_confirmation_template, build_confirmed_data, ALL_FIELDS
)
from metrics.calculator import run_analysis, AnalysisResult, SelfCheckResult
from benchmarks.ato_fetcher import (
    get_industry_list, get_industry_benchmarks, get_benchmark_metadata
)
from commentary.claude_commentary import (
    build_commentary_prompt,
    generate_commentary_streaming,
    check_ollama_status,
)
from exports.pdf_export import generate_pdf_report
from exports.excel_export import generate_excel_report
from exports.word_export import generate_word_report

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ── Page configuration ─────────────────────────────────────────────────────
st.set_page_config(
    page_title="FinSight",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Styling ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Colour scheme: navy, white, grey */
    .stApp { background-color: #F9FAFB; }
    .main .block-container { padding-top: 1.5rem; }

    /* Metric cards */
    .metric-card {
        background: white;
        border-radius: 8px;
        padding: 12px 16px;
        border-left: 4px solid #6B7280;
        margin-bottom: 8px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }
    .metric-card.green { border-left-color: #16A34A; }
    .metric-card.amber { border-left-color: #D97706; }
    .metric-card.red { border-left-color: #DC2626; }
    .metric-label { font-size: 0.8rem; color: #6B7280; margin-bottom: 2px; }
    .metric-value { font-size: 1.2rem; font-weight: 600; color: #1B2A4A; }
    .metric-prior { font-size: 0.75rem; color: #9CA3AF; }

    /* Status dots */
    .dot-green { color: #16A34A; }
    .dot-amber { color: #D97706; }
    .dot-red { color: #DC2626; }
    .dot-grey { color: #9CA3AF; }

    /* Red flag */
    .red-flag {
        background: #FEE2E2;
        border: 1px solid #FECACA;
        border-radius: 6px;
        padding: 8px 12px;
        margin-bottom: 6px;
        color: #DC2626;
        font-size: 0.9rem;
    }

    /* Section header */
    .section-header {
        font-size: 1.1rem;
        font-weight: 600;
        color: #1B2A4A;
        border-bottom: 2px solid #1B2A4A;
        padding-bottom: 4px;
        margin: 16px 0 12px 0;
    }

    /* Sidebar — navy background */
    [data-testid="stSidebar"] { background-color: #1B2A4A; }
    [data-testid="stSidebar"] .stMarkdown,
    [data-testid="stSidebar"] .stMarkdown p,
    [data-testid="stSidebar"] .stMarkdown h1,
    [data-testid="stSidebar"] .stMarkdown h2,
    [data-testid="stSidebar"] .stMarkdown h3,
    [data-testid="stSidebar"] .stMarkdown li { color: #F3F4F6 !important; }
    [data-testid="stSidebar"] label { color: #D1D5DB !important; }
    [data-testid="stSidebar"] .stRadio label,
    [data-testid="stSidebar"] .stRadio p { color: #D1D5DB !important; }
    [data-testid="stSidebar"] input,
    [data-testid="stSidebar"] textarea { color: #1a1a1a !important; }
    [data-testid="stSidebar"] [data-baseweb="select"] span { color: #1a1a1a !important; }
    [data-testid="stSidebar"] small,
    [data-testid="stSidebar"] .stCaption { color: #9CA3AF !important; }
    [data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.2) !important; }

    /* Self-check rows */
    .check-pass { background: #F0FDF4; border-left: 3px solid #16A34A; padding: 6px 10px; margin: 3px 0; border-radius: 4px; }
    .check-warn { background: #FFFBEB; border-left: 3px solid #D97706; padding: 6px 10px; margin: 3px 0; border-radius: 4px; }
    .check-fail { background: #FEF2F2; border-left: 3px solid #DC2626; padding: 6px 10px; margin: 3px 0; border-radius: 4px; }

    /* Data quality warn banner */
    .dq-warn-banner {
        background: #FFFBEB;
        border: 1px solid #D97706;
        border-radius: 6px;
        padding: 10px 14px;
        margin-bottom: 10px;
        color: #92400E;
        font-size: 0.9rem;
    }
</style>
""", unsafe_allow_html=True)

# ── Status helpers ─────────────────────────────────────────────────────────

STATUS_ICON = {"green": "🟢", "amber": "🟡", "red": "🔴", "grey": "⚪"}
STATUS_LABEL = {"green": "Good", "amber": "Review", "red": "Concern", "grey": "N/A"}


def _status_badge(status: str) -> str:
    icons = {"green": "🟢", "amber": "🟡", "red": "🔴", "grey": "⚪"}
    return icons.get(status, "⚪")


def _metric_card(label: str, value: str, prior: str, trend: str, status: str, tooltip: str = ""):
    icon = _status_badge(status)
    st.markdown(f"""
    <div class="metric-card {status}" title="{tooltip}">
        <div class="metric-label">{icon} {label}</div>
        <div class="metric-value">{value} <span style="font-size:0.9rem;color:#9CA3AF">{trend}</span></div>
        <div class="metric-prior">Prior: {prior}</div>
    </div>
    """, unsafe_allow_html=True)


def _fmt_aud(val) -> str:
    if val is None:
        return "N/A"
    return f"${val:,.0f}"


# ── Bug 3: Self-check display helper ──────────────────────────────────────────

def _render_self_checks(self_checks: list, show_title: bool = True):
    """Render the self-check results panel."""
    if not self_checks:
        return

    if show_title:
        st.markdown('<div class="section-header">Data Integrity Self-Checks</div>', unsafe_allow_html=True)

    fails = [c for c in self_checks if c.status == "fail"]
    warns = [c for c in self_checks if c.status == "warn"]
    passes = [c for c in self_checks if c.status == "pass"]

    if fails:
        st.error(
            f"**{len(fails)} check(s) FAILED** — review before proceeding. "
            "Results may be unreliable."
        )
    elif warns:
        st.warning(
            f"**{len(warns)} data quality warning(s)** — results are available but "
            "review warnings before relying on the analysis."
        )
    else:
        st.success("All data integrity checks passed.")

    check_icon = {"pass": "✅", "warn": "⚠️", "fail": "❌"}
    check_css = {"pass": "check-pass", "warn": "check-warn", "fail": "check-fail"}

    for chk in self_checks:
        icon = check_icon.get(chk.status, "⚪")
        css = check_css.get(chk.status, "")
        with st.expander(
            f"{icon} **{chk.check_name}** — {chk.status.upper()}",
            expanded=(chk.status == "fail"),
        ):
            st.markdown(f"**Check:** {chk.description}")
            st.markdown(f"**Result:** {chk.detail}")
            st.info(f"**What does this mean?** {chk.what_it_means}")


# ── Bug 3: Blocking FAIL override state ───────────────────────────────────────

def _render_fail_override_checkbox(self_checks: list) -> bool:
    """
    If any checks are FAIL, show a confirmation checkbox.
    Returns True if user has confirmed override or no FAILs exist.
    """
    fails = [c for c in self_checks if c.status == "fail"]
    if not fails:
        return True

    st.markdown("---")
    st.error(
        "**Data integrity FAILs detected.** The analysis may be unreliable. "
        "Please review the discrepancies above before proceeding."
    )
    override = st.checkbox(
        "I have reviewed the discrepancies and wish to proceed with this data.",
        key="self_check_fail_override",
    )
    return override


# ── Session state init ─────────────────────────────────────────────────────

def _init_state():
    defaults = {
        "analysis_result": None,
        "financial_data": None,
        "session_info": {},
        "commentary": "",
        "pdf_extracted": None,
        "pdf_confirmed": False,
        "pdf_confirmed_data": None,
        "data_source": None,
        "firm_name": "Your Accounting Firm Pty Ltd",
        "ollama_model": "llama3.2",
        "ollama_running": False,
        "debug_mode": False,              # Bug: parsedData debug mode
        "self_check_override": False,     # Bug 3: FAIL override
        "parsing_metadata": {},           # Bug 4/5: column and source info
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


_init_state()

# ── Sidebar ────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## FinSight")
    st.markdown("*SME Financial Analysis Tool*")
    st.markdown("---")

    st.markdown("### Session Setup")

    client_name = st.text_input("Client / Business Name", value=st.session_state.session_info.get("client_name", ""))
    abn = st.text_input("ABN (optional)", value=st.session_state.session_info.get("abn", ""))

    industries = get_industry_list()
    industry_idx = 0
    stored_industry = st.session_state.session_info.get("industry", "")
    if stored_industry in industries:
        industry_idx = industries.index(stored_industry)
    industry = st.selectbox("Industry", industries, index=industry_idx)

    fy_end = st.text_input(
        "Financial Year End",
        value=st.session_state.session_info.get("financial_year_end", "30 June 2024")
    )
    currency = st.selectbox("Currency", ["AUD", "NZD", "USD", "GBP"], index=0)
    firm_name = st.text_input("Firm Name (for reports)", value=st.session_state.firm_name)

    st.markdown("---")
    st.markdown("### Upload Financial Statements")

    upload_type = st.radio(
        "Source Type",
        ["Xero CSV/Excel", "PDF Financial Statements", "Demo Mode"],
        index=0
    )

    uploaded_pl = uploaded_bs = uploaded_cf = uploaded_pdf = None

    if upload_type == "Xero CSV/Excel":
        uploaded_pl = st.file_uploader("P&L Statement (.xlsx/.csv)", type=["xlsx", "csv"])
        uploaded_bs = st.file_uploader("Balance Sheet (.xlsx/.csv)", type=["xlsx", "csv"])
        uploaded_cf = st.file_uploader("Cash Flow (optional, .xlsx/.csv)", type=["xlsx", "csv"])

    elif upload_type == "PDF Financial Statements":
        uploaded_pdf = st.file_uploader("Financial Statements PDF", type=["pdf"])

    run_analysis_btn = st.button("Run Analysis", type="primary", use_container_width=True)

    st.markdown("---")
    st.markdown("### AI Commentary (Ollama)")

    _ollama_ok, _ollama_msg = check_ollama_status()
    st.session_state.ollama_running = _ollama_ok
    _status_icon = "🟢" if _ollama_ok else "🔴"
    st.markdown(f"{_status_icon} **Ollama:** {_ollama_msg}")

    ollama_model = st.selectbox(
        "Model",
        ["llama3.2", "mistral", "qwen2.5", "llama3.1"],
        index=["llama3.2", "mistral", "qwen2.5", "llama3.1"].index(
            st.session_state.get("ollama_model", "llama3.2")
        ),
        help="Local Ollama model to use. Pull with: ollama pull <model>",
    )
    st.session_state.ollama_model = ollama_model

    if not _ollama_ok:
        st.caption("Start Ollama to enable AI commentary generation.")

    st.markdown("---")
    # Bug: parsedData debug mode toggle
    st.markdown("### Developer Options")
    debug_mode = st.checkbox(
        "Debug Mode (show parsed data)",
        value=st.session_state.debug_mode,
        help="Show full parsed data structure with section tags and classifications. Off by default.",
    )
    st.session_state.debug_mode = debug_mode

# ── Update session info ────────────────────────────────────────────────────

st.session_state.session_info = {
    "client_name": client_name,
    "abn": abn,
    "industry": industry,
    "financial_year_end": fy_end,
    "currency": currency,
}
st.session_state.firm_name = firm_name

# ── Run Analysis ───────────────────────────────────────────────────────────

if run_analysis_btn:
    st.session_state.analysis_result = None
    st.session_state.commentary = ""
    st.session_state.pdf_confirmed = False
    st.session_state.self_check_override = False

    industry_benchmarks = get_industry_benchmarks(industry)

    if upload_type == "Demo Mode":
        with st.spinner("Loading demo data..."):
            demo = get_demo_data()
            st.session_state.financial_data = demo
            st.session_state.session_info.update({
                "client_name": demo["client_name"],
                "industry": demo["industry"],
                "abn": demo["abn"],
                "financial_year_end": demo["financial_year_end"],
            })
            result = run_analysis(demo, industry_benchmarks)
            st.session_state.analysis_result = result
            st.session_state.data_source = "demo"
            st.session_state.parsing_metadata = {
                "reference_columns": demo.get("reference_columns", []),
                "inventory_source": demo["data"]["current"].get("_inventory_source", "demo"),
                "data_source": "Demo Mode",
            }
        st.success("Demo data loaded. Navigate the tabs to explore the analysis.")

    elif upload_type == "Xero CSV/Excel":
        if not uploaded_pl and not uploaded_bs:
            st.error("Please upload at least a P&L or Balance Sheet file.")
        else:
            with st.spinner("Parsing Xero files..."):
                try:
                    pl_data = parse_xero_pl(uploaded_pl) if uploaded_pl else {"current": {}, "prior": None}
                    bs_data = parse_xero_balance_sheet(uploaded_bs) if uploaded_bs else {"current": {}, "prior": None}
                    cf_data = parse_xero_cashflow(uploaded_cf) if uploaded_cf else None
                    merged = merge_financial_data(pl_data, bs_data, cf_data)
                    merged["period_labels"] = pl_data.get("period_labels", ["Current", "Prior"])[:2]
                    st.session_state.financial_data = merged
                    result = run_analysis(merged, industry_benchmarks)
                    st.session_state.analysis_result = result
                    st.session_state.data_source = "xero"

                    ref_cols = merged.get("reference_columns", [])
                    inv_source = merged.get("data", {}).get("current", {}).get("_inventory_source", "")
                    st.session_state.parsing_metadata = {
                        "reference_columns": ref_cols,
                        "inventory_source": inv_source,
                        "data_source": "Xero CSV/Excel",
                    }

                    labels = merged["period_labels"]
                    if pl_data.get("period_fallback_warning"):
                        st.warning(
                            f"⚠️ Could not detect date patterns in column headers — "
                            f"using positional order. Detected periods: **{' | '.join(labels)}**. "
                            "Please verify this is correct (newest column first)."
                        )
                    else:
                        st.success(
                            f"Xero files parsed. Detected periods: **{' | '.join(labels)}**"
                        )

                    if ref_cols:
                        st.info(
                            f"Note reference columns detected and excluded from data extraction: "
                            f"**{', '.join(str(c) for c in ref_cols)}**"
                        )

                except Exception as e:
                    st.error(f"Error parsing Xero files: {e}")
                    logger.exception("Xero parse error")

    elif upload_type == "PDF Financial Statements":
        if not uploaded_pdf:
            st.error("Please upload a PDF file.")
        else:
            with st.spinner("Extracting data from PDF..."):
                try:
                    extracted, notes = parse_pdf(uploaded_pdf)
                    st.session_state.pdf_extracted = extracted
                    st.session_state.data_source = "pdf"
                    st.info(notes)
                    st.warning(
                        "PDF parsing is approximate. Please review and correct the extracted "
                        "figures in the confirmation form below before proceeding."
                    )
                except Exception as e:
                    st.error(f"Error parsing PDF: {e}")
                    logger.exception("PDF parse error")

# ── PDF Confirmation Step ─────────────────────────────────────────────────

if (st.session_state.data_source == "pdf"
        and st.session_state.pdf_extracted is not None
        and not st.session_state.pdf_confirmed):

    st.markdown("## Data Confirmation — Review Extracted Figures")
    st.markdown(
        "The figures below were extracted from the PDF. **Please review each value carefully** "
        "and correct any errors before running the analysis. Enter `0` for items that are not "
        "applicable. Leave blank or `0` if unknown."
    )

    # Bug 4/5: Show column and source info if available
    extracted = st.session_state.pdf_extracted
    col_info = extracted.get("_column_info", {})
    if col_info.get("excluded_ref_cols"):
        st.info(
            f"Note reference columns excluded from extraction: "
            f"**{', '.join(str(c) for c in col_info['excluded_ref_cols'])}**"
        )
    if col_info.get("value_cols_used"):
        st.info(
            f"Value columns used for data extraction: "
            f"**{', '.join(str(c) for c in col_info['value_cols_used'])}**"
        )

    # Bug 5: Show inventory source
    inv_source = extracted.get("_inventory_source", "")
    if inv_source and inv_source != "not_found":
        st.success(f"Inventory source: **{inv_source}**")
    elif inv_source == "not_found":
        st.warning(
            "Inventory was not identified in the Balance Sheet current assets section. "
            "If the entity holds inventory, please enter the Balance Sheet value below."
        )

    template = get_confirmation_template(st.session_state.pdf_extracted)
    confirmed = {}

    current_section = None
    with st.form("pdf_confirmation"):
        for field, info in template.items():
            section = info["section"]
            if section != current_section:
                st.markdown(f"### {section}")
                current_section = section

            default_val = info["value"]
            default_str = f"{default_val:.0f}" if default_val is not None else ""

            help_text = f"Enter the {info['label']} figure from the financial statements"
            if field == "inventory":
                help_text = "IMPORTANT: Enter the Inventory/Stock value from the Balance Sheet current assets section only — not from the P&L."

            val = st.text_input(
                label=info["label"],
                value=default_str,
                key=f"pdf_confirm_{field}",
                help=help_text,
            )
            confirmed[field] = val

        # Prior year
        st.markdown("### Prior Year (if available)")
        st.markdown(
            "If the PDF contains prior year figures, enter them below. "
            "Leave blank if no prior period data."
        )
        prior_confirmed = {}
        for field, info in template.items():
            val_p = st.text_input(
                label=f"{info['label']} (Prior Year)",
                value="",
                key=f"pdf_confirm_prior_{field}",
            )
            prior_confirmed[field] = val_p

        col1, col2 = st.columns([1, 3])
        with col1:
            submit_confirmation = st.form_submit_button("Confirm & Run Analysis", type="primary")

    if submit_confirmation:
        cur_data = build_confirmed_data(confirmed)
        pri_data = build_confirmed_data(prior_confirmed)
        has_prior = any(v is not None and v != 0 for v in pri_data.values())

        financial_data = {
            "data": {
                "current": cur_data,
                "prior": pri_data if has_prior else None,
            },
            "period_labels": [fy_end, "Prior Year"],
        }
        st.session_state.financial_data = financial_data
        industry_benchmarks = get_industry_benchmarks(industry)
        result = run_analysis(financial_data, industry_benchmarks)
        st.session_state.analysis_result = result
        st.session_state.pdf_confirmed = True
        inv_src = cur_data.get("_inventory_source", "user confirmed")
        st.session_state.parsing_metadata = {
            "reference_columns": [],
            "inventory_source": inv_src,
            "data_source": "PDF (user confirmed)",
        }
        st.success("Analysis complete. Navigate the tabs above to view results.")
        st.rerun()

# ── Main Content Area ──────────────────────────────────────────────────────

result: AnalysisResult = st.session_state.analysis_result

if result is None:
    st.markdown("""
    # FinSight — SME Financial Analysis Tool

    Welcome to FinSight, an internal tool for analysing SME client financial statements.

    ### Getting Started

    **Step 1:** Complete the **Session Setup** in the sidebar (client name, industry, period).

    **Step 2:** Upload your files:
    - **Xero CSV/Excel**: Upload P&L and Balance Sheet exports from Xero
    - **PDF**: Upload a PDF containing the financial statements
    - **Demo Mode**: Explore the tool with sample data (no file needed)

    **Step 3:** Click **Run Analysis** to calculate metrics, benchmarks, and generate AI commentary.

    ---

    ### Features
    - 📊 20+ financial ratios with traffic-light status
    - 🎯 ATO small business benchmark comparison
    - 🤖 AI-generated meeting commentary via **Ollama** (100% local — no data leaves your machine)
    - ⚠️ Automatic red flag detection
    - 🔍 Financial data integrity self-checks before analysis
    - 📄 Export to PDF, Excel, or Word

    *Ensure [Ollama](https://ollama.com) is running locally to enable AI commentary generation.*
    """)

else:
    # ── Analysis results ────────────────────────────────────────────────

    client_display = st.session_state.session_info.get("client_name") or "Client"
    fy_display = st.session_state.session_info.get("financial_year_end") or ""
    industry_display = st.session_state.session_info.get("industry") or ""

    st.markdown(f"## {client_display} — Financial Analysis")
    st.markdown(
        f"**Industry:** {industry_display} &nbsp;|&nbsp; "
        f"**Period:** {fy_display} &nbsp;|&nbsp; "
        f"**Analysed:** {date.today().strftime('%d %b %Y')}"
    )

    # ── Bug 3: Persistent data quality WARN banner ────────────────────────
    if result.has_self_check_fails:
        st.markdown(
            '<div class="dq-warn-banner">'
            '❌ <strong>Data integrity FAILs detected</strong> — analysis results may be '
            'unreliable. Review the Data Quality panel in the Overview tab before relying '
            'on these figures.'
            '</div>',
            unsafe_allow_html=True,
        )
    elif result.has_self_check_warns:
        st.markdown(
            '<div class="dq-warn-banner">'
            '⚠️ <strong>Data quality warnings exist</strong> — review before relying on results. '
            'See the Data Quality panel in the Overview tab for details.'
            '</div>',
            unsafe_allow_html=True,
        )

    if result.red_flags:
        for flag in result.red_flags:
            st.markdown(f'<div class="red-flag">{flag}</div>', unsafe_allow_html=True)

    tabs = st.tabs(["Overview", "Profitability", "Liquidity", "Efficiency", "Leverage", "Benchmarks", "Commentary", "Export"])

    # ── TAB 0: OVERVIEW ────────────────────────────────────────────────────
    with tabs[0]:
        st.markdown('<div class="section-header">Key Metrics Scorecard</div>', unsafe_allow_html=True)

        cur_data = result.raw_data.get("current", {}) or {}
        pri_data = result.raw_data.get("prior", {}) or {}
        period_labels = result.period_labels

        # Top KPIs
        kpi_cols = st.columns(4)
        kpi_items = [
            ("revenue", "Revenue", "currency"),
            ("gross_profit", "Gross Profit", "currency"),
            ("net_profit", "Net Profit", "currency"),
            ("operating_cash_flow", "Operating Cash Flow", "currency"),
        ]
        for col, (key, label, _) in zip(kpi_cols, kpi_items):
            cur_v = cur_data.get(key)
            pri_v = pri_data.get(key)
            change = ""
            if cur_v is not None and pri_v is not None and pri_v != 0:
                pct = (cur_v - pri_v) / abs(pri_v) * 100
                change = f" ({'+' if pct >= 0 else ''}{pct:.1f}%)"
            with col:
                st.metric(
                    label=label,
                    value=_fmt_aud(cur_v),
                    delta=f"{_fmt_aud((cur_v or 0) - (pri_v or 0))}{change}" if pri_v is not None else None,
                )

        # ── Bug 3: Data Quality / Self-Check Panel ────────────────────────
        st.markdown("---")
        with st.expander(
            "📋 Data Quality Checks — "
            + ("✅ All passed" if not result.has_self_check_fails and not result.has_self_check_warns
               else ("❌ FAILs detected" if result.has_self_check_fails else "⚠️ Warnings present")),
            expanded=(result.has_self_check_fails or result.has_self_check_warns),
        ):
            _render_self_checks(result.self_checks, show_title=False)

            # Parsing metadata (column classification + inventory source)
            meta = st.session_state.parsing_metadata
            if meta:
                st.markdown("**Parsing Details:**")
                inv_src = meta.get("inventory_source", "")
                if inv_src and inv_src != "not_found":
                    st.markdown(f"- Inventory source: `{inv_src}`")
                elif inv_src == "not_found":
                    st.warning(
                        "Inventory not identified on Balance Sheet — "
                        "Quick Ratio = Current Ratio. Inventory Days not calculated."
                    )
                ref_cols = meta.get("reference_columns", [])
                if ref_cols:
                    st.markdown(
                        f"- Note reference columns excluded from extraction: "
                        f"`{', '.join(str(c) for c in ref_cols)}`"
                    )
                else:
                    st.markdown("- No note reference columns detected")

        st.markdown("---")

        # Scorecard grid
        st.markdown('<div class="section-header">All Metrics Summary</div>', unsafe_allow_html=True)

        all_cats = [
            ("liquidity", "Liquidity"),
            ("profitability", "Profitability"),
            ("efficiency", "Efficiency"),
            ("leverage", "Leverage"),
            ("growth", "Growth"),
        ]
        for cat_key, cat_label in all_cats:
            cat_metrics = [(k, m) for k, m in result.metrics.items() if m.category == cat_key]
            if not cat_metrics:
                continue
            st.markdown(f"**{cat_label}**")
            cols = st.columns(min(4, len(cat_metrics)))
            for col, (key, m) in zip(cols, cat_metrics):
                with col:
                    _metric_card(m.label, m.current_fmt, m.prior_fmt, m.trend, m.status, m.tooltip)
                    if m.notes:
                        st.caption(m.notes)

        # Revenue/GP/NP chart
        st.markdown("---")
        st.markdown('<div class="section-header">Revenue, Gross Profit & Net Profit</div>', unsafe_allow_html=True)

        rev_vals, gp_vals, np_vals = [], [], []
        chart_labels = []
        for p, lbl in [("prior", period_labels[1] if len(period_labels) > 1 else "Prior"),
                       ("current", period_labels[0] if period_labels else "Current")]:
            d = result.raw_data.get(p) or {}
            if any(d.get(k) for k in ["revenue", "gross_profit", "net_profit"]):
                chart_labels.append(lbl)
                rev_vals.append(d.get("revenue") or 0)
                gp_vals.append(d.get("gross_profit") or 0)
                np_vals.append(d.get("net_profit") or 0)

        if chart_labels:
            fig = go.Figure()
            fig.add_trace(go.Bar(name="Revenue", x=chart_labels, y=rev_vals, marker_color="#1B2A4A"))
            fig.add_trace(go.Bar(name="Gross Profit", x=chart_labels, y=gp_vals, marker_color="#3B82F6"))
            fig.add_trace(go.Bar(name="Net Profit", x=chart_labels, y=np_vals, marker_color="#10B981"))
            fig.update_layout(
                barmode="group", height=350,
                legend=dict(orientation="h", y=-0.2),
                margin=dict(l=0, r=0, t=20, b=60),
                yaxis_tickprefix="$", yaxis_tickformat=",.0f",
                plot_bgcolor="white", paper_bgcolor="white",
            )
            st.plotly_chart(fig, use_container_width=True)

        # ── Debug mode: full parsed data structure ────────────────────────
        if st.session_state.debug_mode:
            st.markdown("---")
            st.markdown("#### Debug Mode — Parsed Data Structure")
            st.caption(
                "This panel shows the full parsed data with section tags, component lists, "
                "and inventory source. Toggle off in the sidebar Developer Options."
            )
            import json

            def _make_serializable(obj):
                if isinstance(obj, dict):
                    return {k: _make_serializable(v) for k, v in obj.items()}
                elif isinstance(obj, list):
                    return [_make_serializable(i) for i in obj]
                elif isinstance(obj, (int, float)):
                    return obj
                else:
                    return str(obj)

            debug_data = _make_serializable(result.raw_data)
            st.json(debug_data)

    # ── TAB 1: PROFITABILITY ───────────────────────────────────────────────
    with tabs[1]:
        st.markdown('<div class="section-header">Profitability Metrics</div>', unsafe_allow_html=True)

        prof_metrics = [(k, m) for k, m in result.metrics.items() if m.category == "profitability"]
        cols = st.columns(min(3, len(prof_metrics)))
        for col, (key, m) in zip(cols, prof_metrics):
            with col:
                _metric_card(m.label, m.current_fmt, m.prior_fmt, m.trend, m.status, m.tooltip)
                if m.notes:
                    st.caption(m.notes)

        # ── Bug 2: EBIT/EBITDA component breakdown ────────────────────────
        ebit_m = result.metrics.get("ebit_margin")
        ebitda_m = result.metrics.get("ebitda_margin")
        if ebit_m and ebit_m.components:
            comp = ebit_m.components
            with st.expander("EBIT & EBITDA — Component Breakdown", expanded=False):
                st.markdown(
                    "EBIT and EBITDA are always calculated from components "
                    "(Net Profit + Tax + Interest + D&A). The figures used are:"
                )
                cur_data = result.raw_data.get("current", {}) or {}
                ebit_val = cur_data.get("_ebit_computed")
                ebitda_val = cur_data.get("_ebitda_computed")

                col_a, col_b = st.columns(2)
                with col_a:
                    st.markdown("**EBIT Calculation:**")
                    np_v = comp.get("net_profit")
                    int_v = comp.get("interest_expense", 0)
                    tax_v = comp.get("tax_expense", 0)
                    st.markdown(f"- Net Profit: **{_fmt_aud(np_v)}**")

                    int_items = comp.get("interest_items", [])
                    if int_items:
                        st.markdown(f"- Interest Expense: **{_fmt_aud(int_v)}**")
                        for lbl, val in int_items:
                            st.caption(f"  · {lbl}: {_fmt_aud(val)}")
                    else:
                        st.markdown(f"- Interest Expense: **{_fmt_aud(int_v)}** *(not identified)*")

                    tax_items = comp.get("tax_items", [])
                    if tax_items:
                        st.markdown(f"- Tax Expense: **{_fmt_aud(tax_v)}**")
                        for lbl, val in tax_items:
                            st.caption(f"  · {lbl}: {_fmt_aud(val)}")
                    else:
                        st.markdown(f"- Tax Expense: **{_fmt_aud(tax_v)}** *(not identified)*")

                    st.markdown(f"**= EBIT: {_fmt_aud(ebit_val)}**")

                with col_b:
                    st.markdown("**EBITDA Calculation:**")
                    dep_v = comp.get("depreciation", 0)
                    dep_items = comp.get("dep_items", [])
                    st.markdown(f"- EBIT: **{_fmt_aud(ebit_val)}**")
                    if dep_items:
                        st.markdown(f"- Depreciation & Amortisation: **{_fmt_aud(dep_v)}**")
                        for lbl, val in dep_items:
                            st.caption(f"  · {lbl}: {_fmt_aud(val)}")
                    else:
                        st.markdown(f"- D&A: **{_fmt_aud(dep_v)}** *(not identified — EBITDA may be understated)*")
                    st.markdown(f"**= EBITDA: {_fmt_aud(ebitda_val)}**")

                notes = comp.get("assumption_notes", [])
                if notes:
                    for note in notes:
                        st.caption(f"⚠️ {note}")

        # Margin trend chart
        if any(m.prior is not None for _, m in prof_metrics):
            st.markdown("---")
            st.markdown('<div class="section-header">Margin Trends</div>', unsafe_allow_html=True)

            trend_data = []
            for p_key, p_lbl in [
                ("prior2", period_labels[2] if len(period_labels) > 2 else "Prior 2"),
                ("prior", period_labels[1] if len(period_labels) > 1 else "Prior"),
                ("current", period_labels[0] if period_labels else "Current"),
            ]:
                d = result.raw_data.get(p_key) or {}
                if not any(d.values()):
                    continue
                rev = d.get("revenue")
                if not rev:
                    continue
                gp = d.get("gross_profit")
                np_ = d.get("net_profit")
                eb = d.get("_ebitda_computed") or d.get("ebitda")
                trend_data.append({
                    "Period": p_lbl,
                    "Gross Margin %": (gp / rev * 100) if gp is not None else None,
                    "Net Margin %": (np_ / rev * 100) if np_ is not None else None,
                    "EBITDA Margin %": (eb / rev * 100) if eb is not None else None,
                })

            if trend_data:
                df_trend = pd.DataFrame(trend_data)
                fig2 = go.Figure()
                for col_name, colour in [
                    ("Gross Margin %", "#1B2A4A"),
                    ("Net Margin %", "#10B981"),
                    ("EBITDA Margin %", "#3B82F6"),
                ]:
                    if col_name in df_trend.columns:
                        fig2.add_trace(go.Scatter(
                            x=df_trend["Period"], y=df_trend[col_name],
                            name=col_name, mode="lines+markers",
                            line=dict(color=colour, width=2),
                            marker=dict(size=8),
                        ))
                fig2.update_layout(
                    height=320, yaxis_ticksuffix="%",
                    margin=dict(l=0, r=0, t=10, b=30),
                    plot_bgcolor="white", paper_bgcolor="white",
                    legend=dict(orientation="h", y=-0.25),
                )
                st.plotly_chart(fig2, use_container_width=True)

        # Waterfall chart
        st.markdown('<div class="section-header">Waterfall — Revenue to Net Profit</div>', unsafe_allow_html=True)
        cur_d = result.raw_data.get("current") or {}
        rev = cur_d.get("revenue") or 0
        cogs = cur_d.get("cogs") or 0
        gp = cur_d.get("gross_profit") or (rev - cogs)
        opex = cur_d.get("operating_expenses") or 0
        ebit = cur_d.get("_ebit_computed") or cur_d.get("ebit") or (gp - opex)
        interest = cur_d.get("interest_expense") or 0
        tax = cur_d.get("tax_expense") or 0
        net = cur_d.get("net_profit") or 0

        if rev:
            fig3 = go.Figure(go.Waterfall(
                name="Waterfall",
                orientation="v",
                measure=["absolute", "relative", "total", "relative", "total", "relative", "relative", "total"],
                x=["Revenue", "COGS", "Gross Profit", "Opex", "EBIT", "Interest", "Tax", "Net Profit"],
                y=[rev, -cogs, gp, -opex, ebit, -interest, -tax, net],
                connector={"line": {"color": "#9CA3AF"}},
                increasing={"marker": {"color": "#16A34A"}},
                decreasing={"marker": {"color": "#DC2626"}},
                totals={"marker": {"color": "#1B2A4A"}},
            ))
            fig3.update_layout(
                height=380, yaxis_tickprefix="$", yaxis_tickformat=",.0f",
                margin=dict(l=0, r=0, t=10, b=30),
                plot_bgcolor="white", paper_bgcolor="white",
            )
            st.plotly_chart(fig3, use_container_width=True)

    # ── TAB 2: LIQUIDITY ───────────────────────────────────────────────────
    with tabs[2]:
        st.markdown('<div class="section-header">Liquidity Metrics</div>', unsafe_allow_html=True)

        liq_metrics = [(k, m) for k, m in result.metrics.items() if m.category == "liquidity"]
        cols = st.columns(min(3, len(liq_metrics)))
        for col, (key, m) in zip(cols, liq_metrics):
            with col:
                _metric_card(m.label, m.current_fmt, m.prior_fmt, m.trend, m.status, m.tooltip)
                if m.notes:
                    st.caption(m.notes)

        st.markdown("---")
        st.markdown('<div class="section-header">Ratio Gauges</div>', unsafe_allow_html=True)

        gauge_cols = st.columns(2)
        cr = result.metrics.get("current_ratio")
        qr = result.metrics.get("quick_ratio")

        def _gauge(value, title, green_min, amber_min, max_val=4):
            if value is None:
                return go.Figure()
            fig = go.Figure(go.Indicator(
                mode="gauge+number",
                value=round(value, 2),
                title={"text": title, "font": {"size": 13, "color": "#1B2A4A"}},
                number={"suffix": "x", "font": {"size": 20, "color": "#1B2A4A"}},
                gauge={
                    "axis": {"range": [0, max_val], "tickformat": ".1f"},
                    "bar": {"color": "#1B2A4A"},
                    "steps": [
                        {"range": [0, amber_min], "color": "#FEE2E2"},
                        {"range": [amber_min, green_min], "color": "#FEF3C7"},
                        {"range": [green_min, max_val], "color": "#DCFCE7"},
                    ],
                    "threshold": {
                        "line": {"color": "#1B2A4A", "width": 2},
                        "thickness": 0.75,
                        "value": value,
                    },
                }
            ))
            fig.update_layout(height=250, margin=dict(l=20, r=20, t=40, b=0), paper_bgcolor="white")
            return fig

        if cr:
            with gauge_cols[0]:
                st.plotly_chart(_gauge(cr.current, "Current Ratio", 2.0, 1.0), use_container_width=True)
        if qr:
            with gauge_cols[1]:
                st.plotly_chart(_gauge(qr.current, "Quick Ratio", 1.0, 0.5, max_val=3), use_container_width=True)

    # ── TAB 3: EFFICIENCY ─────────────────────────────────────────────────
    with tabs[3]:
        st.markdown('<div class="section-header">Operational Efficiency</div>', unsafe_allow_html=True)

        eff_metrics = [(k, m) for k, m in result.metrics.items() if m.category == "efficiency"]
        cols = st.columns(min(4, len(eff_metrics)))
        for col, (key, m) in zip(cols, eff_metrics):
            with col:
                _metric_card(m.label, m.current_fmt, m.prior_fmt, m.trend, m.status, m.tooltip)
                if m.notes:
                    st.markdown(f'<div class="red-flag">{m.notes}</div>', unsafe_allow_html=True)

        # Working capital days chart
        st.markdown("---")
        st.markdown('<div class="section-header">Working Capital Days Comparison</div>', unsafe_allow_html=True)

        day_metrics = ["debtor_days", "creditor_days", "inventory_days"]
        day_names = ["Debtor Days", "Creditor Days", "Inventory Days"]
        cur_vals = [result.metrics.get(k).current if result.metrics.get(k) else None for k in day_metrics]
        pri_vals = [result.metrics.get(k).prior if result.metrics.get(k) else None for k in day_metrics]

        cur_vals_clean = [v if v is not None else 0 for v in cur_vals]
        pri_vals_clean = [v if v is not None else 0 for v in pri_vals]

        if any(cur_vals_clean):
            fig4 = go.Figure()
            fig4.add_trace(go.Bar(
                name=period_labels[0] if period_labels else "Current",
                x=day_names, y=cur_vals_clean,
                marker_color="#1B2A4A"
            ))
            if any(pri_vals_clean):
                fig4.add_trace(go.Bar(
                    name=period_labels[1] if len(period_labels) > 1 else "Prior",
                    x=day_names, y=pri_vals_clean,
                    marker_color="#93C5FD"
                ))
            fig4.update_layout(
                barmode="group", height=320,
                yaxis_title="Days", yaxis_ticksuffix=" d",
                margin=dict(l=0, r=0, t=10, b=30),
                plot_bgcolor="white", paper_bgcolor="white",
                legend=dict(orientation="h", y=-0.25),
            )
            st.plotly_chart(fig4, use_container_width=True)

        # Bug 5: Inventory source confirmation
        meta = st.session_state.parsing_metadata
        inv_src = meta.get("inventory_source", "")
        if inv_src and inv_src not in ("not_found", ""):
            st.info(f"Inventory source: **{inv_src}**")
        elif inv_src == "not_found":
            st.warning(
                "Inventory not found on Balance Sheet. "
                "Quick Ratio = Current Ratio. Inventory Days not calculated."
            )

    # ── TAB 4: LEVERAGE ────────────────────────────────────────────────────
    with tabs[4]:
        st.markdown('<div class="section-header">Leverage & Solvency</div>', unsafe_allow_html=True)

        lev_metrics = [(k, m) for k, m in result.metrics.items() if m.category == "leverage"]
        cols = st.columns(min(3, len(lev_metrics)))
        for col, (key, m) in zip(cols, lev_metrics):
            with col:
                _metric_card(m.label, m.current_fmt, m.prior_fmt, m.trend, m.status, m.tooltip)

        grow_metrics = [(k, m) for k, m in result.metrics.items() if m.category == "growth"]
        if grow_metrics:
            st.markdown("---")
            st.markdown('<div class="section-header">Growth Metrics (Year-on-Year)</div>', unsafe_allow_html=True)
            cols2 = st.columns(min(4, len(grow_metrics)))
            for col, (key, m) in zip(cols2, grow_metrics):
                with col:
                    _metric_card(m.label, m.current_fmt, m.prior_fmt, m.trend, m.status, m.tooltip)
                    if m.notes:
                        st.markdown(f'<div class="red-flag">{m.notes}</div>', unsafe_allow_html=True)

    # ── TAB 5: BENCHMARKS ─────────────────────────────────────────────────
    with tabs[5]:
        st.markdown('<div class="section-header">ATO Small Business Benchmark Comparison</div>', unsafe_allow_html=True)

        bm_meta = get_benchmark_metadata()
        st.info(
            f"ATO benchmarks for **{industry}**. "
            f"Note: {bm_meta.get('note', 'Benchmarks updated annually and may lag by one year.')} "
            f"Source: {bm_meta.get('url', 'ato.gov.au')}"
        )

        if result.benchmark_comparisons:
            bm_rows = []
            for key, comp in result.benchmark_comparisons.items():
                actual = comp.get("actual_pct")
                low = comp.get("benchmark_low")
                high = comp.get("benchmark_high")
                in_range = actual is not None and low is not None and high is not None and low <= actual <= high
                bm_rows.append({
                    "Expense Category": comp["label"],
                    "Client % of Turnover": f"{actual:.1f}%" if actual is not None else "N/A",
                    "ATO Benchmark": f"{low:.0f}%–{high:.0f}%" if low is not None else "N/A",
                    "Status": "✅ Within range" if in_range else ("⚠️ Outside range" if actual else "N/A"),
                })

            df_bm = pd.DataFrame(bm_rows)
            st.dataframe(df_bm, use_container_width=True, hide_index=True)

            st.markdown("---")
            bm_labels, bm_actuals, bm_lows, bm_highs = [], [], [], []
            for key, comp in result.benchmark_comparisons.items():
                actual = comp.get("actual_pct")
                low = comp.get("benchmark_low")
                high = comp.get("benchmark_high")
                if actual is not None:
                    bm_labels.append(comp["label"])
                    bm_actuals.append(actual)
                    bm_lows.append(low or 0)
                    bm_highs.append(high or 0)

            if bm_labels:
                fig5 = go.Figure()
                fig5.add_trace(go.Bar(
                    name="ATO Benchmark Range",
                    x=[h - l for l, h in zip(bm_lows, bm_highs)],
                    y=bm_labels,
                    orientation="h",
                    base=bm_lows,
                    marker=dict(color="rgba(59,130,246,0.2)", line=dict(color="#3B82F6", width=1)),
                ))
                fig5.add_trace(go.Scatter(
                    name="Client Value",
                    x=bm_actuals,
                    y=bm_labels,
                    mode="markers",
                    marker=dict(size=12, color="#DC2626", symbol="diamond"),
                ))
                fig5.update_layout(
                    height=320, xaxis_ticksuffix="%", xaxis_title="% of Turnover",
                    margin=dict(l=0, r=0, t=10, b=30),
                    plot_bgcolor="white", paper_bgcolor="white",
                    legend=dict(orientation="h", y=-0.25),
                    barmode="overlay",
                )
                st.plotly_chart(fig5, use_container_width=True)
        else:
            st.info("No benchmark comparison data available. Ensure revenue and expense figures are present in the financial data.")

    # ── TAB 6: COMMENTARY ─────────────────────────────────────────────────
    with tabs[6]:
        st.markdown('<div class="section-header">AI Commentary (Ollama Local LLM)</div>', unsafe_allow_html=True)
        st.markdown(
            "Generate professional accountant commentary using a local Ollama model. "
            "**Client data never leaves your machine.** "
            "Edit the commentary below before including it in an export."
        )

        _is_running = st.session_state.get("ollama_running", False)
        _model = st.session_state.get("ollama_model", "llama3.2")

        col_gen, col_info = st.columns([2, 3])
        with col_gen:
            generate_btn = st.button(
                "Generate AI Commentary",
                type="primary",
                disabled=not _is_running,
                help="Ollama must be running. Start with: ollama serve" if not _is_running else f"Using model: {_model}",
            )
        with col_info:
            if not _is_running:
                st.warning("Ollama is not running. Start it with: `ollama serve`")
            else:
                st.info(f"Model: **{_model}** | Data stays local")

        if generate_btn:
            prompt = build_commentary_prompt(
                financial_data=result.raw_data,
                metrics=result.metrics,
                red_flags=result.red_flags,
                benchmark_comparisons=result.benchmark_comparisons,
                session_info=st.session_state.session_info,
                period_labels=result.period_labels,
            )
            commentary_container = st.empty()
            commentary_text = ""
            with st.spinner(f"Generating commentary with {_model}... (this may take 30-60 seconds)"):
                try:
                    for chunk in generate_commentary_streaming(prompt, model=_model):
                        commentary_text += chunk
                        commentary_container.markdown(commentary_text)
                    st.session_state.commentary = commentary_text
                    st.success("Commentary generated. Edit below before exporting.")
                except Exception as e:
                    st.error(f"Error generating commentary: {e}")
                    logger.exception("Commentary generation error")

        if st.session_state.commentary:
            st.markdown("---")
            st.markdown("**Edit Commentary** (changes are preserved for export):")
            edited = st.text_area(
                "Commentary",
                value=st.session_state.commentary,
                height=500,
                label_visibility="collapsed",
            )
            if edited != st.session_state.commentary:
                st.session_state.commentary = edited

    # ── TAB 7: EXPORT ─────────────────────────────────────────────────────
    with tabs[7]:
        st.markdown('<div class="section-header">Export Report</div>', unsafe_allow_html=True)
        st.markdown(
            "Download the full analysis report in your preferred format. "
            "AI commentary will be included if generated."
        )

        # Bug 3: Warn about data quality in exports
        if result.has_self_check_fails:
            st.error(
                "Data integrity FAILs are present. Exports will include a data quality warning. "
                "Review the Overview tab before distributing any reports."
            )
        elif result.has_self_check_warns:
            st.warning(
                "Data quality warnings exist. Review the Overview tab before distributing reports."
            )

        client_name_export = st.session_state.session_info.get("client_name", "Client")
        safe_name = "".join(c if c.isalnum() or c in " -_" else "_" for c in client_name_export)[:40]
        today_str = date.today().strftime("%Y%m%d")

        exp_col1, exp_col2, exp_col3 = st.columns(3)

        with exp_col1:
            st.markdown("#### PDF Report")
            st.markdown("Professional formatted PDF with all metrics, charts summary, benchmarks, and commentary.")
            if st.button("Generate PDF", type="primary", use_container_width=True):
                with st.spinner("Generating PDF report..."):
                    try:
                        pdf_bytes = generate_pdf_report(
                            analysis_result=result,
                            session_info=st.session_state.session_info,
                            commentary=st.session_state.commentary,
                            firm_name=st.session_state.firm_name,
                        )
                        st.download_button(
                            label="Download PDF",
                            data=pdf_bytes,
                            file_name=f"FinSight_{safe_name}_{today_str}.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                        )
                    except Exception as e:
                        st.error(f"PDF generation error: {e}")
                        logger.exception("PDF export error")

        with exp_col2:
            st.markdown("#### Excel Workbook")
            st.markdown("Multi-tab Excel file with raw data, colour-coded metrics, benchmarks, and commentary.")
            if st.button("Generate Excel", type="primary", use_container_width=True):
                with st.spinner("Generating Excel workbook..."):
                    try:
                        xl_bytes = generate_excel_report(
                            analysis_result=result,
                            session_info=st.session_state.session_info,
                            commentary=st.session_state.commentary,
                        )
                        st.download_button(
                            label="Download Excel",
                            data=xl_bytes,
                            file_name=f"FinSight_{safe_name}_{today_str}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )
                    except Exception as e:
                        st.error(f"Excel generation error: {e}")
                        logger.exception("Excel export error")

        with exp_col3:
            st.markdown("#### Word Document")
            st.markdown("Editable .docx report suitable for personalising before client delivery.")
            if st.button("Generate Word", type="primary", use_container_width=True):
                with st.spinner("Generating Word document..."):
                    try:
                        docx_bytes = generate_word_report(
                            analysis_result=result,
                            session_info=st.session_state.session_info,
                            commentary=st.session_state.commentary,
                            firm_name=st.session_state.firm_name,
                        )
                        st.download_button(
                            label="Download Word",
                            data=docx_bytes,
                            file_name=f"FinSight_{safe_name}_{today_str}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                        )
                    except Exception as e:
                        st.error(f"Word generation error: {e}")
                        logger.exception("Word export error")

        st.markdown("---")
        st.markdown(
            "*All reports include the footer: "
            "\"Prepared for internal use only — not for distribution without review\"*"
        )
