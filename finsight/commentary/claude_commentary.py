"""
AI commentary generation using the Anthropic API (claude-sonnet-4-6).
Generates professional accountant commentary for client meetings.
"""

import os
import logging
from typing import Optional
from dotenv import load_dotenv

import anthropic

load_dotenv()
logger = logging.getLogger(__name__)

SYSTEM_PROMPT = """You are a senior Australian CPA providing internal analysis notes for a public accounting firm. Your commentary is fact-based, professional, and structured as talking points for the accountant to use in a client meeting. Use plain English. Do not use jargon unnecessarily. Flag risks clearly but constructively. Reference specific figures. Do not give investment advice. Use Australian English spelling throughout."""

COMMENTARY_TEMPLATE = """Analyse the following financial data for {client_name} ({industry} industry, {financial_year}) and produce structured commentary for an accountant to use in a client meeting.

## FINANCIAL SUMMARY

**Period Labels:** {period_labels}

### Profit & Loss
{pl_summary}

### Balance Sheet
{bs_summary}

### Cash Flow
{cf_summary}

## KEY METRICS
{metrics_summary}

## ATO BENCHMARK COMPARISONS ({industry})
{benchmark_summary}

## RED FLAGS DETECTED
{red_flags}

---

Please provide commentary in the following exact structure:

### 1. Executive Summary
(3–5 sentences summarising the overall financial position and year-on-year trend)

### 2. Key Strengths
(Bullet points with specific figures — what is working well)

### 3. Areas of Concern
(Bullet points with specific figures and suggested discussion questions for the accountant to raise with the client)

### 4. ATO Benchmark Observations
(Where the client sits relative to ATO benchmarks for their industry and what this may indicate)

### 5. Suggested Focus Areas for Next Period
(Actionable priorities the business should focus on)

Keep each section concise and practical. Reference dollar amounts and percentages where relevant. Write in a style suitable for an accountant briefing document, not a formal audit report."""


def _format_currency(value) -> str:
    if value is None:
        return "N/A"
    return f"${value:,.0f}"


def _format_pct(value) -> str:
    if value is None:
        return "N/A"
    return f"{value:.1f}%"


def _format_ratio(value) -> str:
    if value is None:
        return "N/A"
    return f"{value:.2f}x"


def _build_pl_summary(data: dict, period_labels: list) -> str:
    cur = data.get("current", {}) or {}
    prior = data.get("prior", {}) or {}
    label_cur = period_labels[0] if period_labels else "Current"
    label_pri = period_labels[1] if len(period_labels) > 1 else "Prior"

    lines = [f"| Metric | {label_cur} | {label_pri} |", "|---|---|---|"]
    fields = [
        ("revenue", "Revenue"),
        ("cogs", "COGS"),
        ("gross_profit", "Gross Profit"),
        ("operating_expenses", "Operating Expenses"),
        ("ebit", "EBIT"),
        ("ebitda", "EBITDA"),
        ("interest_expense", "Interest Expense"),
        ("tax_expense", "Tax Expense"),
        ("net_profit", "Net Profit"),
    ]
    for key, label in fields:
        c_val = _format_currency(cur.get(key))
        p_val = _format_currency(prior.get(key))
        lines.append(f"| {label} | {c_val} | {p_val} |")
    return "\n".join(lines)


def _build_bs_summary(data: dict, period_labels: list) -> str:
    cur = data.get("current", {}) or {}
    prior = data.get("prior", {}) or {}
    label_cur = period_labels[0] if period_labels else "Current"
    label_pri = period_labels[1] if len(period_labels) > 1 else "Prior"

    lines = [f"| Metric | {label_cur} | {label_pri} |", "|---|---|---|"]
    fields = [
        ("cash", "Cash"),
        ("accounts_receivable", "Accounts Receivable"),
        ("inventory", "Inventory"),
        ("current_assets", "Current Assets"),
        ("total_assets", "Total Assets"),
        ("accounts_payable", "Accounts Payable"),
        ("current_liabilities", "Current Liabilities"),
        ("total_liabilities", "Total Liabilities"),
        ("equity", "Total Equity"),
    ]
    for key, label in fields:
        c_val = _format_currency(cur.get(key))
        p_val = _format_currency(prior.get(key))
        lines.append(f"| {label} | {c_val} | {p_val} |")
    return "\n".join(lines)


def _build_cf_summary(data: dict, period_labels: list) -> str:
    cur = data.get("current", {}) or {}
    prior = data.get("prior", {}) or {}
    label_cur = period_labels[0] if period_labels else "Current"
    label_pri = period_labels[1] if len(period_labels) > 1 else "Prior"

    ocf_cur = cur.get("operating_cash_flow")
    ocf_pri = prior.get("operating_cash_flow")
    if ocf_cur is None and ocf_pri is None:
        return "Cash flow statement not available."

    lines = [f"| Metric | {label_cur} | {label_pri} |", "|---|---|---|"]
    lines.append(f"| Operating Cash Flow | {_format_currency(ocf_cur)} | {_format_currency(ocf_pri)} |")
    lines.append(f"| Investing Cash Flow | {_format_currency(cur.get('investing_cash_flow'))} | {_format_currency(prior.get('investing_cash_flow'))} |")
    lines.append(f"| Financing Cash Flow | {_format_currency(cur.get('financing_cash_flow'))} | {_format_currency(prior.get('financing_cash_flow'))} |")
    return "\n".join(lines)


def _build_metrics_summary(metrics: dict) -> str:
    lines = []
    category_order = ["liquidity", "profitability", "efficiency", "leverage", "growth"]
    category_labels = {
        "liquidity": "LIQUIDITY",
        "profitability": "PROFITABILITY",
        "efficiency": "EFFICIENCY",
        "leverage": "LEVERAGE & SOLVENCY",
        "growth": "GROWTH",
    }

    for cat in category_order:
        cat_metrics = [(k, v) for k, v in metrics.items() if v.category == cat]
        if not cat_metrics:
            continue
        lines.append(f"\n**{category_labels[cat]}**")
        for key, m in cat_metrics:
            status_icon = {"green": "✅", "amber": "🟡", "red": "🔴", "grey": "⬜"}.get(m.status, "⬜")
            trend = m.trend
            lines.append(f"- {status_icon} {m.label}: {m.current_fmt} (Prior: {m.prior_fmt}) {trend}")
            if m.notes:
                lines.append(f"  {m.notes}")

    return "\n".join(lines)


def _build_benchmark_summary(benchmark_comparisons: dict, industry: str) -> str:
    if not benchmark_comparisons:
        return "No benchmark data available."

    lines = [f"Industry: {industry}", ""]
    for key, comp in benchmark_comparisons.items():
        label = comp["label"]
        actual = comp.get("actual_pct")
        low = comp.get("benchmark_low")
        high = comp.get("benchmark_high")
        if actual is not None:
            in_range = low is not None and high is not None and low <= actual <= high
            status = "✅ Within range" if in_range else "⚠️ Outside range"
            lines.append(f"- {label}: {actual:.1f}% (ATO benchmark: {low}–{high}%) — {status}")

    return "\n".join(lines)


def build_commentary_prompt(
    client_name: str,
    industry: str,
    financial_year: str,
    financial_data: dict,
    metrics: dict,
    red_flags: list,
    benchmark_comparisons: dict,
    period_labels: list,
) -> str:
    """Build the user prompt for Claude commentary."""
    pl_summary = _build_pl_summary(financial_data, period_labels)
    bs_summary = _build_bs_summary(financial_data, period_labels)
    cf_summary = _build_cf_summary(financial_data, period_labels)
    metrics_summary = _build_metrics_summary(metrics)
    benchmark_summary = _build_benchmark_summary(benchmark_comparisons, industry)
    red_flags_text = "\n".join(red_flags) if red_flags else "No major red flags detected."

    return COMMENTARY_TEMPLATE.format(
        client_name=client_name,
        industry=industry,
        financial_year=financial_year,
        period_labels=", ".join(period_labels),
        pl_summary=pl_summary,
        bs_summary=bs_summary,
        cf_summary=cf_summary,
        metrics_summary=metrics_summary,
        benchmark_summary=benchmark_summary,
        red_flags=red_flags_text,
    )


def generate_commentary(
    client_name: str,
    industry: str,
    financial_year: str,
    financial_data: dict,
    metrics: dict,
    red_flags: list,
    benchmark_comparisons: dict,
    period_labels: list,
    api_key: Optional[str] = None,
) -> str:
    """
    Call the Anthropic API to generate professional commentary.
    Uses claude-sonnet-4-6 with streaming for reliability.
    Returns the generated commentary as a string.
    """
    key = api_key or os.getenv("ANTHROPIC_API_KEY")
    if not key:
        raise ValueError("ANTHROPIC_API_KEY not set. Please add it to your .env file.")

    client = anthropic.Anthropic(api_key=key)

    user_prompt = build_commentary_prompt(
        client_name=client_name,
        industry=industry,
        financial_year=financial_year,
        financial_data=financial_data,
        metrics=metrics,
        red_flags=red_flags,
        benchmark_comparisons=benchmark_comparisons,
        period_labels=period_labels,
    )

    # Use streaming to handle long responses and avoid timeouts
    with client.messages.stream(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_prompt}],
    ) as stream:
        commentary = stream.get_final_message()

    return commentary.content[0].text


def generate_commentary_streaming(
    client_name: str,
    industry: str,
    financial_year: str,
    financial_data: dict,
    metrics: dict,
    red_flags: list,
    benchmark_comparisons: dict,
    period_labels: list,
    api_key: Optional[str] = None,
):
    """
    Generator version of generate_commentary.
    Yields text chunks as they arrive for real-time display.
    """
    key = api_key or os.getenv("ANTHROPIC_API_KEY")
    if not key:
        raise ValueError("ANTHROPIC_API_KEY not set. Please add it to your .env file.")

    client = anthropic.Anthropic(api_key=key)

    user_prompt = build_commentary_prompt(
        client_name=client_name,
        industry=industry,
        financial_year=financial_year,
        financial_data=financial_data,
        metrics=metrics,
        red_flags=red_flags,
        benchmark_comparisons=benchmark_comparisons,
        period_labels=period_labels,
    )

    with client.messages.stream(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_prompt}],
    ) as stream:
        for text in stream.text_stream:
            yield text
