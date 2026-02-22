"""
AI commentary generation using Ollama local LLM.
All client data stays on-machine — no external API calls are made.

Usage:
  1. Install Ollama: https://ollama.com
  2. Pull a model: ollama pull llama3.2
  3. Ensure Ollama is running: ollama serve
"""

import json
import logging
from typing import Iterator

import requests

logger = logging.getLogger(__name__)

OLLAMA_BASE_URL = "http://localhost:11434"
DEFAULT_MODEL = "llama3.2"

SYSTEM_PROMPT = """You are an experienced Australian CPA (Chartered Professional Accountant) \
specialising in small-to-medium enterprise (SME) financial analysis. Your role is to provide \
clear, practical financial commentary that helps accountants prepare for client meetings.

Write in plain English. Be specific with numbers. Flag concerns clearly but constructively. \
Use Australian English spelling and conventions. Reference ATO benchmarks where relevant. \
Keep the commentary practical and actionable."""

COMMENTARY_TEMPLATE = """
{system_prompt}

Please analyse the financial data below and provide structured commentary under each heading:

## Executive Summary
2-3 sentences on overall financial health and the most important themes.

## Trading Performance
Analysis of revenue, gross profit, and profitability trends. Note any margin compression or improvement.

## Cashflow & Liquidity
Commentary on the liquidity position and quality of cash generation.

## Balance Sheet Strength
Review of asset quality, debt levels, and equity position.

## Key Risks & Opportunities
3-5 specific bullet points identifying risks or opportunities from the data.

## Talking Points for Client Meeting
3-5 specific questions or discussion items to raise with the client.

---

CLIENT DATA:
{data_summary}
"""


# ── Ollama connectivity ────────────────────────────────────────────────────

def check_ollama_status(base_url: str = OLLAMA_BASE_URL) -> tuple:
    """
    Check whether Ollama is running and reachable.
    Returns (is_running: bool, status_message: str).
    """
    try:
        resp = requests.get(f"{base_url}/api/tags", timeout=3)
        if resp.status_code == 200:
            models = resp.json().get("models", [])
            model_names = [m.get("name", "").split(":")[0] for m in models]
            if model_names:
                return True, f"Running — models: {', '.join(model_names[:4])}"
            return True, "Running — no models pulled yet (run: ollama pull llama3.2)"
        return False, f"Unexpected HTTP status {resp.status_code}"
    except requests.exceptions.ConnectionError:
        return False, "Not running — start with: ollama serve"
    except requests.exceptions.Timeout:
        return False, "Connection timed out"
    except Exception as exc:
        return False, str(exc)


# ── Prompt builders ────────────────────────────────────────────────────────

def _fmt(v) -> str:
    return f"${v:,.0f}" if v is not None else "N/A"


def _chg(cur, prior) -> str:
    if cur is not None and prior is not None and prior != 0:
        pct = (cur - prior) / abs(prior) * 100
        return f" ({pct:+.1f}%)"
    return ""


def _build_pl_summary(financial_data: dict, period_labels: list) -> str:
    cur = financial_data.get("current") or {}
    pri = financial_data.get("prior") or {}
    label_cur = period_labels[0] if period_labels else "Current"
    label_pri = period_labels[1] if len(period_labels) > 1 else "Prior"

    lines = [f"PROFIT & LOSS ({label_cur} vs {label_pri}):"]
    fields = [
        ("revenue", "Revenue"),
        ("cogs", "Cost of Goods Sold"),
        ("gross_profit", "Gross Profit"),
        ("operating_expenses", "Operating Expenses"),
        ("ebit", "EBIT"),
        ("ebitda", "EBITDA"),
        ("net_profit", "Net Profit"),
        ("interest_expense", "Interest Expense"),
        ("depreciation", "Depreciation"),
    ]
    for key, label in fields:
        c = cur.get(key)
        p = pri.get(key)
        if c is not None:
            lines.append(f"  {label}: {_fmt(c)} (Prior: {_fmt(p)}){_chg(c, p)}")
    return "\n".join(lines)


def _build_bs_summary(financial_data: dict, period_labels: list) -> str:
    cur = financial_data.get("current") or {}
    label_cur = period_labels[0] if period_labels else "Current"

    lines = [f"BALANCE SHEET ({label_cur}):"]
    fields = [
        ("cash", "Cash & Bank"),
        ("accounts_receivable", "Accounts Receivable"),
        ("inventory", "Inventory"),
        ("current_assets", "Total Current Assets"),
        ("total_assets", "Total Assets"),
        ("accounts_payable", "Accounts Payable"),
        ("current_liabilities", "Total Current Liabilities"),
        ("total_liabilities", "Total Liabilities"),
        ("equity", "Total Equity"),
        ("total_debt", "Total Debt"),
    ]
    for key, label in fields:
        v = cur.get(key)
        if v is not None:
            lines.append(f"  {label}: {_fmt(v)}")
    return "\n".join(lines)


def _build_cf_summary(financial_data: dict) -> str:
    cur = financial_data.get("current") or {}
    lines = ["CASH FLOW:"]
    for key, label in [
        ("operating_cash_flow", "Operating CF"),
        ("investing_cash_flow", "Investing CF"),
        ("financing_cash_flow", "Financing CF"),
    ]:
        v = cur.get(key)
        if v is not None:
            lines.append(f"  {label}: {_fmt(v)}")
    return "\n".join(lines) if len(lines) > 1 else ""


def _build_metrics_summary(metrics: dict) -> str:
    lines = ["KEY CALCULATED METRICS:"]
    status_map = {"green": "OK", "amber": "REVIEW", "red": "CONCERN", "grey": "N/A"}
    for key, m in metrics.items():
        flag = status_map.get(m.status, "")
        lines.append(
            f"  {m.label}: {m.current_fmt} [{flag}]"
            f" (Prior: {m.prior_fmt})"
        )
    return "\n".join(lines)


def _build_benchmark_summary(benchmark_comparisons: dict, industry: str) -> str:
    if not benchmark_comparisons:
        return ""
    lines = [f"ATO SMALL BUSINESS BENCHMARKS (Industry: {industry}):"]
    for key, comp in benchmark_comparisons.items():
        actual = comp.get("actual_pct")
        low = comp.get("benchmark_low")
        high = comp.get("benchmark_high")
        if actual is not None:
            in_range = low is not None and high is not None and low <= actual <= high
            status = "IN RANGE" if in_range else "OUTSIDE RANGE"
            lines.append(
                f"  {comp['label']}: {actual:.1f}%"
                f" (ATO range {low or '?'}%–{high or '?'}%) [{status}]"
            )
    return "\n".join(lines)


def build_commentary_prompt(
    financial_data: dict,
    metrics: dict,
    red_flags: list,
    benchmark_comparisons: dict,
    session_info: dict,
    period_labels: list,
) -> str:
    """
    Assemble the full prompt string for the local LLM.

    Parameters mirror the AnalysisResult fields so the caller can pass them
    directly from the result object.
    """
    client_name = session_info.get("client_name", "the client")
    industry = session_info.get("industry", "")
    fy_end = session_info.get("financial_year_end", "")
    currency = session_info.get("currency", "AUD")

    header = (
        f"Client: {client_name}\n"
        f"Industry: {industry}\n"
        f"Financial Year End: {fy_end}\n"
        f"Reporting Currency: {currency}\n"
    )

    sections = [header]
    sections.append(_build_pl_summary(financial_data, period_labels))
    sections.append(_build_bs_summary(financial_data, period_labels))
    cf = _build_cf_summary(financial_data)
    if cf:
        sections.append(cf)
    sections.append(_build_metrics_summary(metrics))
    bm = _build_benchmark_summary(benchmark_comparisons, industry)
    if bm:
        sections.append(bm)
    if red_flags:
        sections.append(
            "RED FLAGS DETECTED:\n"
            + "\n".join(f"  * {f}" for f in red_flags)
        )

    data_summary = "\n\n".join(s for s in sections if s)
    return COMMENTARY_TEMPLATE.format(
        system_prompt=SYSTEM_PROMPT,
        data_summary=data_summary,
    )


# ── Generation functions ───────────────────────────────────────────────────

def generate_commentary(
    prompt: str,
    model: str = DEFAULT_MODEL,
    base_url: str = OLLAMA_BASE_URL,
) -> str:
    """
    Generate commentary using Ollama (blocking, non-streaming).
    Returns the full commentary text string.

    Raises ConnectionError if Ollama is not reachable.
    Raises TimeoutError if the request exceeds 120 seconds.
    """
    payload = {
        "model": model,
        "prompt": prompt,
        "stream": False,
    }
    try:
        resp = requests.post(
            f"{base_url}/api/generate",
            json=payload,
            timeout=120,
        )
        resp.raise_for_status()
        return resp.json()["response"]
    except requests.exceptions.ConnectionError:
        raise ConnectionError(
            "Cannot connect to Ollama. Ensure it is running: ollama serve"
        )
    except requests.exceptions.Timeout:
        raise TimeoutError(
            "Ollama request timed out after 120 s. Try a smaller model."
        )
    except (KeyError, ValueError) as exc:
        raise ValueError(f"Unexpected response format from Ollama: {exc}")


def generate_commentary_streaming(
    prompt: str,
    model: str = DEFAULT_MODEL,
    base_url: str = OLLAMA_BASE_URL,
) -> Iterator[str]:
    """
    Generate commentary using Ollama with streaming enabled.
    Yields text chunks as they are received.

    Raises ConnectionError if Ollama is not reachable.
    Raises TimeoutError if the first response exceeds 120 seconds.
    """
    payload = {
        "model": model,
        "prompt": prompt,
        "stream": True,
    }
    try:
        with requests.post(
            f"{base_url}/api/generate",
            json=payload,
            stream=True,
            timeout=120,
        ) as resp:
            resp.raise_for_status()
            for line in resp.iter_lines():
                if line:
                    try:
                        chunk = json.loads(line)
                        if "response" in chunk:
                            yield chunk["response"]
                        if chunk.get("done"):
                            break
                    except json.JSONDecodeError:
                        continue
    except requests.exceptions.ConnectionError:
        raise ConnectionError(
            "Cannot connect to Ollama. Ensure it is running: ollama serve"
        )
    except requests.exceptions.Timeout:
        raise TimeoutError(
            "Ollama request timed out. Try a smaller model."
        )
