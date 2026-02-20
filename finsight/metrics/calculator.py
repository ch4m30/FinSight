"""
Financial metrics calculator.
Calculates all KPIs, applies traffic-light thresholds, and detects red flags.
"""

from dataclasses import dataclass, field
from typing import Optional
import logging

logger = logging.getLogger(__name__)


# ── Data structures ───────────────────────────────────────────────────────────

@dataclass
class MetricResult:
    """A single calculated metric with status and formatting."""
    name: str
    label: str
    current: Optional[float]
    prior: Optional[float]
    prior2: Optional[float]
    status: str              # 'green', 'amber', 'red', 'grey'
    format_type: str         # 'percentage', 'ratio', 'currency', 'days'
    category: str            # 'liquidity', 'profitability', 'efficiency', 'leverage', 'growth'
    tooltip: str = ""
    trend: str = "–"         # '↑', '↓', '→', '–'
    benchmark_low: Optional[float] = None
    benchmark_high: Optional[float] = None
    benchmark_status: str = "grey"
    notes: str = ""

    def formatted(self, value: Optional[float]) -> str:
        """Return formatted string for display."""
        if value is None:
            return "N/A"
        if self.format_type == "percentage":
            return f"{value:.1f}%"
        elif self.format_type == "ratio":
            return f"{value:.2f}x"
        elif self.format_type == "currency":
            return f"${value:,.0f}"
        elif self.format_type == "days":
            return f"{value:.0f} days"
        return str(value)

    @property
    def current_fmt(self) -> str:
        return self.formatted(self.current)

    @property
    def prior_fmt(self) -> str:
        return self.formatted(self.prior)


@dataclass
class AnalysisResult:
    """Complete analysis result for all periods."""
    metrics: dict[str, MetricResult] = field(default_factory=dict)
    red_flags: list[str] = field(default_factory=list)
    period_labels: list[str] = field(default_factory=list)
    raw_data: dict = field(default_factory=dict)
    benchmark_comparisons: dict = field(default_factory=dict)


# ── Helper functions ──────────────────────────────────────────────────────────

def _safe_div(numerator, denominator) -> Optional[float]:
    """Safe division returning None if denominator is zero or None."""
    if numerator is None or denominator is None:
        return None
    if denominator == 0:
        return None
    return numerator / denominator


def _pct(numerator, denominator) -> Optional[float]:
    """Return value as percentage (0–100 scale)."""
    result = _safe_div(numerator, denominator)
    return result * 100 if result is not None else None


def _trend(current: Optional[float], prior: Optional[float], higher_better: bool = True) -> str:
    """Return trend arrow."""
    if current is None or prior is None:
        return "–"
    if abs(current - prior) < 0.001:
        return "→"
    if current > prior:
        return "↑" if higher_better else "↓"
    return "↓" if higher_better else "↑"


def _traffic_light(value: Optional[float], thresholds: dict) -> str:
    """
    Apply traffic light thresholds.
    thresholds = {'green_min': ..., 'green_max': ..., 'amber_min': ..., 'amber_max': ...}
    or {'green_above': ..., 'amber_above': ...} for simple high-good metrics
    or {'green_below': ..., 'amber_below': ...} for simple low-good metrics
    """
    if value is None:
        return "grey"

    # Simple threshold style: green_above / amber_above
    if "green_above" in thresholds:
        if value >= thresholds["green_above"]:
            return "green"
        elif value >= thresholds["amber_above"]:
            return "amber"
        return "red"

    # Simple threshold style: green_below / amber_below (lower is better)
    if "green_below" in thresholds:
        if value <= thresholds["green_below"]:
            return "green"
        elif value <= thresholds["amber_below"]:
            return "amber"
        return "red"

    # Range style
    if "green_min" in thresholds and "green_max" in thresholds:
        if thresholds["green_min"] <= value <= thresholds["green_max"]:
            return "green"
        elif thresholds.get("amber_min", float("-inf")) <= value <= thresholds.get("amber_max", float("inf")):
            return "amber"
        return "red"

    return "grey"


def _get(data: dict, key: str) -> Optional[float]:
    """Get value from data dict, returning None if missing or zero-ish for non-zero fields."""
    return data.get(key)


# ── Metric calculators ────────────────────────────────────────────────────────

def calculate_liquidity(cur: dict, prior: dict, prior2: dict) -> dict[str, MetricResult]:
    metrics = {}

    # Current Ratio
    cr_cur = _safe_div(_get(cur, "current_assets"), _get(cur, "current_liabilities"))
    cr_pri = _safe_div(_get(prior, "current_assets"), _get(prior, "current_liabilities"))
    cr_p2 = _safe_div(_get(prior2, "current_assets"), _get(prior2, "current_liabilities"))
    metrics["current_ratio"] = MetricResult(
        name="current_ratio",
        label="Current Ratio",
        current=cr_cur, prior=cr_pri, prior2=cr_p2,
        status=_traffic_light(cr_cur, {"green_above": 2.0, "amber_above": 1.0}),
        format_type="ratio",
        category="liquidity",
        trend=_trend(cr_cur, cr_pri),
        tooltip="Current Assets ÷ Current Liabilities. Green ≥2.0, Amber 1.0–1.99, Red <1.0",
    )

    # Quick Ratio
    inv = _get(cur, "inventory") or 0
    inv_p = _get(prior, "inventory") or 0
    inv_p2 = _get(prior2, "inventory") or 0
    qr_cur = _safe_div(
        (_get(cur, "current_assets") or 0) - inv,
        _get(cur, "current_liabilities")
    )
    qr_pri = _safe_div(
        (_get(prior, "current_assets") or 0) - inv_p,
        _get(prior, "current_liabilities")
    )
    qr_p2 = _safe_div(
        (_get(prior2, "current_assets") or 0) - inv_p2,
        _get(prior2, "current_liabilities")
    )
    metrics["quick_ratio"] = MetricResult(
        name="quick_ratio",
        label="Quick Ratio",
        current=qr_cur, prior=qr_pri, prior2=qr_p2,
        status=_traffic_light(qr_cur, {"green_above": 1.0, "amber_above": 0.5}),
        format_type="ratio",
        category="liquidity",
        trend=_trend(qr_cur, qr_pri),
        tooltip="(Current Assets − Inventory) ÷ Current Liabilities. Green ≥1.0, Amber 0.5–0.99",
    )

    # Days Cash on Hand
    opex = _get(cur, "operating_expenses")
    cash_cur = _get(cur, "cash")
    dcoh_cur = _safe_div(cash_cur, _safe_div(opex, 365)) if opex else None

    opex_p = _get(prior, "operating_expenses")
    cash_pri = _get(prior, "cash")
    dcoh_pri = _safe_div(cash_pri, _safe_div(opex_p, 365)) if opex_p else None

    metrics["days_cash_on_hand"] = MetricResult(
        name="days_cash_on_hand",
        label="Days Cash on Hand",
        current=dcoh_cur, prior=dcoh_pri, prior2=None,
        status=_traffic_light(dcoh_cur, {"green_above": 30, "amber_above": 15}),
        format_type="days",
        category="liquidity",
        trend=_trend(dcoh_cur, dcoh_pri),
        tooltip="Cash ÷ (Operating Expenses ÷ 365). Green ≥30 days, Amber 15–29 days",
    )

    return metrics


def calculate_profitability(cur: dict, prior: dict, prior2: dict) -> dict[str, MetricResult]:
    metrics = {}

    # Gross Profit Margin
    gpm_cur = _pct(_get(cur, "gross_profit"), _get(cur, "revenue"))
    gpm_pri = _pct(_get(prior, "gross_profit"), _get(prior, "revenue"))
    gpm_p2 = _pct(_get(prior2, "gross_profit"), _get(prior2, "revenue"))
    metrics["gross_profit_margin"] = MetricResult(
        name="gross_profit_margin",
        label="Gross Profit Margin %",
        current=gpm_cur, prior=gpm_pri, prior2=gpm_p2,
        status="grey",  # Benchmarked against ATO — set externally
        format_type="percentage",
        category="profitability",
        trend=_trend(gpm_cur, gpm_pri),
        tooltip="Gross Profit ÷ Revenue × 100. Benchmarked against ATO industry data.",
    )

    # Net Profit Margin
    npm_cur = _pct(_get(cur, "net_profit"), _get(cur, "revenue"))
    npm_pri = _pct(_get(prior, "net_profit"), _get(prior, "revenue"))
    npm_p2 = _pct(_get(prior2, "net_profit"), _get(prior2, "revenue"))
    metrics["net_profit_margin"] = MetricResult(
        name="net_profit_margin",
        label="Net Profit Margin %",
        current=npm_cur, prior=npm_pri, prior2=npm_p2,
        status="grey",  # Benchmarked against ATO — set externally
        format_type="percentage",
        category="profitability",
        trend=_trend(npm_cur, npm_pri),
        tooltip="Net Profit ÷ Revenue × 100. Benchmarked against ATO industry data.",
    )

    # EBITDA Margin
    ebitda_m_cur = _pct(_get(cur, "ebitda"), _get(cur, "revenue"))
    ebitda_m_pri = _pct(_get(prior, "ebitda"), _get(prior, "revenue"))
    ebitda_m_p2 = _pct(_get(prior2, "ebitda"), _get(prior2, "revenue"))
    metrics["ebitda_margin"] = MetricResult(
        name="ebitda_margin",
        label="EBITDA Margin %",
        current=ebitda_m_cur, prior=ebitda_m_pri, prior2=ebitda_m_p2,
        status=_traffic_light(ebitda_m_cur, {"green_above": 15, "amber_above": 5}),
        format_type="percentage",
        category="profitability",
        trend=_trend(ebitda_m_cur, ebitda_m_pri),
        tooltip="EBITDA ÷ Revenue × 100. Green ≥15%, Amber 5–14.9%, Red <5%",
    )

    # Return on Assets
    roa_cur = _pct(_get(cur, "net_profit"), _get(cur, "total_assets"))
    roa_pri = _pct(_get(prior, "net_profit"), _get(prior, "total_assets"))
    roa_p2 = _pct(_get(prior2, "net_profit"), _get(prior2, "total_assets"))
    metrics["return_on_assets"] = MetricResult(
        name="return_on_assets",
        label="Return on Assets %",
        current=roa_cur, prior=roa_pri, prior2=roa_p2,
        status=_traffic_light(roa_cur, {"green_above": 10, "amber_above": 3}),
        format_type="percentage",
        category="profitability",
        trend=_trend(roa_cur, roa_pri),
        tooltip="Net Profit ÷ Total Assets × 100. Green ≥10%, Amber 3–9.9%, Red <3%",
    )

    # Return on Equity
    roe_cur = _pct(_get(cur, "net_profit"), _get(cur, "equity"))
    roe_pri = _pct(_get(prior, "net_profit"), _get(prior, "equity"))
    roe_p2 = _pct(_get(prior2, "net_profit"), _get(prior2, "equity"))
    metrics["return_on_equity"] = MetricResult(
        name="return_on_equity",
        label="Return on Equity %",
        current=roe_cur, prior=roe_pri, prior2=roe_p2,
        status=_traffic_light(roe_cur, {"green_above": 15, "amber_above": 5}),
        format_type="percentage",
        category="profitability",
        trend=_trend(roe_cur, roe_pri),
        tooltip="Net Profit ÷ Total Equity × 100. Green ≥15%, Amber 5–14.9%, Red <5%",
    )

    return metrics


def calculate_efficiency(cur: dict, prior: dict, prior2: dict) -> dict[str, MetricResult]:
    metrics = {}

    # Debtor Days
    dd_cur = _safe_div(
        (_get(cur, "accounts_receivable") or 0) * 365,
        _get(cur, "revenue")
    )
    dd_pri = _safe_div(
        (_get(prior, "accounts_receivable") or 0) * 365,
        _get(prior, "revenue")
    )
    dd_p2 = _safe_div(
        (_get(prior2, "accounts_receivable") or 0) * 365,
        _get(prior2, "revenue")
    )
    metrics["debtor_days"] = MetricResult(
        name="debtor_days",
        label="Debtor Days",
        current=dd_cur, prior=dd_pri, prior2=dd_p2,
        status=_traffic_light(dd_cur, {"green_below": 30, "amber_below": 60}),
        format_type="days",
        category="efficiency",
        trend=_trend(dd_cur, dd_pri, higher_better=False),
        tooltip="Accounts Receivable ÷ Revenue × 365. Green ≤30 days, Amber 31–60 days, Red >60",
    )

    # Creditor Days
    cogs = _get(cur, "cogs") or _get(cur, "revenue")
    cogs_p = _get(prior, "cogs") or _get(prior, "revenue")
    cogs_p2 = _get(prior2, "cogs") or _get(prior2, "revenue")
    cd_cur = _safe_div((_get(cur, "accounts_payable") or 0) * 365, cogs)
    cd_pri = _safe_div((_get(prior, "accounts_payable") or 0) * 365, cogs_p)
    cd_p2 = _safe_div((_get(prior2, "accounts_payable") or 0) * 365, cogs_p2)
    metrics["creditor_days"] = MetricResult(
        name="creditor_days",
        label="Creditor Days",
        current=cd_cur, prior=cd_pri, prior2=cd_p2,
        status="grey",  # Informational
        format_type="days",
        category="efficiency",
        trend=_trend(cd_cur, cd_pri),
        tooltip="Accounts Payable ÷ COGS × 365. Informational — flag if shorter than Debtor Days.",
        notes="⚠️ Creditor days below debtor days creates working capital pressure." if (
            cd_cur is not None and dd_cur is not None and cd_cur < dd_cur
        ) else "",
    )

    # Inventory Days
    id_cur = _safe_div((_get(cur, "inventory") or 0) * 365, cogs)
    id_pri = _safe_div((_get(prior, "inventory") or 0) * 365, cogs_p)
    id_p2 = _safe_div((_get(prior2, "inventory") or 0) * 365, cogs_p2)
    metrics["inventory_days"] = MetricResult(
        name="inventory_days",
        label="Inventory Days",
        current=id_cur, prior=id_pri, prior2=id_p2,
        status=_traffic_light(id_cur, {"green_below": 45, "amber_below": 90}) if id_cur else "grey",
        format_type="days",
        category="efficiency",
        trend=_trend(id_cur, id_pri, higher_better=False),
        tooltip="Inventory ÷ COGS × 365. Green ≤45 days, Amber 46–90 days, Red >90 days.",
    )

    # Cash Conversion Cycle
    ccc_cur = None
    if dd_cur is not None and cd_cur is not None:
        inv_days = id_cur or 0
        ccc_cur = dd_cur + inv_days - cd_cur
    ccc_pri = None
    if dd_pri is not None and cd_pri is not None:
        inv_days_p = id_pri or 0
        ccc_pri = dd_pri + inv_days_p - cd_pri
    metrics["cash_conversion_cycle"] = MetricResult(
        name="cash_conversion_cycle",
        label="Cash Conversion Cycle",
        current=ccc_cur, prior=ccc_pri, prior2=None,
        status=_traffic_light(ccc_cur, {"green_below": 30, "amber_below": 60}) if ccc_cur is not None else "grey",
        format_type="days",
        category="efficiency",
        trend=_trend(ccc_cur, ccc_pri, higher_better=False),
        tooltip="Debtor Days + Inventory Days − Creditor Days. Lower is better.",
    )

    return metrics


def calculate_leverage(cur: dict, prior: dict, prior2: dict) -> dict[str, MetricResult]:
    metrics = {}

    # Debt to Equity
    dte_cur = _safe_div(_get(cur, "total_liabilities"), _get(cur, "equity"))
    dte_pri = _safe_div(_get(prior, "total_liabilities"), _get(prior, "equity"))
    dte_p2 = _safe_div(_get(prior2, "total_liabilities"), _get(prior2, "equity"))
    metrics["debt_to_equity"] = MetricResult(
        name="debt_to_equity",
        label="Debt-to-Equity Ratio",
        current=dte_cur, prior=dte_pri, prior2=dte_p2,
        status=_traffic_light(dte_cur, {"green_below": 1.0, "amber_below": 2.0}) if dte_cur is not None else "grey",
        format_type="ratio",
        category="leverage",
        trend=_trend(dte_cur, dte_pri, higher_better=False),
        tooltip="Total Liabilities ÷ Total Equity. Green ≤1.0, Amber 1.01–2.0, Red >2.0",
    )

    # Interest Coverage
    ic_cur = _safe_div(_get(cur, "ebit"), _get(cur, "interest_expense"))
    ic_pri = _safe_div(_get(prior, "ebit"), _get(prior, "interest_expense"))
    ic_p2 = _safe_div(_get(prior2, "ebit"), _get(prior2, "interest_expense"))
    metrics["interest_coverage"] = MetricResult(
        name="interest_coverage",
        label="Interest Coverage Ratio",
        current=ic_cur, prior=ic_pri, prior2=ic_p2,
        status=_traffic_light(ic_cur, {"green_above": 3.0, "amber_above": 1.5}) if ic_cur is not None else "grey",
        format_type="ratio",
        category="leverage",
        trend=_trend(ic_cur, ic_pri),
        tooltip="EBIT ÷ Interest Expense. Green ≥3.0x, Amber 1.5–2.99x, Red <1.5x",
    )

    # Net Debt
    nd_cur = None
    if _get(cur, "total_debt") is not None:
        nd_cur = (_get(cur, "total_debt") or 0) - (_get(cur, "cash") or 0)
    elif _get(cur, "total_liabilities") is not None:
        nd_cur = (_get(cur, "total_liabilities") or 0) - (_get(cur, "cash") or 0)
    nd_pri = None
    if _get(prior, "total_debt") is not None:
        nd_pri = (_get(prior, "total_debt") or 0) - (_get(prior, "cash") or 0)
    metrics["net_debt"] = MetricResult(
        name="net_debt",
        label="Net Debt",
        current=nd_cur, prior=nd_pri, prior2=None,
        status="grey",  # Informational
        format_type="currency",
        category="leverage",
        trend=_trend(nd_cur, nd_pri, higher_better=False),
        tooltip="Total Debt − Cash. Informational — shows net borrowing position.",
    )

    return metrics


def calculate_growth(cur: dict, prior: dict, prior2: dict) -> dict[str, MetricResult]:
    """Growth metrics require at least 2 periods."""
    metrics = {}

    if not prior or not any(v is not None for v in prior.values()):
        return metrics

    # Revenue Growth
    rev_growth = _pct(
        (_get(cur, "revenue") or 0) - (_get(prior, "revenue") or 0),
        _get(prior, "revenue")
    )
    metrics["revenue_growth"] = MetricResult(
        name="revenue_growth",
        label="Revenue Growth % YoY",
        current=rev_growth, prior=None, prior2=None,
        status=_traffic_light(rev_growth, {"green_above": 10, "amber_above": 0}) if rev_growth is not None else "grey",
        format_type="percentage",
        category="growth",
        trend="↑" if (rev_growth or 0) > 0 else "↓",
        tooltip="(Current Revenue − Prior Revenue) ÷ Prior Revenue × 100.",
    )

    # Gross Profit Growth
    gp_growth = _pct(
        (_get(cur, "gross_profit") or 0) - (_get(prior, "gross_profit") or 0),
        _get(prior, "gross_profit")
    )
    metrics["gross_profit_growth"] = MetricResult(
        name="gross_profit_growth",
        label="Gross Profit $ Growth % YoY",
        current=gp_growth, prior=None, prior2=None,
        status=_traffic_light(gp_growth, {"green_above": 10, "amber_above": 0}) if gp_growth is not None else "grey",
        format_type="percentage",
        category="growth",
        trend="↑" if (gp_growth or 0) > 0 else "↓",
        tooltip="Year-on-year growth in gross profit dollars.",
    )

    # Expense Growth vs Revenue Growth
    exp_growth = _pct(
        (_get(cur, "operating_expenses") or 0) - (_get(prior, "operating_expenses") or 0),
        _get(prior, "operating_expenses")
    )
    expense_flag = ""
    if exp_growth is not None and rev_growth is not None:
        if exp_growth > rev_growth + 2:
            expense_flag = "⚠️ Expenses growing faster than revenue."
    metrics["expense_growth"] = MetricResult(
        name="expense_growth",
        label="Expense Growth % YoY",
        current=exp_growth, prior=None, prior2=None,
        status="red" if expense_flag else ("green" if (exp_growth or 0) <= (rev_growth or 0) else "amber"),
        format_type="percentage",
        category="growth",
        trend="↑" if (exp_growth or 0) > 0 else "↓",
        tooltip="Operating Expense growth YoY. Flag if growing faster than revenue.",
        notes=expense_flag,
    )

    # Net Profit Growth
    np_growth = _pct(
        (_get(cur, "net_profit") or 0) - (_get(prior, "net_profit") or 0),
        _get(prior, "net_profit")
    )
    metrics["net_profit_growth"] = MetricResult(
        name="net_profit_growth",
        label="Net Profit Growth % YoY",
        current=np_growth, prior=None, prior2=None,
        status=_traffic_light(np_growth, {"green_above": 10, "amber_above": 0}) if np_growth is not None else "grey",
        format_type="percentage",
        category="growth",
        trend="↑" if (np_growth or 0) > 0 else "↓",
        tooltip="Year-on-year growth in net profit.",
    )

    return metrics


def detect_red_flags(cur: dict, prior: dict, prior2: dict, metrics: dict) -> list[str]:
    """Detect and return list of red flag warnings."""
    flags = []

    # Current ratio below 1.0
    cr = metrics.get("current_ratio")
    if cr and cr.current is not None and cr.current < 1.0:
        flags.append(f"⚠️ Current Ratio is {cr.current:.2f}x — below 1.0x signals potential liquidity issues.")

    # Interest coverage below 1.5x
    ic = metrics.get("interest_coverage")
    if ic and ic.current is not None and ic.current < 1.5:
        flags.append(f"⚠️ Interest Coverage is {ic.current:.2f}x — below 1.5x indicates earnings may not cover interest.")

    # Net loss
    net_profit = _get(cur, "net_profit")
    if net_profit is not None and net_profit < 0:
        flags.append(f"⚠️ Net Loss of ${abs(net_profit):,.0f} recorded in current period.")

    # AR growing faster than revenue
    if prior:
        rev_cur = _get(cur, "revenue") or 0
        rev_pri = _get(prior, "revenue") or 1
        ar_cur = _get(cur, "accounts_receivable") or 0
        ar_pri = _get(prior, "accounts_receivable") or 0
        if rev_pri > 0 and ar_pri > 0:
            rev_growth = (rev_cur - rev_pri) / rev_pri
            ar_growth = (ar_cur - ar_pri) / ar_pri
            if ar_growth > rev_growth + 0.05 and ar_growth > 0.05:
                flags.append(
                    f"⚠️ Accounts Receivable grew {ar_growth*100:.1f}% vs Revenue growth of "
                    f"{rev_growth*100:.1f}% — possible collection issues."
                )

    # Inventory growing faster than COGS
    if prior:
        cogs_cur = _get(cur, "cogs") or 0
        cogs_pri = _get(prior, "cogs") or 1
        inv_cur = _get(cur, "inventory") or 0
        inv_pri = _get(prior, "inventory") or 0
        if cogs_pri > 0 and inv_pri > 0:
            cogs_growth = (cogs_cur - cogs_pri) / cogs_pri
            inv_growth = (inv_cur - inv_pri) / inv_pri
            if inv_growth > cogs_growth + 0.05 and inv_growth > 0.05:
                flags.append(
                    f"⚠️ Inventory grew {inv_growth*100:.1f}% vs COGS growth of "
                    f"{cogs_growth*100:.1f}% — possible slow-moving stock."
                )

    # Revenue growth with declining operating cash flow
    if prior:
        rev_cur = _get(cur, "revenue") or 0
        rev_pri = _get(prior, "revenue") or 0
        ocf_cur = _get(cur, "operating_cash_flow")
        ocf_pri = _get(prior, "operating_cash_flow")
        if ocf_cur is not None and ocf_pri is not None:
            if rev_cur > rev_pri and ocf_cur < ocf_pri:
                flags.append(
                    f"⚠️ Revenue growing but Operating Cash Flow declined from "
                    f"${ocf_pri:,.0f} to ${ocf_cur:,.0f} — quality of earnings concern."
                )

    # Expenses growing faster than revenue
    exp_metric = metrics.get("expense_growth")
    rev_metric = metrics.get("revenue_growth")
    if exp_metric and rev_metric:
        if (exp_metric.current or 0) > (rev_metric.current or 0) + 2:
            flags.append(
                f"⚠️ Operating expenses ({exp_metric.current:.1f}% growth) growing faster than "
                f"revenue ({rev_metric.current:.1f}% growth) — margin pressure."
            )

    return flags


def apply_ato_benchmarks(metrics: dict, industry_benchmarks: dict) -> dict:
    """Apply ATO benchmark status to relevant metrics."""
    from benchmarks.ato_fetcher import benchmark_status

    if not industry_benchmarks:
        return metrics

    # Gross profit margin
    gpm = metrics.get("gross_profit_margin")
    if gpm and gpm.current is not None:
        bm = industry_benchmarks.get("gross_profit_margin", {})
        if bm:
            gpm.benchmark_low = bm.get("low")
            gpm.benchmark_high = bm.get("high")
            gpm.status = benchmark_status(gpm.current, bm["low"], bm["high"])
            gpm.benchmark_status = gpm.status

    # Net profit margin
    npm = metrics.get("net_profit_margin")
    if npm and npm.current is not None:
        bm = industry_benchmarks.get("net_profit_margin", {})
        if bm:
            npm.benchmark_low = bm.get("low")
            npm.benchmark_high = bm.get("high")
            npm.status = benchmark_status(npm.current, bm["low"], bm["high"])
            npm.benchmark_status = npm.status

    return metrics


def calculate_benchmark_comparisons(cur: dict, industry_benchmarks: dict) -> dict:
    """
    Calculate ATO benchmark comparisons (as % of revenue).
    Returns dict of comparison data for each benchmark metric.
    """
    comparisons = {}
    revenue = _get(cur, "revenue")
    if not revenue or not industry_benchmarks:
        return comparisons

    benchmark_map = {
        "cost_of_sales": ("cogs", "Cost of Sales"),
        "labour": ("operating_expenses", "Labour / Wages"),  # Approximation
        "rent": (None, "Rent"),
        "motor_vehicle": (None, "Motor Vehicle Expenses"),
    }

    for bm_key, (data_key, label) in benchmark_map.items():
        bm = industry_benchmarks.get(bm_key, {})
        if not bm:
            continue
        actual_val = _get(cur, data_key) if data_key else None
        actual_pct = _pct(actual_val, revenue) if actual_val else None
        comparisons[bm_key] = {
            "label": label,
            "actual_pct": actual_pct,
            "benchmark_low": bm.get("low"),
            "benchmark_high": bm.get("high"),
        }

    return comparisons


def run_analysis(financial_data: dict, industry_benchmarks: dict = None) -> AnalysisResult:
    """
    Main entry point: run full analysis on financial data.
    financial_data: {'data': {'current': {...}, 'prior': {...}}, 'period_labels': [...]}
    """
    data = financial_data.get("data", {})
    cur = data.get("current") or {}
    prior = data.get("prior") or {}
    prior2 = data.get("prior2") or {}

    result = AnalysisResult()
    result.period_labels = financial_data.get("period_labels", ["Current", "Prior"])
    result.raw_data = data

    # Calculate all metric groups
    all_metrics = {}
    all_metrics.update(calculate_liquidity(cur, prior, prior2))
    all_metrics.update(calculate_profitability(cur, prior, prior2))
    all_metrics.update(calculate_efficiency(cur, prior, prior2))
    all_metrics.update(calculate_leverage(cur, prior, prior2))
    all_metrics.update(calculate_growth(cur, prior, prior2))

    # Apply ATO benchmarks
    if industry_benchmarks:
        all_metrics = apply_ato_benchmarks(all_metrics, industry_benchmarks)

    result.metrics = all_metrics
    result.red_flags = detect_red_flags(cur, prior, prior2, all_metrics)
    result.benchmark_comparisons = calculate_benchmark_comparisons(cur, industry_benchmarks or {})

    return result
