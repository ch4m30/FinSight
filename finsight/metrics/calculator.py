"""
Financial metrics calculator.
Calculates all KPIs, applies traffic-light thresholds, and detects red flags.

Bug fixes applied:
- Bug 2: EBIT and EBITDA always calculated from components (never from parsed line)
- Bug 3: Financial integrity self-check layer added
- Bug 5: Inventory in ratio calculations only from balance_sheet source
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
    components: dict = field(default_factory=dict)  # Bug 2: component breakdown

    def formatted(self, value: Optional[float]) -> str:
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
class SelfCheckResult:
    """Result of a single financial integrity self-check."""
    check_name: str
    description: str
    status: str          # 'pass', 'warn', 'fail'
    detail: str
    what_it_means: str
    values: dict = field(default_factory=dict)


@dataclass
class AnalysisResult:
    """Complete analysis result for all periods."""
    metrics: dict[str, MetricResult] = field(default_factory=dict)
    red_flags: list[str] = field(default_factory=list)
    period_labels: list[str] = field(default_factory=list)
    raw_data: dict = field(default_factory=dict)
    benchmark_comparisons: dict = field(default_factory=dict)
    self_checks: list[SelfCheckResult] = field(default_factory=list)  # Bug 3
    has_self_check_fails: bool = False   # Bug 3: True if any check is FAIL
    has_self_check_warns: bool = False   # Bug 3: True if any check is WARN


# ── Helper functions ──────────────────────────────────────────────────────────

def _safe_div(numerator, denominator) -> Optional[float]:
    if numerator is None or denominator is None:
        return None
    if denominator == 0:
        return None
    return numerator / denominator


def _pct(numerator, denominator) -> Optional[float]:
    result = _safe_div(numerator, denominator)
    return result * 100 if result is not None else None


def _trend(current: Optional[float], prior: Optional[float], higher_better: bool = True) -> str:
    if current is None or prior is None:
        return "–"
    if abs(current - prior) < 0.001:
        return "→"
    if current > prior:
        return "↑" if higher_better else "↓"
    return "↓" if higher_better else "↑"


def _traffic_light(value: Optional[float], thresholds: dict) -> str:
    if value is None:
        return "grey"
    if "green_above" in thresholds:
        if value >= thresholds["green_above"]:
            return "green"
        elif value >= thresholds["amber_above"]:
            return "amber"
        return "red"
    if "green_below" in thresholds:
        if value <= thresholds["green_below"]:
            return "green"
        elif value <= thresholds["amber_below"]:
            return "amber"
        return "red"
    if "green_min" in thresholds and "green_max" in thresholds:
        if thresholds["green_min"] <= value <= thresholds["green_max"]:
            return "green"
        elif thresholds.get("amber_min", float("-inf")) <= value <= thresholds.get("amber_max", float("inf")):
            return "amber"
        return "red"
    return "grey"


def _get(data: dict, key: str) -> Optional[float]:
    return data.get(key)


# ── Bug 2: EBIT/EBITDA component calculation ──────────────────────────────────

def _compute_ebit_from_components(period_data: dict) -> tuple:
    """
    Always calculate EBIT from: Net Profit + Income Tax Expense + Interest Expense.
    Never rely on a parsed EBIT line.

    Returns (ebit, ebitda, components_dict, assumption_notes).
    """
    net_profit = period_data.get("net_profit")
    if net_profit is None:
        return None, None, {}, ["Net Profit not available — EBIT cannot be calculated"]

    interest = period_data.get("interest_expense") or 0
    tax = period_data.get("tax_expense") or 0
    dep = period_data.get("depreciation") or 0

    assumption_notes = []
    if interest == 0:
        assumption_notes.append("Interest expense not identified — EBIT = Net Profit + Tax only")
    if tax == 0:
        assumption_notes.append("Tax expense not identified (may be partnership/trust) — EBIT = Net Profit + Interest only")
    if dep == 0:
        assumption_notes.append("D&A not identified — EBITDA may be understated")

    ebit = net_profit + interest + tax
    ebitda = ebit + dep

    # Build component detail for UI display
    int_components = period_data.get("_interest_components", [])
    tax_components = period_data.get("_tax_components", [])
    dep_components = period_data.get("_dep_components", [])

    components = {
        "net_profit": net_profit,
        "interest_expense": interest,
        "interest_items": int_components,
        "tax_expense": tax,
        "tax_items": tax_components,
        "depreciation": dep,
        "dep_items": dep_components,
        "assumption_notes": assumption_notes,
    }

    return ebit, ebitda, components, assumption_notes


# ── Bug 3: Financial integrity self-checks ────────────────────────────────────

def run_self_checks(financial_data: dict) -> list[SelfCheckResult]:
    """
    Run all financial integrity self-checks on parsed data.
    Returns list of SelfCheckResult objects with PASS/WARN/FAIL status.
    """
    checks = []
    data = financial_data.get("data", {})
    cur = data.get("current") or {}
    prior = data.get("prior") or {}

    # ── CHECK 1: P&L Balance ──────────────────────────────────────────────────
    revenue = cur.get("revenue")
    cogs = cur.get("cogs")
    opex = cur.get("operating_expenses")
    net_profit = cur.get("net_profit")
    interest = cur.get("interest_expense") or 0
    tax = cur.get("tax_expense") or 0

    if all(v is not None for v in [revenue, cogs, opex, net_profit]):
        calc_np = revenue - (cogs or 0) - (opex or 0) - interest - tax
        diff = abs(calc_np - net_profit)
        tolerance = max(1.0, abs(net_profit) * 0.001)  # ±$1 or 0.1%
        if diff <= 1:
            status = "pass"
        elif diff <= tolerance:
            status = "warn"
        else:
            status = "fail"
        checks.append(SelfCheckResult(
            check_name="P&L Balance",
            description="Revenue − COGS − Operating Expenses − Interest − Tax = Net Profit",
            status=status,
            detail=(
                f"Calculated Net Profit: ${calc_np:,.0f} | "
                f"Reported Net Profit: ${net_profit:,.0f} | "
                f"Difference: ${diff:,.0f}"
            ),
            what_it_means=(
                "Verifies that P&L components add up to the reported Net Profit. "
                "A FAIL indicates missing line items (e.g. unidentified expenses) or a parsing error."
            ),
            values={"calculated": calc_np, "reported": net_profit, "difference": diff},
        ))
    else:
        missing = [k for k, v in [
            ("Revenue", revenue), ("COGS", cogs),
            ("Operating Expenses", opex), ("Net Profit", net_profit)
        ] if v is None]
        checks.append(SelfCheckResult(
            check_name="P&L Balance",
            description="Revenue − COGS − Operating Expenses − Interest − Tax = Net Profit",
            status="warn",
            detail=f"Cannot perform check — missing: {', '.join(missing)}",
            what_it_means="Requires Revenue, COGS, Operating Expenses, and Net Profit to verify.",
            values={},
        ))

    # ── CHECK 2: Gross Profit Consistency ─────────────────────────────────────
    gross_profit = cur.get("gross_profit")
    if revenue is not None and cogs is not None and gross_profit is not None:
        calc_gp = revenue - cogs
        diff = abs(gross_profit - calc_gp)
        status = "pass" if diff <= 1 else "fail"
        checks.append(SelfCheckResult(
            check_name="Gross Profit Check",
            description="Gross Profit (parsed) = Revenue − COGS",
            status=status,
            detail=(
                f"Parsed GP: ${gross_profit:,.0f} | "
                f"Calculated (Rev−COGS): ${calc_gp:,.0f} | "
                f"Difference: ${diff:,.0f}"
            ),
            what_it_means=(
                "Confirms the Gross Profit figure matches the Revenue minus COGS calculation. "
                "A FAIL may mean the COGS or Revenue figure is a line item instead of a section total."
            ),
            values={"parsed": gross_profit, "calculated": calc_gp, "difference": diff},
        ))
    elif revenue is not None and cogs is not None:
        # No parsed GP — derive it, no check needed
        pass
    else:
        checks.append(SelfCheckResult(
            check_name="Gross Profit Check",
            description="Gross Profit (parsed) = Revenue − COGS",
            status="warn",
            detail="Cannot perform check — Revenue or COGS not available",
            what_it_means="Requires Revenue, COGS, and Gross Profit to verify consistency.",
            values={},
        ))

    # ── CHECK 3: Balance Sheet Equation ───────────────────────────────────────
    total_assets = cur.get("total_assets")
    total_liabilities = cur.get("total_liabilities")
    equity = cur.get("equity")

    if all(v is not None for v in [total_assets, total_liabilities, equity]):
        liab_plus_eq = total_liabilities + equity
        diff = abs(total_assets - liab_plus_eq)
        if diff <= 1:
            status = "pass"
        elif diff <= total_assets * 0.005:
            status = "warn"
        else:
            status = "fail"
        checks.append(SelfCheckResult(
            check_name="Balance Sheet Equation",
            description="Total Assets = Total Liabilities + Equity",
            status=status,
            detail=(
                f"Total Assets: ${total_assets:,.0f} | "
                f"Liabilities + Equity: ${liab_plus_eq:,.0f} | "
                f"Difference: ${diff:,.0f}"
            ),
            what_it_means=(
                "The fundamental accounting equation. If this fails, there may be missing "
                "balance sheet items or a parsing error that will affect all balance sheet ratios."
            ),
            values={"assets": total_assets, "liabilities_plus_equity": liab_plus_eq, "difference": diff},
        ))
    else:
        missing = [k for k, v in [
            ("Total Assets", total_assets),
            ("Total Liabilities", total_liabilities),
            ("Equity", equity)
        ] if v is None]
        checks.append(SelfCheckResult(
            check_name="Balance Sheet Equation",
            description="Total Assets = Total Liabilities + Equity",
            status="warn",
            detail=f"Cannot perform check — missing: {', '.join(missing)}",
            what_it_means="Requires Total Assets, Total Liabilities, and Equity.",
            values={},
        ))

    # ── CHECK 4: Equity Movement (prior year required, WARN only) ─────────────
    if prior:
        prior_equity = prior.get("equity")
        cur_equity = cur.get("equity")
        cur_np = cur.get("net_profit")

        if all(v is not None for v in [prior_equity, cur_equity, cur_np]):
            expected_eq = prior_equity + cur_np
            diff = abs(cur_equity - expected_eq)
            # Always WARN — capital movements are a normal explanation
            status = "warn" if diff > 1 else "pass"
            checks.append(SelfCheckResult(
                check_name="Equity Movement",
                description="Closing Equity ≈ Opening Equity + Net Profit (±capital movements)",
                status=status,
                detail=(
                    f"Closing Equity: ${cur_equity:,.0f} | "
                    f"Opening + Net Profit: ${expected_eq:,.0f} | "
                    f"Unexplained movement: ${diff:,.0f}"
                ),
                what_it_means=(
                    "Checks equity movement. Differences are normal if dividends, drawings, "
                    "or capital contributions occurred during the year. This is always WARN at most — "
                    "review if the unexplained movement is unexpected."
                ),
                values={"closing": cur_equity, "expected": expected_eq, "difference": diff},
            ))
        else:
            checks.append(SelfCheckResult(
                check_name="Equity Movement",
                description="Closing Equity ≈ Opening Equity + Net Profit (±capital movements)",
                status="warn",
                detail="Prior year equity or current net profit not available",
                what_it_means="Requires current and prior year equity, and current net profit.",
                values={},
            ))
    # (Skip check 4 entirely if no prior data — don't add a warn for missing data)

    # ── CHECK 5: Current Assets/Liabilities Subtotals ─────────────────────────
    current_assets = cur.get("current_assets")
    cash = cur.get("cash") or 0
    ar = cur.get("accounts_receivable") or 0
    inv = cur.get("inventory") or 0

    if current_assets is not None and (cash or ar or inv):
        component_sum = cash + ar + inv
        # Components might be a subset (other CA items not captured) — only warn if > total
        if component_sum > current_assets + 1:
            diff = component_sum - current_assets
            checks.append(SelfCheckResult(
                check_name="Current Assets Subtotal",
                description="Sum of identified current assets ≤ Total Current Assets",
                status="warn",
                detail=(
                    f"Cash+AR+Inventory: ${component_sum:,.0f} | "
                    f"Total Current Assets: ${current_assets:,.0f} | "
                    f"Excess: ${diff:,.0f}"
                ),
                what_it_means=(
                    "The sum of Cash, AR, and Inventory exceeds Total Current Assets — "
                    "possible parsing error where a line item total was picked up instead of a subtotal."
                ),
                values={"components": component_sum, "total": current_assets, "excess": diff},
            ))
        else:
            checks.append(SelfCheckResult(
                check_name="Current Assets Subtotal",
                description="Sum of identified current assets ≤ Total Current Assets",
                status="pass",
                detail=(
                    f"Cash+AR+Inventory: ${component_sum:,.0f} | "
                    f"Total Current Assets: ${current_assets:,.0f}"
                ),
                what_it_means=(
                    "The identified current assets are consistent with the total. "
                    "The difference (if any) represents other current assets not individually captured."
                ),
                values={"components": component_sum, "total": current_assets},
            ))

    # ── CHECK 6: Revenue Reasonableness (YoY > 50%) ───────────────────────────
    if prior:
        prior_revenue = prior.get("revenue")
        cur_revenue = cur.get("revenue")
        if prior_revenue and cur_revenue and prior_revenue > 0:
            pct_change = (cur_revenue - prior_revenue) / prior_revenue * 100
            abs_change = abs(pct_change)
            if abs_change > 50:
                checks.append(SelfCheckResult(
                    check_name="Revenue Reasonableness",
                    description=f"Revenue changed {pct_change:+.1f}% YoY — verify this is correct",
                    status="warn",
                    detail=(
                        f"Current: ${cur_revenue:,.0f} | "
                        f"Prior: ${prior_revenue:,.0f} | "
                        f"Change: {pct_change:+.1f}%"
                    ),
                    what_it_means=(
                        "Revenue has changed by more than 50% year-on-year. "
                        "This could be a genuine business event (major client win/loss, COVID impact) "
                        "or a parsing error where the wrong column was used."
                    ),
                    values={"current": cur_revenue, "prior": prior_revenue, "pct_change": pct_change},
                ))
            else:
                checks.append(SelfCheckResult(
                    check_name="Revenue Reasonableness",
                    description="Revenue change within normal range (<50% YoY)",
                    status="pass",
                    detail=(
                        f"Current: ${cur_revenue:,.0f} | "
                        f"Prior: ${prior_revenue:,.0f} | "
                        f"Change: {pct_change:+.1f}%"
                    ),
                    what_it_means="Year-on-year revenue movement is within expected bounds.",
                    values={"current": cur_revenue, "prior": prior_revenue, "pct_change": pct_change},
                ))

    return checks


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

    # Quick Ratio — Bug 5: inventory already sourced from BS current assets only
    inv = _get(cur, "inventory") or 0
    inv_p = _get(prior, "inventory") or 0
    inv_p2 = _get(prior2, "inventory") or 0

    # Note if inventory was not found on BS
    inv_note = ""
    inv_source = cur.get("_inventory_source", "")
    if not inv_source or inv_source == "not_found":
        inv_note = "Inventory not identified on Balance Sheet — Quick Ratio equals Current Ratio. Inventory Days cannot be calculated."

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
        notes=inv_note,
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

    # ── Bug 2: Use component-computed EBIT/EBITDA ─────────────────────────────
    ebit_cur = cur.get("_ebit_computed")
    ebit_pri = prior.get("_ebit_computed") if prior else None
    ebit_p2 = prior2.get("_ebit_computed") if prior2 else None

    ebitda_cur = cur.get("_ebitda_computed")
    ebitda_pri = prior.get("_ebitda_computed") if prior else None
    ebitda_p2 = prior2.get("_ebitda_computed") if prior2 else None

    ebit_components_cur = cur.get("_ebit_components", {})

    # ── Gross Profit Margin ───────────────────────────────────────────────────
    gpm_cur = _pct(_get(cur, "gross_profit"), _get(cur, "revenue"))
    gpm_pri = _pct(_get(prior, "gross_profit"), _get(prior, "revenue"))
    gpm_p2 = _pct(_get(prior2, "gross_profit"), _get(prior2, "revenue"))
    metrics["gross_profit_margin"] = MetricResult(
        name="gross_profit_margin",
        label="Gross Profit Margin %",
        current=gpm_cur, prior=gpm_pri, prior2=gpm_p2,
        status="grey",
        format_type="percentage",
        category="profitability",
        trend=_trend(gpm_cur, gpm_pri),
        tooltip="Gross Profit ÷ Revenue × 100. Benchmarked against ATO industry data.",
    )

    # ── Net Profit Margin ─────────────────────────────────────────────────────
    npm_cur = _pct(_get(cur, "net_profit"), _get(cur, "revenue"))
    npm_pri = _pct(_get(prior, "net_profit"), _get(prior, "revenue"))
    npm_p2 = _pct(_get(prior2, "net_profit"), _get(prior2, "revenue"))
    metrics["net_profit_margin"] = MetricResult(
        name="net_profit_margin",
        label="Net Profit Margin %",
        current=npm_cur, prior=npm_pri, prior2=npm_p2,
        status="grey",
        format_type="percentage",
        category="profitability",
        trend=_trend(npm_cur, npm_pri),
        tooltip="Net Profit ÷ Revenue × 100. Benchmarked against ATO industry data.",
    )

    # ── EBIT Margin (Bug 2: from component-computed EBIT) ─────────────────────
    ebit_m_cur = _pct(ebit_cur, _get(cur, "revenue"))
    ebit_m_pri = _pct(ebit_pri, _get(prior, "revenue"))
    ebit_m_p2 = _pct(ebit_p2, _get(prior2, "revenue"))

    # Build tooltip with component breakdown
    ebit_tooltip = "EBIT ÷ Revenue × 100. EBIT = Net Profit + Tax + Interest (always calculated from components)."
    if ebit_components_cur:
        notes_list = ebit_components_cur.get("assumption_notes", [])
        if notes_list:
            ebit_tooltip += " Note: " + " | ".join(notes_list)

    metrics["ebit_margin"] = MetricResult(
        name="ebit_margin",
        label="EBIT Margin %",
        current=ebit_m_cur, prior=ebit_m_pri, prior2=ebit_m_p2,
        status=_traffic_light(ebit_m_cur, {"green_above": 10, "amber_above": 3}),
        format_type="percentage",
        category="profitability",
        trend=_trend(ebit_m_cur, ebit_m_pri),
        tooltip=ebit_tooltip,
        components=ebit_components_cur,
        notes="; ".join(ebit_components_cur.get("assumption_notes", [])),
    )

    # ── EBITDA Margin (Bug 2: from component-computed EBITDA) ─────────────────
    ebitda_m_cur = _pct(ebitda_cur, _get(cur, "revenue"))
    ebitda_m_pri = _pct(ebitda_pri, _get(prior, "revenue"))
    ebitda_m_p2 = _pct(ebitda_p2, _get(prior2, "revenue"))

    dep_note = ""
    if ebit_components_cur and (ebit_components_cur.get("depreciation") or 0) == 0:
        dep_note = "D&A not identified — EBITDA may be understated"

    metrics["ebitda_margin"] = MetricResult(
        name="ebitda_margin",
        label="EBITDA Margin %",
        current=ebitda_m_cur, prior=ebitda_m_pri, prior2=ebitda_m_p2,
        status=_traffic_light(ebitda_m_cur, {"green_above": 15, "amber_above": 5}),
        format_type="percentage",
        category="profitability",
        trend=_trend(ebitda_m_cur, ebitda_m_pri),
        tooltip="EBITDA ÷ Revenue × 100. Green ≥15%, Amber 5–14.9%, Red <5%. EBITDA = EBIT + D&A.",
        notes=dep_note,
        components=ebit_components_cur,
    )

    # ── Return on Assets ──────────────────────────────────────────────────────
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

    # ── Return on Equity ──────────────────────────────────────────────────────
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
    dd_cur = _safe_div((_get(cur, "accounts_receivable") or 0) * 365, _get(cur, "revenue"))
    dd_pri = _safe_div((_get(prior, "accounts_receivable") or 0) * 365, _get(prior, "revenue"))
    dd_p2 = _safe_div((_get(prior2, "accounts_receivable") or 0) * 365, _get(prior2, "revenue"))
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
        status="grey",
        format_type="days",
        category="efficiency",
        trend=_trend(cd_cur, cd_pri),
        tooltip="Accounts Payable ÷ COGS × 365. Informational — flag if shorter than Debtor Days.",
        notes="⚠️ Creditor days below debtor days creates working capital pressure." if (
            cd_cur is not None and dd_cur is not None and cd_cur < dd_cur
        ) else "",
    )

    # Inventory Days — Bug 5: inventory already sourced from BS only
    inv_source = cur.get("_inventory_source", "")
    inv_note = ""
    if not inv_source or inv_source == "not_found":
        inv_note = "Inventory not identified on Balance Sheet — Inventory Days cannot be calculated."

    id_cur = _safe_div((_get(cur, "inventory") or 0) * 365, cogs) if not inv_note else None
    id_pri = _safe_div((_get(prior, "inventory") or 0) * 365, cogs_p) if not inv_note else None
    id_p2 = _safe_div((_get(prior2, "inventory") or 0) * 365, cogs_p2) if not inv_note else None
    metrics["inventory_days"] = MetricResult(
        name="inventory_days",
        label="Inventory Days",
        current=id_cur, prior=id_pri, prior2=id_p2,
        status=_traffic_light(id_cur, {"green_below": 45, "amber_below": 90}) if id_cur else "grey",
        format_type="days",
        category="efficiency",
        trend=_trend(id_cur, id_pri, higher_better=False),
        tooltip="Inventory ÷ COGS × 365 (inventory from Balance Sheet only). Green ≤45, Amber 46–90.",
        notes=inv_note,
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

    # Interest Coverage — Bug 2: use component-computed EBIT
    ebit_cur = cur.get("_ebit_computed") or _get(cur, "ebit")
    ebit_pri = (prior.get("_ebit_computed") if prior else None) or _get(prior, "ebit")
    ebit_p2 = (prior2.get("_ebit_computed") if prior2 else None) or _get(prior2, "ebit")

    ic_cur = _safe_div(ebit_cur, _get(cur, "interest_expense"))
    ic_pri = _safe_div(ebit_pri, _get(prior, "interest_expense"))
    ic_p2 = _safe_div(ebit_p2, _get(prior2, "interest_expense"))
    metrics["interest_coverage"] = MetricResult(
        name="interest_coverage",
        label="Interest Coverage Ratio",
        current=ic_cur, prior=ic_pri, prior2=ic_p2,
        status=_traffic_light(ic_cur, {"green_above": 3.0, "amber_above": 1.5}) if ic_cur is not None else "grey",
        format_type="ratio",
        category="leverage",
        trend=_trend(ic_cur, ic_pri),
        tooltip=(
            "EBIT ÷ Interest Expense. Green ≥3.0x, Amber 1.5–2.99x, Red <1.5x. "
            "EBIT calculated from Net Profit + Tax + Interest."
        ),
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
        status="grey",
        format_type="currency",
        category="leverage",
        trend=_trend(nd_cur, nd_pri, higher_better=False),
        tooltip="Total Debt − Cash. Informational — shows net borrowing position.",
    )

    return metrics


def calculate_growth(cur: dict, prior: dict, prior2: dict) -> dict[str, MetricResult]:
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

    # Expense Growth
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
    flags = []

    cr = metrics.get("current_ratio")
    if cr and cr.current is not None and cr.current < 1.0:
        flags.append(f"⚠️ Current Ratio is {cr.current:.2f}x — below 1.0x signals potential liquidity issues.")

    ic = metrics.get("interest_coverage")
    if ic and ic.current is not None and ic.current < 1.5:
        flags.append(f"⚠️ Interest Coverage is {ic.current:.2f}x — below 1.5x indicates earnings may not cover interest.")

    net_profit = _get(cur, "net_profit")
    if net_profit is not None and net_profit < 0:
        flags.append(f"⚠️ Net Loss of ${abs(net_profit):,.0f} recorded in current period.")

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
    from benchmarks.ato_fetcher import benchmark_status

    if not industry_benchmarks:
        return metrics

    gpm = metrics.get("gross_profit_margin")
    if gpm and gpm.current is not None:
        bm = industry_benchmarks.get("gross_profit_margin", {})
        if bm:
            gpm.benchmark_low = bm.get("low")
            gpm.benchmark_high = bm.get("high")
            gpm.status = benchmark_status(gpm.current, bm["low"], bm["high"])
            gpm.benchmark_status = gpm.status

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
    comparisons = {}
    revenue = _get(cur, "revenue")
    if not revenue or not industry_benchmarks:
        return comparisons

    benchmark_map = {
        "cost_of_sales": ("cogs", "Cost of Sales"),
        "labour": ("operating_expenses", "Labour / Wages"),
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

    # ── Bug 2: Pre-compute component-based EBIT/EBITDA for all periods ────────
    for period_data in [cur, prior, prior2]:
        if not period_data:
            continue
        ebit, ebitda, components, _ = _compute_ebit_from_components(period_data)
        if ebit is not None:
            period_data["_ebit_computed"] = ebit
            period_data["_ebitda_computed"] = ebitda
            period_data["_ebit_components"] = components

    result = AnalysisResult()
    result.period_labels = financial_data.get("period_labels", ["Current", "Prior"])
    result.raw_data = data

    # ── Calculate all metric groups ───────────────────────────────────────────
    all_metrics = {}
    all_metrics.update(calculate_liquidity(cur, prior, prior2))
    all_metrics.update(calculate_profitability(cur, prior, prior2))
    all_metrics.update(calculate_efficiency(cur, prior, prior2))
    all_metrics.update(calculate_leverage(cur, prior, prior2))
    all_metrics.update(calculate_growth(cur, prior, prior2))

    if industry_benchmarks:
        all_metrics = apply_ato_benchmarks(all_metrics, industry_benchmarks)

    result.metrics = all_metrics
    result.red_flags = detect_red_flags(cur, prior, prior2, all_metrics)
    result.benchmark_comparisons = calculate_benchmark_comparisons(cur, industry_benchmarks or {})

    # ── Bug 3: Run financial integrity self-checks ────────────────────────────
    self_checks = run_self_checks(financial_data)
    result.self_checks = self_checks
    result.has_self_check_fails = any(c.status == "fail" for c in self_checks)
    result.has_self_check_warns = any(c.status == "warn" for c in self_checks)

    return result
