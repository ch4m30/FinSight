"""
Xero CSV/Excel financial statement parser.
Handles Xero P&L, Balance Sheet, and Cash Flow exports.

Bug fixes applied:
- Bug 1: Hierarchical parsing to capture subtotals/totals for Revenue and COGS
- Bug 2: Expanded keyword matching and sum-all-matching for EBIT components
- Bug 4: Note reference column detection and exclusion
- Bug 5: Balance sheet section tracking so inventory is only sourced from Current Assets
"""

import pandas as pd
import numpy as np
import re
import logging
from io import BytesIO

logger = logging.getLogger(__name__)

# ── Subtotal/total row detection ──────────────────────────────────────────────
SUBTOTAL_KEYWORDS = [
    "total", "subtotal", "gross profit", "net profit", "net loss", "net income",
    "net revenue", "net sales",
]

# ── Revenue keywords ──────────────────────────────────────────────────────────
REVENUE_KEYWORDS = [
    "revenue", "income", "sales", "turnover", "fees", "service income",
    "trading income", "gross receipts", "total income", "total revenue",
    "grant income", "other income",
]
# Explicit total-row patterns for revenue (checked first — Bug 1)
REVENUE_TOTAL_KEYWORDS = [
    "total revenue", "total income", "total sales", "total trading income",
    "total fees", "total service income", "total turnover", "gross income",
    "total receipts", "total gross receipts",
]

# ── COGS keywords ─────────────────────────────────────────────────────────────
COGS_KEYWORDS = [
    "cost of sales", "cost of goods", "cogs", "direct costs", "direct expenses",
    "purchases", "cost of revenue", "materials", "subcontractors",
    "direct labour", "direct wages", "opening stock", "closing stock",
    "freight in", "freight-in",
]
# Explicit total-row patterns for COGS (checked first — Bug 1)
COGS_TOTAL_KEYWORDS = [
    "total cost of sales", "total cost of goods", "total cogs",
    "total direct costs", "total direct expenses", "total purchases",
    "total cost of revenue", "total materials",
]

GROSS_PROFIT_KEYWORDS = ["gross profit", "gross margin"]
OPERATING_EXPENSES_KEYWORDS = [
    "operating expenses", "expenses", "overheads", "administrative", "general expenses",
    "selling expenses", "total expenses", "total overheads",
]
EBIT_KEYWORDS = ["ebit", "operating profit", "profit from operations", "operating income"]
EBITDA_KEYWORDS = ["ebitda"]
NET_PROFIT_KEYWORDS = [
    "net profit", "net income", "profit after tax", "profit before tax",
    "net loss", "net earnings", "profit for the year", "surplus",
]

# ── Bug 2: Expanded component keywords ───────────────────────────────────────
DEPRECIATION_KEYWORDS = [
    "depreciation", "amortisation", "amortization", "dep &",
    "depreciation and amortisation", "depreciation and amortization",
    "d&a", "right of use", "rou asset depreciation", "right-of-use",
    "amortisation of intangibles", "amortization of intangibles",
]
INTEREST_KEYWORDS = [
    "interest expense", "finance charge", "bank charge", "loan interest",
    "interest on loan", "borrowing cost", "finance costs", "interest paid",
    "bank interest", "interest on overdraft", "hire purchase interest",
]
TAX_KEYWORDS = [
    "income tax", "tax expense", "taxation", "company tax",
    "income tax expense", "provision for tax", "corporate tax", "tax payable",
    "fringe benefits tax", "fbt",
]

# ── Balance sheet keywords ────────────────────────────────────────────────────
CURRENT_ASSETS_KEYWORDS = ["current assets", "total current assets"]
CASH_KEYWORDS = ["cash", "bank", "cash and cash equivalents", "cash at bank"]
RECEIVABLES_KEYWORDS = [
    "accounts receivable", "debtors", "trade receivables", "receivables",
    "trade debtors", "sundry debtors",
]
# Bug 5: Inventory keyword list — only used when in current_assets section of BS
INVENTORY_KEYWORDS = [
    "inventory", "stock on hand", "closing stock", "finished goods",
    "raw materials", "work in progress", "wip", "trading stock", "stock",
]
OTHER_CURRENT_ASSETS_KEYWORDS = ["other current assets", "prepayments", "accrued income"]
NON_CURRENT_ASSETS_KEYWORDS = [
    "non-current assets", "fixed assets", "total non-current assets",
    "plant and equipment", "property plant",
]
TOTAL_ASSETS_KEYWORDS = ["total assets"]
CURRENT_LIABILITIES_KEYWORDS = ["current liabilities", "total current liabilities"]
PAYABLES_KEYWORDS = [
    "accounts payable", "creditors", "trade payables", "trade creditors", "sundry creditors",
]
OTHER_CURRENT_LIAB_KEYWORDS = ["other current liabilities", "accrued liabilities", "gst payable"]
NON_CURRENT_LIABILITIES_KEYWORDS = [
    "non-current liabilities", "long-term liabilities", "total non-current liabilities",
]
TOTAL_LIABILITIES_KEYWORDS = ["total liabilities"]
EQUITY_KEYWORDS = ["equity", "total equity", "shareholders equity", "net assets", "owners equity"]
DEBT_KEYWORDS = ["loans", "borrowings", "bank loan", "term loan", "line of credit", "overdraft"]

# Cash flow
OPERATING_CF_KEYWORDS = [
    "cash from operations", "operating cash flow", "net cash from operating",
    "cash generated from operations", "net cash provided by operating",
]
INVESTING_CF_KEYWORDS = [
    "cash from investing", "investing activities", "net cash from investing",
]
FINANCING_CF_KEYWORDS = [
    "cash from financing", "financing activities", "net cash from financing",
]

# P&L section header detection (labels that introduce a section, usually without a value)
PL_SECTION_HEADERS = {
    "revenue": ["revenue", "income", "sales", "trading income"],
    "cogs": ["cost of sales", "cost of goods", "direct costs", "purchases"],
    "operating_expenses": ["operating expenses", "expenses", "overheads", "administration"],
}


# ── Period column detection ───────────────────────────────────────────────────

_RELATIVE_PRIORITY = {
    "current year": 100, "current": 100, "this year": 100,
    "prior year": 50, "prior": 50, "previous year": 50,
    "comparative": 50, "last year": 50,
}


def _extract_year_from_col(text: str) -> int | None:
    s = str(text).strip()
    m = re.search(r'\bFY\s*(\d{4})\b', s, re.IGNORECASE)
    if m:
        return int(m.group(1))
    m = re.search(r'\bFY\s*(\d{2})\b', s, re.IGNORECASE)
    if m:
        return 2000 + int(m.group(1))
    m = re.search(r'year\s+ended.*?(20\d{2})', s, re.IGNORECASE)
    if m:
        return int(m.group(1))
    m = re.search(r'\b(20\d{2})/(\d{2,4})\b', s)
    if m:
        suffix = m.group(2)
        if len(suffix) == 2:
            return int(m.group(1)[:2] + suffix)
        return int(suffix)
    m = re.search(r'[A-Za-z]{3}\s+\d{4}\s*[-\u2013]\s*[A-Za-z]{3}\s+(20\d{2})', s)
    if m:
        return int(m.group(1))
    m = re.search(r'\b(20\d{2})\b', s)
    if m:
        return int(m.group(1))
    return None


def _detect_and_sort_periods(df: pd.DataFrame) -> tuple:
    label_col = df.columns[0]
    value_cols = [c for c in df.columns if c != label_col]
    if not value_cols:
        return df, [], False

    col_years = {}
    for col in value_cols:
        yr = _extract_year_from_col(str(col))
        if yr is not None:
            col_years[col] = yr

    if col_years:
        sorted_cols = sorted(value_cols, key=lambda c: col_years.get(c, 0), reverse=True)
        used_fallback = len(col_years) < len(value_cols)
        df_sorted = df[[label_col] + sorted_cols]
        return df_sorted, [str(c).strip() for c in sorted_cols[:3]], used_fallback

    col_rel = {}
    for col in value_cols:
        col_lower = str(col).lower().strip()
        for term, priority in _RELATIVE_PRIORITY.items():
            if term in col_lower:
                col_rel[col] = priority
                break

    if col_rel:
        sorted_cols = sorted(value_cols, key=lambda c: col_rel.get(c, 0), reverse=True)
        df_sorted = df[[label_col] + sorted_cols]
        return df_sorted, [str(c).strip() for c in sorted_cols[:3]], False

    return df, [str(c).strip() for c in value_cols[:3]], True


def _clean_amount(val) -> float | None:
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val) if not np.isnan(float(val)) else None
    s = str(val).strip()
    if not s or s in ("-", "", "n/a", "N/A", "—"):
        return None
    negative = s.startswith("(") and s.endswith(")")
    s = s.replace("(", "").replace(")", "")
    s = s.replace("$", "").replace(",", "").replace(" ", "")
    try:
        result = float(s)
        return -result if negative else result
    except ValueError:
        return None


def _matches_keywords(text: str, keywords: list) -> bool:
    t = str(text).lower().strip()
    return any(kw.lower() in t for kw in keywords)


def _is_subtotal_row(label: str) -> bool:
    """Return True if the label indicates a subtotal, total, or net row."""
    label_lower = str(label).lower().strip()
    return any(kw in label_lower for kw in SUBTOTAL_KEYWORDS)


# ── Bug 4: Note reference column detection ───────────────────────────────────

def _detect_reference_columns(df: pd.DataFrame) -> set:
    """
    Classify numeric columns as reference (note index) or value columns.
    A reference column contains only small integers in range 1–50 with no decimals.
    Returns a set of column names classified as reference columns.
    """
    label_col = df.columns[0]
    ref_cols = set()

    for col in df.columns[1:]:
        numeric_vals = []
        has_decimal = False
        for v in df[col]:
            c = _clean_amount(v)
            if c is not None:
                numeric_vals.append(c)
                if c != int(c):
                    has_decimal = True

        if not numeric_vals or has_decimal:
            continue

        # Check if all values are small integers in 1–50
        all_small_int = all(1 <= v <= 50 and v == int(v) for v in numeric_vals)
        median_abs = float(np.median([abs(v) for v in numeric_vals]))

        if all_small_int and median_abs <= 50:
            ref_cols.add(col)
            logger.info(f"Column '{col}' detected as note reference column — excluded from extraction")

    return ref_cols


# ── Value search helpers (Bug 1, Bug 2) ──────────────────────────────────────

def _find_value(df: pd.DataFrame, keywords: list, col_idx: int = 0,
                skip_cols: set = None) -> float | None:
    """
    Search df for the first row matching keywords and return the value.
    Prefers subtotal rows over regular rows.
    """
    label_col = df.columns[0]
    value_cols = [c for c in df.columns if c != label_col]
    if skip_cols:
        value_cols = [c for c in value_cols if c not in skip_cols]
    if not value_cols or col_idx >= len(value_cols):
        return None

    vcol = value_cols[col_idx]
    first_match = None
    first_subtotal = None

    for _, row in df.iterrows():
        label = str(row[label_col]).strip()
        if _matches_keywords(label, keywords):
            val = _clean_amount(row[vcol])
            if val is not None:
                if first_match is None:
                    first_match = val
                if _is_subtotal_row(label) and first_subtotal is None:
                    first_subtotal = val

    # Prefer subtotal rows (Bug 1)
    return first_subtotal if first_subtotal is not None else first_match


def _find_value_prefer_subtotal(df: pd.DataFrame, total_keywords: list,
                                fallback_keywords: list, col_idx: int = 0,
                                skip_cols: set = None) -> float | None:
    """
    Bug 1: Find value by first trying explicit total keywords, then fallback keywords.
    Also prefers subtotal-flagged rows within each search.
    """
    # 1. Try explicit total keywords first (e.g. "Total Revenue")
    result = _find_value(df, total_keywords, col_idx, skip_cols)
    if result is not None:
        return result
    # 2. Fall back to general section keywords (prefers subtotal rows internally)
    return _find_value(df, fallback_keywords, col_idx, skip_cols)


def _sum_all_matching(df: pd.DataFrame, keywords: list, col_idx: int = 0,
                      skip_cols: set = None) -> tuple:
    """
    Bug 2: Sum ALL non-subtotal rows matching keywords.
    Returns (total_or_None, list_of_(label, value)_components).
    Deduplicates by label to avoid double-counting.
    """
    label_col = df.columns[0]
    value_cols = [c for c in df.columns if c != label_col]
    if skip_cols:
        value_cols = [c for c in value_cols if c not in skip_cols]
    if not value_cols or col_idx >= len(value_cols):
        return None, []

    vcol = value_cols[col_idx]
    total = 0.0
    components = []
    seen_labels: set = set()

    # First check if there's an explicit subtotal for this component group
    for _, row in df.iterrows():
        label = str(row[label_col]).strip()
        if _matches_keywords(label, keywords) and _is_subtotal_row(label):
            val = _clean_amount(row[vcol])
            if val is not None:
                return val, [(label, val)]

    # No subtotal found — sum all matching line items
    for _, row in df.iterrows():
        label = str(row[label_col]).strip()
        label_lower = label.lower()
        if _is_subtotal_row(label):
            continue
        if _matches_keywords(label, keywords) and label_lower not in seen_labels:
            val = _clean_amount(row[vcol])
            if val is not None:
                total += val
                components.append((label, val))
                seen_labels.add(label_lower)

    return (total if components else None), components


def _sum_section_lines(df: pd.DataFrame, section_header_keywords: list,
                       col_idx: int = 0, skip_cols: set = None) -> float | None:
    """
    Sum all non-subtotal line items that appear between a section header
    (a row matching section_header_keywords with no value) and the next
    subtotal or major section boundary.
    Used as fallback when no explicit total row is found.
    """
    label_col = df.columns[0]
    value_cols = [c for c in df.columns if c != label_col]
    if skip_cols:
        value_cols = [c for c in value_cols if c not in skip_cols]
    if not value_cols or col_idx >= len(value_cols):
        return None

    vcol = value_cols[col_idx]
    in_section = False
    section_items = []

    for _, row in df.iterrows():
        label = str(row[label_col]).strip()
        val = _clean_amount(row[vcol])

        if not in_section:
            # Section header: matches section keywords AND has no numeric value
            if _matches_keywords(label, section_header_keywords) and val is None:
                in_section = True
        else:
            if val is None:
                continue
            # Subtotal row for this section → use it directly if it matches
            if (_is_subtotal_row(label) and
                    _matches_keywords(label, section_header_keywords)):
                return val
            # Any subtotal row stops section scanning
            if _is_subtotal_row(label):
                break
            section_items.append(val)

    return sum(section_items) if section_items else None


# ── P&L extraction ────────────────────────────────────────────────────────────

def _extract_period_data(df: pd.DataFrame, col_idx: int = 0,
                         skip_cols: set = None) -> dict:
    """
    Extract all P&L line items for a given period column index.
    Bug 1: Uses hierarchical subtotal preference for Revenue and COGS.
    Bug 2: Sums ALL matching lines for interest, tax, depreciation.
    Stores line_items list and ebit_components dict for downstream use.
    """
    data = {}

    def g(keywords):
        return _find_value(df, keywords, col_idx, skip_cols)

    # ── Revenue: prefer explicit total row (Bug 1) ────────────────────────────
    revenue = _find_value_prefer_subtotal(
        df, REVENUE_TOTAL_KEYWORDS, REVENUE_KEYWORDS, col_idx, skip_cols
    )
    if revenue is None:
        # Final fallback: sum section lines
        revenue = _sum_section_lines(df, ["revenue", "income", "sales"], col_idx, skip_cols)
    data["revenue"] = revenue

    # ── COGS: prefer explicit total row (Bug 1) ───────────────────────────────
    cogs = _find_value_prefer_subtotal(
        df, COGS_TOTAL_KEYWORDS, COGS_KEYWORDS, col_idx, skip_cols
    )
    if cogs is None:
        cogs = _sum_section_lines(df, ["cost of sales", "cost of goods", "direct costs"],
                                  col_idx, skip_cols)
    data["cogs"] = cogs

    data["gross_profit"] = g(GROSS_PROFIT_KEYWORDS)
    data["operating_expenses"] = g(OPERATING_EXPENSES_KEYWORDS)
    data["ebit"] = g(EBIT_KEYWORDS)
    data["ebitda"] = g(EBITDA_KEYWORDS)
    data["net_profit"] = g(NET_PROFIT_KEYWORDS)

    # ── Bug 2: Sum ALL matching component lines ───────────────────────────────
    dep_total, dep_items = _sum_all_matching(df, DEPRECIATION_KEYWORDS, col_idx, skip_cols)
    interest_total, interest_items = _sum_all_matching(df, INTEREST_KEYWORDS, col_idx, skip_cols)
    tax_total, tax_items = _sum_all_matching(df, TAX_KEYWORDS, col_idx, skip_cols)

    data["depreciation"] = dep_total
    data["interest_expense"] = interest_total
    data["tax_expense"] = tax_total

    # Store component detail for EBIT/EBITDA breakdown display
    data["_dep_components"] = dep_items
    data["_interest_components"] = interest_items
    data["_tax_components"] = tax_items

    # ── Derive gross profit if missing ────────────────────────────────────────
    if data["gross_profit"] is None and data["revenue"] and data["cogs"] is not None:
        data["gross_profit"] = data["revenue"] - data["cogs"]

    # ── Validate gross profit consistency (Bug 3 pre-check) ──────────────────
    if (data["gross_profit"] is not None and data["revenue"] is not None
            and data["cogs"] is not None):
        calc_gp = data["revenue"] - data["cogs"]
        if abs(data["gross_profit"] - calc_gp) > 1:
            logger.warning(
                f"Gross profit mismatch: parsed={data['gross_profit']:.0f}, "
                f"calculated={calc_gp:.0f}"
            )

    # ── Derive EBIT from components (Bug 2) — don't rely on parsed EBIT ──────
    # Store parsed EBIT separately; calculator will always recompute
    data["_ebit_parsed"] = data["ebit"]
    # (Calculator will override with component formula — see calculator.py)

    # ── Derive EBITDA if not explicit ─────────────────────────────────────────
    if data["ebitda"] is None and data["ebit"] is not None:
        dep = data["depreciation"] or 0
        data["ebitda"] = data["ebit"] + dep

    return data


# ── Balance sheet extraction (Bug 5: section tracking for inventory) ─────────

def _extract_balance_sheet_data(df: pd.DataFrame, col_idx: int = 0,
                                 skip_cols: set = None) -> dict:
    """
    Extract balance sheet line items for a given period.
    Bug 5: Tracks BS section so inventory is only captured from Current Assets.
    """
    label_col = df.columns[0]
    value_cols = [c for c in df.columns if c != label_col]
    if skip_cols:
        value_cols = [c for c in value_cols if c not in skip_cols]
    if not value_cols or col_idx >= len(value_cols):
        return {}

    vcol = value_cols[col_idx]

    # ── Section tracking state ────────────────────────────────────────────────
    in_current_assets = False
    in_non_current_assets = False
    in_current_liabilities = False
    in_equity = False

    data = {}
    bs_line_items = []  # For debug mode and self-checks

    for _, row in df.iterrows():
        raw_label = row[label_col]
        label = str(raw_label).strip()
        label_lower = label.lower()
        val = _clean_amount(row[vcol])

        # ── Detect section transitions from header rows (no value) ────────────
        if val is None:
            # Non-current assets (check before current to avoid substring match)
            if (("non-current assets" in label_lower or "noncurrent assets" in label_lower
                    or "fixed assets" in label_lower or "plant and equipment" in label_lower)
                    and "total" not in label_lower):
                in_current_assets = False
                in_non_current_assets = True
                in_current_liabilities = False
                in_equity = False
                continue
            # Current assets
            if ("current assets" in label_lower and "non" not in label_lower
                    and "total" not in label_lower):
                in_current_assets = True
                in_non_current_assets = False
                in_current_liabilities = False
                in_equity = False
                continue
            # Current liabilities
            if ("current liabilities" in label_lower and "non" not in label_lower
                    and "total" not in label_lower):
                in_current_assets = False
                in_non_current_assets = False
                in_current_liabilities = True
                in_equity = False
                continue
            # Non-current liabilities — exit current liabilities
            if ("non-current liabilities" in label_lower
                    or "long-term liabilities" in label_lower):
                in_current_liabilities = False
                continue
            # Equity
            if ("equity" in label_lower or "net assets" in label_lower
                    or "shareholders" in label_lower) and "total" not in label_lower:
                in_current_assets = False
                in_non_current_assets = False
                in_current_liabilities = False
                in_equity = True
                continue
            continue

        # ── Section transitions from total rows (rows with values) ───────────
        if "total current assets" in label_lower:
            data["current_assets"] = val
            in_current_assets = False  # Past the CA section
            continue
        if ("total non-current assets" in label_lower
                or "total fixed assets" in label_lower):
            data["non_current_assets"] = val
            in_non_current_assets = False
            continue
        if "total current liabilities" in label_lower:
            data["current_liabilities"] = val
            in_current_liabilities = False
            continue
        if "total non-current liabilities" in label_lower:
            data.setdefault("non_current_liabilities", val)
            continue
        if "total assets" in label_lower and "non" not in label_lower:
            data["total_assets"] = val
            continue
        if "total liabilities" in label_lower and "non-current" not in label_lower and "current" not in label_lower:
            data["total_liabilities"] = val
            continue

        # ── Record line item ──────────────────────────────────────────────────
        if in_current_assets:
            subsection = "current_assets"
        elif in_non_current_assets:
            subsection = "non_current_assets"
        elif in_current_liabilities:
            subsection = "current_liabilities"
        elif in_equity:
            subsection = "equity"
        else:
            subsection = "unknown"

        bs_line_items.append({
            "label": label,
            "value": val,
            "source": "balance_sheet",
            "subsection": subsection,
            "is_subtotal": _is_subtotal_row(label),
        })

        # ── Cash ──────────────────────────────────────────────────────────────
        if "cash" not in data and _matches_keywords(label, CASH_KEYWORDS):
            data["cash"] = val

        # ── Accounts Receivable ───────────────────────────────────────────────
        if "accounts_receivable" not in data and _matches_keywords(label, RECEIVABLES_KEYWORDS):
            data["accounts_receivable"] = val

        # ── Inventory: ONLY from current_assets section (Bug 5) ───────────────
        if "inventory" not in data and in_current_assets and _matches_keywords(label, INVENTORY_KEYWORDS):
            data["inventory"] = val
            data["_inventory_source"] = f"balance_sheet/current_assets ({label})"
            logger.info(f"Inventory captured from balance_sheet/current_assets: {label} = {val}")

        # ── Current assets subtotal (not caught by "total current assets" above) ──
        if "current_assets" not in data and _matches_keywords(label, CURRENT_ASSETS_KEYWORDS):
            data["current_assets"] = val

        # ── Non-current assets ────────────────────────────────────────────────
        if "non_current_assets" not in data and _matches_keywords(label, NON_CURRENT_ASSETS_KEYWORDS):
            if _is_subtotal_row(label) or in_non_current_assets:
                data["non_current_assets"] = val

        # ── Total Assets ──────────────────────────────────────────────────────
        if "total_assets" not in data and _matches_keywords(label, TOTAL_ASSETS_KEYWORDS):
            data["total_assets"] = val

        # ── Accounts Payable ──────────────────────────────────────────────────
        if "accounts_payable" not in data and _matches_keywords(label, PAYABLES_KEYWORDS):
            data["accounts_payable"] = val

        # ── Current Liabilities ───────────────────────────────────────────────
        if "current_liabilities" not in data and _matches_keywords(label, CURRENT_LIABILITIES_KEYWORDS):
            data["current_liabilities"] = val

        # ── Non-Current Liabilities ───────────────────────────────────────────
        if "non_current_liabilities" not in data and _matches_keywords(label, NON_CURRENT_LIABILITIES_KEYWORDS):
            data["non_current_liabilities"] = val

        # ── Total Liabilities ─────────────────────────────────────────────────
        if "total_liabilities" not in data and _matches_keywords(label, TOTAL_LIABILITIES_KEYWORDS):
            data["total_liabilities"] = val

        # ── Equity ────────────────────────────────────────────────────────────
        if "equity" not in data and _matches_keywords(label, EQUITY_KEYWORDS):
            if _is_subtotal_row(label) or in_equity:
                data["equity"] = val

        # ── Total Debt ────────────────────────────────────────────────────────
        if "total_debt" not in data and _matches_keywords(label, DEBT_KEYWORDS):
            data["total_debt"] = val

    # ── Fallback derivations ──────────────────────────────────────────────────
    # If inventory was not found in current_assets section, do NOT use P&L Closing Stock
    if "inventory" not in data:
        data["inventory"] = None
        data["_inventory_source"] = "not_found"
        logger.info("Inventory not found in Balance Sheet current assets section")

    if data.get("current_assets") is None:
        parts = [data.get("cash"), data.get("accounts_receivable"), data.get("inventory")]
        parts = [p for p in parts if p is not None]
        if parts:
            data["current_assets"] = sum(parts)

    if data.get("total_assets") is None:
        ca = data.get("current_assets")
        nca = data.get("non_current_assets")
        if ca is not None and nca is not None:
            data["total_assets"] = ca + nca

    if data.get("total_liabilities") is None:
        cl = data.get("current_liabilities")
        ncl = data.get("non_current_liabilities")
        if cl is not None and ncl is not None:
            data["total_liabilities"] = cl + ncl

    data["_bs_line_items"] = bs_line_items
    return data


# ── Cash flow extraction ──────────────────────────────────────────────────────

def _extract_cashflow_data(df: pd.DataFrame, col_idx: int = 0,
                           skip_cols: set = None) -> dict:
    data = {}

    def g(keywords):
        return _find_value(df, keywords, col_idx, skip_cols)

    data["operating_cash_flow"] = g(OPERATING_CF_KEYWORDS)
    data["investing_cash_flow"] = g(INVESTING_CF_KEYWORDS)
    data["financing_cash_flow"] = g(FINANCING_CF_KEYWORDS)
    return data


# ── File reading ──────────────────────────────────────────────────────────────

def _read_file(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    content = uploaded_file.read()
    uploaded_file.seek(0)

    if name.endswith(".csv"):
        for enc in ["utf-8", "utf-8-sig", "latin1"]:
            try:
                df = pd.read_csv(BytesIO(content), encoding=enc, header=None)
                return df
            except Exception:
                continue
        raise ValueError("Could not read CSV file")
    elif name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(BytesIO(content), header=None, engine="openpyxl")
        return df
    else:
        raise ValueError(f"Unsupported file type: {name}")


def _identify_header_row(df: pd.DataFrame) -> int:
    for i, row in df.iterrows():
        row_str = " ".join(str(v).lower() for v in row.values)
        if any(kw in row_str for kw in ["period", "ytd", "current", "prior", "year"]):
            return i
        years = re.findall(r"\b20\d{2}\b", row_str)
        if len(years) >= 1:
            return i
    return 0


def _clean_dataframe(raw_df: pd.DataFrame) -> pd.DataFrame:
    header_row = _identify_header_row(raw_df)
    df = raw_df.copy()

    if header_row > 0:
        df.columns = df.iloc[header_row].astype(str)
        df = df.iloc[header_row + 1:].reset_index(drop=True)
    else:
        df.columns = [f"col_{i}" for i in range(len(df.columns))]

    df = df.dropna(how="all")
    return df


# ── Public parse functions ────────────────────────────────────────────────────

def parse_xero_pl(uploaded_file) -> dict:
    """
    Parse a Xero P&L export.
    Returns dict with 'current', 'prior', 'prior2', 'period_labels',
    'period_fallback_warning', 'reference_columns'.
    """
    raw_df = _read_file(uploaded_file)
    df = _clean_dataframe(raw_df)
    df_sorted, period_labels, used_fallback = _detect_and_sort_periods(df)

    # Bug 4: Detect and exclude reference columns
    ref_cols = _detect_reference_columns(df_sorted)

    n_periods = len(df_sorted.columns) - 1
    return {
        "current": _extract_period_data(df_sorted, 0, ref_cols),
        "prior": _extract_period_data(df_sorted, 1, ref_cols) if n_periods >= 2 else None,
        "prior2": _extract_period_data(df_sorted, 2, ref_cols) if n_periods >= 3 else None,
        "period_labels": period_labels,
        "period_fallback_warning": used_fallback,
        "reference_columns": list(ref_cols),
        "raw_df": df_sorted,
        "type": "pl",
    }


def parse_xero_balance_sheet(uploaded_file) -> dict:
    """Parse a Xero Balance Sheet export."""
    raw_df = _read_file(uploaded_file)
    df = _clean_dataframe(raw_df)
    df_sorted, period_labels, used_fallback = _detect_and_sort_periods(df)

    ref_cols = _detect_reference_columns(df_sorted)

    n_periods = len(df_sorted.columns) - 1
    return {
        "current": _extract_balance_sheet_data(df_sorted, 0, ref_cols),
        "prior": _extract_balance_sheet_data(df_sorted, 1, ref_cols) if n_periods >= 2 else None,
        "prior2": _extract_balance_sheet_data(df_sorted, 2, ref_cols) if n_periods >= 3 else None,
        "period_labels": period_labels,
        "period_fallback_warning": used_fallback,
        "reference_columns": list(ref_cols),
        "raw_df": df_sorted,
        "type": "bs",
    }


def parse_xero_cashflow(uploaded_file) -> dict:
    """Parse a Xero Cash Flow Statement export."""
    raw_df = _read_file(uploaded_file)
    df = _clean_dataframe(raw_df)
    df_sorted, period_labels, used_fallback = _detect_and_sort_periods(df)

    ref_cols = _detect_reference_columns(df_sorted)

    n_periods = len(df_sorted.columns) - 1
    return {
        "current": _extract_cashflow_data(df_sorted, 0, ref_cols),
        "prior": _extract_cashflow_data(df_sorted, 1, ref_cols) if n_periods >= 2 else None,
        "period_labels": period_labels,
        "period_fallback_warning": used_fallback,
        "raw_df": df_sorted,
        "type": "cf",
    }


def merge_financial_data(pl_data: dict, bs_data: dict, cf_data: dict | None = None) -> dict:
    """
    Merge P&L, Balance Sheet, and optional Cash Flow data into unified structure.
    """
    periods = ["current", "prior", "prior2"]
    merged = {}

    for period in periods:
        pl = pl_data.get(period) or {}
        bs = bs_data.get(period) or {}
        cf = (cf_data.get(period) if cf_data else None) or {}

        if not any(v is not None for v in list(pl.values()) + list(bs.values())):
            continue

        merged[period] = {**pl, **bs, **cf}

    labels = pl_data.get("period_labels", ["Current", "Prior"])

    # Collect column classification info for UI display
    ref_cols_pl = pl_data.get("reference_columns", [])
    ref_cols_bs = bs_data.get("reference_columns", []) if bs_data else []
    all_ref_cols = list(set(ref_cols_pl + ref_cols_bs))

    return {
        "data": merged,
        "period_labels": labels,
        "reference_columns": all_ref_cols,
    }


def get_demo_data() -> dict:
    """Return sample financial data for demonstration purposes."""
    return {
        "data": {
            "current": {
                "revenue": 2_850_000,
                "cogs": 1_425_000,
                "gross_profit": 1_425_000,
                "operating_expenses": 1_100_000,
                "ebit": 325_000,
                "depreciation": 85_000,
                "ebitda": 410_000,
                "interest_expense": 42_000,
                "tax_expense": 84_900,
                "net_profit": 198_100,
                "cash": 180_000,
                "accounts_receivable": 310_000,
                "inventory": 220_000,
                "_inventory_source": "balance_sheet/current_assets (demo data)",
                "current_assets": 755_000,
                "non_current_assets": 890_000,
                "total_assets": 1_645_000,
                "accounts_payable": 185_000,
                "current_liabilities": 380_000,
                "non_current_liabilities": 420_000,
                "total_liabilities": 800_000,
                "equity": 845_000,
                "total_debt": 560_000,
                "operating_cash_flow": 285_000,
                "_dep_components": [("Depreciation & Amortisation (demo)", 85_000)],
                "_interest_components": [("Interest Expense (demo)", 42_000)],
                "_tax_components": [("Income Tax Expense (demo)", 84_900)],
            },
            "prior": {
                "revenue": 2_580_000,
                "cogs": 1_264_200,
                "gross_profit": 1_315_800,
                "operating_expenses": 1_040_000,
                "ebit": 275_800,
                "depreciation": 80_000,
                "ebitda": 355_800,
                "interest_expense": 45_000,
                "tax_expense": 69_240,
                "net_profit": 161_560,
                "cash": 145_000,
                "accounts_receivable": 265_000,
                "inventory": 195_000,
                "_inventory_source": "balance_sheet/current_assets (demo data)",
                "current_assets": 650_000,
                "non_current_assets": 940_000,
                "total_assets": 1_590_000,
                "accounts_payable": 160_000,
                "current_liabilities": 335_000,
                "non_current_liabilities": 480_000,
                "total_liabilities": 815_000,
                "equity": 775_000,
                "total_debt": 610_000,
                "operating_cash_flow": 252_000,
                "_dep_components": [("Depreciation & Amortisation (demo)", 80_000)],
                "_interest_components": [("Interest Expense (demo)", 45_000)],
                "_tax_components": [("Income Tax Expense (demo)", 69_240)],
            },
        },
        "period_labels": ["FY2024", "FY2023"],
        "client_name": "Demo Trading Pty Ltd",
        "industry": "Wholesale trade",
        "abn": "12 345 678 901",
        "financial_year_end": "30 June 2024",
        "reference_columns": [],
    }
