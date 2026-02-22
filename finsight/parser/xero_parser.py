"""
Xero CSV/Excel financial statement parser.
Handles Xero P&L, Balance Sheet, and Cash Flow exports.
"""

import pandas as pd
import numpy as np
import re
import logging
from io import BytesIO

logger = logging.getLogger(__name__)

# ── Account category keyword mappings ────────────────────────────────────────
REVENUE_KEYWORDS = [
    "revenue", "income", "sales", "turnover", "fees", "service income",
    "trading income", "gross receipts", "total income", "total revenue"
]
COGS_KEYWORDS = [
    "cost of sales", "cost of goods", "cogs", "direct costs", "direct expenses",
    "purchases", "cost of revenue", "materials", "subcontractors",
    "direct labour", "direct wages"
]
GROSS_PROFIT_KEYWORDS = ["gross profit", "gross margin"]
OPERATING_EXPENSES_KEYWORDS = [
    "operating expenses", "expenses", "overheads", "administrative", "general expenses",
    "selling expenses", "total expenses", "total overheads"
]
EBIT_KEYWORDS = ["ebit", "operating profit", "profit from operations", "operating income"]
EBITDA_KEYWORDS = ["ebitda"]
NET_PROFIT_KEYWORDS = [
    "net profit", "net income", "profit after tax", "profit before tax",
    "net loss", "net earnings", "profit for the year", "surplus"
]
DEPRECIATION_KEYWORDS = ["depreciation", "amortisation", "amortization", "d&a"]
INTEREST_KEYWORDS = ["interest expense", "finance costs", "borrowing costs", "interest paid"]
TAX_KEYWORDS = ["income tax", "tax expense", "corporate tax", "tax payable"]

# Balance sheet
CURRENT_ASSETS_KEYWORDS = ["current assets", "total current assets"]
CASH_KEYWORDS = ["cash", "bank", "cash and cash equivalents", "cash at bank"]
RECEIVABLES_KEYWORDS = [
    "accounts receivable", "debtors", "trade receivables", "receivables",
    "trade debtors", "sundry debtors"
]
INVENTORY_KEYWORDS = ["inventory", "stock", "goods on hand", "closing stock"]
OTHER_CURRENT_ASSETS_KEYWORDS = ["other current assets", "prepayments", "accrued income"]
NON_CURRENT_ASSETS_KEYWORDS = ["non-current assets", "fixed assets", "total non-current assets"]
TOTAL_ASSETS_KEYWORDS = ["total assets"]
CURRENT_LIABILITIES_KEYWORDS = ["current liabilities", "total current liabilities"]
PAYABLES_KEYWORDS = [
    "accounts payable", "creditors", "trade payables", "trade creditors", "sundry creditors"
]
OTHER_CURRENT_LIAB_KEYWORDS = ["other current liabilities", "accrued liabilities", "gst payable"]
NON_CURRENT_LIABILITIES_KEYWORDS = [
    "non-current liabilities", "long-term liabilities", "total non-current liabilities"
]
TOTAL_LIABILITIES_KEYWORDS = ["total liabilities"]
EQUITY_KEYWORDS = ["equity", "total equity", "shareholders equity", "net assets", "owners equity"]
DEBT_KEYWORDS = ["loans", "borrowings", "bank loan", "term loan", "line of credit", "overdraft"]

# Cash flow
OPERATING_CF_KEYWORDS = [
    "cash from operations", "operating cash flow", "net cash from operating",
    "cash generated from operations", "net cash provided by operating"
]
INVESTING_CF_KEYWORDS = [
    "cash from investing", "investing activities", "net cash from investing"
]
FINANCING_CF_KEYWORDS = [
    "cash from financing", "financing activities", "net cash from financing"
]


# ── Period column detection ───────────────────────────────────────────────

# Relative-term priority: higher number = more recent
_RELATIVE_PRIORITY = {
    "current year": 100, "current": 100, "this year": 100,
    "prior year": 50, "prior": 50, "previous year": 50,
    "comparative": 50, "last year": 50,
}


def _extract_year_from_col(text: str) -> int | None:
    """
    Try to extract a fiscal year (integer) from a column header string.
    Returns None if no year-like pattern is found.
    """
    s = str(text).strip()

    # FY2024 or FY 2024
    m = re.search(r'\bFY\s*(\d{4})\b', s, re.IGNORECASE)
    if m:
        return int(m.group(1))

    # FY24 → 20xx
    m = re.search(r'\bFY\s*(\d{2})\b', s, re.IGNORECASE)
    if m:
        return 2000 + int(m.group(1))

    # Year ended ... 2024
    m = re.search(r'year\s+ended.*?(20\d{2})', s, re.IGNORECASE)
    if m:
        return int(m.group(1))

    # 2023/24 or 2023/2024 — take the later year
    m = re.search(r'\b(20\d{2})/(\d{2,4})\b', s)
    if m:
        suffix = m.group(2)
        if len(suffix) == 2:
            return int(m.group(1)[:2] + suffix)  # 2023 + 24 → 2024
        return int(suffix)

    # Month-year range: "Jul 2023 - Jun 2024" — take the ending year
    m = re.search(r'[A-Za-z]{3}\s+\d{4}\s*[-\u2013]\s*[A-Za-z]{3}\s+(20\d{2})', s)
    if m:
        return int(m.group(1))

    # Bare 4-digit year: 2024
    m = re.search(r'\b(20\d{2})\b', s)
    if m:
        return int(m.group(1))

    return None


def _detect_and_sort_periods(df: pd.DataFrame) -> tuple:
    """
    Detect period column labels from header text and sort chronologically,
    most-recent period first (index 0 = current).

    Returns:
        df_sorted     — DataFrame with value columns reordered newest→oldest
        period_labels — list of column name strings (up to 3)
        used_fallback — True if date detection failed (positional order used)
    """
    label_col = df.columns[0]
    value_cols = [c for c in df.columns if c != label_col]

    if not value_cols:
        return df, [], False

    # 1. Try year extraction
    col_years = {}
    for col in value_cols:
        yr = _extract_year_from_col(str(col))
        if yr is not None:
            col_years[col] = yr

    if col_years:
        # Sort all columns; unrecognised columns go to the end
        sorted_cols = sorted(
            value_cols,
            key=lambda c: col_years.get(c, 0),
            reverse=True,
        )
        used_fallback = len(col_years) < len(value_cols)
        df_sorted = df[[label_col] + sorted_cols]
        return df_sorted, [str(c).strip() for c in sorted_cols[:3]], used_fallback

    # 2. Try relative terms (Current Year, Prior Year, etc.)
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

    # 3. Positional fallback
    return df, [str(c).strip() for c in value_cols[:3]], True


def _clean_amount(val) -> float | None:
    """Parse various number formats to float, returning None if unparseable."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val) if not np.isnan(float(val)) else None
    s = str(val).strip()
    if not s or s in ("-", "", "n/a", "N/A", "—"):
        return None
    # Handle parentheses as negative: (1,234.56)
    negative = s.startswith("(") and s.endswith(")")
    s = s.replace("(", "").replace(")", "")
    s = s.replace("$", "").replace(",", "").replace(" ", "")
    try:
        result = float(s)
        return -result if negative else result
    except ValueError:
        return None


def _matches_keywords(text: str, keywords: list) -> bool:
    """Case-insensitive check if text matches any keyword."""
    t = str(text).lower().strip()
    return any(kw.lower() in t for kw in keywords)


def _find_value(df: pd.DataFrame, keywords: list, col_idx: int = 0) -> float | None:
    """
    Search a dataframe for the first row matching keywords and return the value
    at col_idx (0 = current period, 1 = prior period).
    """
    label_col = df.columns[0]
    value_cols = [c for c in df.columns if c != label_col]
    if not value_cols or col_idx >= len(value_cols):
        return None

    for _, row in df.iterrows():
        label = str(row[label_col])
        if _matches_keywords(label, keywords):
            val = _clean_amount(row[value_cols[col_idx]])
            if val is not None:
                return val
    return None


def _extract_period_data(df: pd.DataFrame, col_idx: int) -> dict:
    """Extract all P&L line items for a given period column index."""
    data = {}

    def g(keywords):
        return _find_value(df, keywords, col_idx)

    data["revenue"] = g(REVENUE_KEYWORDS)
    data["cogs"] = g(COGS_KEYWORDS)
    data["gross_profit"] = g(GROSS_PROFIT_KEYWORDS)
    data["operating_expenses"] = g(OPERATING_EXPENSES_KEYWORDS)
    data["ebit"] = g(EBIT_KEYWORDS)
    data["ebitda"] = g(EBITDA_KEYWORDS)
    data["depreciation"] = g(DEPRECIATION_KEYWORDS)
    data["interest_expense"] = g(INTEREST_KEYWORDS)
    data["tax_expense"] = g(TAX_KEYWORDS)
    data["net_profit"] = g(NET_PROFIT_KEYWORDS)

    # Derive gross profit if not explicit
    if data["gross_profit"] is None and data["revenue"] and data["cogs"]:
        data["gross_profit"] = data["revenue"] - data["cogs"]

    # Derive EBIT if not explicit
    if data["ebit"] is None and data["net_profit"] is not None:
        tax = data["tax_expense"] or 0
        interest = data["interest_expense"] or 0
        data["ebit"] = data["net_profit"] + tax + interest

    # Derive EBITDA if not explicit
    if data["ebitda"] is None and data["ebit"] is not None:
        dep = data["depreciation"] or 0
        data["ebitda"] = data["ebit"] + dep

    return data


def _extract_balance_sheet_data(df: pd.DataFrame, col_idx: int) -> dict:
    """Extract balance sheet line items for a given period."""
    data = {}

    def g(keywords):
        return _find_value(df, keywords, col_idx)

    data["cash"] = g(CASH_KEYWORDS)
    data["accounts_receivable"] = g(RECEIVABLES_KEYWORDS)
    data["inventory"] = g(INVENTORY_KEYWORDS)
    data["current_assets"] = g(CURRENT_ASSETS_KEYWORDS)
    data["non_current_assets"] = g(NON_CURRENT_ASSETS_KEYWORDS)
    data["total_assets"] = g(TOTAL_ASSETS_KEYWORDS)
    data["accounts_payable"] = g(PAYABLES_KEYWORDS)
    data["current_liabilities"] = g(CURRENT_LIABILITIES_KEYWORDS)
    data["non_current_liabilities"] = g(NON_CURRENT_LIABILITIES_KEYWORDS)
    data["total_liabilities"] = g(TOTAL_LIABILITIES_KEYWORDS)
    data["equity"] = g(EQUITY_KEYWORDS)
    data["total_debt"] = g(DEBT_KEYWORDS)

    # Derive totals where missing
    if data["current_assets"] is None:
        parts = [data["cash"], data["accounts_receivable"], data["inventory"]]
        parts = [p for p in parts if p is not None]
        if parts:
            data["current_assets"] = sum(parts)

    if data["total_assets"] is None and data["current_assets"] and data["non_current_assets"]:
        data["total_assets"] = data["current_assets"] + data["non_current_assets"]

    return data


def _extract_cashflow_data(df: pd.DataFrame, col_idx: int) -> dict:
    """Extract cash flow statement data."""
    data = {}

    def g(keywords):
        return _find_value(df, keywords, col_idx)

    data["operating_cash_flow"] = g(OPERATING_CF_KEYWORDS)
    data["investing_cash_flow"] = g(INVESTING_CF_KEYWORDS)
    data["financing_cash_flow"] = g(FINANCING_CF_KEYWORDS)

    return data


def _read_file(uploaded_file) -> pd.DataFrame:
    """Read uploaded file into DataFrame."""
    name = uploaded_file.name.lower()
    content = uploaded_file.read()
    uploaded_file.seek(0)

    if name.endswith(".csv"):
        # Try multiple encodings
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
    """Find the row index that looks like a header (contains 'period' or year)."""
    for i, row in df.iterrows():
        row_str = " ".join(str(v).lower() for v in row.values)
        if any(kw in row_str for kw in ["period", "ytd", "current", "prior", "year"]):
            return i
        # Detect year patterns like 2023, 2024
        years = re.findall(r"\b20\d{2}\b", row_str)
        if len(years) >= 1:
            return i
    return 0


def _clean_dataframe(raw_df: pd.DataFrame) -> pd.DataFrame:
    """Set header row and clean the dataframe."""
    header_row = _identify_header_row(raw_df)
    df = raw_df.copy()

    if header_row > 0:
        df.columns = df.iloc[header_row].astype(str)
        df = df.iloc[header_row + 1:].reset_index(drop=True)
    else:
        df.columns = [f"col_{i}" for i in range(len(df.columns))]

    # Remove completely empty rows
    df = df.dropna(how="all")
    return df


def parse_xero_pl(uploaded_file) -> dict:
    """
    Parse a Xero P&L export.
    Returns dict with 'current', 'prior', 'prior2', 'period_labels',
    'period_fallback_warning' (True if positional ordering was used).
    """
    raw_df = _read_file(uploaded_file)
    df = _clean_dataframe(raw_df)
    df_sorted, period_labels, used_fallback = _detect_and_sort_periods(df)

    n_periods = len(df_sorted.columns) - 1  # exclude label column
    return {
        "current": _extract_period_data(df_sorted, 0),
        "prior": _extract_period_data(df_sorted, 1) if n_periods >= 2 else None,
        "prior2": _extract_period_data(df_sorted, 2) if n_periods >= 3 else None,
        "period_labels": period_labels,
        "period_fallback_warning": used_fallback,
        "raw_df": df_sorted,
        "type": "pl",
    }


def parse_xero_balance_sheet(uploaded_file) -> dict:
    """Parse a Xero Balance Sheet export."""
    raw_df = _read_file(uploaded_file)
    df = _clean_dataframe(raw_df)
    df_sorted, period_labels, used_fallback = _detect_and_sort_periods(df)

    n_periods = len(df_sorted.columns) - 1
    return {
        "current": _extract_balance_sheet_data(df_sorted, 0),
        "prior": _extract_balance_sheet_data(df_sorted, 1) if n_periods >= 2 else None,
        "prior2": _extract_balance_sheet_data(df_sorted, 2) if n_periods >= 3 else None,
        "period_labels": period_labels,
        "period_fallback_warning": used_fallback,
        "raw_df": df_sorted,
        "type": "bs",
    }


def parse_xero_cashflow(uploaded_file) -> dict:
    """Parse a Xero Cash Flow Statement export."""
    raw_df = _read_file(uploaded_file)
    df = _clean_dataframe(raw_df)
    df_sorted, period_labels, used_fallback = _detect_and_sort_periods(df)

    n_periods = len(df_sorted.columns) - 1
    return {
        "current": _extract_cashflow_data(df_sorted, 0),
        "prior": _extract_cashflow_data(df_sorted, 1) if n_periods >= 2 else None,
        "period_labels": period_labels,
        "period_fallback_warning": used_fallback,
        "raw_df": df_sorted,
        "type": "cf",
    }


def merge_financial_data(pl_data: dict, bs_data: dict, cf_data: dict | None = None) -> dict:
    """
    Merge P&L, Balance Sheet, and optional Cash Flow data into unified structure.
    Returns dict with 'periods' list and 'data' dict keyed by period index.
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
    return {"data": merged, "period_labels": labels}


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
            },
        },
        "period_labels": ["FY2024", "FY2023"],
        "client_name": "Demo Trading Pty Ltd",
        "industry": "Wholesale trade",
        "abn": "12 345 678 901",
        "financial_year_end": "30 June 2024",
    }
