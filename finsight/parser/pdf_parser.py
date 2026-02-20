"""
PDF financial statement parser using pdfplumber.
Extracts P&L, Balance Sheet, and Cash Flow data from PDFs.
Includes a data confirmation/editing layer for accuracy review.
"""

import re
import logging
from io import BytesIO
from typing import Optional

try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except Exception:
    PDFPLUMBER_AVAILABLE = False

import pandas as pd

logger = logging.getLogger(__name__)

# ── Field definitions for manual confirmation ─────────────────────────────────
PL_FIELDS = [
    ("revenue", "Revenue / Total Income", "P&L"),
    ("cogs", "Cost of Goods Sold / Cost of Sales", "P&L"),
    ("gross_profit", "Gross Profit", "P&L"),
    ("operating_expenses", "Total Operating Expenses", "P&L"),
    ("ebit", "EBIT (Operating Profit)", "P&L"),
    ("depreciation", "Depreciation & Amortisation", "P&L"),
    ("ebitda", "EBITDA", "P&L"),
    ("interest_expense", "Interest Expense", "P&L"),
    ("tax_expense", "Income Tax Expense", "P&L"),
    ("net_profit", "Net Profit / Net Income", "P&L"),
]

BS_FIELDS = [
    ("cash", "Cash & Bank Balances", "Balance Sheet"),
    ("accounts_receivable", "Accounts Receivable / Debtors", "Balance Sheet"),
    ("inventory", "Inventory / Stock", "Balance Sheet"),
    ("current_assets", "Total Current Assets", "Balance Sheet"),
    ("non_current_assets", "Total Non-Current Assets", "Balance Sheet"),
    ("total_assets", "Total Assets", "Balance Sheet"),
    ("accounts_payable", "Accounts Payable / Creditors", "Balance Sheet"),
    ("current_liabilities", "Total Current Liabilities", "Balance Sheet"),
    ("non_current_liabilities", "Total Non-Current Liabilities", "Balance Sheet"),
    ("total_liabilities", "Total Liabilities", "Balance Sheet"),
    ("equity", "Total Equity", "Balance Sheet"),
    ("total_debt", "Total Borrowings / Debt", "Balance Sheet"),
]

CF_FIELDS = [
    ("operating_cash_flow", "Net Cash from Operating Activities", "Cash Flow"),
    ("investing_cash_flow", "Net Cash from Investing Activities", "Cash Flow"),
    ("financing_cash_flow", "Net Cash from Financing Activities", "Cash Flow"),
]

ALL_FIELDS = PL_FIELDS + BS_FIELDS + CF_FIELDS


def _extract_text_from_pdf(uploaded_file) -> str:
    """Extract all text from a PDF file."""
    if not PDFPLUMBER_AVAILABLE:
        raise ImportError("pdfplumber is not available. Install it with: pip install pdfplumber")

    content = uploaded_file.read()
    uploaded_file.seek(0)
    full_text = []

    with pdfplumber.open(BytesIO(content)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text.append(text)

    return "\n".join(full_text)


def _extract_tables_from_pdf(uploaded_file) -> list[pd.DataFrame]:
    """Extract tabular data from PDF pages."""
    if not PDFPLUMBER_AVAILABLE:
        return []

    content = uploaded_file.read()
    uploaded_file.seek(0)
    tables = []

    with pdfplumber.open(BytesIO(content)) as pdf:
        for page in pdf.pages:
            page_tables = page.extract_tables()
            for tbl in page_tables:
                if tbl and len(tbl) > 1:
                    try:
                        df = pd.DataFrame(tbl[1:], columns=tbl[0])
                        tables.append(df)
                    except Exception:
                        # Fallback: treat first row as data
                        df = pd.DataFrame(tbl)
                        tables.append(df)

    return tables


def _clean_amount(val: str) -> Optional[float]:
    """Parse a string amount to float."""
    if val is None:
        return None
    s = str(val).strip()
    if not s or s in ("-", "", "n/a", "N/A", "—", "nil"):
        return None
    negative = s.startswith("(") and s.endswith(")")
    s = re.sub(r"[$()\s,]", "", s)
    try:
        result = float(s)
        return -result if negative else result
    except ValueError:
        return None


def _find_amount_in_line(line: str) -> Optional[float]:
    """Extract a dollar amount from a text line."""
    # Match patterns like: 1,234,567.89 or (1,234,567) or $1,234,567
    patterns = [
        r"\([\d,]+(?:\.\d{1,2})?\)",  # Parenthesised negatives
        r"-[\d,]+(?:\.\d{1,2})?",     # Negative with minus
        r"[\d,]+(?:\.\d{1,2})?",      # Plain positive
    ]
    for pattern in patterns:
        match = re.search(pattern, line)
        if match:
            return _clean_amount(match.group())
    return None


def _keyword_match(text: str, keywords: list) -> bool:
    """Check if text contains any of the keywords (case-insensitive)."""
    t = text.lower()
    return any(kw.lower() in t for kw in keywords)


# Keyword sets for section detection
SECTION_KEYWORDS = {
    "revenue": ["revenue", "total income", "sales", "turnover", "total revenue", "gross receipts"],
    "cogs": ["cost of sales", "cost of goods", "cogs", "direct costs", "purchases"],
    "gross_profit": ["gross profit"],
    "operating_expenses": ["total expenses", "operating expenses", "total overheads"],
    "ebit": ["ebit", "operating profit", "profit from operations"],
    "depreciation": ["depreciation", "amortisation", "amortization"],
    "ebitda": ["ebitda"],
    "interest_expense": ["interest expense", "finance costs", "interest paid"],
    "tax_expense": ["income tax", "tax expense"],
    "net_profit": ["net profit", "net income", "profit after tax", "profit before tax", "net loss"],
    "cash": ["cash at bank", "cash and cash equivalents", "bank balances"],
    "accounts_receivable": ["accounts receivable", "trade receivables", "debtors"],
    "inventory": ["inventory", "closing stock", "goods on hand"],
    "current_assets": ["total current assets"],
    "non_current_assets": ["total non-current assets", "total fixed assets"],
    "total_assets": ["total assets"],
    "accounts_payable": ["accounts payable", "trade payables", "creditors"],
    "current_liabilities": ["total current liabilities"],
    "non_current_liabilities": ["total non-current liabilities"],
    "total_liabilities": ["total liabilities"],
    "equity": ["total equity", "net assets", "shareholders equity"],
    "total_debt": ["total loans", "total borrowings", "bank loans"],
    "operating_cash_flow": ["net cash from operating", "cash from operations", "operating cash flow"],
    "investing_cash_flow": ["net cash from investing", "investing activities"],
    "financing_cash_flow": ["net cash from financing", "financing activities"],
}


def _parse_text_to_data(text: str) -> dict:
    """
    Parse raw text extracted from PDF into financial line items.
    Returns dict with extracted values (may be incomplete/inaccurate).
    """
    data = {}
    lines = text.split("\n")

    for line in lines:
        line_stripped = line.strip()
        if not line_stripped:
            continue

        for field, keywords in SECTION_KEYWORDS.items():
            if field not in data and _keyword_match(line_stripped, keywords):
                amount = _find_amount_in_line(line_stripped)
                if amount is not None:
                    data[field] = amount
                    break

    # Derive gross profit if missing
    if "gross_profit" not in data:
        if "revenue" in data and "cogs" in data:
            data["gross_profit"] = data["revenue"] - data["cogs"]

    # Derive EBIT if missing
    if "ebit" not in data and "net_profit" in data:
        tax = data.get("tax_expense", 0) or 0
        interest = data.get("interest_expense", 0) or 0
        data["ebit"] = data["net_profit"] + tax + interest

    # Derive EBITDA if missing
    if "ebitda" not in data and "ebit" in data:
        dep = data.get("depreciation", 0) or 0
        data["ebitda"] = data["ebit"] + dep

    return data


def _parse_tables_to_data(tables: list[pd.DataFrame]) -> dict:
    """
    Attempt to parse structured tables into financial data.
    Combines results from all tables.
    """
    data = {}

    for df in tables:
        if df.empty or len(df.columns) < 2:
            continue

        label_col = df.columns[0]
        value_col = df.columns[1]

        for _, row in df.iterrows():
            label = str(row[label_col]).strip()
            if not label:
                continue

            for field, keywords in SECTION_KEYWORDS.items():
                if field not in data and _keyword_match(label, keywords):
                    val = _clean_amount(str(row[value_col]))
                    if val is not None:
                        data[field] = val
                        break

    return data


def parse_pdf(uploaded_file) -> tuple[dict, str]:
    """
    Main PDF parsing function.
    Returns (extracted_data, extraction_notes).
    extracted_data is a dict of field_name -> value.
    """
    notes = []

    # Try table extraction first (more structured)
    try:
        tables = _extract_tables_from_pdf(uploaded_file)
        table_data = _parse_tables_to_data(tables)
        notes.append(f"Extracted {len(tables)} table(s) from PDF.")
    except Exception as e:
        table_data = {}
        notes.append(f"Table extraction failed: {e}")

    # Text extraction as supplementary
    try:
        text = _extract_text_from_pdf(uploaded_file)
        text_data = _parse_text_to_data(text)
        notes.append("Text extraction completed.")
    except Exception as e:
        text_data = {}
        notes.append(f"Text extraction failed: {e}")

    # Merge: table data takes precedence, fill gaps with text data
    merged = {**text_data, **table_data}

    # Count extracted fields
    extracted_count = sum(1 for v in merged.values() if v is not None)
    total_fields = len(SECTION_KEYWORDS)
    notes.append(
        f"Extracted {extracted_count}/{total_fields} fields. "
        "Please review and correct values below before proceeding."
    )

    return merged, "\n".join(notes)


def get_confirmation_template(extracted_data: dict) -> dict:
    """
    Build a confirmation template with all expected fields,
    pre-populated with any extracted values.
    """
    template = {}
    for field, label, section in ALL_FIELDS:
        template[field] = {
            "label": label,
            "section": section,
            "value": extracted_data.get(field),
            "field_key": field,
        }
    return template


def build_confirmed_data(confirmed_values: dict) -> dict:
    """
    Convert confirmed field values (from user edits) into the
    standard financial data dict.
    """
    result = {}
    for field, value in confirmed_values.items():
        if isinstance(value, (int, float)):
            result[field] = float(value)
        elif isinstance(value, str) and value.strip():
            cleaned = _clean_amount(value)
            result[field] = cleaned
        else:
            result[field] = None

    # Derive missing calculated fields
    if result.get("gross_profit") is None and result.get("revenue") and result.get("cogs"):
        result["gross_profit"] = result["revenue"] - result["cogs"]
    if result.get("ebit") is None and result.get("net_profit") is not None:
        tax = result.get("tax_expense") or 0
        interest = result.get("interest_expense") or 0
        result["ebit"] = result["net_profit"] + tax + interest
    if result.get("ebitda") is None and result.get("ebit") is not None:
        dep = result.get("depreciation") or 0
        result["ebitda"] = result["ebit"] + dep

    return result
