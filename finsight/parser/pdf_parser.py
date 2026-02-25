"""
PDF financial statement parser using pdfplumber.
Extracts P&L, Balance Sheet, and Cash Flow data from PDFs.
Includes a data confirmation/editing layer for accuracy review.

Bug fixes applied:
- Bug 1: Section-aware parsing to capture totals/subtotals
- Bug 2: Expanded keyword matching for EBIT components
- Bug 4: Note reference detection using positional analysis and column classification
- Bug 5: Inventory only sourced from Balance Sheet current assets section
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
    ("inventory", "Inventory / Stock on Hand (Balance Sheet only)", "Balance Sheet"),
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

# ── Section markers used to track position in the financial statements ────────
# P&L sections
PL_SECTION_HEADERS = {
    "revenue": [
        "revenue", "income", "total income", "total revenue", "sales",
        "turnover", "trading income", "gross receipts",
    ],
    "cogs": [
        "cost of sales", "cost of goods sold", "direct costs", "purchases",
        "cost of revenue",
    ],
    "gross_profit": ["gross profit"],
    "operating_expenses": [
        "operating expenses", "expenses", "overheads", "administrative",
        "general and administrative",
    ],
    "below_ebit": [
        "finance costs", "interest", "income tax", "taxation",
    ],
}

# P&L subtotal / total row patterns (Bug 1)
PL_TOTAL_PATTERNS = {
    "revenue": [
        "total revenue", "total income", "total sales", "total trading income",
        "total fees", "total turnover", "gross income",
    ],
    "cogs": [
        "total cost of sales", "total cost of goods", "total cogs",
        "total direct costs", "total direct expenses",
    ],
}

# Balance Sheet sections
BS_SECTION_HEADERS = {
    "current_assets": ["current assets"],
    "non_current_assets": ["non-current assets", "fixed assets", "plant and equipment"],
    "current_liabilities": ["current liabilities"],
    "non_current_liabilities": ["non-current liabilities", "long-term liabilities"],
    "equity": ["equity", "shareholders equity", "net assets"],
}

# Inventory keywords — for use in BS current_assets section ONLY (Bug 5)
INVENTORY_KEYWORDS = [
    "inventory", "stock on hand", "closing stock", "finished goods",
    "raw materials", "work in progress", "wip", "trading stock", "stock",
]

# Subtotal keywords
SUBTOTAL_KEYWORDS = ["total", "subtotal", "gross profit", "net profit", "net loss", "net income"]


def _is_subtotal_row(label: str) -> bool:
    """Return True if the label is a subtotal/total row."""
    return any(kw in label.lower() for kw in SUBTOTAL_KEYWORDS)


# ── Keyword sets for field extraction ─────────────────────────────────────────
SECTION_KEYWORDS = {
    # P&L
    "revenue": ["total revenue", "total income", "total sales", "revenue", "sales", "turnover", "total trading income"],
    "cogs": ["total cost of sales", "total cost of goods", "total direct costs", "cost of sales", "cost of goods", "direct costs"],
    "gross_profit": ["gross profit"],
    "operating_expenses": ["total expenses", "total operating expenses", "operating expenses", "total overheads"],
    "ebit": ["ebit", "operating profit", "profit from operations"],
    "depreciation": [
        "depreciation", "amortisation", "amortization", "dep &",
        "depreciation and amortisation", "d&a", "right of use",
        "rou asset depreciation",
    ],
    "ebitda": ["ebitda"],
    "interest_expense": [
        "interest expense", "finance costs", "interest paid", "finance charge",
        "loan interest", "bank charge", "borrowing cost",
    ],
    "tax_expense": ["income tax", "tax expense", "taxation", "company tax", "provision for tax"],
    "net_profit": ["net profit", "net income", "profit after tax", "profit before tax", "net loss"],
    # Balance Sheet
    "cash": ["cash at bank", "cash and cash equivalents", "bank balances", "cash"],
    "accounts_receivable": ["accounts receivable", "trade receivables", "debtors"],
    "current_assets": ["total current assets"],
    "non_current_assets": ["total non-current assets", "total fixed assets"],
    "total_assets": ["total assets"],
    "accounts_payable": ["accounts payable", "trade payables", "creditors"],
    "current_liabilities": ["total current liabilities"],
    "non_current_liabilities": ["total non-current liabilities"],
    "total_liabilities": ["total liabilities"],
    "equity": ["total equity", "net assets", "shareholders equity"],
    "total_debt": ["total loans", "total borrowings", "bank loans"],
    # Cash Flow
    "operating_cash_flow": ["net cash from operating", "cash from operations", "operating cash flow"],
    "investing_cash_flow": ["net cash from investing", "investing activities"],
    "financing_cash_flow": ["net cash from financing", "financing activities"],
}

# inventory handled separately — only from BS current assets


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


def _extract_words_from_pdf(uploaded_file) -> list:
    """
    Extract word-level data with bounding boxes for positional note-ref detection.
    Returns list of {text, x0, x1, y0, y1, page} dicts.
    Bug 4: Used to positionally identify note reference numbers.
    """
    if not PDFPLUMBER_AVAILABLE:
        return []

    content = uploaded_file.read()
    uploaded_file.seek(0)
    words = []

    try:
        with pdfplumber.open(BytesIO(content)) as pdf:
            for page_num, page in enumerate(pdf.pages):
                page_words = page.extract_words()
                for w in (page_words or []):
                    words.append({
                        "text": w.get("text", ""),
                        "x0": w.get("x0", 0),
                        "x1": w.get("x1", 0),
                        "top": w.get("top", 0),
                        "page": page_num,
                    })
    except Exception as e:
        logger.warning(f"Word extraction failed: {e}")

    return words


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
    patterns = [
        r"\([\d,]+(?:\.\d{1,2})?\)",  # Parenthesised negatives
        r"-[\d,]+(?:\.\d{1,2})?",      # Negative with minus
        r"[\d,]+(?:\.\d{1,2})?",       # Plain positive
    ]
    for pattern in patterns:
        match = re.search(pattern, line)
        if match:
            val = _clean_amount(match.group())
            # Bug 4: Ignore small integers (1-50) that are likely note references
            if val is not None and not _is_likely_note_ref_value(val, line):
                return val
    return None


def _is_likely_note_ref_value(val: float, line: str) -> bool:
    """
    Bug 4: Return True if val is likely a note reference number, not a financial figure.
    Heuristics:
    - Small integer between 1 and 50 with no decimal
    - Appears as the ONLY number in a short line (label + note ref with no dollar amount)
    - No dollar sign or comma near it in the line
    """
    if val is None:
        return False
    if not (1 <= val <= 50 and val == int(val)):
        return False
    # If there's also a large number in the line, this small int is likely a note ref
    large_numbers = re.findall(r"[\d,]{4,}", line)
    if large_numbers:
        return True  # There's a real financial figure alongside the small integer
    # If there's a dollar sign, likely a real value
    if "$" in line:
        return False
    return True


def _keyword_match(text: str, keywords: list) -> bool:
    t = text.lower()
    return any(kw.lower() in t for kw in keywords)


# ── Bug 4: Column classification for table-based extraction ──────────────────

def _classify_table_columns(df: pd.DataFrame) -> tuple[list, list]:
    """
    Classify columns in a table as VALUE columns or REFERENCE columns.
    Returns (value_cols, reference_cols).
    A REFERENCE column contains only small integers (1–50) with no decimals.
    A VALUE column contains larger numbers or decimals.
    """
    if df.empty or len(df.columns) < 2:
        return list(df.columns), []

    value_cols = []
    ref_cols = []

    for col in df.columns[1:]:  # Skip label column
        numeric_vals = []
        has_decimal = False
        for v in df[col]:
            c = _clean_amount(str(v))
            if c is not None:
                numeric_vals.append(c)
                if c != int(c):
                    has_decimal = True

        if not numeric_vals:
            continue

        all_small_int = all(1 <= v <= 50 and v == int(v) for v in numeric_vals)
        import numpy as np
        median_abs = float(np.median([abs(v) for v in numeric_vals]))

        if all_small_int and median_abs <= 50 and not has_decimal:
            ref_cols.append(col)
        else:
            value_cols.append(col)

    # Ensure first column (labels) is included
    label_col = df.columns[0]
    if label_col not in value_cols:
        value_cols = [label_col] + value_cols

    return value_cols, ref_cols


# ── Section-aware text parsing (Bug 1, Bug 5) ─────────────────────────────────

def _parse_text_to_data(text: str) -> dict:
    """
    Parse raw text extracted from PDF into financial line items.
    Bug 1: Prefers subtotal/total rows for revenue and COGS.
    Bug 5: Tracks balance sheet sections so inventory only comes from current assets.
    Returns dict with extracted values (may be incomplete/inaccurate).
    """
    data = {}
    notes = []
    lines = text.split("\n")

    # ── Track which statement section we're in ────────────────────────────────
    current_pl_section = None   # 'revenue', 'cogs', 'operating_expenses', etc.
    current_bs_section = None   # 'current_assets', 'non_current_assets', etc.
    in_balance_sheet = False
    in_pl = True  # Start assuming P&L first

    # Candidates for subtotal preference (Bug 1)
    revenue_candidates = []  # (is_subtotal, value)
    cogs_candidates = []

    for line in lines:
        line_stripped = line.strip()
        if not line_stripped:
            continue
        line_lower = line_stripped.lower()

        # ── Detect which statement we're reading ──────────────────────────────
        if "balance sheet" in line_lower or "statement of financial position" in line_lower:
            in_balance_sheet = True
            in_pl = False
            current_pl_section = None
            continue
        if ("profit and loss" in line_lower or "income statement" in line_lower
                or "statement of profit" in line_lower or "statement of comprehensive income" in line_lower):
            in_pl = True
            in_balance_sheet = False
            current_bs_section = None
            continue

        # ── P&L section tracking ──────────────────────────────────────────────
        if in_pl:
            # Detect P&L section headers (lines with no amount)
            amount_in_line = _find_amount_in_line(line_stripped)

            if amount_in_line is None:
                # Could be a section header
                for section, keywords in PL_SECTION_HEADERS.items():
                    if _keyword_match(line_lower, keywords):
                        current_pl_section = section
                        break
                continue

            # Line has an amount — extract based on current section
            if current_pl_section == "revenue" or _keyword_match(line_lower, SECTION_KEYWORDS["revenue"]):
                is_total = _is_subtotal_row(line_stripped)
                revenue_candidates.append((is_total, amount_in_line))
                # Also try other fields
            elif current_pl_section == "cogs" or _keyword_match(line_lower, SECTION_KEYWORDS["cogs"]):
                is_total = _is_subtotal_row(line_stripped)
                cogs_candidates.append((is_total, amount_in_line))

        # ── Balance sheet section tracking (Bug 5) ────────────────────────────
        if in_balance_sheet:
            amount_in_line = _find_amount_in_line(line_stripped)

            if amount_in_line is None:
                # Section header detection
                if ("non-current assets" in line_lower or "noncurrent assets" in line_lower
                        or "fixed assets" in line_lower):
                    current_bs_section = "non_current_assets"
                elif "current assets" in line_lower and "total" not in line_lower:
                    current_bs_section = "current_assets"
                elif "current liabilities" in line_lower and "non" not in line_lower and "total" not in line_lower:
                    current_bs_section = "current_liabilities"
                elif "non-current liabilities" in line_lower or "long-term liabilities" in line_lower:
                    current_bs_section = "non_current_liabilities"
                elif "equity" in line_lower and "total" not in line_lower:
                    current_bs_section = "equity"
                continue

            # Bug 5: Inventory only from current_assets BS section
            if ("inventory" in data and current_bs_section != "current_assets"):
                pass  # Don't update inventory if we already have it from CA section
            elif current_bs_section == "current_assets" and _keyword_match(line_lower, INVENTORY_KEYWORDS):
                if "inventory" not in data:
                    data["inventory"] = amount_in_line
                    data["_inventory_source"] = f"balance_sheet/current_assets ({line_stripped[:40]})"
                    notes.append(f"Inventory sourced from Balance Sheet current assets: {amount_in_line}")

        # ── General field extraction (for fields not section-tracked) ─────────
        for field, keywords in SECTION_KEYWORDS.items():
            if field == "inventory":
                continue  # Handled above with section tracking
            if field not in data and _keyword_match(line_stripped, keywords):
                amount = _find_amount_in_line(line_stripped)
                if amount is not None:
                    # Bug 1: For revenue/cogs, prefer subtotal rows
                    if field in ("revenue", "cogs"):
                        is_total = _is_subtotal_row(line_stripped)
                        if is_total or field not in data:
                            data[field] = amount
                    else:
                        data[field] = amount
                    break

    # ── Bug 1: Apply subtotal preference for revenue and COGS ─────────────────
    if revenue_candidates:
        subtotals = [v for is_total, v in revenue_candidates if is_total]
        if subtotals:
            data["revenue"] = subtotals[-1]  # Last subtotal is most likely the section total
        elif "revenue" not in data:
            data["revenue"] = revenue_candidates[-1][1]  # Last matching line

    if cogs_candidates:
        subtotals = [v for is_total, v in cogs_candidates if is_total]
        if subtotals:
            data["cogs"] = subtotals[-1]
        elif "cogs" not in data:
            data["cogs"] = cogs_candidates[-1][1]

    # ── If inventory not found in BS current assets, mark as not found ─────────
    if "inventory" not in data:
        data["_inventory_source"] = "not_found"
        data["_inventory_note"] = (
            "Inventory not identified on Balance Sheet — "
            "Quick Ratio equals Current Ratio. Inventory Days cannot be calculated."
        )

    # ── Derive missing calculated fields ──────────────────────────────────────
    if "gross_profit" not in data:
        if "revenue" in data and "cogs" in data:
            data["gross_profit"] = data["revenue"] - data["cogs"]

    if "ebit" not in data and "net_profit" in data:
        tax = data.get("tax_expense", 0) or 0
        interest = data.get("interest_expense", 0) or 0
        data["ebit"] = data["net_profit"] + tax + interest

    if "ebitda" not in data and "ebit" in data:
        dep = data.get("depreciation", 0) or 0
        data["ebitda"] = data["ebit"] + dep

    return data


def _parse_tables_to_data(tables: list[pd.DataFrame]) -> dict:
    """
    Attempt to parse structured tables into financial data.
    Bug 4: Classifies and excludes reference columns.
    Bug 5: Tracks table context for inventory sourcing.
    """
    data = {}
    column_info = {"excluded_ref_cols": [], "value_cols_used": []}

    for df in tables:
        if df.empty or len(df.columns) < 2:
            continue

        # Bug 4: Classify columns
        value_cols, ref_cols = _classify_table_columns(df)
        if ref_cols:
            column_info["excluded_ref_cols"].extend(ref_cols)
            logger.info(f"Table: excluded reference columns {ref_cols}")

        label_col = df.columns[0]
        # Use first value column (after excluding ref cols)
        non_ref_value_cols = [c for c in df.columns[1:] if c not in ref_cols]
        if not non_ref_value_cols:
            continue
        value_col = non_ref_value_cols[0]
        column_info["value_cols_used"].append(str(value_col))

        for _, row in df.iterrows():
            label = str(row[label_col]).strip()
            if not label or label.lower() == "nan":
                continue

            # Bug 5: Inventory only from balance sheet current assets context
            if _keyword_match(label, INVENTORY_KEYWORDS):
                # Only capture if not already from a better source
                if "inventory" not in data:
                    val = _clean_amount(str(row[value_col]))
                    if val is not None and val > 0:
                        data["inventory"] = val
                        data["_inventory_source"] = f"table_extraction ({label[:40]})"
                continue

            for field, keywords in SECTION_KEYWORDS.items():
                if field == "inventory":
                    continue
                if field not in data and _keyword_match(label, keywords):
                    val = _clean_amount(str(row[value_col]))
                    if val is not None:
                        # Bug 1: prefer subtotal rows for revenue/cogs
                        if field in ("revenue", "cogs"):
                            if _is_subtotal_row(label) or field not in data:
                                data[field] = val
                        else:
                            data[field] = val
                    break

    data["_column_info"] = column_info
    return data


def parse_pdf(uploaded_file) -> tuple[dict, str]:
    """
    Main PDF parsing function.
    Returns (extracted_data, extraction_notes).
    """
    notes = []

    # Try table extraction first
    try:
        tables = _extract_tables_from_pdf(uploaded_file)
        table_data = _parse_tables_to_data(tables)
        ref_cols = table_data.get("_column_info", {}).get("excluded_ref_cols", [])
        if ref_cols:
            notes.append(
                f"Note reference columns detected and excluded from table extraction: {ref_cols}"
            )
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

    # Inventory: prefer balance-sheet-sourced value
    if text_data.get("_inventory_source", "").startswith("balance_sheet") and "inventory" in text_data:
        merged["inventory"] = text_data["inventory"]
        merged["_inventory_source"] = text_data["_inventory_source"]

    # If inventory not found or is negative (likely P&L closing stock), clear it
    inv = merged.get("inventory")
    if inv is not None and inv < 0:
        notes.append(
            "Inventory value was negative — likely captured from P&L Closing Stock line. "
            "Setting to None. Please enter Balance Sheet inventory value in the form below."
        )
        merged["inventory"] = None
        merged["_inventory_source"] = "cleared_negative"

    # Count extracted fields
    extracted_count = sum(1 for k, v in merged.items()
                          if v is not None and not k.startswith("_"))
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
            "source_note": extracted_data.get(f"_{field}_source", ""),
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

    # Ensure inventory source is tagged as balance_sheet (user confirmed it)
    if result.get("inventory") is not None:
        result["_inventory_source"] = "balance_sheet/current_assets (user confirmed)"

    # Derive missing calculated fields
    if result.get("gross_profit") is None and result.get("revenue") and result.get("cogs"):
        result["gross_profit"] = result["revenue"] - result["cogs"]

    # Bug 2: Always derive EBIT from components (never trust a parsed EBIT line)
    if result.get("net_profit") is not None:
        tax = result.get("tax_expense") or 0
        interest = result.get("interest_expense") or 0
        computed_ebit = result["net_profit"] + tax + interest
        result["ebit"] = computed_ebit
        result["_ebit_components"] = {
            "net_profit": result["net_profit"],
            "interest_expense": interest,
            "tax_expense": tax,
        }

    if result.get("ebit") is not None:
        dep = result.get("depreciation") or 0
        result["ebitda"] = result["ebit"] + dep

    return result
