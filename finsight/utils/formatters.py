"""
Shared number formatting utilities for consistent display across the app and exports.

Conventions:
  Currency  →  $1,234,567      (whole dollars, no decimals)
  Percent   →  42.3%           (1 decimal place)
  Ratio     →  1.85x           (2 decimal places)
  Days      →  47 days         (whole number)
"""


def format_currency(v, currency: str = "AUD") -> str:
    """Format a numeric value as whole-dollar currency: $1,234,567"""
    if v is None:
        return "N/A"
    try:
        return f"${float(v):,.0f}"
    except (TypeError, ValueError):
        return "N/A"


def format_percent(v) -> str:
    """Format a numeric value as a percentage with 1 decimal place: 42.3%"""
    if v is None:
        return "N/A"
    try:
        return f"{float(v):.1f}%"
    except (TypeError, ValueError):
        return "N/A"


def format_ratio(v) -> str:
    """Format a numeric value as a ratio with 2 decimal places: 1.85x"""
    if v is None:
        return "N/A"
    try:
        return f"{float(v):.2f}x"
    except (TypeError, ValueError):
        return "N/A"


def format_days(v) -> str:
    """Format a numeric value as whole-number days: 47 days"""
    if v is None:
        return "N/A"
    try:
        return f"{int(round(float(v)))} days"
    except (TypeError, ValueError):
        return "N/A"


def format_metric(v, format_type: str) -> str:
    """
    Dispatch to the correct formatter based on format_type string.
    Recognised types: 'currency', 'percentage', 'ratio', 'days'
    """
    if format_type == "currency":
        return format_currency(v)
    elif format_type == "percentage":
        return format_percent(v)
    elif format_type == "ratio":
        return format_ratio(v)
    elif format_type == "days":
        return format_days(v)
    return str(v) if v is not None else "N/A"
