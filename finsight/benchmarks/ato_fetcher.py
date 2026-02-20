"""
ATO Small Business Benchmark fetcher.
Attempts to load live data from ATO; falls back to bundled JSON.
"""

import json
import os
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

BENCHMARK_FILE = Path(__file__).parent / "ato_benchmarks.json"


def load_benchmarks() -> dict:
    """Load ATO benchmark data. Returns bundled JSON (live fetch not available via API)."""
    try:
        with open(BENCHMARK_FILE, "r") as f:
            data = json.load(f)
        return data
    except Exception as e:
        logger.error(f"Failed to load benchmark file: {e}")
        return {"industries": {}, "_metadata": {"note": "Benchmark data unavailable"}}


def get_industry_list() -> list:
    """Return sorted list of industry names for dropdown."""
    data = load_benchmarks()
    industries = list(data.get("industries", {}).keys())
    industries.sort()
    return industries


def get_industry_benchmarks(industry: str) -> dict:
    """Return benchmark ranges for a specific industry."""
    data = load_benchmarks()
    industries = data.get("industries", {})
    return industries.get(industry, industries.get("Other", {}))


def benchmark_status(actual_pct: float, low: float, high: float, higher_is_better: bool = True) -> str:
    """
    Determine traffic-light status vs ATO benchmark range.
    Returns 'green', 'amber', or 'red'.
    """
    if actual_pct is None:
        return "grey"

    in_range = low <= actual_pct <= high
    if in_range:
        return "green"

    # Calculate how far outside the range (as % of range width)
    range_width = max(high - low, 1)
    if actual_pct < low:
        deviation_pct = (low - actual_pct) / range_width * 100
    else:
        deviation_pct = (actual_pct - high) / range_width * 100

    # Within 20% of range width outside = amber, beyond = red
    if deviation_pct <= 20:
        return "amber"
    return "red"


def get_benchmark_metadata() -> dict:
    """Return metadata about the benchmark dataset."""
    data = load_benchmarks()
    return data.get("_metadata", {})
