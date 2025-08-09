"""Utilities for normalising fiscal year end strings.

The design document (section 16.2) specifies that dates extracted from
various textual representations should be normalised to the ISO format
``YYYY-MM-DD``.  Only a couple of common patterns are implemented here:

* ``For the fiscal year ended March 31, 2024`` -> ``2024-03-31``
* ``FY2024 (April 1, 2023–March 31, 2024)`` -> ``2024-03-31``
* ``2024年3月31日`` -> ``2024-03-31``
"""
from __future__ import annotations

import re
from typing import Optional

_MONTHS = {
    "january": 1,
    "february": 2,
    "march": 3,
    "april": 4,
    "may": 5,
    "june": 6,
    "july": 7,
    "august": 8,
    "september": 9,
    "october": 10,
    "november": 11,
    "december": 12,
}


def _format_date(year: int, month: int, day: int) -> str:
    return f"{year:04d}-{month:02d}-{day:02d}"


def _parse_en_date(text: str) -> Optional[str]:
    m = re.search(r"([A-Za-z]+)\s+(\d{1,2}),\s*(\d{4})", text)
    if not m:
        return None
    month = _MONTHS.get(m.group(1).lower())
    if not month:
        return None
    day = int(m.group(2))
    year = int(m.group(3))
    return _format_date(year, month, day)


def normalize_fiscal_year_end(text: str) -> Optional[str]:
    """Normalise various representations of fiscal period end dates.

    Parameters
    ----------
    text:
        Source string containing a date expression.

    Returns
    -------
    str | None
        Normalised date in ``YYYY-MM-DD`` format or ``None`` if the
        string does not contain a recognised pattern.
    """
    if not text:
        return None

    # Japanese "YYYY年M月D日"
    m = re.search(r"(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日", text)
    if m:
        year, month, day = map(int, m.groups())
        return _format_date(year, month, day)

    # ``For the fiscal year ended March 31, 2024``
    m = re.search(r"fiscal year ended\s+([A-Za-z]+\s+\d{1,2},\s*\d{4})", text, re.I)
    if m:
        return _parse_en_date(m.group(1))

    # ``FY2024 (April 1, 2023–March 31, 2024)`` -> take the rightmost date
    m = re.search(r"\(([^)]+)\)", text)
    if m:
        parts = re.split(r"[\u2013-]", m.group(1))
        if parts:
            date = _parse_en_date(parts[-1])
            if date:
                return date

    return None


__all__ = ["normalize_fiscal_year_end"]
