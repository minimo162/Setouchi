"""Basic QA utilities for document consistency checks.

This module implements simple helpers outlined in ``design.md`` section 16.4
for automated validation of translated Annual Securities Reports.  The
functions focus on three aspects:

* extracting numeric tokens (including percentages)
* detecting note references (``Note 5`` / ``注 5``)
* verifying that Markdown tables have consistent column counts
"""
from __future__ import annotations

import re
from typing import List

_NUMERIC_RE = re.compile(r"[-+]?\d[\d,]*(?:\.\d+)?%?")
_NOTE_RE = re.compile(r"(?:注\s*\d+|Note\s*\d+)", re.I)


def extract_numeric_tokens(text: str) -> List[str]:
    """Return a list of numeric tokens found in ``text``.

    The regex roughly mirrors the pattern suggested in the design document,
    capturing optional sign, commas, decimal parts and trailing percentage
    symbols.
    """
    if not text:
        return []
    return _NUMERIC_RE.findall(text)


def find_note_references(text: str) -> List[str]:
    """Detect note references such as ``Note 5`` or ``注 5`` within ``text``."""
    if not text:
        return []
    return _NOTE_RE.findall(text)


def check_table_column_consistency(table: str) -> bool:
    """Check whether all rows in a Markdown table have the same column count."""
    counts: List[int] = []
    for line in table.splitlines():
        if "|" not in line:
            continue
        parts = [c.strip() for c in line.strip().strip("|").split("|")]
        if not any(parts):
            continue
        counts.append(len(parts))
    return len(set(counts)) <= 1


__all__ = [
    "extract_numeric_tokens",
    "find_note_references",
    "check_table_column_consistency",
]
