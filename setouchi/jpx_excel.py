"""Utilities for parsing JPX English disclosure Excel list.

This module follows the design specified in ``design.md`` for detecting the header
row and extracting companies whose "Annual Securities Reports" disclosure status is
``Available``.
"""
from __future__ import annotations

from pathlib import Path
from typing import List

import pandas as pd


def _normalize_columns(columns: List[str]) -> List[str]:
    """Normalize column names by stripping, lowering and converting to string."""
    return [str(c).strip().lower() for c in columns]


def detect_header_row(path: str | Path, max_rows: int = 10) -> int:
    """Detect the header row in the Excel file.

    The header row is defined as the first row within ``max_rows`` that contains both
    ``"annual securities reports"`` and ``"disclosure status"`` in its column names.

    Parameters
    ----------
    path:
        Path to the Excel file.
    max_rows:
        Number of initial rows to search for the header.

    Returns
    -------
    int
        Zero-based index of the header row.

    Raises
    ------
    ValueError
        If no suitable header row is found within ``max_rows`` rows.
    """
    path = Path(path)
    for r in range(max_rows):
        df = pd.read_excel(path, header=r, nrows=0)
        cols = _normalize_columns(df.columns)
        joined = " ".join(cols)
        if "annual securities reports" in joined and "disclosure status" in joined:
            return r
    raise ValueError("Could not detect header row containing required columns")


def extract_available_companies(path: str | Path, max_rows: int = 10) -> pd.DataFrame:
    """Extract rows where Annual Securities Reports are marked as available.

    Parameters
    ----------
    path:
        Path to the JPX Excel file.
    max_rows:
        Number of initial rows to search for the header.

    Returns
    -------
    pandas.DataFrame
        DataFrame containing rows with ``Disclosure Status`` of ``Available``.
    """
    header_row = detect_header_row(path, max_rows=max_rows)
    df = pd.read_excel(path, header=header_row)
    # Normalize column names and rows
    cols = _normalize_columns(df.columns)
    df.columns = cols
    if "disclosure status" not in df.columns:
        raise ValueError("'Disclosure Status' column not found after normalization")
    rows: List[List[str]] = []
    for row in getattr(df, "_data", []):
        status = str(row.get("disclosure status", "")).strip().lower()
        if status == "available":
            rows.append([row.get(col, "") for col in df.columns])
    return pd.DataFrame(rows, columns=df.columns)


if __name__ == "__main__":  # pragma: no cover - CLI utility
    import argparse

    parser = argparse.ArgumentParser(description="Extract companies with available ASR")
    parser.add_argument("excel_path", help="Path to JPX English disclosure Excel file")
    args = parser.parse_args()

    available = extract_available_companies(args.excel_path)
    if available.empty:
        print("No companies with available Annual Securities Reports found.")
    else:
        print(available.to_csv(index=False))
