"""Setouchi package for financial document processing.

Only a tiny subset of the full design is implemented for the kata.  The
public API re-exports helpers for working with JPX Excel files, fiscal
year normalisation and section heading detection.
"""

from .fiscal import normalize_fiscal_year_end
from .jpx_excel import extract_available_companies, detect_header_row
from .section import normalise_section_heading

__all__ = [
    "normalize_fiscal_year_end",
    "extract_available_companies",
    "detect_header_row",
    "normalise_section_heading",
]

