import pytest

from setouchi.section import normalise_section_heading


def test_exact_match_business_overview():
    assert normalise_section_heading("Business Overview") == "business_overview"


def test_fuzzy_match_management_analysis():
    heading = "Managements Discussion & Analysis"
    assert normalise_section_heading(heading) == "management_analysis"


def test_no_match_returns_none():
    assert normalise_section_heading("Completely unrelated heading") is None
