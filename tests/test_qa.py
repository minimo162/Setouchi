import pytest

from setouchi import (
    extract_numeric_tokens,
    find_note_references,
    check_table_column_consistency,
)


def test_extract_numeric_tokens():
    text = "Revenue increased to 1,234 million yen, up 10% from 1,100."
    assert extract_numeric_tokens(text) == ["1,234", "10%", "1,100"]


def test_find_note_references():
    text = "See Note 5 and 注 7 for details."
    assert find_note_references(text) == ["Note 5", "注 7"]


def test_check_table_column_consistency():
    table = """| A | B | C |\n|---|---|---|\n|1|2|3|"""
    assert check_table_column_consistency(table)
    bad = """| A | B |\n|---|---|\n|1|2|3|"""
    assert not check_table_column_consistency(bad)
