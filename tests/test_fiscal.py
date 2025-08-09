from setouchi.fiscal import normalize_fiscal_year_end


def test_normalize_en_statement():
    assert normalize_fiscal_year_end("For the fiscal year ended March 31, 2024") == "2024-03-31"


def test_normalize_en_range():
    text = "FY2024 (April 1, 2023–March 31, 2024)"
    assert normalize_fiscal_year_end(text) == "2024-03-31"


def test_normalize_japanese_date():
    assert normalize_fiscal_year_end("2024年3月31日") == "2024-03-31"
