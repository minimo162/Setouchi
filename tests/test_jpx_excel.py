import pandas as pd
from pathlib import Path

from setouchi.jpx_excel import detect_header_row, extract_available_companies


def create_sample_excel(path: Path) -> None:
    """Create a sample Excel file with header not on the first row."""
    # Rows before header
    pre_header = pd.DataFrame({"A": ["foo"], "B": ["bar"]})
    header = ["Company", "Annual Securities Reports", "Disclosure Status"]
    data = [
        ["Alpha Corp", "2023", "Available"],
        ["Beta Inc", "2023", "Not Available"],
    ]
    with pd.ExcelWriter(path) as writer:
        pre_header.to_excel(writer, index=False, header=False)
        pd.DataFrame(data, columns=header).to_excel(writer, index=False)


def test_detect_header_row(tmp_path: Path) -> None:
    excel = tmp_path / "sample.xlsx"
    create_sample_excel(excel)
    assert detect_header_row(excel) == 1


def test_extract_available_companies(tmp_path: Path) -> None:
    excel = tmp_path / "sample.xlsx"
    create_sample_excel(excel)
    result = extract_available_companies(excel)
    assert len(result) == 1
    assert result.iloc[0]["company"] == "Alpha Corp"
