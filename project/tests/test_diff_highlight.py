import unittest
from typing import Any, Dict, List, Optional

from excel_copilot.tools.actions import ExcelActions


class DummyCharacterFont:
    def __init__(self, record: Dict[str, Any], key: str):
        self._record = record
        self._key = key

    def __getattr__(self, name: str):
        if name.lower() == "color":
            return self._record.get(self._key)
        raise AttributeError(name)

    def __setattr__(self, name: str, value: Any) -> None:
        if name in {"_record", "_key"}:
            super().__setattr__(name, value)
            return
        if name.lower() == "color":
            self._record[self._key] = value
        else:
            raise AttributeError(name)


class DummyCharacterAPI:
    def __init__(self, record: Dict[str, Any]):
        self.Font = DummyCharacterFont(record, "color")


class DummyCharacter:
    def __init__(self, record: Dict[str, Any]):
        self.api = DummyCharacterAPI(record)
        self.font = DummyCharacterFont(record, "tuple_color")


class DummyCharactersAccessor:
    def __init__(self):
        self.requests: List[Dict[str, Any]] = []

    def __getitem__(self, key):
        if isinstance(key, tuple):
            start, length = key
        else:
            start, length = key, 1
        record = {"start": start, "length": length, "color": None, "tuple_color": None}
        self.requests.append(record)
        return DummyCharacter(record)


class DummyCell:
    def __init__(self, value: str):
        self.value = value
        self.characters = DummyCharactersAccessor()


class DummyRangeAxis:
    def __init__(self, count: int):
        self.count = count


class DummyRange:
    def __init__(self, cells: List[List[DummyCell]]):
        self._cells = cells
        self.rows = DummyRangeAxis(len(cells))
        self.columns = DummyRangeAxis(len(cells[0]) if cells else 0)

    def __getitem__(self, idx):
        r_idx, c_idx = idx
        return self._cells[r_idx][c_idx]


class DummySheet:
    def __init__(self, range_map: Dict[str, DummyRange]):
        self._range_map = range_map

    def range(self, cell_range: str) -> DummyRange:
        return self._range_map[cell_range]


class DummySheets:
    def __init__(self, sheet: DummySheet):
        self._sheet = sheet

    @property
    def active(self) -> DummySheet:
        return self._sheet

    def __getitem__(self, key: Any) -> DummySheet:
        return self._sheet


class DummyWorkbook:
    def __init__(self, sheet: DummySheet):
        self.sheets = DummySheets(sheet)


class DummyExcelManager:
    def __init__(self, sheet: DummySheet):
        self._book = DummyWorkbook(sheet)

    def get_active_workbook(self) -> DummyWorkbook:
        return self._book


class ApplyDiffHighlightColorsTests(unittest.TestCase):
    def setUp(self):
        self.cell = DummyCell(value="X" * 2500)
        dummy_range = DummyRange([[self.cell]])
        dummy_sheet = DummySheet({"A1:A1": dummy_range})
        manager = DummyExcelManager(dummy_sheet)
        self.actions = ExcelActions(manager)

        def _get_sheet(_self, name: Optional[str] = None) -> DummySheet:  # type: ignore[override]
            return dummy_sheet

        self.actions._get_sheet = _get_sheet.__get__(self.actions, ExcelActions)

    def test_highlights_with_japanese_types(self):
        spans = [
            {"start": 10, "length": 5, "type": "削除"},
            {"start": 20, "length": 7, "type": "追加"},
        ]
        style_matrix = [[spans]]

        self.actions.apply_diff_highlight_colors("A1:A1", style_matrix)

        self.assertEqual(len(self.cell.characters.requests), 12)

        deletion_records = self.cell.characters.requests[:5]
        addition_records = self.cell.characters.requests[5:]

        for offset, record in enumerate(deletion_records):
            self.assertEqual(record["start"], 10 + 1 + offset)
            self.assertEqual(record["length"], 1)
            self.assertEqual(record["color"], 0x2828C6)
            self.assertIsNone(record["tuple_color"])  # COM path used

        for offset, record in enumerate(addition_records):
            self.assertEqual(record["start"], 20 + 1 + offset)
            self.assertEqual(record["length"], 1)
            self.assertEqual(record["color"], 0x205E1B)

    def test_skips_unknown_types(self):
        spans = [
            {"start": 5, "length": 3, "type": "unknown"},
        ]
        style_matrix = [[spans]]

        self.actions.apply_diff_highlight_colors("A1:A1", style_matrix)

        self.assertEqual(len(self.cell.characters.requests), 0)


if __name__ == "__main__":
    unittest.main()
