import unittest
from typing import Optional

import xlwings as xw

from excel_copilot.tools.actions import ExcelActions


class _ManagerStub:
    def __init__(self, workbook):
        self._workbook = workbook

    def get_active_workbook(self):
        return self._workbook


class ApplyDiffHighlightColorsIntegrationTests(unittest.TestCase):
    def setUp(self):
        self.app = xw.App(visible=False, add_book=True)
        self.book = self.app.books.active
        self.sheet = self.book.sheets[0]
        self.sheet.name = "HighlightTest"
        self.sheet.range("A1").value = "削除追加テストサンプル"
        manager = _ManagerStub(self.book)
        self.actions = ExcelActions(manager)

    def tearDown(self):
        try:
            self.book.close()
        finally:
            self.app.quit()

    def test_apply_colors_in_excel(self):
        sheet_name: Optional[str] = self.sheet.name
        cell_range = "A1:A1"
        spans = [
            {"start": 0, "length": 3, "type": "削除"},
            {"start": 3, "length": 3, "type": "追加"},
        ]
        style_matrix = [[spans]]

        self.actions.apply_diff_highlight_colors(cell_range, style_matrix, sheet_name=sheet_name)

        deletion_colors = {self.sheet.range("A1").characters[i].font.color for i in range(1, 4)}
        addition_colors = {self.sheet.range("A1").characters[i].font.color for i in range(4, 7)}

        self.assertEqual(deletion_colors, {(198, 40, 40)})
        self.assertEqual(addition_colors, {(27, 94, 32)})


if __name__ == "__main__":
    unittest.main()
