import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import xlwings as xw

from excel_copilot.tools.actions import ExcelActions
from excel_copilot.tools.excel_tools import _build_diff_highlight


class DummyManager:
    def __init__(self, book):
        self._book = book

    def get_active_workbook(self):
        return self._book


def _color_to_rgb(color_value):
    if color_value is None:
        return None
    if isinstance(color_value, tuple):
        return tuple(int(c) for c in color_value)
    if isinstance(color_value, int):
        b = (color_value >> 16) & 0xFF
        g = (color_value >> 8) & 0xFF
        r = color_value & 0xFF
        return (r, g, b)
    raise TypeError(f"Unsupported color type: {type(color_value)}")


def verify_span_colors(text, spans, sheet, cell_label):
    color_map = {
        "addition": (21, 101, 192),
        "追加": (21, 101, 192),
        "deletion": (198, 40, 40),
        "削除": (198, 40, 40),
    }
    cell = sheet.range(cell_label)

    def get_char_rgb(one_based_index: int):
        try:
            value = cell.characters[one_based_index - 1].font.color
            if value is not None:
                return tuple(int(c) for c in value)
        except Exception:
            pass
        value = cell.api.Characters(one_based_index, 1).Font.Color
        return _color_to_rgb(value)

    for span in spans:
        expected_rgb = color_map.get(span.get("type"))
        if expected_rgb is None:
            raise AssertionError(f"Unexpected span type: {span.get('type')}")
        start = span["start"]
        length = span["length"]
        for offset in range(length):
            idx = start + offset + 1
            rgb = get_char_rgb(idx)
            if rgb != expected_rgb:
                raise AssertionError(
                    f"Color mismatch at index {idx} (0-based {start + offset}): expected {expected_rgb}, got {rgb}"
                )

    if spans:
        first_span = spans[0]
        if first_span["start"] > 0:
            before_rgb = get_char_rgb(first_span["start"])
            if before_rgb != (0, 0, 0):
                raise AssertionError(f"Expected black before span; got {before_rgb}")
        last_span = spans[-1]
        end_index = last_span["start"] + last_span["length"]
        if end_index < len(text):
            after_rgb = get_char_rgb(end_index + 1)
            if after_rgb != (0, 0, 0):
                raise AssertionError(f"Expected black after span; got {after_rgb}")


def main():
    app = xw.App(visible=False, add_book=False)
    book = None
    try:
        book = xw.Book()
        sheet = book.sheets[0]
        manager = DummyManager(book)
        actions = ExcelActions(manager)

        scenarios = [
            ("abc", "abXc", "A1"),
            ("こんにちは世界", "こんXYにちは世界", "A2"),
            ("テスト", "テスト", "A3"),
        ]
        for before, after, cell_label in scenarios:
            highlight_text, spans = _build_diff_highlight(before, after)
            sheet.range(cell_label).value = highlight_text
            actions.apply_diff_highlight_colors(
                cell_label,
                [[spans]],
                sheet_name=sheet.name,
                addition_color_hex="#1565C0",
                deletion_color_hex="#C62828",
            )
            if spans:
                verify_span_colors(highlight_text, spans, sheet, cell_label)

        print("All highlight span color checks passed.")
    finally:
        if book is not None:
            book.close()
        app.quit()


if __name__ == "__main__":
    main()
