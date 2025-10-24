import json
import pathlib
import sys
import types
import unittest

sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))

from excel_copilot.tools import excel_tools


def _split_range(cell_range: str) -> tuple[str, str]:
    cleaned = cell_range.strip()
    if "!" in cleaned:
        _, cleaned = cleaned.split("!", 1)
    cleaned = cleaned.replace("$", "")
    if ":" in cleaned:
        start, end = cleaned.split(":", 1)
    else:
        start = end = cleaned
    return start.upper(), end.upper()


class FakeActions:
    def __init__(self) -> None:
        self.book = types.SimpleNamespace(fullname="C:\\fake\\book.xlsx")
        self._data: dict[str, dict[str, str]] = {
            "Sheet1": {
                "A1": "テスト  ",
                "B1": "",
            }
        }
        self.writes: dict[tuple[str, str], list[list[str]]] = {}
        self.logs: list[str] = []

    def read_range(self, cell_range: str, sheet_name: str | None = None) -> list[list[str]]:
        sheet = sheet_name or "Sheet1"
        start, end = _split_range(cell_range)
        if start != end:
            raise NotImplementedError("FakeActions only supports single-cell ranges in tests.")
        value = self._data.get(sheet, {}).get(start, "")
        return [[value]]

    def write_range(
        self,
        cell_range: str,
        data: list[list[str]],
        sheet_name: str | None = None,
        apply_formatting: bool = True,
    ) -> str:
        sheet = sheet_name or "Sheet1"
        if not data or not data[0]:
            raise AssertionError("Expected data for write_range.")
        if len(data) != 1 or len(data[0]) != 1:
            raise NotImplementedError("FakeActions only supports single-cell writes in tests.")
        start, end = _split_range(cell_range)
        if start != end:
            raise NotImplementedError("FakeActions only supports single-cell ranges in tests.")
        value = data[0][0]
        self._data.setdefault(sheet, {})[start] = value
        self.writes[(sheet, cell_range)] = data
        return f"range '{cell_range}' updated"

    def log_progress(self, message: str) -> None:
        self.logs.append(message)


class FakeBrowserManager:
    def __init__(self) -> None:
        self.prompts: list[str] = []

    def ask(self, prompt: str, stop_event=None) -> str:
        self.prompts.append(prompt)
        payload = [
            {
                "translated_text": "Test",
                "translated_length": 4,
                "length_ratio": 0.8,
                "length_verification": {
                    "result": json.dumps({"source_length": 5, "translated_length": 4, "ratio": 0.8}),
                    "status": "ok",
                },
            }
        ]
        return json.dumps(payload, ensure_ascii=False)


class TranslationLengthRatioTests(unittest.TestCase):
    def test_trailing_whitespace_does_not_trigger_length_limit(self) -> None:
        actions = FakeActions()
        browser = FakeBrowserManager()

        result = excel_tools.translate_range_without_references(
            actions=actions,
            browser_manager=browser,
            cell_range="Sheet1!A1:A1",
            sheet_name="Sheet1",
            translation_output_range="Sheet1!B1:B1",
            overwrite_source=False,
            length_ratio_limit=1.3,
            rows_per_batch=1,
        )

        self.assertTrue(result)
        self.assertIn(("Sheet1", "B1"), actions.writes)
        self.assertEqual(actions.writes[("Sheet1", "B1")], [["Test"]])
        self.assertFalse(any("Length adjustment task" in prompt for prompt in browser.prompts))


if __name__ == "__main__":
    unittest.main()
