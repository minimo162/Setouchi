import json
import pathlib
import re
import sys
import types
import unittest

if "playwright" not in sys.modules:
    playwright_stub = types.ModuleType("playwright")
    sync_api_stub = types.ModuleType("playwright.sync_api")

    def _missing_playwright(*args, **kwargs):
        raise ModuleNotFoundError("playwright is not available in the test environment")

    class _PlaywrightStub:
        pass

    class _PlaywrightTimeoutError(RuntimeError):
        pass

    sync_api_stub.sync_playwright = _missing_playwright
    sync_api_stub.Page = _PlaywrightStub
    sync_api_stub.BrowserContext = _PlaywrightStub
    sync_api_stub.Playwright = _PlaywrightStub
    sync_api_stub.TimeoutError = _PlaywrightTimeoutError
    sync_api_stub.Locator = _PlaywrightStub
    sync_api_stub.ElementHandle = _PlaywrightStub

    playwright_stub.sync_api = sync_api_stub

    sys.modules["playwright"] = playwright_stub
    sys.modules["playwright.sync_api"] = sync_api_stub

if "pyperclip" not in sys.modules:
    pyperclip_stub = types.ModuleType("pyperclip")

    def _missing_pyperclip(*args, **kwargs):
        raise ModuleNotFoundError("pyperclip is not available in the test environment")

    pyperclip_stub.copy = _missing_pyperclip
    pyperclip_stub.paste = _missing_pyperclip
    sys.modules["pyperclip"] = pyperclip_stub

if "xlwings" not in sys.modules:
    def _col_name(index: int) -> str:
        if index <= 0:
            return ""
        name = ""
        while index:
            index, remainder = divmod(index - 1, 26)
            name = chr(65 + remainder) + name
        return name

    xlwings_stub = types.ModuleType("xlwings")
    xlwings_stub.Range = type("Range", (), {})
    xlwings_stub.Sheet = type("Sheet", (), {})
    xlwings_stub.Book = type("Book", (), {})
    xlwings_stub.App = type("App", (), {})
    xlwings_stub.apps = []
    xlwings_stub.utils = types.SimpleNamespace(col_name=_col_name)

    sys.modules["xlwings"] = xlwings_stub

sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))

from excel_copilot.tools import excel_tools
from excel_copilot.core.exceptions import ToolExecutionError


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


_CELL_PATTERN = re.compile(r"([A-Z]+)(\d+)")


def _column_to_index(column_label: str) -> int:
    index = 0
    for char in column_label:
        index = index * 26 + (ord(char) - 64)
    return index - 1


def _index_to_column_label(index: int) -> str:
    index += 1
    label_chars: list[str] = []
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        label_chars.append(chr(65 + remainder))
    return "".join(reversed(label_chars))


def _parse_cell(cell_ref: str) -> tuple[int, int]:
    match = _CELL_PATTERN.fullmatch(cell_ref)
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_ref}")
    column_label, row_part = match.groups()
    return int(row_part), _column_to_index(column_label)


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
        start_ref, end_ref = _split_range(cell_range)
        start_row, start_col = _parse_cell(start_ref)
        end_row, end_col = _parse_cell(end_ref)
        rows: list[list[str]] = []
        for row_number in range(start_row, end_row + 1):
            row_values: list[str] = []
            for col_index in range(start_col, end_col + 1):
                cell_id = f"{_index_to_column_label(col_index)}{row_number}"
                row_values.append(self._data.get(sheet, {}).get(cell_id, ""))
            rows.append(row_values)
        return rows

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

        start_ref, end_ref = _split_range(cell_range)
        start_row, start_col = _parse_cell(start_ref)
        end_row, end_col = _parse_cell(end_ref)

        expected_rows = end_row - start_row + 1
        expected_cols = end_col - start_col + 1

        if len(data) != expected_rows:
            raise AssertionError("Row count mismatch in FakeActions.write_range.")
        for row in data:
            if len(row) != expected_cols:
                raise AssertionError("Column count mismatch in FakeActions.write_range.")

        for row_offset, row_values in enumerate(data):
            for col_offset, value in enumerate(row_values):
                row_number = start_row + row_offset
                col_index = start_col + col_offset
                cell_id = f"{_index_to_column_label(col_index)}{row_number}"
                self._data.setdefault(sheet, {})[cell_id] = value

        self.writes[(sheet, cell_range)] = data
        return f"range '{cell_range}' updated"

    def log_progress(self, message: str) -> None:
        self.logs.append(message)


class FakeBrowserManager:
    def __init__(self, responses: list[str] | None = None) -> None:
        self.prompts: list[str] = []
        self._responses: list[str] = list(responses) if responses else []

    def ask(self, prompt: str, stop_event=None) -> str:
        self.prompts.append(prompt)
        if self._responses:
            return self._responses.pop(0)
        payload = [
            {
                "translated_text": "Test  ",
                "source_length": 5,
                "translated_length": 6,
                "length_ratio": 1.2,
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
        self.assertEqual(actions.writes[("Sheet1", "B1")], [["Test  "]])
        self.assertEqual(len(browser.prompts), 1)
        self.assertFalse(any("Length adjustment task" in prompt for prompt in browser.prompts))

    def test_unescaped_length_verification_result_is_repaired(self) -> None:
        actions = FakeActions()
        malformed_response = (
            '[{"translated_text": "Test  ", "source_length": 5, "translated_length": 6, "length_ratio": 1.2, '
            '"length_verification": {"result": "{"source_length": 5, "translated_length": 6, "length_ratio": 1.2}", '
            '"status": "ok"}}]'
        )
        browser = FakeBrowserManager(responses=[malformed_response])

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
        self.assertEqual(actions.writes[("Sheet1", "B1")], [["Test  "]])
        self.assertEqual(len(browser.prompts), 1)

    def test_invalid_escape_sequences_are_sanitized(self) -> None:
        actions = FakeActions()
        actions._data["Sheet1"]["A1"] = "テスト"
        invalid_response = r"""
[
    {
        "translated_text": "Excluding credit applied",
        "source_length": 12,
        "translated_length": 28,
        "length_ratio": 2.33,
        "length_verification": {
            "method": "utf16-le",
            "translated_length_computed": 28,
            "length_ratio_computed": 2.33,
            "status": "ok"
        }
    },
    {
        "translated_text": "Allocation shortfall",
        "source_length": 7,
        "translated_length": 17,
        "length_ratio": 2.43,
        "length_verification": {
            "method": "utf16-le",
            "translated_length_computed": 17,
            "length_ratio_computed": 2.43,
            "status": "ok"
        }
    },
    {
        "translated_text": "Planned credit for shortfall",
        "source_length": 11,
        "translated_length": 27,
        "length_ratio": 2.45,
        "length_verification": {
            "method": "utf16-le",
            "translated_length_computed": 27,
            "length_ratio_computed": 2.45,
            "status": "ok"
        }
    },
    {
        "translated_text": "Shipped vehicle MY(\*)",
        "source_length": 10,
        "translated_length": 24,
        "length_ratio": 2.40,
        "length_verification": {
            "method": "utf16-le",
            "translated_length_computed": 24,
            "length_ratio_computed": 2.40,
            "status": "ok"
        }
    }
]
"""
        browser = FakeBrowserManager(responses=[invalid_response])

        result = excel_tools.translate_range_without_references(
            actions=actions,
            browser_manager=browser,
            cell_range="Sheet1!A1:A1",
            sheet_name="Sheet1",
            translation_output_range="Sheet1!B1:B1",
            overwrite_source=False,
            rows_per_batch=1,
        )

        self.assertTrue(result)
        self.assertIn(("Sheet1", "B1"), actions.writes)

    def test_prompt_includes_explicit_json_encoding_guidance(self) -> None:
        actions = FakeActions()
        browser = FakeBrowserManager()

        excel_tools.translate_range_without_references(
            actions=actions,
            browser_manager=browser,
            cell_range="Sheet1!A1:A1",
            sheet_name="Sheet1",
            translation_output_range="Sheet1!B1:B1",
            overwrite_source=False,
            length_ratio_limit=1.3,
            rows_per_batch=1,
        )

        self.assertTrue(
            any(
                '"source_length"' in prompt
                and '"translated_length"' in prompt
                and '"length_ratio"' in prompt
                for prompt in browser.prompts
            ),
            "Translation prompt should instruct the model to output source_length/translated_length/length_ratio fields.",
        )

    def test_playwright_prompt_for_tariff_impact_contains_ratio_guidance(self) -> None:
        actions = FakeActions()
        actions._data["Sheet1"]["A1"] = "関税影響"
        browser = FakeBrowserManager(
            responses=[
                json.dumps(
                    [
                        {
                            "translated_text": "Tariff hit",
                            "source_length": 4,
                            "translated_length": 10,
                            "length_ratio": 2.5,
                        }
                    ],
                    ensure_ascii=False,
                )
            ]
        )

        result = excel_tools.translate_range_without_references(
            actions=actions,
            browser_manager=browser,
            cell_range="Sheet1!A1:A1",
            sheet_name="Sheet1",
            translation_output_range="Sheet1!B1:B1",
            overwrite_source=False,
            length_ratio_limit=2.5,
            length_ratio_min=2.0,
        )

        self.assertTrue(result)
        self.assertIn(("Sheet1", "B1"), actions.writes)
        self.assertEqual(actions.writes[("Sheet1", "B1")], [["Tariff hit"]])
        self.assertGreaterEqual(len(browser.prompts), 1)

        prompt_text = browser.prompts[0]
        self.assertIn("許容文字数倍率レンジ: 2.00〜2.50。", prompt_text)
        self.assertIn("最小長 = ceil(source_length × 2.00)、最大長 = floor(source_length × 2.50)", prompt_text)
        self.assertIn("目標長 = round(source_length × 2.25)", prompt_text)
        self.assertIn("translated_length と length_verification.translated_length_computed は必ず len(translated_text.encode(\"utf-16-le\")) // 2 の実測値と完全一致させ", prompt_text)
        self.assertIn("length_ratio と length_verification.length_ratio_computed は実測 translated_length / source_length を基に再計算し", prompt_text)
        self.assertIn("例示や過去応答の数値をコピーせず、毎回 translated_text の実測値から計算した数値のみを記入してください", prompt_text)
        self.assertIn("len(translated_text.encode(\"utf-16-le\")) // 2 を再測定し、その実測値で translated_length と length_verification.translated_length_computed を上書きし", prompt_text)
        self.assertIn("length_verification.status は translated_length・length_ratio が実測値と一致し許容レンジ内である場合に限り \"ok\"", prompt_text)
        self.assertIn("列挙は 'and' ではなくコンマやスラッシュを用いて簡潔に区切ってください", prompt_text)
        self.assertIn('Source sentences:\n["関税影響"]', prompt_text)

    def test_last_json_array_is_selected_when_multiple_payloads_returned(self) -> None:
        actions = FakeActions()
        actions._data["Sheet1"]["A1"] = "関税影響"
        intermediate_payload = [
            {
                "translated_text": "Tariff impact",
                "source_length": 4,
                "translated_length": 13,
                "length_ratio": 3.25,
                "length_verification": {
                    "method": "utf16-le",
                    "translated_length_computed": 13,
                    "length_ratio_computed": 3.25,
                    "status": "mismatch",
                },
            }
        ]
        final_payload = [
            {
                "translated_text": "Tariff fee",
                "source_length": 4,
                "translated_length": 10,
                "length_ratio": 2.5,
                "length_verification": {
                    "method": "utf16-le",
                    "translated_length_computed": 10,
                    "length_ratio_computed": 2.5,
                    "status": "ok",
                },
            }
        ]
        browser = FakeBrowserManager(
            responses=[
                json.dumps(intermediate_payload, ensure_ascii=False)
                + "\n"
                + json.dumps(final_payload, ensure_ascii=False)
            ]
        )

        result = excel_tools.translate_range_without_references(
            actions=actions,
            browser_manager=browser,
            cell_range="Sheet1!A1:A1",
            sheet_name="Sheet1",
            translation_output_range="Sheet1!B1:B1",
            overwrite_source=False,
            length_ratio_limit=2.5,
            length_ratio_min=2.0,
        )

        self.assertTrue(result)
        self.assertIn(("Sheet1", "B1"), actions.writes)
        self.assertEqual(actions.writes[("Sheet1", "B1")], [["Tariff fee"]])
        self.assertEqual(len(browser.prompts), 1)

    def test_length_metadata_mismatch_is_auto_corrected(self) -> None:
        actions = FakeActions()
        actions._data["Sheet1"]["A1"] = "テスト  "
        source_units = len(actions._data["Sheet1"]["A1"])

        bad_payload = [
            {
                "translated_text": "abcde",
                "source_length": source_units,
                "translated_length": 3,  # incorrect metadata
                "length_ratio": 1.0,
            }
        ]
        browser = FakeBrowserManager(
            responses=[
                json.dumps(bad_payload, ensure_ascii=False),
            ]
        )

        result = excel_tools.translate_range_without_references(
            actions=actions,
            browser_manager=browser,
            cell_range="Sheet1!A1:A1",
            sheet_name="Sheet1",
            translation_output_range="Sheet1!B1:B1",
            overwrite_source=False,
            length_ratio_limit=2.5,
            rows_per_batch=1,
        )

        self.assertTrue(result)
        self.assertIn(("Sheet1", "B1"), actions.writes)
        self.assertEqual(actions.writes[("Sheet1", "B1")], [["abcde"]])
        self.assertEqual(
            len(browser.prompts),
            1,
            "Metadata mismatch should be auto-corrected without requesting a retry.",
        )
        self.assertTrue(
            any("metadata auto-corrected" in log for log in actions.logs),
            "Auto-correction event should be logged for traceability.",
        )

    def test_length_ratio_violation_continues_without_retry(self) -> None:
        actions = FakeActions()
        actions._data["Sheet1"]["A1"] = "テスト  "
        source_units = len(actions._data["Sheet1"]["A1"])

        bad_payload = [
            {
                "translated_text": "abcdefghij",
                "source_length": source_units,
                "translated_length": 3,
                "length_ratio": 1.0,
            }
        ]
        good_ratio = 4 / source_units
        good_payload = [
            {
                "translated_text": "abcd",
                "source_length": source_units,
                "translated_length": 4,
                "length_ratio": good_ratio,
            }
        ]
        browser = FakeBrowserManager(
            responses=[
                json.dumps(bad_payload, ensure_ascii=False),
                json.dumps(good_payload, ensure_ascii=False),
            ]
        )

        result = excel_tools.translate_range_without_references(
            actions=actions,
            browser_manager=browser,
            cell_range="Sheet1!A1:A1",
            sheet_name="Sheet1",
            translation_output_range="Sheet1!B1:B1",
            overwrite_source=False,
            length_ratio_limit=1.5,
            rows_per_batch=1,
        )

        self.assertTrue(result)
        self.assertIn(("Sheet1", "B1"), actions.writes)
        self.assertEqual(actions.writes[("Sheet1", "B1")], [["abcdefghij"]])
        self.assertEqual(len(browser.prompts), 1, "Length ratio violation should no longer trigger a retry.")
        self.assertTrue(
            any("文字数倍率制約を逸脱した訳文を検出しました" in log for log in actions.logs),
            "Progress log should note that the limit was violated but processing continued.",
        )
        self.assertIn("文字数倍率制約の警告", result)

    def test_length_ratio_violation_with_persistent_payload_does_not_retry(self) -> None:
        actions = FakeActions()
        actions._data["Sheet1"]["A1"] = "テスト  "
        source_units = len(actions._data["Sheet1"]["A1"])

        bad_payload = [
            {
                "translated_text": "abcdefghij",
                "source_length": source_units,
                "translated_length": 3,
                "length_ratio": 1.0,
            }
        ]
        responses = [json.dumps(bad_payload, ensure_ascii=False)] * 4
        browser = FakeBrowserManager(responses=responses)

        result = excel_tools.translate_range_without_references(
            actions=actions,
            browser_manager=browser,
            cell_range="Sheet1!A1:A1",
            sheet_name="Sheet1",
            translation_output_range="Sheet1!B1:B1",
            overwrite_source=False,
            length_ratio_limit=1.5,
            rows_per_batch=1,
        )

        self.assertTrue(result)
        self.assertIn(("Sheet1", "B1"), actions.writes)
        self.assertEqual(actions.writes[("Sheet1", "B1")], [["abcdefghij"]])
        self.assertEqual(len(browser.prompts), 1, "Retries should not be attempted even when violations persist.")
        self.assertFalse(
            any("最も近い訳文を採用します" in log for log in actions.logs),
            "Fallback adoption log should no longer be emitted.",
        )
        self.assertTrue(
            any("文字数倍率制約を逸脱した訳文を検出しました" in log for log in actions.logs),
            "Violation warning should be logged even when retries are skipped.",
        )
        self.assertIn("文字数倍率制約の警告", result)

    def test_existing_translation_is_replaced_by_cached_translation(self) -> None:
        actions = FakeActions()
        actions._data["Sheet1"].update(
            {
                "A1": "テスト",
                "A2": "テスト",
                "B1": "",
                "B2": "This translated sentence is intentionally far longer than the original text.",
            }
        )
        source_units = len(actions._data["Sheet1"]["A1"])
        translation_payload = [
            {
                "translated_text": "Test",
                "source_length": source_units,
                "translated_length": len("Test"),
                "length_ratio": len("Test") / max(source_units, 1),
            }
        ]
        browser = FakeBrowserManager(
            responses=[json.dumps(translation_payload, ensure_ascii=False)]
        )

        result = excel_tools.translate_range_without_references(
            actions=actions,
            browser_manager=browser,
            cell_range="Sheet1!A1:A2",
            sheet_name="Sheet1",
            translation_output_range="Sheet1!B1:B2",
            overwrite_source=False,
            length_ratio_limit=1.8,
            rows_per_batch=1,
        )

        self.assertTrue(result)
        self.assertEqual(len(browser.prompts), 1, "Only the first row should trigger a translation prompt when cache is reused.")
        self.assertEqual(actions._data["Sheet1"]["B1"], "Test")
        self.assertEqual(actions._data["Sheet1"]["B2"], "Test")

    def test_existing_translation_is_retranslated_when_stale(self) -> None:
        actions = FakeActions()
        actions._data["Sheet1"]["A1"] = "テスト"
        actions._data["Sheet1"]["B1"] = "Old translation"
        source_units = len(actions._data["Sheet1"]["A1"])
        new_payload = [
            {
                "translated_text": "New translation",
                "source_length": source_units,
                "translated_length": len("New translation"),
                "length_ratio": len("New translation") / max(source_units, 1),
            }
        ]
        browser = FakeBrowserManager(
            responses=[json.dumps(new_payload, ensure_ascii=False)]
        )

        result = excel_tools.translate_range_without_references(
            actions=actions,
            browser_manager=browser,
            cell_range="Sheet1!A1:A1",
            sheet_name="Sheet1",
            translation_output_range="Sheet1!B1:B1",
            overwrite_source=False,
            rows_per_batch=1,
        )

        self.assertTrue(result)
        self.assertIn(("Sheet1", "B1"), actions.writes)
        self.assertEqual(actions.writes[("Sheet1", "B1")], [["New translation"]])
        self.assertEqual(len(browser.prompts), 1, "Stale translations should trigger a fresh prompt.")
        self.assertTrue(
            any("既存の訳文が見つかりましたが、新しい訳文を取得します" in log for log in actions.logs),
            "Progress log should note that a stale translation was replaced.",
        )

if __name__ == "__main__":
    unittest.main()
