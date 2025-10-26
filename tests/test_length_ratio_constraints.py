import json
import pathlib
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

    def test_length_ratio_violation_triggers_retry(self) -> None:
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
        self.assertEqual(actions.writes[("Sheet1", "B1")], [["abcd"]])
        self.assertEqual(len(browser.prompts), 2, "Length ratio violation should trigger exactly one retry.")
        self.assertTrue(
            any("translated_length と length_ratio" in prompt for prompt in browser.prompts[1:]),
            "Retry prompt should mention length ratio adjustment.",
        )

    def test_length_ratio_mismatch_raises_after_max_retries(self) -> None:
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

        with self.assertRaises(excel_tools.ToolExecutionError):
            excel_tools.translate_range_without_references(
                actions=actions,
                browser_manager=browser,
                cell_range="Sheet1!A1:A1",
                sheet_name="Sheet1",
                translation_output_range="Sheet1!B1:B1",
                overwrite_source=False,
                length_ratio_limit=1.5,
                rows_per_batch=1,
            )

        self.assertNotIn(("Sheet1", "B1"), actions.writes)
        self.assertEqual(len(browser.prompts), 4, "Should attempt initial call plus three retries.")

    def test_existing_translation_outside_length_limit_raises(self) -> None:
        actions = FakeActions()
        actions._data["Sheet1"]["A1"] = "テスト"
        actions._data["Sheet1"]["B1"] = "This translated sentence is intentionally far longer than the original text."
        browser = FakeBrowserManager()

        with self.assertRaises(ToolExecutionError) as ctx:
            excel_tools.translate_range_without_references(
                actions=actions,
                browser_manager=browser,
                cell_range="Sheet1!A1:A1",
                sheet_name="Sheet1",
                translation_output_range="Sheet1!B1:B1",
                overwrite_source=False,
                length_ratio_limit=1.2,
                rows_per_batch=1,
            )

        self.assertIn("Length ratio constraint violation", str(ctx.exception))
        self.assertEqual(len(browser.prompts), 0, "Reused translations should not trigger new prompts.")

if __name__ == "__main__":
    unittest.main()
