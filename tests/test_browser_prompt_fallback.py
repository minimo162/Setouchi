import logging
import pathlib
import sys
import unittest

sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))

from excel_copilot.core.browser_copilot_manager import BrowserCopilotManager


class FakeLocator:
    def __init__(self, should_fail: bool = False) -> None:
        self.should_fail = should_fail
        self.last_payload: dict[str, str] | None = None
        self.focus_called = False

    def focus(self) -> None:
        self.focus_called = True

    def evaluate(self, script: str, payload: dict[str, str]) -> bool:
        if self.should_fail:
            raise RuntimeError("evaluate failed")
        self.last_payload = payload
        return True


class ForceSetChatInputTests(unittest.TestCase):
    def setUp(self) -> None:
        self.manager = BrowserCopilotManager.__new__(BrowserCopilotManager)  # type: ignore[call-arg]
        self.manager.page = None
        self.manager._logger = logging.getLogger(__name__)

    def test_force_set_chat_input_text_records_sanitized_prompt(self) -> None:
        locator = FakeLocator()
        prompt = "Line1\r\nLine2  "

        result = self.manager._force_set_chat_input_text(locator, prompt)

        self.assertTrue(result)
        self.assertTrue(locator.focus_called)
        assert locator.last_payload is not None
        self.assertEqual(locator.last_payload.get("text"), "Line1\nLine2  ")

    def test_force_set_chat_input_text_handles_failure(self) -> None:
        locator = FakeLocator(should_fail=True)

        result = self.manager._force_set_chat_input_text(locator, "prompt")

        self.assertFalse(result)


if __name__ == "__main__":
    unittest.main()
