import queue
import types
import unittest

from desktop_app import (
    CopilotApp,
    CopilotMode,
    FORM_TOOL_NAMES,
    MISSING_CONTEXT_ERROR_MESSAGE,
)
from excel_copilot.ui.messages import AppState


def _build_app_for_mode(mode: CopilotMode, control_values: dict[str, str]) -> CopilotApp:
    app = CopilotApp.__new__(CopilotApp)  # type: ignore[call-arg]
    app.mode = mode
    app.form_controls = {name: types.SimpleNamespace(value=value) for name, value in control_values.items()}
    app.current_workbook_name = None
    app.current_sheet_name = None
    return app


class CollectFormPayloadTests(unittest.TestCase):
    def test_translation_mode_provides_default_target_language(self) -> None:
        control_values = {
            "cell_range": "A2:A5",
            "translation_output_range": "B2:B5",
        }
        app = _build_app_for_mode(CopilotMode.TRANSLATION, control_values)

        payload, error_message, summary_arguments = CopilotApp._collect_form_payload(app)

        self.assertIsNone(error_message)
        self.assertIsNotNone(payload)
        assert payload is not None  # help type checkers
        self.assertEqual(payload["tool_name"], FORM_TOOL_NAMES[CopilotMode.TRANSLATION])
        self.assertEqual(payload["arguments"]["cell_range"], "A2:A5")
        self.assertEqual(payload["arguments"]["translation_output_range"], "B2:B5")
        self.assertEqual(payload["arguments"]["target_language"], "English")
        self.assertIsNotNone(summary_arguments)
        assert summary_arguments is not None
        self.assertEqual(summary_arguments["translation_output_range"], "B2:B5")

    def test_reference_mode_requires_reference_urls(self) -> None:
        control_values = {
            "cell_range": "A2:A5",
            "translation_output_range": "B2:D5",
            "source_reference_urls": "",
            "target_reference_urls": "",
        }
        app = _build_app_for_mode(CopilotMode.TRANSLATION_WITH_REFERENCES, control_values)

        payload, error_message, summary_arguments = CopilotApp._collect_form_payload(app)

        self.assertIsNone(payload)
        self.assertIsNone(summary_arguments)
        self.assertEqual(error_message, "参照URLを1件以上入力してください。")

    def test_review_mode_maps_to_quality_tool(self) -> None:
        control_values = {
            "source_range": "B2:B5",
            "translated_range": "C2:C5",
            "review_output_range": "D2:G5",
        }
        app = _build_app_for_mode(CopilotMode.REVIEW, control_values)

        payload, error_message, summary_arguments = CopilotApp._collect_form_payload(app)

        self.assertIsNone(error_message)
        self.assertIsNotNone(payload)
        assert payload is not None
        arguments = payload["arguments"]
        self.assertEqual(payload["tool_name"], FORM_TOOL_NAMES[CopilotMode.REVIEW])
        self.assertEqual(arguments["source_range"], "B2:B5")
        self.assertEqual(arguments["status_output_range"], "D2:D5")
        self.assertEqual(arguments["issue_output_range"], "E2:E5")
        self.assertEqual(arguments["highlight_output_range"], "F2:F5")
        self.assertEqual(arguments["corrected_output_range"], "G2:G5")
        self.assertIsNotNone(summary_arguments)
        assert summary_arguments is not None
        self.assertEqual(summary_arguments["review_output_range"], "D2:G5")

    def test_review_mode_allows_three_column_range(self) -> None:
        control_values = {
            "source_range": "B2:B5",
            "translated_range": "C2:C5",
            "review_output_range": "D2:F5",
        }
        app = _build_app_for_mode(CopilotMode.REVIEW, control_values)

        payload, error_message, summary_arguments = CopilotApp._collect_form_payload(app)

        self.assertIsNone(error_message)
        self.assertIsNotNone(payload)
        assert payload is not None
        arguments = payload["arguments"]
        self.assertEqual(arguments["status_output_range"], "D2:D5")
        self.assertEqual(arguments["issue_output_range"], "E2:E5")
        self.assertEqual(arguments["highlight_output_range"], "F2:F5")
        self.assertNotIn("corrected_output_range", arguments)
        self.assertEqual(summary_arguments["review_output_range"], "D2:F5")

    def test_submit_form_requires_active_context(self) -> None:
        app = CopilotApp.__new__(CopilotApp)  # type: ignore[call-arg]
        app.app_state = AppState.READY
        app.current_workbook_name = None
        app.current_sheet_name = None
        app.request_queue = queue.Queue()
        app._last_error = ""

        def fake_set_form_error(message: str) -> None:
            app._last_error = message

        collect_called = {"value": False}

        def fake_collect_form_payload():
            collect_called["value"] = True
            return {}, None, {}

        app._set_form_error = fake_set_form_error  # type: ignore[attr-defined]
        app._update_ui = lambda: None  # type: ignore[attr-defined]
        app._collect_form_payload = fake_collect_form_payload  # type: ignore[attr-defined]

        app._submit_form(None)  # type: ignore[attr-defined]

        self.assertFalse(collect_called["value"])
        self.assertEqual(app._last_error, MISSING_CONTEXT_ERROR_MESSAGE)
        self.assertTrue(app.request_queue.empty())


if __name__ == "__main__":
    unittest.main()
