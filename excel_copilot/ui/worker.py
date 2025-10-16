"""Background worker that coordinates Excel Copilot tasks."""

from __future__ import annotations

import inspect
import queue
import threading
import traceback
from typing import Any, Dict, List, Optional, Union

from excel_copilot.agent.prompts import CopilotMode
from excel_copilot.agent.react_agent import ReActAgent
from excel_copilot.config import (
    COPILOT_BROWSER_CHANNELS,
    COPILOT_HEADLESS,
    COPILOT_PAGE_GOTO_TIMEOUT_MS,
    COPILOT_SLOW_MO_MS,
    COPILOT_USER_DATA_DIR,
)
from excel_copilot.core.browser_copilot_manager import BrowserCopilotManager
from excel_copilot.core.excel_manager import ExcelConnectionError
from excel_copilot.tools import excel_tools
from excel_copilot.tools.schema_builder import create_tool_schema

from .messages import RequestMessage, RequestType, ResponseMessage, ResponseType


class CopilotWorker:
    def __init__(
        self,
        request_q: queue.Queue,
        response_q: queue.Queue,
        sheet_name: Optional[str] = None,
        workbook_name: Optional[str] = None,
    ):
        self.request_queue = request_q
        self.response_queue = response_q
        self.browser_manager: Optional[BrowserCopilotManager] = None
        self.agent: Optional[ReActAgent] = None
        self.stop_event = threading.Event()
        self.sheet_name = sheet_name
        self.workbook_name = workbook_name
        self.mode = CopilotMode.TRANSLATION_WITH_REFERENCES
        self.tool_functions: List[callable] = []
        self.tool_schemas: List[Dict[str, Any]] = []

    def run(self):
        try:
            self._initialize()
            if self.agent and self.browser_manager:
                self._main_loop()
        except Exception as e:
            print(f"Critical error in worker run method: {e}")
            traceback.print_exc()
            self._emit_response(ResponseMessage(ResponseType.ERROR, f"\u81f4\u547d\u7684\u306a\u5b9f\u884c\u6642\u30a8\u30e9\u30fc: {e}"))
        finally:
            self._cleanup()

    def _emit_response(self, message: Union[ResponseMessage, Dict[str, Any]]):
        try:
            self.response_queue.put(ResponseMessage.from_raw(message))
        except Exception as err:
            print(f"Failed to enqueue response: {err}")

    def _build_agent(self):
        if not self.browser_manager or not self.tool_functions or not self.tool_schemas:
            return
        self.agent = ReActAgent(
            tools=self.tool_functions,
            tool_schemas=self.tool_schemas,
            browser_manager=self.browser_manager,
            sheet_name=self.sheet_name,
            workbook_name=self.workbook_name,
            mode=self.mode,
            progress_callback=lambda msg: self._emit_response(ResponseMessage(ResponseType.OBSERVATION, msg)),
        )

    def _format_user_prompt(self, user_input: str) -> str:
        trimmed_input = (user_input or "").strip()
        if self.mode is CopilotMode.TRANSLATION:
            prefix_lines = [
                "[Translation (No References) Mode Request]",
                "- Solve this by calling `translate_range_without_references` with explicit source and output ranges.",
                "- Keep the translation column aligned with the specified output range (one column per source column).",
                "- Do not request workbook uploads; Excel is already connected.",
                "- Treat this as a single-run request and avoid proposing follow-up tasks once you finish.",
            ]
        elif self.mode is CopilotMode.TRANSLATION_WITH_REFERENCES:
            prefix_lines = [
                "[Translation (With References) Mode Request]",
                "- Solve this by calling `translate_range_with_references` and pass the user's `source_reference_urls` (原文側) and `target_reference_urls` (翻訳先側) or the provided reference ranges.",
                "- Translate the entire requested range in one call and rely on the tool's batching; only adjust `rows_per_batch` when necessary for very large jobs.",
                "- Provide citation output when evidence is expected and reserve columns for: translated text, translation process explanation, and one reference pair per column starting from the specified output column (e.g., \"XX列以降\").",
                "- Do not request workbook uploads; Excel is already connected.",
                "- Treat this as a single-run request and avoid proposing follow-up tasks once you finish.",
            ]
        else:
            prefix_lines = [
                "[Translation Review Mode Request]",
                "- Use `check_translation_quality` with ranges for status, issues, and highlight only (three columns total).",
                "- Clearly identify which range contains the Japanese source text and which range contains the English translation under review.",
                "- Keep outputs aligned with the ranges specified in the instructions.",
                "- Do not request workbook uploads; Excel is already connected.",
                "- Treat this as a single-run request and avoid proposing follow-up tasks once you finish.",
            ]
        prefix = "\n".join(prefix_lines)
        if not trimmed_input:
            return prefix
        return f"{prefix}\n\nUser instruction:\n{trimmed_input}"


    def _load_tools(self, mode: Optional[CopilotMode] = None):
        target_mode = mode or self.mode
        allowed_by_mode: Dict[CopilotMode, List[str]] = {
            CopilotMode.TRANSLATION: ["translate_range_without_references"],
            CopilotMode.TRANSLATION_WITH_REFERENCES: ["translate_range_with_references"],
            CopilotMode.REVIEW: ["check_translation_quality"],
        }
        allowed_tool_names = allowed_by_mode.get(target_mode, [])

        selected = []
        for name in allowed_tool_names:
            func = getattr(excel_tools, name, None)
            if inspect.isfunction(func):
                selected.append(func)

        if not selected:
            raise RuntimeError(f"No tools available for mode '{target_mode.value}'.")

        self.tool_functions = selected
        self.tool_schemas = [create_tool_schema(func) for func in self.tool_functions]

    def _restart_browser_session(self) -> bool:
        if not self.browser_manager:
            return True

        self._emit_response(ResponseMessage(ResponseType.STATUS, "Copilot セッションをリセットしています..."))
        try:
            reset_ok = self.browser_manager.reset_chat_session()
        except Exception as e:
            error_message = f"Copilot セッションのリセットに失敗しました: {e}"
            print(error_message)
            traceback.print_exc()
            self._emit_response(ResponseMessage(ResponseType.ERROR, error_message))
            return False

        if not reset_ok:
            self._emit_response(ResponseMessage(ResponseType.STATUS, "Copilot セッションのリセットに失敗したためブラウザを再初期化しています..."))
            try:
                self.browser_manager.restart()
            except Exception as e:
                error_message = f"ブラウザの再初期化に失敗しました: {e}"
                print(error_message)
                traceback.print_exc()
                try:
                    self.browser_manager.close()
                except Exception:
                    pass
                self.browser_manager = None
                self.agent = None
                self.tool_functions = []
                self.tool_schemas = []
                self._emit_response(ResponseMessage(ResponseType.ERROR, error_message))
                return False


        if self.agent:
            try:
                self.agent.reset()
            except Exception as reset_err:
                print(f"エージェントのリセットに失敗しましたが続行します: {reset_err}")

        self._emit_response(
            ResponseMessage(
                ResponseType.STATUS,
                "ブラウザの初期化が完了しました。",
                "Copilot セッションの準備が整いました。",
            )
        )
        return True

    def _initialize(self):
        try:
            print("Worker\u306e\u521d\u671f\u5316\u3092\u958b\u59cb\u3057\u307e\u3059...")
            self._emit_response(ResponseMessage(ResponseType.STATUS, "\u30d6\u30e9\u30a6\u30b6 (Playwright) \u3092\u8d77\u52d5\u4e2d..."))
            self.browser_manager = BrowserCopilotManager(
                user_data_dir=COPILOT_USER_DATA_DIR,
                headless=COPILOT_HEADLESS,
                browser_channels=COPILOT_BROWSER_CHANNELS,
                goto_timeout_ms=COPILOT_PAGE_GOTO_TIMEOUT_MS,
                slow_mo_ms=COPILOT_SLOW_MO_MS,
            )
            self.browser_manager.start()
            print("BrowserManager \u306e\u8d77\u52d5\u304c\u5b8c\u4e86\u3057\u307e\u3057\u305f\u3002")

            self._emit_response(ResponseMessage(ResponseType.STATUS, "AI\u30a8\u30fc\u30b8\u30a7\u30f3\u30c8\u3092\u6e96\u5099\u4e2d..."))
            self._load_tools(self.mode)
            self._build_agent()
            print("AI\u30a8\u30fc\u30b8\u30a7\u30f3\u30c8\u306e\u6e96\u5099\u304c\u5b8c\u4e86\u3057\u307e\u3057\u305f\u3002")

            self._emit_response(ResponseMessage(ResponseType.INITIALIZATION_COMPLETE, "\u521d\u671f\u5316\u304c\u5b8c\u4e86\u3057\u307e\u3057\u305f\u3002\u6307\u793a\u3092\u3069\u3046\u305e\u3002"))
            print("Worker\u306e\u521d\u671f\u5316\u304c\u5b8c\u4e86\u3057\u307e\u3057\u305f\u3002")
        except Exception as e:
            print(f"Worker\u306e\u521d\u671f\u5316\u4e2d\u306b\u30a8\u30e9\u30fc\u304c\u767a\u751f\u3057\u307e\u3057\u305f: {e}")
            traceback.print_exc()
            self._emit_response(ResponseMessage(ResponseType.ERROR, f"\u521d\u671f\u5316\u30a8\u30e9\u30fc: {e}"))

    def _main_loop(self):
        print("\u30e1\u30a4\u30f3\u30eb\u30fc\u30d7\u3092\u958b\u59cb\u3057\u307e\u3059...")
        while True:
            raw_request = self.request_queue.get()
            try:
                request = RequestMessage.from_raw(raw_request)
            except ValueError as exc:
                self._emit_response(ResponseMessage(ResponseType.ERROR, f"\u7121\u52b9\u306a\u30ea\u30af\u30a8\u30b9\u30c8\u3092\u53d7\u4fe1\u3057\u307e\u3057\u305f: {exc}"))
                continue

            if request.type is RequestType.QUIT:
                print("\u7d42\u4e86\u30ea\u30af\u30a8\u30b9\u30c8\u3092\u53d7\u4fe1\u3057\u307e\u3057\u305f\u3002\u30e1\u30a4\u30f3\u30eb\u30fc\u30d7\u3092\u7d42\u4e86\u3057\u307e\u3059\u3002")
                break
            if request.type is RequestType.STOP:
                self.stop_event.set()
                if self.browser_manager:
                    try:
                        self.browser_manager.request_stop()
                    except Exception as stop_err:
                        print(f"\u505c\u6b62\u30ea\u30af\u30a8\u30b9\u30c8\u306e\u8ee2\u9001\u306b\u5931\u6557\u3057\u307e\u3057\u305f: {stop_err}")
                continue

            if request.type is RequestType.UPDATE_CONTEXT:
                self._update_context(request.payload)
                continue
            if request.type is RequestType.RESET_BROWSER:
                self._handle_browser_reset_request()
                continue
            if request.type is RequestType.USER_INPUT:
                if isinstance(request.payload, str):
                    self._execute_task(request.payload)
                else:
                    self._emit_response(ResponseMessage(ResponseType.ERROR, "\u30e6\u30fc\u30b6\u30fc\u5165\u529b\u306e\u5f62\u5f0f\u304c\u4e0d\u6b63\u3067\u3059\u3002"))

    def _update_context(self, payload: Optional[Dict[str, Any]]):
        if not isinstance(payload, dict):
            return

        new_workbook_name = payload.get("workbook_name")
        if isinstance(new_workbook_name, str) and new_workbook_name.strip():
            normalized_workbook = new_workbook_name.strip()
            if normalized_workbook != self.workbook_name:
                self.workbook_name = normalized_workbook
                if self.agent:
                    self.agent.set_workbook(normalized_workbook)
                self._emit_response(
                    ResponseMessage(
                        ResponseType.INFO,
                        f"\u64cd\u4f5c\u5bfe\u8c61\u306e\u30d6\u30c3\u30af\u3092\u300e{normalized_workbook}\u300f\u306b\u5909\u66f4\u3057\u307e\u3057\u305f\u3002",
                    )
                )

        new_sheet_name = payload.get("sheet_name")
        if new_sheet_name:
            self.sheet_name = new_sheet_name
            if self.agent:
                self.agent.sheet_name = new_sheet_name
            sheet_label = new_sheet_name or "\u672a\u9078\u629e"
            self._emit_response(ResponseMessage(ResponseType.INFO, f"\u64cd\u4f5c\u5bfe\u8c61\u306e\u30b7\u30fc\u30c8\u3092\u300c{sheet_label}\u300d\u306b\u5909\u66f4\u3057\u307e\u3057\u305f\u3002"))

        mode_value = payload.get("mode")
        if mode_value is not None:
            try:
                new_mode = CopilotMode(mode_value)
            except ValueError:
                self._emit_response(ResponseMessage(ResponseType.ERROR, f"\u30e2\u30fc\u30c9\u5024\u304c\u4e0d\u6b63\u3067\u3059: {mode_value}"))
            else:
                if new_mode != self.mode:
                    self.mode = new_mode
                    try:
                        self._load_tools(new_mode)
                    except Exception as tool_err:
                        self.tool_functions = []
                        self.tool_schemas = []
                        self.agent = None
                        self._emit_response(ResponseMessage(ResponseType.ERROR, f"\u5229\u7528\u53ef\u80fd\u306a\u30c4\u30fc\u30eb\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093: {tool_err}"))
                        return
                    self._build_agent()
                    mode_label_map = {
                        CopilotMode.TRANSLATION: "\u7ffb\u8a33\uff08\u901a\u5e38\uff09",
                        CopilotMode.TRANSLATION_WITH_REFERENCES: "\u7ffb\u8a33\uff08\u53c2\u7167\u3042\u308a\uff09",
                        CopilotMode.REVIEW: "\u7ffb\u8a33\u30c1\u30a7\u30c3\u30af",
                    }
                    mode_label = mode_label_map.get(new_mode, new_mode.value)
                    self._emit_response(ResponseMessage(ResponseType.INFO, f"\u30e2\u30fc\u30c9\u3092{mode_label}\u306b\u5207\u308a\u66ff\u3048\u307e\u3057\u305f\u3002"))

    def _execute_task(self, user_input: str):
        self.stop_event.clear()
        if not self.agent:
            self._emit_response(ResponseMessage(ResponseType.ERROR, "AI\u30a8\u30fc\u30b8\u30a7\u30f3\u30c8\u304c\u521d\u671f\u5316\u3055\u308c\u3066\u3044\u307e\u305b\u3093\u3002"))
            return

        try:
            formatted_input = self._format_user_prompt(user_input)
            for message_dict in self.agent.run(formatted_input, self.stop_event):
                self._emit_response(message_dict)
        except ExcelConnectionError as e:
            self._emit_response(ResponseMessage(ResponseType.ERROR, str(e)))
        except Exception as e:
            self._emit_response(ResponseMessage(ResponseType.ERROR, f"\u30bf\u30b9\u30af\u5b9f\u884c\u30a8\u30e9\u30fc: {e}"))
        finally:
            stop_requested = self.stop_event.is_set()
            if stop_requested:
                self._emit_response(ResponseMessage(ResponseType.INFO, "\u30e6\u30fc\u30b6\u30fc\u306b\u3088\u3063\u3066\u30bf\u30b9\u30af\u304c\u4e2d\u65ad\u3055\u308c\u307e\u3057\u305f\u3002"))
            if stop_requested:
                restart_ok = self._restart_browser_session()
                if restart_ok:
                    metadata = {"action": "focus_app_window"}
                    self._emit_response(
                        ResponseMessage(
                            ResponseType.INFO,
                            "",
                            metadata=metadata,
                        )
                    )
            self._emit_response(ResponseMessage(ResponseType.END_OF_TASK))

    def _handle_browser_reset_request(self):
        restart_ok = self._restart_browser_session()
        if restart_ok:
            self._emit_response(
                ResponseMessage(
                    ResponseType.INFO,
                    "",
                    metadata={"action": "focus_app_window"},
                )
            )

    def _cleanup(self):
        print("\u30af\u30ea\u30fc\u30f3\u30a2\u30c3\u30d7\u3092\u958b\u59cb\u3057\u307e\u3059...")
        if self.browser_manager:
            self.browser_manager.close()
        print("Worker\u306e\u30af\u30ea\u30fc\u30f3\u30a2\u30c3\u30d7\u304c\u5b8c\u4e86\u3057\u307e\u3057\u305f\u3002")
