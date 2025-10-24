"""Background worker that coordinates Excel Copilot tasks."""

from __future__ import annotations

import inspect
import logging
import queue
import threading
from typing import Any, Dict, List, Optional, Union

from excel_copilot.agent.prompts import CopilotMode
from excel_copilot.config import (
    COPILOT_BROWSER_CHANNELS,
    COPILOT_HEADLESS,
    COPILOT_PAGE_GOTO_TIMEOUT_MS,
    COPILOT_SLOW_MO_MS,
    COPILOT_USER_DATA_DIR,
)
from excel_copilot.core.browser_copilot_manager import BrowserCopilotManager
from excel_copilot.core.exceptions import UserStopRequested
from excel_copilot.core.excel_manager import ExcelConnectionError, ExcelManager
from excel_copilot.tools import excel_tools
from excel_copilot.tools.actions import ExcelActions

from .messages import RequestMessage, RequestType, ResponseMessage, ResponseType

_logger = logging.getLogger(__name__)

INVALID_TASK_PAYLOAD_MESSAGE = "タスク要求の形式が正しくありません。"
INVALID_ARGUMENTS_MESSAGE = "引数の形式が JSON オブジェクトではありません。"
BROWSER_NOT_READY_MESSAGE = "Copilot ブラウザーセッションが初期化されていません。"
TOOL_NOT_AVAILABLE_MESSAGE = "現在のモードで利用できるツールが見つかりません。"
MISSING_WORKBOOK_ERROR_MESSAGE = (
    "対象ブックを特定できず Excel 操作を開始できません。Excel 上で対象ブックを開き、一覧から選択してください。"
)


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
        self.stop_event = threading.Event()
        self.sheet_name = sheet_name
        self.workbook_name = workbook_name
        self.mode = CopilotMode.TRANSLATION_WITH_REFERENCES
        self.tool_functions: List[callable] = []
        self.current_tool: Optional[callable] = None

    def run(self):
        _logger.info("Copilot worker thread started.")
        try:
            self._initialize()
            if self.browser_manager:
                self._main_loop()
        except Exception as e:
            _logger.exception("Critical error in worker run method: %s", e)
            self._emit_response(ResponseMessage(ResponseType.ERROR, f"\u81f4\u547d\u7684\u306a\u5b9f\u884c\u6642\u30a8\u30e9\u30fc: {e}"))
        finally:
            self._cleanup()

    def _emit_response(self, message: Union[ResponseMessage, Dict[str, Any]]):
        try:
            self.response_queue.put(ResponseMessage.from_raw(message))
        except Exception as err:
            _logger.exception("Failed to enqueue response: %s", err)

    def _load_tools(self, mode: Optional[CopilotMode] = None) -> None:
        target_mode = mode or self.mode
        allowed_by_mode: Dict[CopilotMode, List[str]] = {
            CopilotMode.TRANSLATION: ["translate_range_without_references"],
            CopilotMode.TRANSLATION_WITH_REFERENCES: ["translate_range_with_references"],
            CopilotMode.REVIEW: ["check_translation_quality"],
        }
        allowed_tool_names = allowed_by_mode.get(target_mode, [])

        selected: List[callable] = []
        for name in allowed_tool_names:
            func = getattr(excel_tools, name, None)
            if callable(func):
                selected.append(func)

        if not selected:
            raise RuntimeError(f"No tools available for mode '{target_mode.value}'.")

        self.tool_functions = selected
        self.current_tool = selected[0]

    def _restart_browser_session(self) -> bool:
        if not self.browser_manager:
            return True

        self._emit_response(ResponseMessage(ResponseType.STATUS, "Copilot セッションをリセットしています..."))
        try:
            reset_ok = self.browser_manager.reset_chat_session()
        except Exception as e:
            error_message = f"Copilot セッションのリセットに失敗しました: {e}"
            _logger.exception(error_message)
            self._emit_response(ResponseMessage(ResponseType.ERROR, error_message))
            return False

        if not reset_ok:
            self._emit_response(ResponseMessage(ResponseType.STATUS, "セッション再利用に失敗したためブラウザを再起動します。"))
            try:
                self.browser_manager.restart()
            except Exception as e:
                error_message = f"ブラウザの再起動に失敗しました: {e}"
                _logger.exception(error_message)
                try:
                    self.browser_manager.close()
                except Exception:
                    pass
                self.browser_manager = None
                self.tool_functions = []
                self.current_tool = None
                self._emit_response(ResponseMessage(ResponseType.ERROR, error_message))
                return False

        return True

    def _initialize(self):
        try:
            _logger.info("Worker initialization started.")
            self._emit_response(ResponseMessage(ResponseType.STATUS, "\u30d6\u30e9\u30a6\u30b6 (Playwright) \u3092\u8d77\u52d5\u4e2d..."))
            self.browser_manager = BrowserCopilotManager(
                user_data_dir=COPILOT_USER_DATA_DIR,
                headless=COPILOT_HEADLESS,
                browser_channels=COPILOT_BROWSER_CHANNELS,
                goto_timeout_ms=COPILOT_PAGE_GOTO_TIMEOUT_MS,
                slow_mo_ms=COPILOT_SLOW_MO_MS,
            )
            self.browser_manager.start()
            self.browser_manager.set_chat_transcript_sink(self._handle_chat_transcript_event)
            _logger.info("BrowserCopilotManager start completed.")

            self._emit_response(ResponseMessage(ResponseType.STATUS, "\u30d5\u30a9\u30fc\u30e0\u7528\u30c4\u30fc\u30eb\u3092\u521d\u671f\u5316\u3057\u3066\u3044\u307e\u3059..."))
            self._load_tools(self.mode)
            _logger.info("Structured tool preparation completed.")

            self._emit_response(ResponseMessage(ResponseType.INITIALIZATION_COMPLETE, "\u521d\u671f\u5316\u304c\u5b8c\u4e86\u3057\u307e\u3057\u305f\u3002\u6307\u793a\u3092\u3069\u3046\u305e\u3002"))
            _logger.info("Worker initialization completed.")
        except Exception as e:
            _logger.exception("Worker initialization failed: %s", e)
            self._emit_response(ResponseMessage(ResponseType.ERROR, f"\u521d\u671f\u5316\u30a8\u30e9\u30fc: {e}"))

    def _main_loop(self):
        _logger.info("Worker main loop started.")
        while True:
            raw_request = self.request_queue.get()
            try:
                request = RequestMessage.from_raw(raw_request)
            except ValueError as exc:
                self._emit_response(ResponseMessage(ResponseType.ERROR, f"\u7121\u52b9\u306a\u30ea\u30af\u30a8\u30b9\u30c8\u3092\u53d7\u4fe1\u3057\u307e\u3057\u305f: {exc}"))
                continue

            if request.type is RequestType.QUIT:
                _logger.info("Quit request received; shutting down worker loop.")
                break
            if request.type is RequestType.STOP:
                self.stop_event.set()
                if self.browser_manager:
                    try:
                        self.browser_manager.request_stop()
                    except Exception as stop_err:
                        _logger.exception("Failed to forward stop request: %s", stop_err)
                continue

            if request.type is RequestType.UPDATE_CONTEXT:
                self._update_context(request.payload)
                continue
            if request.type is RequestType.RESET_BROWSER:
                self._handle_browser_reset_request()
                continue
            if request.type is RequestType.USER_INPUT:
                self._execute_task(request.payload)

    def _update_context(self, payload: Optional[Dict[str, Any]]):
        if not isinstance(payload, dict):
            return

        new_workbook_name = payload.get("workbook_name")
        if isinstance(new_workbook_name, str) and new_workbook_name.strip():
            normalized_workbook = new_workbook_name.strip()
            if normalized_workbook != self.workbook_name:
                self.workbook_name = normalized_workbook
                self._emit_response(
                    ResponseMessage(
                        ResponseType.INFO,
                        f"操作対象のブックを『{normalized_workbook}』に変更しました。",
                    )
                )

        new_sheet_name = payload.get("sheet_name")
        if isinstance(new_sheet_name, str) and new_sheet_name.strip():
            sheet_label = new_sheet_name.strip()
            if sheet_label != self.sheet_name:
                self.sheet_name = sheet_label
                self._emit_response(
                    ResponseMessage(ResponseType.INFO, f"操作対象のシートを「{sheet_label}」に変更しました。")
                )

        mode_value = payload.get("mode")
        if mode_value is None:
            return
        try:
            new_mode = CopilotMode(mode_value)
        except ValueError:
            self._emit_response(ResponseMessage(ResponseType.ERROR, f"モード値が不正です: {mode_value}"))
            return

        if new_mode == self.mode:
            return

        self.mode = new_mode
        try:
            self._load_tools(new_mode)
        except Exception as tool_err:
            self.tool_functions = []
            self.current_tool = None
            self._emit_response(ResponseMessage(ResponseType.ERROR, f"利用可能なツールが見つかりません: {tool_err}"))
            return

        # モード変更時のチャット通知は省略し、UI 側で直接反映する。




    def _execute_task(self, payload: Any):
        self.stop_event.clear()

        if not isinstance(payload, dict):
            self._emit_response(
                ResponseMessage(ResponseType.ERROR, INVALID_TASK_PAYLOAD_MESSAGE)
            )
            self._finalize_task()
            return

        try:
            self._run_structured_task(payload)
        finally:
            self._finalize_task()

    def _run_structured_task(self, payload: Dict[str, Any]) -> None:
        if not isinstance(payload, dict):
            self._emit_response(
                ResponseMessage(ResponseType.ERROR, INVALID_TASK_PAYLOAD_MESSAGE)
            )
            return
        if not self.browser_manager:
            self._emit_response(
                ResponseMessage(ResponseType.ERROR, BROWSER_NOT_READY_MESSAGE)
            )
            return
        if not self.current_tool:
            self._emit_response(
                ResponseMessage(ResponseType.ERROR, TOOL_NOT_AVAILABLE_MESSAGE)
            )
            return

        tool_function = self.current_tool
        tool_name = tool_function.__name__

        requested_tool = payload.get("tool_name")
        if isinstance(requested_tool, str) and requested_tool and requested_tool != tool_name:
            self._emit_response(
                ResponseMessage(
                    ResponseType.ERROR,
                    f"リクエストされたツール '{requested_tool}' は現在のモードでは使用できません。",
                )
            )
            return

        mode_hint = payload.get("mode")
        if isinstance(mode_hint, str) and mode_hint and mode_hint != self.mode.value:
            self._emit_response(
                ResponseMessage(
                    ResponseType.ERROR,
                    f"リクエストされたモード '{mode_hint}' は現在のモード '{self.mode.value}' と一致しません。",
                )
            )
            return

        workbook_override = payload.get("workbook_name")
        if isinstance(workbook_override, str) and workbook_override.strip():
            self.workbook_name = workbook_override.strip()

        arguments = payload.get("arguments")
        if arguments is None:
            arguments = {k: v for k, v in payload.items() if k not in {"tool_name", "mode", "workbook_name"}}

        if not isinstance(arguments, dict):
            self._emit_response(
                ResponseMessage(ResponseType.ERROR, INVALID_ARGUMENTS_MESSAGE)
            )
            return

        if not self.workbook_name:
            self._emit_response(
                ResponseMessage(ResponseType.ERROR, MISSING_WORKBOOK_ERROR_MESSAGE)
            )
            return

        arguments = dict(arguments)
        sheet_override = payload.get("sheet_name")
        if isinstance(sheet_override, str) and sheet_override.strip():
            arguments.setdefault("sheet_name", sheet_override.strip())

        self._emit_response(ResponseMessage(ResponseType.STATUS, f"{tool_name} を実行しています..."))

        excel_actions: Optional[ExcelActions] = None
        try:
            with ExcelManager(self.workbook_name) as manager:
                excel_actions = ExcelActions(manager, progress_callback=self._handle_progress_update)
                try:
                    call_args = self._prepare_tool_arguments(tool_function, arguments, excel_actions)
                except (ValueError, RuntimeError) as exc:
                    self._emit_response(ResponseMessage(ResponseType.ERROR, str(exc)))
                    return

                result = tool_function(**call_args)

                if not self.stop_event.is_set():
                    message = result.strip() if isinstance(result, str) else ""
                    final_message = message or f"{tool_name} が完了しました。"
                    self._emit_response(ResponseMessage(ResponseType.FINAL_ANSWER, final_message))
        except UserStopRequested:
            pass
        except ExcelConnectionError as exc:
            self._emit_response(ResponseMessage(ResponseType.ERROR, str(exc)))
        except Exception as exc:
            self._emit_response(
                ResponseMessage(ResponseType.ERROR, f"ツール実行中にエラーが発生しました: {exc}")
            )
        finally:
            if excel_actions and hasattr(excel_actions, "consume_progress_messages"):
                try:
                    excel_actions.consume_progress_messages()
                except Exception:
                    pass

    def _prepare_tool_arguments(
        self,
        tool_function: callable,
        arguments: Dict[str, Any],
        excel_actions: ExcelActions,
    ) -> Dict[str, Any]:
        if not isinstance(arguments, dict):
            raise ValueError("ツール引数は辞書形式で指定してください。")

        prepared = dict(arguments)
        signature = inspect.signature(tool_function)

        if "actions" in signature.parameters:
            prepared["actions"] = excel_actions
        if "browser_manager" in signature.parameters:
            if not self.browser_manager:
                raise RuntimeError("Copilot ブラウザが初期化されていません。")
            prepared["browser_manager"] = self.browser_manager
        if "sheet_name" in signature.parameters and "sheet_name" not in prepared and self.sheet_name:
            prepared["sheet_name"] = self.sheet_name
        if "sheetname" in signature.parameters and "sheetname" not in prepared and self.sheet_name:
            prepared["sheetname"] = self.sheet_name
        if "stop_event" in signature.parameters:
            prepared["stop_event"] = self.stop_event

        return prepared

    def _handle_progress_update(self, message: str) -> None:
        text = (message or "").strip()
        if text:
            self._emit_response(ResponseMessage(ResponseType.STATUS, text))

    def _handle_chat_transcript_event(
        self,
        role: str,
        text: str,
        metadata: Optional[Dict[str, Any]] = None,
    ) -> None:
        normalized_role = (role or "").strip().lower()
        if normalized_role == "prompt":
            event_type = ResponseType.CHAT_PROMPT
        elif normalized_role == "response":
            event_type = ResponseType.CHAT_RESPONSE
        else:
            return
        payload_text = text if isinstance(text, str) else ("" if text is None else str(text))
        payload_metadata = dict(metadata or {}) if metadata else {}
        self._emit_response(ResponseMessage(event_type, payload_text, payload_metadata))


    def _finalize_task(self) -> None:
        stop_requested = self.stop_event.is_set()
        if stop_requested:
            self._emit_response(
                ResponseMessage(ResponseType.INFO, "ユーザーによってタスクが中断されました。")
            )
            restart_ok = self._restart_browser_session()
            if restart_ok:
                metadata = {"action": "focus_app_window"}
                self._emit_response(ResponseMessage(ResponseType.INFO, "", metadata=metadata))
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
        _logger.info("Worker cleanup starting...")
        if self.browser_manager:
            self.browser_manager.set_chat_transcript_sink(None)
            self.browser_manager.close()
        _logger.info("Worker cleanup completed.")
        try:
            self._emit_response(ResponseMessage(ResponseType.SHUTDOWN_COMPLETE))
        except Exception:
            _logger.debug("Failed to emit shutdown confirmation.", exc_info=True)
