# desktop_app.py

import flet as ft
import argparse
import threading
import queue
import inspect
import json
import time
import traceback
import os
import logging
from datetime import datetime
from pathlib import Path
from dataclasses import dataclass, field
from typing import Dict, Optional, Any, Union, List
from enum import Enum, auto

from excel_copilot.core.excel_manager import ExcelManager, ExcelConnectionError
from excel_copilot.agent.react_agent import ReActAgent
from excel_copilot.agent.prompts import CopilotMode
from excel_copilot.core.browser_copilot_manager import BrowserCopilotManager
from excel_copilot.tools import excel_tools
from excel_copilot.tools.schema_builder import create_tool_schema
from excel_copilot.config import (
    COPILOT_USER_DATA_DIR,
    COPILOT_HEADLESS,
    COPILOT_BROWSER_CHANNELS,
    COPILOT_PAGE_GOTO_TIMEOUT_MS,
    COPILOT_SLOW_MO_MS,
)

if not logging.getLogger().handlers:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )

FOCUS_WAIT_TIMEOUT_SECONDS = 15.0
PREFERENCE_LAST_WORKBOOK_KEY = "__last_workbook__"

class AppState(Enum):
    INITIALIZING = auto()
    READY = auto()
    TASK_IN_PROGRESS = auto()
    STOPPING = auto()
    ERROR = auto()

class RequestType(str, Enum):
    USER_INPUT = "USER_INPUT"
    STOP = "STOP"
    QUIT = "QUIT"
    UPDATE_CONTEXT = "UPDATE_CONTEXT"
    RESET_BROWSER = "RESET_BROWSER"

class ResponseType(str, Enum):
    STATUS = "status"
    ERROR = "error"
    INFO = "info"
    END_OF_TASK = "end_of_task"
    INITIALIZATION_COMPLETE = "initialization_complete"
    THOUGHT = "thought"
    ACTION = "action"
    OBSERVATION = "observation"
    FINAL_ANSWER = "final_answer"

@dataclass(frozen=True)
class RequestMessage:
    type: RequestType
    payload: Optional[Any] = None

    @classmethod
    def from_raw(cls, raw: Union["RequestMessage", Dict[str, Any]]) -> "RequestMessage":
        if isinstance(raw, cls):
            return raw
        if not isinstance(raw, dict):
            raise ValueError(f"Unsupported request payload type: {type(raw)}")
        raw_type = raw.get("type")
        if isinstance(raw_type, RequestType):
            request_type = raw_type
        else:
            try:
                request_type = RequestType(str(raw_type))
            except ValueError as exc:
                raise ValueError(f"Unsupported request type: {raw_type}") from exc
        return cls(type=request_type, payload=raw.get("payload"))

@dataclass(frozen=True)
class ResponseMessage:
    type: ResponseType
    content: str = ""
    metadata: Dict[str, Any] = field(default_factory=dict)

    @classmethod
    def from_raw(cls, raw: Union["ResponseMessage", Dict[str, Any]]) -> "ResponseMessage":
        if isinstance(raw, cls):
            return raw
        if not isinstance(raw, dict):
            raise ValueError(f"Unsupported response payload type: {type(raw)}")
        raw_type = raw.get("type")
        if isinstance(raw_type, ResponseType):
            response_type = raw_type
        else:
            try:
                response_type = ResponseType(str(raw_type))
            except ValueError:
                response_type = ResponseType.INFO
        content = raw.get("content", "")
        metadata = {k: v for k, v in raw.items() if k not in {"type", "content"}}
        if raw_type and (not isinstance(raw_type, ResponseType)) and raw_type != response_type.value:
            metadata.setdefault("source_type", raw_type)
        return cls(type=response_type, content=content, metadata=metadata)

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
                "- Solve this by calling `translate_range_with_references` and include the reference ranges or URLs provided by the user.",
                "- Work one cell at a time without `rows_per_batch`; split multi-row ranges across multiple calls.",
                "- Provide citation output when evidence is expected and keep translation, quote, and explanation columns aligned.",
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

        self._emit_response(ResponseMessage(ResponseType.STATUS, "ブラウザを初期化しています..."))
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
                metadata={"browser_ready": True},
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

class ChatMessage(ft.ResponsiveRow):
    def __init__(self, msg_type: Union[ResponseType, str], msg_content: str):
        super().__init__()
        self.vertical_alignment = ft.CrossAxisAlignment.START
        self.opacity = 0
        self.animate_opacity = 300
        self.offset = ft.Offset(0, 0.1)
        self.animate_offset = 300

        palette = {
            "primary_container": "#4F378B",
            "on_primary_container": "#EADDFF",
            "secondary_container": "#4A4458",
            "on_secondary_container": "#E8DEF8",
            "tertiary_container": "#633B48",
            "on_tertiary_container": "#FFD8E4",
            "neutral_container": "#332D41",
            "on_neutral_container": "#E8DEF8",
            "surface_variant": "#49454F",
            "on_surface_variant": "#CAC4D0",
            "error_container": "#8C1D18",
            "on_error_container": "#F9DEDC",
        }

        type_map = {
            "user": {
                "bgcolor": palette["primary_container"],
                "icon": ft.Icons.PERSON_ROUNDED,
                "icon_color": palette["on_primary_container"],
                "text_style": {"color": palette["on_primary_container"], "size": 14},
            },
            "thought": {
                "bgcolor": palette["secondary_container"],
                "icon": ft.Icons.LIGHTBULB_OUTLINE,
                "icon_color": palette["on_secondary_container"],
                "text_style": {"color": palette["on_secondary_container"], "size": 13},
                "title": "AI Thought",
            },
            "action": {
                "bgcolor": palette["surface_variant"],
                "icon": ft.Icons.CODE,
                "icon_color": palette["on_surface_variant"],
                "text_style": {"font_family": "monospace", "color": palette["on_surface_variant"], "size": 13},
                "title": "Action",
            },
            "observation": {
                "bgcolor": palette["neutral_container"],
                "icon": ft.Icons.FIND_IN_PAGE_OUTLINED,
                "icon_color": palette["on_neutral_container"],
                "text_style": {"color": palette["on_neutral_container"], "size": 13},
                "title": "Observation",
            },
            "final_answer": {
                "bgcolor": palette["tertiary_container"],
                "icon": ft.Icons.CHECK_CIRCLE_OUTLINE,
                "icon_color": palette["on_tertiary_container"],
                "text_style": {"color": palette["on_tertiary_container"], "size": 14},
                "title": "Answer",
            },
            "info": {
                "text_style": {"color": palette["on_surface_variant"], "size": 12},
                "icon": ft.Icons.INFO_OUTLINE,
                "icon_color": palette["on_surface_variant"],
            },
            "status": {
                "text_style": {"color": palette["on_surface_variant"], "size": 12},
            },
            "error": {
                "icon": ft.Icons.ERROR_OUTLINE_ROUNDED,
                "icon_color": palette["on_error_container"],
                "bgcolor": palette["error_container"],
                "text_style": {"color": palette["on_error_container"], "size": 13},
                "title": "Error",
            },
        }

        msg_type_value = msg_type.value if isinstance(msg_type, ResponseType) else msg_type
        config = type_map.get(msg_type_value, type_map["info"])

        if msg_type_value in ["info", "status"]:
            self.controls = [
                ft.Column(
                    [ft.Text(msg_content, **config.get("text_style", {}))],
                    col=12,
                    alignment=ft.MainAxisAlignment.CENTER,
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                )
            ]
            return

        content_controls = []
        if config.get("title"):
            content_controls.append(ft.Text(config["title"], weight=ft.FontWeight.BOLD, size=12, color=config.get("icon_color")))

        text_style = dict(config.get("text_style", {}))
        line_controls = []
        icon_color = config.get("icon_color", text_style.get("color"))
        size = text_style.get("size")
        normalized_content = (msg_content or "").replace("\r\n", "\n")
        for raw_line in normalized_content.split("\n"):
            if raw_line.strip() == "":
                line_controls.append(ft.Container(height=6))
                continue

            stripped = raw_line.strip()
            if stripped.startswith("\u5f15\u7528"):
                label, sep, remainder = stripped.partition(":")
                bullet = ft.Text("\u2022", size=size or 13, color=icon_color)
                label_text = ft.Text(label.strip() + (sep if sep else ""), weight=ft.FontWeight.BOLD, size=size or 13, color=icon_color)
                remainder_texts = []
                remainder_value = remainder.strip() if remainder else ""
                if remainder_value:
                    remainder_texts.append(ft.Text(remainder_value, **text_style, selectable=True))
                line_controls.append(
                    ft.Row(
                        [
                            bullet,
                            ft.Column([label_text] + remainder_texts, spacing=2, tight=True),
                        ],
                        alignment=ft.MainAxisAlignment.START,
                        vertical_alignment=ft.CrossAxisAlignment.START,
                        spacing=6,
                    )
                )
            else:
                line_controls.append(ft.Text(raw_line, **text_style, selectable=True))

        content_controls.extend(line_controls if line_controls else [ft.Text(msg_content, **text_style, selectable=True)])

        message_bubble = ft.Container(
            content=ft.Column(content_controls, spacing=6, tight=True),
            bgcolor=config.get("bgcolor"),
            border_radius=16,
            padding=16,
            expand=True,
            shadow=ft.BoxShadow(
                spread_radius=1,
                blur_radius=18,
                color="#33000000",
                offset=ft.Offset(2, 4),
            ),
        )

        icon_name = config.get("icon", ft.Icons.SMART_BUTTON)
        icon_color = config.get("icon_color", "#CFD8DC")
        icon = ft.Icon(name=icon_name, color=icon_color, size=20)
        icon_container = ft.Container(icon, margin=ft.margin.only(right=8, left=8, top=3))

        bubble_and_icon_row = ft.Row(
            [icon_container, message_bubble] if msg_type_value != "user" else [message_bubble, icon_container],
            vertical_alignment=ft.CrossAxisAlignment.START,
        )

        if msg_type_value == "user":
            self.controls = [
                ft.Column(col={"sm": 2, "md": 4}),
                ft.Column(col={"sm": 10, "md": 8}, controls=[bubble_and_icon_row]),
            ]
        else:
            self.controls = [
                ft.Column(col={"sm": 10, "md": 8}, controls=[bubble_and_icon_row]),
            ]

    def appear(self):
        self.opacity = 1
        self.offset = ft.Offset(0, 0)
        self.update()

class CopilotApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.request_queue: "queue.Queue[RequestMessage]" = queue.Queue()
        self.response_queue: "queue.Queue[ResponseMessage]" = queue.Queue()
        self.worker_thread: Optional[threading.Thread] = None
        self.queue_thread: Optional[threading.Thread] = None
        self.worker: Optional[CopilotWorker] = None
        self.app_state: Optional[AppState] = None
        self.ui_loop_running = True
        self.shutdown_requested = False
        self.window_closed_event = threading.Event()
        self.current_workbook_name: Optional[str] = None
        self.current_sheet_name: Optional[str] = None
        self.sheet_selection_updating = False
        self.workbook_selection_updating = False

        self.mode = CopilotMode.TRANSLATION_WITH_REFERENCES
        self.mode_selector: Optional[ft.RadioGroup] = None

        self.title_label: Optional[ft.Text] = None
        self.status_label: Optional[ft.Text] = None
        self.workbook_selector: Optional[ft.Dropdown] = None
        self.sheet_selector: Optional[ft.Dropdown] = None
        self.chat_list: Optional[ft.ListView] = None
        self.user_input: Optional[ft.TextField] = None
        self.action_button: Optional[ft.Container] = None
        self.save_log_button: Optional[ft.TextButton] = None
        self.workbook_refresh_button: Optional[ft.TextButton] = None
        self.browser_reset_button: Optional[ft.TextButton] = None

        self.chat_history: list[dict[str, str]] = []
        self.history_lock = threading.Lock()
        self.log_dir = Path(COPILOT_USER_DATA_DIR) / "setouchi_logs"
        self.preference_file = Path(COPILOT_USER_DATA_DIR) / "setouchi_state.json"
        self.preference_lock = threading.Lock()

        self._browser_ready_for_focus = False
        self._pending_focus_action: Optional[str] = None
        self._pending_focus_deadline: Optional[float] = None
        self._focus_wait_timeout_sec = FOCUS_WAIT_TIMEOUT_SECONDS
        self._status_message_override: Optional[str] = None
        self._status_color_override: Optional[str] = None
        self._excel_refresh_lock = threading.Lock()
        self._last_excel_snapshot: Dict[str, Any] = {}
        self._excel_poll_thread: Optional[threading.Thread] = None
        self._excel_poll_stop_event = threading.Event()
        self._excel_refresh_event = threading.Event()
        self._excel_poll_interval = 0.8
        self._browser_reset_in_progress = False
        self._manual_refresh_in_progress = False
        self._workbook_refresh_button_default_text = "\u30d6\u30c3\u30af\u4e00\u89a7\u3092\u66f4\u65b0"

        self._configure_page()
        self._build_layout()
        self._register_window_handlers()

    def mount(self):
        self._set_state(AppState.INITIALIZING)
        self._update_ui()
        sheet_name = self._refresh_excel_context(is_initial_start=True)

        self.worker = CopilotWorker(
            self.request_queue,
            self.response_queue,
            sheet_name,
            self.current_workbook_name,
        )
        self.worker_thread = threading.Thread(target=self.worker.run, daemon=True)
        self.worker_thread.start()

        self.queue_thread = threading.Thread(target=self._process_response_queue_loop, daemon=True)
        self.queue_thread.start()

        self.request_queue.put(RequestMessage(RequestType.UPDATE_CONTEXT, {"mode": self.mode.value}))

    def _configure_page(self):
        self.page.title = "Excel Co-pilot"
        self.page.window.width = 1280
        self.page.window.height = 768
        self.page.window.min_width = 960
        self.page.window.min_height = 600
        self.page.theme = ft.Theme(color_scheme_seed="#6750A4", use_material3=True)
        self.page.theme_mode = ft.ThemeMode.DARK
        self.page.bgcolor = "#1D1B20"
        self.page.padding = ft.Padding(24, 24, 24, 24)
        self.page.window.center()
        self.page.window.prevent_close = True

    def _focus_app_window(self):
        try:
            window = getattr(self.page, "window", None)
            if not window:
                return
            bring_fn = getattr(window, "to_front", None) or getattr(window, "bring_to_front", None)
            if callable(bring_fn):
                bring_fn()
        except Exception as focus_err:
            print(f"アプリウィンドウの前面表示に失敗しました: {focus_err}")

    def _focus_excel_window(self):
        try:
            with ExcelManager(self.current_workbook_name) as manager:
                manager.focus_application_window()
        except Exception as focus_err:
            print(f"Excelウィンドウの前面表示に失敗しました: {focus_err}")

    def _build_layout(self):
        self.title_label = ft.Text(
            "Excel Co-pilot",
            size=24,
            weight=ft.FontWeight.BOLD,
            color="#EADDFF",
        )
        self.status_label = ft.Text(
            "\u521d\u671f\u5316\u4e2d...",
            size=12,
            color="#CAC4D0",
            animate_opacity=300,
            animate_scale=600,
        )

        self.page.appbar = ft.AppBar(
            leading=ft.Icon(ft.Icons.TABLE_CHART_OUTLINED, color="#EADDFF"),
            title=ft.Column(
                [self.title_label, self.status_label],
                spacing=4,
                alignment=ft.MainAxisAlignment.CENTER,
                horizontal_alignment=ft.CrossAxisAlignment.START,
            ),
            center_title=False,
            bgcolor="#211F26",
            elevation=2,
        )

        button_shape = ft.RoundedRectangleBorder(radius=12)
        self.save_log_button = ft.FilledTonalButton(
            text="\u4f1a\u8a71\u30ed\u30b0\u3092\u4fdd\u5b58",
            icon=ft.Icons.SAVE_OUTLINED,
            on_click=self._handle_save_log_click,
            disabled=True,
            style=ft.ButtonStyle(shape=button_shape),
        )

        self.browser_reset_button = ft.OutlinedButton(
            text="\u30d6\u30e9\u30a6\u30b6\u3092\u518d\u521d\u671f\u5316",
            icon=ft.Icons.REFRESH,
            on_click=self._handle_browser_reset_click,
            disabled=True,
            style=ft.ButtonStyle(shape=button_shape),
        )

        dropdown_style = {
            "width": 240,
            "border_radius": 12,
            "border_color": "#4F378B",
            "focused_border_color": "#D0BCFF",
            "fill_color": "#2B2930",
            "text_style": ft.TextStyle(color="#E6E0E9"),
            "hint_style": ft.TextStyle(color="#9A8FAE"),
            "disabled": True,
            "filled": True,
        }

        self.workbook_selector = ft.Dropdown(
            options=[],
            on_change=self._on_workbook_change,
            on_focus=self._on_workbook_dropdown_focus,
            hint_text="\u30d6\u30c3\u30af\u3092\u9078\u629e",
            **dropdown_style,
        )

        self.workbook_selector_wrapper = ft.GestureDetector(
            content=self.workbook_selector,
            on_tap_down=self._on_workbook_dropdown_tap,
        )

        self.sheet_selector = ft.Dropdown(
            options=[],
            on_change=self._on_sheet_change,
            on_focus=self._on_sheet_dropdown_focus,
            hint_text="\u30b7\u30fc\u30c8\u3092\u9078\u629e",
            **dropdown_style,
        )

        self.sheet_selector_wrapper = ft.GestureDetector(
            content=self.sheet_selector,
            on_tap_down=self._on_sheet_dropdown_tap,
        )

        self.workbook_refresh_button = ft.FilledTonalButton(
            text=self._workbook_refresh_button_default_text,
            icon=ft.Icons.SYNC,
            on_click=self._handle_workbook_refresh_click,
            disabled=True,
            style=ft.ButtonStyle(shape=button_shape),
        )

        context_card = ft.Card(
            content=ft.Container(
                padding=ft.Padding(20, 20, 20, 20),
                content=ft.Column(
                    [
                        ft.Text(
                            "\u30b3\u30f3\u30c6\u30ad\u30b9\u30c8",
                            size=16,
                            weight=ft.FontWeight.BOLD,
                            color="#E8DEF8",
                        ),
                        ft.Text(
                            "\u51e6\u7406\u5bfe\u8c61\u3092\u9078\u629e\u3057\u307e\u3059",
                            size=12,
                            color="#CAC4D0",
                        ),
                        ft.Divider(color="#332D41", height=24),
                        ft.Column(
                            [
                                ft.Text("\u30d6\u30c3\u30af", size=13, color="#CAC4D0"),
                                self.workbook_selector_wrapper,
                                ft.Text("\u30b7\u30fc\u30c8", size=13, color="#CAC4D0"),
                                self.sheet_selector_wrapper,
                            ],
                            spacing=12,
                        ),
                        ft.Row(
                            [self.workbook_refresh_button],
                            alignment=ft.MainAxisAlignment.END,
                        ),
                        ft.Row(
                            [
                                ft.Container(content=self.save_log_button, expand=True),
                                ft.Container(content=self.browser_reset_button, expand=True),
                            ],
                            spacing=12,
                        ),
                    ],
                    spacing=16,
                    tight=True,
                ),
            ),
        )

        self.chat_list = ft.ListView(
            expand=True,
            spacing=16,
            auto_scroll=True,
            padding=ft.Padding(0, 12, 0, 12),
        )

        self.user_input = ft.TextField(
            hint_text="",
            expand=True,
            multiline=True,
            min_lines=3,
            max_lines=5,
            on_submit=self._run_copilot,
            border_radius=14,
            border_color="#4F378B",
            focused_border_color="#D0BCFF",
            filled=True,
            fill_color="#2B2930",
        )
        self._apply_mode_to_input_placeholder()

        self.mode_selector = ft.RadioGroup(
            value=self.mode.value,
            on_change=self._on_mode_change,
            content=ft.Row(
                controls=[
                    ft.Radio(value=CopilotMode.TRANSLATION_WITH_REFERENCES.value, label="\u7ffb\u8a33\uff08\u53c2\u7167\u3042\u308a\uff09"),
                    ft.Radio(value=CopilotMode.TRANSLATION.value, label="\u7ffb\u8a33\uff08\u901a\u5e38\uff09"),
                    ft.Radio(value=CopilotMode.REVIEW.value, label="\u7ffb\u8a33\u30c1\u30a7\u30c3\u30af"),
                ],
                alignment=ft.MainAxisAlignment.START,
                spacing=18,
            ),
        )

        action_button_content = self._make_send_button()
        self.action_button = ft.Container(
            content=action_button_content,
            width=56,
            height=56,
            bgcolor="#6750A4",
            border_radius=28,
            alignment=ft.alignment.center,
            ink=True,
            on_hover=self._handle_button_hover,
            animate_scale=100,
            scale=1,
        )

        chat_card = ft.Card(
            expand=True,
            content=ft.Container(
                expand=True,
                padding=ft.Padding(20, 20, 20, 20),
                content=ft.Column(
                    [
                        ft.Row(
                            [
                                ft.Icon(ft.Icons.CHAT_BUBBLE_OUTLINE, size=20, color="#E8DEF8"),
                                ft.Text(
                                    "\u30c1\u30e3\u30c3\u30c8",
                                    size=16,
                                    weight=ft.FontWeight.BOLD,
                                    color="#E8DEF8",
                                ),
                            ],
                            spacing=8,
                        ),
                        ft.Divider(color="#332D41"),
                        self.chat_list,
                    ],
                    spacing=16,
                    expand=True,
                ),
            ),
        )

        composer_card = ft.Card(
            content=ft.Container(
                padding=ft.Padding(20, 20, 20, 20),
                content=ft.Column(
                    [
                        ft.Row(
                            [
                                ft.Icon(ft.Icons.TUNE_ROUNDED, size=20, color="#E8DEF8"),
                                ft.Text(
                                    "\u30e2\u30fc\u30c9\u3068\u6307\u793a",
                                    size=16,
                                    weight=ft.FontWeight.BOLD,
                                    color="#E8DEF8",
                                ),
                            ],
                            spacing=8,
                        ),
                        ft.Container(
                            content=self.mode_selector,
                            bgcolor="#2B2930",
                            border_radius=18,
                            padding=ft.Padding(12, 8, 12, 8),
                        ),
                        self.user_input,
                        ft.Row([self.action_button], alignment=ft.MainAxisAlignment.END),
                    ],
                    spacing=18,
                ),
            ),
        )

        main_column = ft.Column(
            controls=[chat_card, composer_card],
            expand=True,
            spacing=16,
        )

        layout = ft.ResponsiveRow(
            controls=[
                ft.Container(
                    content=ft.Column([context_card], spacing=16),
                    col={"sm": 12, "md": 4, "lg": 3},
                ),
                ft.Container(
                    content=main_column,
                    col={"sm": 12, "md": 8, "lg": 9},
                ),
            ],
            spacing=20,
            run_spacing=20,
            expand=True,
        )

        self.page.add(layout)

    def _register_window_handlers(self):
        self.page.window.on_event = self._on_window_event
        self.page.on_disconnect = self._on_page_disconnect

    def _make_send_button(self) -> ft.IconButton:
        return ft.IconButton(
            icon=ft.Icons.SEND_ROUNDED,
            icon_color="#FFFFFF",
            icon_size=24,
            tooltip="\u9001\u4fe1",
            on_click=self._run_copilot,
            style=ft.ButtonStyle(
                shape=ft.CircleBorder(),
                padding=ft.Padding(0, 0, 0, 0),
            ),
        )

    def _make_stop_button(self) -> ft.IconButton:
        return ft.IconButton(
            icon=ft.Icons.STOP_ROUNDED,
            icon_color="#FFFFFF",
            icon_size=24,
            tooltip="\u51e6\u7406\u3092\u505c\u6b62",
            on_click=self._stop_task,
            style=ft.ButtonStyle(
                shape=ft.CircleBorder(),
                padding=ft.Padding(0, 0, 0, 0),
            ),
        )

    def _handle_button_hover(self, e: ft.ControlEvent):
        if e.data == "true":
            e.control.scale = 1.05
        else:
            e.control.scale = 1
        e.control.update()

    def _apply_mode_to_input_placeholder(self):
        if not self.user_input:
            return
        if self.mode is CopilotMode.TRANSLATION:
            self.user_input.hint_text = "\u7ffb\u8a33\uff08\u901a\u5e38\uff09\u7528\u306e\u6307\u793a\u3092\u5165\u529b\u3057\u3066\u304f\u3060\u3055\u3044\u3002\u4f8b: B\u5217\u3092\u7ffb\u8a33\u3057\u3001\u7d50\u679c\u3092C\u5217\u306b\u66f8\u304d\u8fbc\u3093\u3067\u304f\u3060\u3055\u3044\u3002"
        elif self.mode is CopilotMode.TRANSLATION_WITH_REFERENCES:
            self.user_input.hint_text = "\u7ffb\u8a33\uff08\u53c2\u7167\u3042\u308a\uff09\u7528\u306e\u6307\u793a\u3092\u5165\u529b\u3057\u3066\u304f\u3060\u3055\u3044\u3002\u4f8b: B\u5217\u3092\u7ffb\u8a33\u3057\u3001\u6307\u5b9a\u3057\u305f\u53c2\u7167URL\u3092\u4f7f\u3063\u3066C:E\u5217\u306b\u7ffb\u8a33\u30fb\u5f15\u7528\u30fb\u89e3\u8aac\u3092\u66f8\u304d\u8fbc\u3093\u3067\u304f\u3060\u3055\u3044\u3002"
        else:
            self.user_input.hint_text = "\u7ffb\u8a33\u30c1\u30a7\u30c3\u30af\u306e\u6307\u793a\u3092\u5165\u529b\u3057\u3066\u304f\u3060\u3055\u3044\u3002\u4f8b: \u539f\u6587(B\u5217)\u3001\u7ffb\u8a33(C\u5217)\u3001\u30ec\u30d3\u30e5\u30fc\u7d50\u679c\u3092D:F\u5217\u306b\u66f8\u304d\u8fbc\u3093\u3067\u304f\u3060\u3055\u3044\u3002"

    def _on_mode_change(self, e: Optional[ft.ControlEvent]):
        control = getattr(e, "control", None) if e else None
        selected_value = getattr(control, "value", None)
        if not selected_value:
            return
        try:
            new_mode = CopilotMode(selected_value)
        except ValueError:
            return
        if new_mode == self.mode:
            return
        self.mode = new_mode
        self._apply_mode_to_input_placeholder()
        if self.mode_selector:
            self.mode_selector.value = self.mode.value
        self.request_queue.put(RequestMessage(RequestType.UPDATE_CONTEXT, {"mode": self.mode.value}))
        self._update_ui()

    def _set_state(self, new_state: AppState):
        if self.app_state == new_state:
            return

        self.app_state = new_state
        if new_state is AppState.TASK_IN_PROGRESS:
            self._browser_ready_for_focus = False
            self._pending_focus_action = None
            self._pending_focus_deadline = None
            self._status_message_override = None
            self._status_color_override = None
        is_ready = new_state is AppState.READY
        is_task_in_progress = new_state is AppState.TASK_IN_PROGRESS
        is_stopping = new_state is AppState.STOPPING
        is_error = new_state is AppState.ERROR
        can_interact = is_ready or is_error
        status_palette = {
            "base": "#CAC4D0",
            "ready": "#D0BCFF",
            "busy": "#B69DF8",
            "error": "#F2B8B5",
            "stopping": "#B69DF8",
            "info": "#CAC4D0",
        }

        if self.user_input:
            self.user_input.disabled = not can_interact
        if self.mode_selector:
            self.mode_selector.disabled = not can_interact
        if self.workbook_selector:
            if new_state in {AppState.TASK_IN_PROGRESS, AppState.STOPPING}:
                self.workbook_selector.disabled = True
            else:
                self.workbook_selector.disabled = not (can_interact and bool(self.workbook_selector.options))
        if self.sheet_selector:
            if new_state in {AppState.TASK_IN_PROGRESS, AppState.STOPPING}:
                self.sheet_selector.disabled = True
            else:
                self.sheet_selector.disabled = not (can_interact and bool(self.sheet_selector.options))
        if self.browser_reset_button:
            if new_state in {AppState.TASK_IN_PROGRESS, AppState.STOPPING}:
                self.browser_reset_button.disabled = True
            elif not self._browser_reset_in_progress and can_interact:
                self.browser_reset_button.disabled = False

        if self.workbook_refresh_button:
            if self._manual_refresh_in_progress:
                self.workbook_refresh_button.disabled = True
            else:
                self.workbook_refresh_button.disabled = not can_interact
            if not self._manual_refresh_in_progress and can_interact:
                self.workbook_refresh_button.text = self._workbook_refresh_button_default_text

        if self.status_label:
            self.status_label.opacity = 1
            self.status_label.scale = 1
            if new_state is AppState.INITIALIZING:
                self._status_message_override = None
                self._status_color_override = None
                self.status_label.value = "\u521d\u671f\u5316\u4e2d..."
                self.status_label.color = status_palette["base"]
            elif is_ready:
                if self._status_message_override:
                    self.status_label.value = self._status_message_override
                    self.status_label.color = self._status_color_override or status_palette["ready"]
                else:
                    self.status_label.value = "\u5f85\u6a5f\u4e2d"
                    self.status_label.color = status_palette["ready"]
            elif is_error:
                self._status_message_override = None
                self._status_color_override = None
                self.status_label.value = "\u30a8\u30e9\u30fc"
                self.status_label.color = status_palette["error"]
            else:
                self._status_message_override = None
                self._status_color_override = None
                self.status_label.color = status_palette["base"]

        if self.action_button:
            if is_task_in_progress:
                if self.status_label:
                    self.status_label.value = "\u51e6\u7406\u3092\u5b9f\u884c\u4e2d..."
                    self.status_label.color = status_palette["busy"]
                    self.status_label.opacity = 0.5
                    self.status_label.scale = 0.95
                self.action_button.content = self._make_stop_button()
                self.action_button.disabled = False
            elif is_stopping:
                if self.status_label:
                    self.status_label.value = "\u51e6\u7406\u3092\u505c\u6b62\u3057\u3066\u3044\u307e\u3059..."
                    self.status_label.color = status_palette["stopping"]
                self.action_button.content = ft.ProgressRing(width=18, height=18, stroke_width=2)
                self.action_button.disabled = True
            else:
                self.action_button.content = self._make_send_button()
                self.action_button.disabled = not can_interact

        self._update_ui()

    def _update_ui(self):
        try:
            self.page.update()
        except Exception as e:
            print(f"UI\u306e\u66f4\u65b0\u306b\u5931\u6557\u3057\u307e\u3057\u305f: {e}")

    def _add_message(self, msg_type: Union[ResponseType, str], msg_content: str):
        if not msg_content:
            return

        msg_type_value = msg_type.value if isinstance(msg_type, ResponseType) else str(msg_type)
        self._append_history(msg_type_value, msg_content)
        self._update_save_button_state()

        if not self.chat_list:
            return

        msg = ChatMessage(msg_type, msg_content)
        self.chat_list.controls.append(msg)
        self._update_ui()
        time.sleep(0.01)
        msg.appear()

    def _append_history(self, msg_type: str, msg_content: str):
        entry = {
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "type": msg_type,
            "content": msg_content.replace("\r\n", "\n"),
        }
        with self.history_lock:
            self.chat_history.append(entry)

    def _update_save_button_state(self):
        if not self.save_log_button:
            return
        with self.history_lock:
            has_history = bool(self.chat_history)
        self.save_log_button.disabled = not has_history

    def _handle_save_log_click(self, e: Optional[ft.ControlEvent]):
        try:
            file_path = self._export_chat_history()
        except ValueError as info_err:
            self._add_message(ResponseType.INFO, str(info_err))
            return
        except Exception as ex:
            print(f"\u4f1a\u8a71\u30ed\u30b0\u306e\u66f8\u304d\u51fa\u3057\u306b\u5931\u6557\u3057\u307e\u3057\u305f: {ex}")
            self._add_message(ResponseType.ERROR, f"\u4f1a\u8a71\u30ed\u30b0\u306e\u4fdd\u5b58\u306b\u5931\u6557\u3057\u307e\u3057\u305f: {ex}")
            return

        self._add_message(ResponseType.INFO, f"\u4f1a\u8a71\u30ed\u30b0\u3092\u4fdd\u5b58\u3057\u307e\u3057\u305f: {file_path}")

    def _handle_browser_reset_click(self, e: Optional[ft.ControlEvent]):
        if self._browser_reset_in_progress:
            return
        if self.app_state not in {AppState.READY, AppState.ERROR}:
            return
        if not self.worker or not self.request_queue:
            return

        self._browser_reset_in_progress = True
        if self.browser_reset_button:
            self.browser_reset_button.disabled = True
        self._add_message(ResponseType.INFO, "\u30d6\u30e9\u30a6\u30b6\u306e\u518d\u521d\u671f\u5316\u3092\u5b9f\u884c\u3057\u307e\u3059...")
        self.request_queue.put(RequestMessage(RequestType.RESET_BROWSER))
        self._update_ui()

    def _handle_workbook_refresh_click(self, e: Optional[ft.ControlEvent]):
        if self._manual_refresh_in_progress:
            return
        if self.app_state not in {AppState.READY, AppState.ERROR}:
            return

        self._manual_refresh_in_progress = True
        if self.workbook_refresh_button:
            self.workbook_refresh_button.disabled = True
            self.workbook_refresh_button.text = "\u66f4\u65b0\u4e2d..."
        self._update_ui()

        def _run_refresh():
            try:
                self._refresh_excel_context(auto_triggered=False)
            finally:
                self._manual_refresh_in_progress = False
                if self.workbook_refresh_button:
                    self.workbook_refresh_button.text = self._workbook_refresh_button_default_text
                    self.workbook_refresh_button.disabled = self.app_state not in {AppState.READY, AppState.ERROR}
                self._update_ui()

        invoke_later = getattr(self.page, "invoke_later", None)
        if callable(invoke_later):
            try:
                invoke_later(_run_refresh)
                return
            except Exception as invoke_err:
                print(f"invoke_later for manual refresh failed: {invoke_err}")
        _run_refresh()


    def _export_chat_history(self) -> Path:
        with self.history_lock:
            if not self.chat_history:
                raise ValueError("\u4fdd\u5b58\u3067\u304d\u308b\u4f1a\u8a71\u5c65\u6b74\u304c\u3042\u308a\u307e\u305b\u3093\u3002")
            entries = [entry.copy() for entry in self.chat_history]

        self.log_dir.mkdir(parents=True, exist_ok=True)
        export_time = datetime.now()
        file_path = self.log_dir / f"conversation-{export_time.strftime('%Y%m%d-%H%M%S')}.md"

        workbook_display_name = self.current_workbook_name or "\u4e0d\u660e"
        sheet_display_name = self.current_sheet_name or "\u4e0d\u660e"
        lines = [
            "# Excel Co-pilot \u4f1a\u8a71\u30ed\u30b0",
            f"- \u30a8\u30af\u30b9\u30dd\u30fc\u30c8\u6642\u523b: {export_time.isoformat(timespec='seconds')}",
            f"- \u5bfe\u8c61\u30d6\u30c3\u30af: {workbook_display_name}",
            f"- \u5bfe\u8c61\u30b7\u30fc\u30c8: {sheet_display_name}",
            "",
        ]

        for entry in entries:
            lines.append(f"## [{entry['timestamp']}] {entry['type']}")
            lines.append(entry["content"])
            lines.append("")

        file_path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")
        return file_path

    def _load_last_workbook_preference(self) -> Optional[str]:
        with self.preference_lock:
            if not self.preference_file.exists():
                return None
            try:
                raw_text = self.preference_file.read_text(encoding="utf-8")
                data = json.loads(raw_text) if raw_text else {}
            except (OSError, json.JSONDecodeError) as err:
                print(f"Failed to load workbook preference: {err}")
                return None

        if isinstance(data, dict):
            value = data.get(PREFERENCE_LAST_WORKBOOK_KEY)
            if isinstance(value, str) and value.strip():
                return value.strip()
        return None

    def _save_last_workbook_preference(self, workbook_name: Optional[str]):
        if not workbook_name:
            return
        with self.preference_lock:
            try:
                if self.preference_file.exists():
                    raw_text = self.preference_file.read_text(encoding="utf-8")
                    loaded = json.loads(raw_text) if raw_text else {}
                    preferences = dict(loaded) if isinstance(loaded, dict) else {}
                else:
                    preferences = {}
            except (OSError, json.JSONDecodeError) as err:
                print(f"Failed to read workbook preference: {err}")
                preferences = {}

            preferences[PREFERENCE_LAST_WORKBOOK_KEY] = workbook_name

            try:
                self.preference_file.parent.mkdir(parents=True, exist_ok=True)
                self.preference_file.write_text(json.dumps(preferences, ensure_ascii=False, indent=2), encoding="utf-8")
            except OSError as err:
                print(f"Failed to write workbook preference: {err}")

    def _load_last_sheet_preference(self, workbook_name: Optional[str]) -> Optional[str]:
        if not workbook_name:
            return None
        with self.preference_lock:
            if not self.preference_file.exists():
                return None
            try:
                raw_text = self.preference_file.read_text(encoding="utf-8")
                data = json.loads(raw_text) if raw_text else {}
            except (OSError, json.JSONDecodeError) as err:
                print(f"Failed to load sheet preference: {err}")
                return None
            value = data.get(workbook_name) if isinstance(data, dict) else None
        if isinstance(value, str) and value.strip():
            return value.strip()
        return None

    def _save_last_sheet_preference(self, workbook_name: Optional[str], sheet_name: Optional[str]):
        if not workbook_name or not sheet_name:
            return
        with self.preference_lock:
            try:
                if self.preference_file.exists():
                    raw_text = self.preference_file.read_text(encoding="utf-8")
                    loaded = json.loads(raw_text) if raw_text else {}
                    if isinstance(loaded, dict):
                        preferences = dict(loaded)
                    else:
                        preferences = {}
                else:
                    preferences = {}
            except (OSError, json.JSONDecodeError) as err:
                print(f"Failed to read sheet preference: {err}")
                preferences = {}

            reserved_entries = {
                key: value for key, value in preferences.items() if isinstance(key, str) and key.startswith("__")
            }
            mutable_pairs = [
                (key, value)
                for key, value in preferences.items()
                if isinstance(key, str) and not key.startswith("__") and isinstance(value, str)
            ]
            mutable_pairs = [(key, value) for key, value in mutable_pairs if key != workbook_name]
            mutable_pairs.append((workbook_name, sheet_name))
            trimmed_items = mutable_pairs[-50:]
            trimmed_data = dict(trimmed_items)
            trimmed_data.update(reserved_entries)

            try:
                self.preference_file.parent.mkdir(parents=True, exist_ok=True)
                self.preference_file.write_text(json.dumps(trimmed_data, ensure_ascii=False, indent=2), encoding="utf-8")
            except OSError as err:
                print(f"Failed to write sheet preference: {err}")

    def _refresh_excel_context(
        self,
        is_initial_start: bool = False,
        desired_workbook: Optional[str] = None,
        auto_triggered: bool = False,
    ) -> Optional[str]:
        if not self.sheet_selector or not self.workbook_selector or not self.ui_loop_running:
            return None

        with self._excel_refresh_lock:
            selector_value = None
            if self.workbook_selector:
                selector_value = self.workbook_selector.value

            target_workbook = (
                desired_workbook
                or selector_value
                or self.current_workbook_name
                or self._load_last_workbook_preference()
            )

            try:
                with ExcelManager(target_workbook) as manager:
                    workbook_names = manager.list_workbook_names()
                    if not workbook_names:
                        raise ExcelConnectionError("開いているExcelブックが見つかりません。")

                    if (
                        target_workbook
                        and target_workbook in workbook_names
                        and not auto_triggered
                    ):
                        try:
                            manager.activate_workbook(target_workbook)
                        except Exception as activate_err:
                            print(f"対象ブック '{target_workbook}' の選択に失敗しました: {activate_err}")

                    info_dict = manager.get_active_workbook_and_sheet()
                    active_workbook = info_dict["workbook_name"]
                    active_sheet = info_dict["sheet_name"]

                    sheet_names = manager.list_sheet_names()

                    preferred_sheet = self._load_last_sheet_preference(active_workbook)
                    if (
                        preferred_sheet
                        and preferred_sheet in sheet_names
                        and preferred_sheet != active_sheet
                        and not auto_triggered
                    ):
                        try:
                            active_sheet = manager.activate_sheet(preferred_sheet)
                        except Exception as activate_err:
                            print(
                                f"前回選択したシート '{preferred_sheet}' の復元に失敗しました: {activate_err}"
                            )
                            self._add_message(
                                ResponseType.INFO,
                                f"保存済みシート『{preferred_sheet}』を開けませんでした: {activate_err}"
                            )

                snapshot = {
                    "workbooks": tuple(workbook_names),
                    "workbook": active_workbook,
                    "sheet": active_sheet,
                    "sheets": tuple(sheet_names),
                }

                if auto_triggered and snapshot == self._last_excel_snapshot:
                    return active_sheet

                self._last_excel_snapshot = snapshot

                controls_changed = False

                existing_workbook_values = [
                    (opt.key or opt.text) for opt in (self.workbook_selector.options or [])
                ]
                if existing_workbook_values != workbook_names:
                    self.workbook_selection_updating = True
                    self.workbook_selector.options = [ft.dropdown.Option(name) for name in workbook_names]
                    self.workbook_selection_updating = False
                    controls_changed = True

                if not auto_triggered:
                    if self.workbook_selector.value != active_workbook:
                        self.workbook_selection_updating = True
                        self.workbook_selector.value = active_workbook
                        self.workbook_selection_updating = False
                        controls_changed = True
                elif not self.workbook_selector.value and active_workbook:
                    self.workbook_selection_updating = True
                    self.workbook_selector.value = active_workbook
                    self.workbook_selection_updating = False
                    controls_changed = True
                if self.workbook_selector.disabled:
                    self.workbook_selector.disabled = False
                    controls_changed = True

                existing_sheet_values = [
                    (opt.key or opt.text) for opt in (self.sheet_selector.options or [])
                ]
                if existing_sheet_values != sheet_names:
                    self.sheet_selection_updating = True
                    self.sheet_selector.options = [ft.dropdown.Option(name) for name in sheet_names]
                    self.sheet_selection_updating = False
                    controls_changed = True

                if not auto_triggered:
                    if self.sheet_selector.value != active_sheet:
                        self.sheet_selection_updating = True
                        self.sheet_selector.value = active_sheet
                        self.sheet_selection_updating = False
                        controls_changed = True
                elif not self.sheet_selector.value and active_sheet:
                    self.sheet_selection_updating = True
                    self.sheet_selector.value = active_sheet
                    self.sheet_selection_updating = False
                    controls_changed = True
                if self.sheet_selector.disabled:
                    self.sheet_selector.disabled = False
                    controls_changed = True

                context_changed = False
                if not auto_triggered:
                    if active_workbook != self.current_workbook_name:
                        self.current_workbook_name = active_workbook
                        context_changed = True
                    if active_sheet != self.current_sheet_name:
                        self.current_sheet_name = active_sheet
                        context_changed = True
                else:
                    if self.current_workbook_name is None and active_workbook:
                        self.current_workbook_name = active_workbook
                    if self.current_sheet_name is None and active_sheet:
                        self.current_sheet_name = active_sheet

                if self.current_workbook_name:
                    self._save_last_workbook_preference(self.current_workbook_name)
                if self.current_workbook_name and self.current_sheet_name:
                    self._save_last_sheet_preference(self.current_workbook_name, self.current_sheet_name)

                if context_changed and self.request_queue:
                    payload: Dict[str, Any] = {"workbook_name": self.current_workbook_name}
                    if self.current_sheet_name:
                        payload["sheet_name"] = self.current_sheet_name
                    self.request_queue.put(RequestMessage(RequestType.UPDATE_CONTEXT, payload))

                if context_changed or controls_changed or is_initial_start:
                    self._update_ui()

                return active_sheet

            except Exception as ex:
                error_message = f"Excelの情報取得に失敗しました: {ex}"
                self.sheet_selection_updating = True
                self.sheet_selector.options = []
                self.sheet_selector.value = None
                self.sheet_selector.disabled = True
                self.sheet_selection_updating = False

                self.workbook_selection_updating = True
                self.workbook_selector.options = []
                self.workbook_selector.value = None
                self.workbook_selector.disabled = True
                self.workbook_selection_updating = False

                self._last_excel_snapshot = {}
                if not auto_triggered and not is_initial_start:
                    self._add_message(ResponseType.ERROR, error_message)
                self._update_ui()
                return None

    def _start_background_excel_polling(self):
        if self._excel_poll_thread and self._excel_poll_thread.is_alive():
            return
        self._excel_poll_stop_event.clear()
        self._excel_refresh_event.set()
        self._excel_poll_thread = threading.Thread(
            target=self._excel_polling_loop,
            name="excel-context-poll",
            daemon=True,
        )
        self._excel_poll_thread.start()

    def _stop_background_excel_polling(self):
        self._excel_poll_stop_event.set()
        self._excel_refresh_event.set()
        thread = self._excel_poll_thread
        if thread and thread.is_alive():
            try:
                thread.join(timeout=2.0)
            except Exception as join_err:
                print(f"Excel poll thread join failed: {join_err}")
        self._excel_poll_thread = None

    def _excel_polling_loop(self):
        while not self._excel_poll_stop_event.is_set():
            try:
                triggered = self._excel_refresh_event.wait(timeout=self._excel_poll_interval)
                self._excel_refresh_event.clear()
            except Exception as wait_err:
                print(f"Excel poll wait failed: {wait_err}")
            if self._excel_poll_stop_event.is_set():
                break
            if self.app_state in {AppState.TASK_IN_PROGRESS, AppState.STOPPING}:
                if triggered:
                    time.sleep(0.2)
                continue
            self._invoke_excel_refresh(auto_triggered=True)

    def _invoke_excel_refresh(self, auto_triggered: bool):
        if not self.ui_loop_running:
            return

        def _run():
            if not self.ui_loop_running:
                return
            self._refresh_excel_context(auto_triggered=auto_triggered)

        invoke_later = getattr(self.page, "invoke_later", None)
        if callable(invoke_later):
            try:
                invoke_later(_run)
                return
            except Exception as invoke_err:
                print(f"invoke_later failed, running refresh inline: {invoke_err}")
        _run()

    def _request_background_excel_refresh(self):
        # Excel list updates are manual; background refresh triggers are disabled.
        return

    def _run_copilot(self, e: Optional[ft.ControlEvent]):
        if not self.user_input:
            return
        user_text = self.user_input.value
        if not user_text or self.app_state not in {AppState.READY, AppState.ERROR}:
            return

        self._set_state(AppState.TASK_IN_PROGRESS)
        self._add_message("user", user_text)
        self.user_input.value = ""
        self.request_queue.put(RequestMessage(RequestType.USER_INPUT, user_text))
        self._update_ui()

    def _stop_task(self, e: Optional[ft.ControlEvent]):
        if self.app_state is not AppState.TASK_IN_PROGRESS:
            return

        self._set_state(AppState.STOPPING)
        if self.worker:
            # Ensure the worker sees the stop request even while busy executing the task.
            self.worker.stop_event.set()
        self.request_queue.put(RequestMessage(RequestType.STOP))

    def _on_workbook_change(self, e: ft.ControlEvent):
        if self.workbook_selection_updating:
            return
        selected_workbook = e.control.value if e and e.control else None
        if not selected_workbook:
            return
        self._save_last_workbook_preference(selected_workbook)
        self._refresh_excel_context(desired_workbook=selected_workbook)
        if self.user_input and not self.user_input.disabled:
            try:
                self.user_input.focus()
            except Exception:
                pass

    def _refresh_excel_context_before_dropdown(self):
        # Excel list updates are manual; skip automatic refresh on dropdown events.
        return

    def _on_workbook_dropdown_focus(self, e: Optional[ft.ControlEvent]):
        if not self.workbook_selector or self.workbook_selector.disabled:
            return
        self._refresh_excel_context_before_dropdown()

    def _on_workbook_dropdown_tap(self, e: Optional[ft.TapEvent]):
        if not self.workbook_selector or self.workbook_selector.disabled:
            return
        self._refresh_excel_context_before_dropdown()

    def _on_sheet_change(self, e: ft.ControlEvent):
        if self.sheet_selection_updating:
            return
        selected_sheet = e.control.value if e and e.control else None
        if not selected_sheet:
            return

        previous_sheet = self.current_sheet_name
        try:
            with ExcelManager(self.current_workbook_name) as manager:
                if self.current_workbook_name:
                    try:
                        manager.activate_workbook(self.current_workbook_name)
                    except Exception:
                        pass
                manager.activate_sheet(selected_sheet)
        except Exception as ex:
            error_message = f"\u30b7\u30fc\u30c8\u306e\u5207\u308a\u66ff\u3048\u306b\u5931\u6557\u3057\u307e\u3057\u305f: {ex}"
            self.sheet_selection_updating = True
            if self.sheet_selector:
                self.sheet_selector.value = previous_sheet
            self.sheet_selection_updating = False
            self._add_message(ResponseType.ERROR, error_message)
            self._update_ui()
            return

        payload: Dict[str, Any] = {"sheet_name": selected_sheet}
        if self.current_workbook_name:
            payload["workbook_name"] = self.current_workbook_name
        self.request_queue.put(RequestMessage(RequestType.UPDATE_CONTEXT, payload))
        self.current_sheet_name = selected_sheet
        if self.current_workbook_name:
            self._save_last_sheet_preference(self.current_workbook_name, selected_sheet)
            self._save_last_workbook_preference(self.current_workbook_name)
        self._update_ui()
        if self.user_input and not self.user_input.disabled:
            try:
                self.user_input.focus()
            except Exception:
                pass

    def _on_sheet_dropdown_focus(self, e: Optional[ft.ControlEvent]):
        if not self.sheet_selector or self.sheet_selector.disabled:
            return
        self._refresh_excel_context_before_dropdown()

    def _on_sheet_dropdown_tap(self, e: Optional[ft.TapEvent]):
        if not self.sheet_selector or self.sheet_selector.disabled:
            return
        self._refresh_excel_context_before_dropdown()

    def _process_response_queue_loop(self):
        while self.ui_loop_running:
            try:
                raw_message = self.response_queue.get(timeout=0.1)
            except queue.Empty:
                continue
            except Exception as e:
                print(f"\u30ec\u30b9\u30dd\u30f3\u30b9\u30ad\u30e5\u30fc\u51e6\u7406\u4e2d\u306b\u30a8\u30e9\u30fc\u304c\u767a\u751f\u3057\u307e\u3057\u305f: {e}")
                continue

            try:
                response = ResponseMessage.from_raw(raw_message)
            except ValueError as exc:
                print(f"\u30ec\u30b9\u30dd\u30f3\u30b9\u306e\u89e3\u6790\u306b\u5931\u6557\u3057\u307e\u3057\u305f: {exc}")
                continue

            self._display_response(response)

    def _display_response(self, response: ResponseMessage):
        type_value = response.metadata.get("source_type", response.type.value)
        status_palette = {
            "base": "#CAC4D0",
            "info": "#D0BCFF",
            "error": "#F2B8B5",
        }

        if (
            self._pending_focus_action == "focus_excel_window"
            and self._pending_focus_deadline is not None
            and time.monotonic() >= self._pending_focus_deadline
        ):
            print("Excel focus fallback triggered after waiting for browser readiness timeout.")
            self._focus_excel_window()
            self._pending_focus_action = None
            self._pending_focus_deadline = None

        browser_ready = bool(response.metadata.get("browser_ready"))
        if browser_ready:
            self._browser_ready_for_focus = True
            if self._pending_focus_action == "focus_excel_window":
                self._focus_excel_window()
                self._pending_focus_action = None
                self._pending_focus_deadline = None
            if self._browser_reset_in_progress:
                self._browser_reset_in_progress = False
                if self.browser_reset_button and self.app_state in {AppState.READY, AppState.ERROR}:
                    self.browser_reset_button.disabled = False

        if response.type is ResponseType.INITIALIZATION_COMPLETE:
            self._set_state(AppState.READY)
            if self.status_label:
                self.status_label.value = response.content or self.status_label.value
            self._focus_app_window()
        elif response.type is ResponseType.STATUS:
            status_text = (response.content or "").strip()
            if status_text:
                self._status_message_override = status_text
                self._status_color_override = status_palette["info"]
            else:
                self._status_message_override = None
                self._status_color_override = None
            if self.status_label:
                self.status_label.value = status_text
                if status_text:
                    self.status_label.color = self._status_color_override or status_palette["info"]
        elif response.type is ResponseType.ERROR:
            if self.app_state in {AppState.TASK_IN_PROGRESS, AppState.STOPPING}:
                if self.status_label:
                    self.status_label.value = response.content or "\u51e6\u7406\u4e2d\u306b\u30a8\u30e9\u30fc\u304c\u767a\u751f\u3057\u307e\u3057\u305f"
                    self.status_label.color = status_palette["error"]
                    self.status_label.opacity = 0.9
                if response.content:
                    self._add_message(response.type, response.content)
            else:
                self._set_state(AppState.ERROR)
                if response.content:
                    self._add_message(response.type, response.content)
            if self._browser_reset_in_progress:
                self._browser_reset_in_progress = False
                if self.browser_reset_button and self.app_state in {AppState.READY, AppState.ERROR}:
                    self.browser_reset_button.disabled = False
        elif response.type is ResponseType.END_OF_TASK:
            self._set_state(AppState.READY)
        elif response.type is ResponseType.INFO:
            action = response.metadata.get("action") if response.metadata else None
            if action == "focus_excel_window":
                wait_for_browser_ready = bool(response.metadata.get("wait_for_browser_ready")) if response.metadata else False
                if wait_for_browser_ready and not self._browser_ready_for_focus:
                    self._pending_focus_action = "focus_excel_window"
                    self._pending_focus_deadline = time.monotonic() + self._focus_wait_timeout_sec
                else:
                    self._focus_excel_window()
                    self._pending_focus_action = None
                    self._pending_focus_deadline = None
            elif action == "focus_app_window":
                self._pending_focus_action = None
                self._pending_focus_deadline = None
                self._focus_app_window()
                if self._browser_reset_in_progress:
                    self._browser_reset_in_progress = False
                    if self.browser_reset_button and self.app_state in {AppState.READY, AppState.ERROR}:
                        self.browser_reset_button.disabled = False
            elif response.content:
                self._add_message(type_value, response.content)
        else:
            if response.content:
                self._add_message(type_value, response.content)

        self._update_ui()

    def _force_exit(self, reason: str = ""):
        if self.shutdown_requested:
            print("Force exit: shutdown already in progress.")
        else:
            print(f"Force exit triggered. reason={reason}")
            self.shutdown_requested = True
            self._stop_background_excel_polling()
            if self.ui_loop_running:
                self.ui_loop_running = False
                try:
                    self.request_queue.put_nowait(RequestMessage(RequestType.QUIT))
                    print("Force exit: QUIT request posted.")
                except Exception as queue_err:
                    print(f"Force exit: failed to enqueue QUIT: {queue_err}")
                if self.worker_thread:
                    try:
                        self.worker_thread.join(timeout=3.0)
                        print("Force exit: worker thread joined or timeout.")
                    except Exception as join_err:
                        print(f"Force exit: worker join error: {join_err}")
            try:
                self.page.window.prevent_close = False
            except Exception as prevent_err:
                print(f"Force exit: unable to clear prevent_close: {prevent_err}")

            close_requested = False
            try:
                self.page.window.close()
                close_requested = True
                print("Force exit: window.close() called.")
            except AttributeError:
                try:
                    self.page.window.destroy()
                    close_requested = True
                    self.window_closed_event.set()
                    print("Force exit: window.destroy() called.")
                except Exception as destroy_err:
                    print(f"Force exit: window destroy failed: {destroy_err}")
            except Exception as close_err:
                print(f"Force exit: window close failed: {close_err}")

            if close_requested:
                try:
                    self.page.update()
                except Exception as update_err:
                    print(f"Force exit: page update after close failed: {update_err}")

        if not self.window_closed_event.is_set():
            try:
                if self.window_closed_event.wait(timeout=3.0):
                    print("Force exit: window close confirmed.")
                else:
                    print("Force exit: window close wait timed out.")
            except Exception as wait_err:
                print(f"Force exit: waiting for window close failed: {wait_err}")
        os._exit(0)

    def _on_window_event(self, e: ft.ControlEvent):
        event_name = getattr(e, "event", None)
        data = getattr(e, "data", None)
        payload_raw = event_name or data or ""
        payload = str(payload_raw).lower()
        normalized_payload = payload.replace("_", "-")
        print(f"Window event received: event={event_name}, data={data}")

        window_gone_events = {"closed", "close-completed", "destroyed"}
        close_request_events = {"close", "closing", "close-requested"}
        if normalized_payload in window_gone_events:
            self.window_closed_event.set()
        if normalized_payload in close_request_events or (
            normalized_payload in window_gone_events and not self.shutdown_requested
        ):
            self._force_exit(reason=f"window-event:{normalized_payload}")

    def _on_page_disconnect(self, e: ft.ControlEvent):
        print("Page disconnect detected.")
        self.window_closed_event.set()
        self._force_exit(reason="page-disconnect")

def main(page: ft.Page):
    app = CopilotApp(page)
    page.copilot_app = app
    app.mount()

def _parse_cli_args():
    parser = argparse.ArgumentParser(
        description="Launch the Excel Copilot Flet application.")
    parser.add_argument("--host", help="Host interface to bind the Flet web server.")
    parser.add_argument("--port", type=int, help="Port to bind the Flet web server.")
    parser.add_argument("--no-browser", action="store_true",
                        help="Run without launching the bundled Flet viewer.")
    parser.add_argument("--web-renderer", choices=["auto", "html", "canvaskit"],
                        help="Select the Flet web renderer variant.")
    return parser.parse_args()

if __name__ == "__main__":
    args = _parse_cli_args()
    app_kwargs = dict(target=main)
    if args.host:
        app_kwargs["host"] = args.host
    if args.port:
        app_kwargs["port"] = args.port
    if args.web_renderer:
        app_kwargs["web_renderer"] = args.web_renderer
    if args.no_browser or COPILOT_HEADLESS:
        app_kwargs["view"] = None
    else:
        app_kwargs["view"] = ft.AppView.WEB_BROWSER
    ft.app(**app_kwargs)
