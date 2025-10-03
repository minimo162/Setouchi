# desktop_app.py

import flet as ft
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
    def __init__(self, request_q: queue.Queue, response_q: queue.Queue, sheet_name: Optional[str] = None):
        self.request_queue = request_q
        self.response_queue = response_q
        self.browser_manager: Optional[BrowserCopilotManager] = None
        self.agent: Optional[ReActAgent] = None
        self.stop_event = threading.Event()
        self.sheet_name = sheet_name
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
            mode=self.mode,
            progress_callback=lambda msg: self._emit_response(ResponseMessage(ResponseType.OBSERVATION, msg)),
        )

    def _format_user_prompt(self, user_input: str) -> str:
        trimmed_input = (user_input or "").strip()
        if self.mode is CopilotMode.TRANSLATION:
            prefix_lines = [
                "[Translation (No References) Mode Request]",
                "- Solve this by calling `translate_range_without_references` with explicit source and output ranges.",
                "- Keep translation, quote, and explanation columns aligned with the specified output range.",
                "- Do not request workbook uploads; Excel is already connected.",
            ]
        elif self.mode is CopilotMode.TRANSLATION_WITH_REFERENCES:
            prefix_lines = [
                "[Translation (With References) Mode Request]",
                "- Solve this by calling `translate_range_with_references` and include the reference ranges or URLs provided by the user.",
                "- Work one cell at a time without `rows_per_batch`; split multi-row ranges across multiple calls.",
                "- Provide citation output when evidence is expected and keep translation, quote, and explanation columns aligned.",
                "- Do not request workbook uploads; Excel is already connected.",
            ]
        else:
            prefix_lines = [
                "[Translation Review Mode Request]",
                "- Use `check_translation_quality` with all required ranges for status, issues, and corrections.",
                "- Clearly identify which range contains the Japanese source text and which range contains the English translation under review.",
                "- Keep outputs aligned with the ranges specified in the instructions.",
                "- Do not request workbook uploads; Excel is already connected.",
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
            if request.type is RequestType.USER_INPUT:
                if isinstance(request.payload, str):
                    self._execute_task(request.payload)
                else:
                    self._emit_response(ResponseMessage(ResponseType.ERROR, "\u30e6\u30fc\u30b6\u30fc\u5165\u529b\u306e\u5f62\u5f0f\u304c\u4e0d\u6b63\u3067\u3059\u3002"))

    def _update_context(self, payload: Optional[Dict[str, Any]]):
        if not isinstance(payload, dict):
            return

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
                        CopilotMode.TRANSLATION: "\u7ffb\u8a33\uff08\u53c2\u7167\u306a\u3057\uff09",
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
            if self.stop_event.is_set():
                self._emit_response(ResponseMessage(ResponseType.INFO, "\u30e6\u30fc\u30b6\u30fc\u306b\u3088\u3063\u3066\u30bf\u30b9\u30af\u304c\u4e2d\u65ad\u3055\u308c\u307e\u3057\u305f\u3002"))
            self._emit_response(ResponseMessage(ResponseType.END_OF_TASK))

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

        type_map = {
            "user": {
                "bgcolor": "#3C3A4A",
                "icon": ft.Icons.PERSON_ROUNDED,
                "icon_color": "#FFFFFF",
                "text_style": {"color": "#FFFFFF", "size": 14},
            },
            "thought": {
                "icon": ft.Icons.LIGHTBULB_OUTLINE,
                "icon_color": "#D1C4E9",
                "text_style": {"italic": True, "color": "#D1C4E9", "size": 13},
                "bgcolor": "transparent",
            },
            "action": {
                "icon": ft.Icons.CODE,
                "icon_color": "#B39DDB",
                "text_style": {"font_family": "monospace", "color": "#E0E0E0", "size": 13},
                "bgcolor": "#2C2A3A",
                "title": "Action",
            },
            "observation": {
                "icon": ft.Icons.FIND_IN_PAGE_OUTLINED,
                "icon_color": "#B39DDB",
                "bgcolor": "#2C2A3A",
                "title": "Observation",
            },
            "final_answer": {
                "icon": ft.Icons.CHECK_CIRCLE_OUTLINE,
                "icon_color": "#81C784",
                "bgcolor": "#2E4434",
                "title": "Answer",
            },
            "info": {
                "text_style": {"color": "#90A4AE", "size": 12},
            },
            "status": {
                "text_style": {"color": "#90A4AE", "size": 12},
            },
            "error": {
                "icon": ft.Icons.ERROR_OUTLINE_ROUNDED,
                "icon_color": "#E57373",
                "bgcolor": "#5E2A2A",
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
            content=ft.Column(content_controls, spacing=5, tight=True),
            bgcolor=config.get("bgcolor"),
            border_radius=12,
            padding=12,
            expand=True,
            shadow=ft.BoxShadow(
                spread_radius=1,
                blur_radius=10,
                color="#1A000000",
                offset=ft.Offset(2, 2),
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

        self.mode = CopilotMode.TRANSLATION_WITH_REFERENCES
        self.mode_selector: Optional[ft.RadioGroup] = None

        self.title_label: Optional[ft.Text] = None
        self.status_label: Optional[ft.Text] = None
        self.excel_info_label: Optional[ft.Text] = None
        self.refresh_button: Optional[ft.ElevatedButton] = None
        self.sheet_selector: Optional[ft.Dropdown] = None
        self.chat_list: Optional[ft.ListView] = None
        self.user_input: Optional[ft.TextField] = None
        self.action_button: Optional[ft.Container] = None

        self.chat_history: list[dict[str, str]] = []
        self.history_lock = threading.Lock()
        self.log_dir = Path(COPILOT_USER_DATA_DIR) / "setouchi_logs"
        self.preference_file = Path(COPILOT_USER_DATA_DIR) / "setouchi_state.json"
        self.preference_lock = threading.Lock()

        self._configure_page()
        self._build_layout()
        self._register_window_handlers()

    def mount(self):
        self._set_state(AppState.INITIALIZING)
        self._update_ui()
        sheet_name = self._refresh_excel_context(is_initial_start=True)

        self.worker = CopilotWorker(self.request_queue, self.response_queue, sheet_name)
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
        self.page.theme_mode = ft.ThemeMode.DARK
        self.page.bgcolor = "#141218"
        self.page.window.center()
        self.page.window.prevent_close = True

    def _build_layout(self):
        self.title_label = ft.Text("Excel\nCo-pilot", size=26, weight=ft.FontWeight.BOLD, color="#FFFFFF")
        self.status_label = ft.Text("\u521d\u671f\u5316\u4e2d...", size=15, color=ft.Colors.GREY_500, animate_opacity=300, animate_scale=600)

        self.excel_info_label = ft.Text("", size=14, color="#CFD8DC")
        self.refresh_button = ft.ElevatedButton(
            text="\u66f4\u65b0",
            on_click=self._handle_refresh_click,
            bgcolor=ft.Colors.DEEP_PURPLE_500,
            color=ft.Colors.WHITE,
            scale=1,
            animate_scale=100,
            on_hover=self._handle_button_hover,
        )

        self.save_log_button = ft.ElevatedButton(
            text="\u4f1a\u8a71\u30ed\u30b0\u3092\u4fdd\u5b58",
            icon=ft.Icons.SAVE_OUTLINED,
            on_click=self._handle_save_log_click,
            bgcolor="#2C2A3A",
            color=ft.Colors.WHITE,
            disabled=True,
            animate_scale=100,
            scale=1,
            on_hover=self._handle_button_hover,
        )

        self.sheet_selector = ft.Dropdown(
            options=[],
            width=180,
            on_change=self._on_sheet_change,
            hint_text="\u30b7\u30fc\u30c8\u3092\u9078\u629e",
            border_radius=8,
            fill_color="#2C2A3A",
            text_style=ft.TextStyle(color=ft.Colors.WHITE),
            disabled=True,
        )

        sidebar_content = ft.Column(
            [
                self.title_label,
                self.status_label,
                ft.Divider(color="#4A4458"),
                self.excel_info_label,
                self.refresh_button,
                self.save_log_button,
                self.sheet_selector,
            ],
            width=220,
            spacing=20,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        )

        sidebar = ft.Container(
            sidebar_content,
            padding=15,
            border_radius=15,
            gradient=ft.LinearGradient(
                begin=ft.alignment.top_center,
                end=ft.alignment.bottom_center,
                colors=["#2A243A", "#1C1A24"],
            ),
            shadow=ft.BoxShadow(
                spread_radius=1,
                blur_radius=15,
                color="#1A000000",
                offset=ft.Offset(5, 5),
            ),
        )

        self.chat_list = ft.ListView(expand=True, spacing=15, auto_scroll=True, padding=20)
        self.user_input = ft.TextField(
            hint_text="",
            expand=True,
            on_submit=self._run_copilot,
            border_color="#4A4458",
            focused_border_color="#B39DDB",
            border_radius=10,
        )
        self._apply_mode_to_input_placeholder()

        self.mode_selector = ft.RadioGroup(
            value=self.mode.value,
            on_change=self._on_mode_change,
            content=ft.Row(
                controls=[
                    ft.Radio(value=CopilotMode.TRANSLATION_WITH_REFERENCES.value, label="\u7ffb\u8a33\uff08\u53c2\u7167\u3042\u308a\uff09"),
                    ft.Radio(value=CopilotMode.TRANSLATION.value, label="\u7ffb\u8a33\uff08\u53c2\u7167\u306a\u3057\uff09"),
                    ft.Radio(value=CopilotMode.REVIEW.value, label="\u7ffb\u8a33\u30c1\u30a7\u30c3\u30af"),
                ],
                alignment=ft.MainAxisAlignment.START,
                spacing=16,
            ),
        )
        mode_selector_container = ft.Container(
            content=self.mode_selector,
            padding=ft.Padding(left=8, top=4, right=8, bottom=8),
        )

        action_button_content = self._make_send_button()
        self.action_button = ft.Container(
            content=action_button_content,
            scale=1,
            animate_scale=100,
            on_hover=self._handle_button_hover,
            bgcolor="#2C2A3A",
            border_radius=30,
        )

        input_row = ft.Row([self.user_input, self.action_button], alignment=ft.MainAxisAlignment.CENTER)

        main_content = ft.Column(
            [
                self.chat_list,
                mode_selector_container,
                input_row,
            ],
            expand=True,
        )

        self.page.add(
            ft.Row(
                [
                    sidebar,
                    main_content,
                ],
                expand=True,
                spacing=20,
                vertical_alignment=ft.CrossAxisAlignment.STRETCH,
            )
        )

    def _register_window_handlers(self):
        self.page.window.on_event = self._on_window_event
        self.page.on_disconnect = self._on_page_disconnect

    def _make_send_button(self) -> ft.IconButton:
        return ft.IconButton(icon=ft.Icons.SEND_ROUNDED, on_click=self._run_copilot, icon_color="#B39DDB", tooltip="\u9001\u4fe1")

    def _make_stop_button(self) -> ft.IconButton:
        return ft.IconButton(icon=ft.Icons.STOP_ROUNDED, on_click=self._stop_task, icon_color="#B39DDB", tooltip="\u51e6\u7406\u3092\u505c\u6b62")

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
            self.user_input.hint_text = "\u7ffb\u8a33\uff08\u53c2\u7167\u306a\u3057\uff09\u7528\u306e\u6307\u793a\u3092\u5165\u529b\u3057\u3066\u304f\u3060\u3055\u3044\u3002\u4f8b: B\u5217\u3092\u7ffb\u8a33\u3057\u3001\u7d50\u679c\u3092C:E\u5217\u306b\u66f8\u304d\u8fbc\u3093\u3067\u304f\u3060\u3055\u3044\u3002"
        elif self.mode is CopilotMode.TRANSLATION_WITH_REFERENCES:
            self.user_input.hint_text = "\u7ffb\u8a33\uff08\u53c2\u7167\u3042\u308a\uff09\u7528\u306e\u6307\u793a\u3092\u5165\u529b\u3057\u3066\u304f\u3060\u3055\u3044\u3002\u4f8b: B\u5217\u3092\u7ffb\u8a33\u3057\u3001\u6307\u5b9a\u3057\u305f\u53c2\u7167URL\u3092\u4f7f\u3063\u3066C:E\u5217\u306b\u7ffb\u8a33\u30fb\u5f15\u7528\u30fb\u89e3\u8aac\u3092\u66f8\u304d\u8fbc\u3093\u3067\u304f\u3060\u3055\u3044\u3002"
        else:
            self.user_input.hint_text = "\u7ffb\u8a33\u30c1\u30a7\u30c3\u30af\u306e\u6307\u793a\u3092\u5165\u529b\u3057\u3066\u304f\u3060\u3055\u3044\u3002\u4f8b: \u539f\u6587(B2:B20)\u3001\u7ffb\u8a33(C2:C20)\u3001\u30ec\u30d3\u30e5\u30fc\u7d50\u679c\u3092D:G\u5217\u306b\u66f8\u304d\u8fbc\u3093\u3067\u304f\u3060\u3055\u3044\u3002"

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
        is_ready = new_state is AppState.READY
        is_task_in_progress = new_state is AppState.TASK_IN_PROGRESS
        is_stopping = new_state is AppState.STOPPING
        is_error = new_state is AppState.ERROR
        can_interact = is_ready or is_error

        if self.user_input:
            self.user_input.disabled = not can_interact
        if self.mode_selector:
            self.mode_selector.disabled = not can_interact
        if self.refresh_button:
            self.refresh_button.disabled = new_state is AppState.INITIALIZING

        if self.status_label:
            self.status_label.opacity = 1
            self.status_label.scale = 1
            if new_state is AppState.INITIALIZING:
                self.status_label.value = "\u521d\u671f\u5316\u4e2d..."
                self.status_label.color = ft.Colors.GREY_500
            elif is_ready:
                self.status_label.value = "\u5f85\u6a5f\u4e2d"
                self.status_label.color = ft.Colors.GREEN_300
            elif is_error:
                self.status_label.value = "\u30a8\u30e9\u30fc"
                self.status_label.color = ft.Colors.RED_300
            else:
                self.status_label.color = ft.Colors.GREY_500

        if self.action_button:
            if is_task_in_progress:
                if self.status_label:
                    self.status_label.value = "\u51e6\u7406\u3092\u5b9f\u884c\u4e2d..."
                    self.status_label.color = ft.Colors.DEEP_PURPLE_300
                    self.status_label.opacity = 0.5
                    self.status_label.scale = 0.95
                self.action_button.content = self._make_stop_button()
                self.action_button.disabled = False
            elif is_stopping:
                if self.status_label:
                    self.status_label.value = "\u51e6\u7406\u3092\u505c\u6b62\u3057\u3066\u3044\u307e\u3059..."
                    self.status_label.color = ft.Colors.DEEP_PURPLE_200
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
                        preferences = {k: v for k, v in loaded.items() if isinstance(k, str) and isinstance(v, str)}
                    else:
                        preferences = {}
                else:
                    preferences = {}
            except (OSError, json.JSONDecodeError) as err:
                print(f"Failed to read sheet preference: {err}")
                preferences = {}

            preferences[workbook_name] = sheet_name
            trimmed_items = list(preferences.items())[-50:]
            preferences = dict(trimmed_items)

            try:
                self.preference_file.parent.mkdir(parents=True, exist_ok=True)
                self.preference_file.write_text(json.dumps(preferences, ensure_ascii=False, indent=2), encoding="utf-8")
            except OSError as err:
                print(f"Failed to write sheet preference: {err}")

    def _handle_refresh_click(self, e: Optional[ft.ControlEvent]):
        self._refresh_excel_context()

    def _refresh_excel_context(self, is_initial_start: bool = False) -> Optional[str]:
        if not self.sheet_selector or not self.refresh_button:
            return None

        self.sheet_selector.disabled = True
        self.refresh_button.disabled = True
        self.refresh_button.text = "\u66f4\u65b0\u4e2d..."
        self._update_ui()

        try:
            with ExcelManager() as manager:
                info_dict = manager.get_active_workbook_and_sheet()
                sheet_names = manager.list_sheet_names()

                preferred_sheet = self._load_last_sheet_preference(info_dict["workbook_name"])
                if preferred_sheet and preferred_sheet in sheet_names and preferred_sheet != info_dict["sheet_name"]:
                    try:
                        activated_name = manager.activate_sheet(preferred_sheet)
                        info_dict["sheet_name"] = activated_name
                    except Exception as activate_err:
                        print(f"\u524d\u56de\u9078\u629e\u3057\u305f\u30b7\u30fc\u30c8 '{preferred_sheet}' \u306e\u5fa9\u5143\u306b\u5931\u6557\u3057\u307e\u3057\u305f: {activate_err}")
                        self._add_message(ResponseType.INFO, f"\u4fdd\u5b58\u6e08\u307f\u30b7\u30fc\u30c8\u300e{preferred_sheet}\u300f\u3092\u958b\u3051\u307e\u305b\u3093\u3067\u3057\u305f: {activate_err}")

                self.current_workbook_name = info_dict["workbook_name"]
                self.current_sheet_name = info_dict["sheet_name"]

                self.sheet_selection_updating = True
                info_text = f"\u5bfe\u8c61\u30d6\u30c3\u30af: {info_dict['workbook_name']}\n\u5bfe\u8c61\u30b7\u30fc\u30c8: {info_dict['sheet_name']}"
                if self.excel_info_label:
                    self.excel_info_label.value = info_text

                self.sheet_selector.options = [ft.dropdown.Option(name) for name in sheet_names]
                self.sheet_selector.value = info_dict["sheet_name"]
                self.sheet_selector.disabled = False

                if self.current_workbook_name and self.current_sheet_name:
                    self._save_last_sheet_preference(self.current_workbook_name, self.current_sheet_name)

                if not is_initial_start:
                    self.request_queue.put(RequestMessage(RequestType.UPDATE_CONTEXT, {"sheet_name": info_dict["sheet_name"]}))

                return info_dict["sheet_name"]
        except Exception as ex:
            error_message = f"Excel\u306e\u60c5\u5831\u53d6\u5f97\u306b\u5931\u6557\u3057\u307e\u3057\u305f: {ex}"
            if self.excel_info_label:
                self.excel_info_label.value = error_message
            self.sheet_selector.disabled = True
            self.sheet_selector.options = []
            self.sheet_selector.value = None
            if not is_initial_start:
                self._add_message(ResponseType.ERROR, error_message)
            return None
        finally:
            self.sheet_selection_updating = False
            self.refresh_button.disabled = False
            self.refresh_button.text = "\u66f4\u65b0"
            self._update_ui()

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

    def _on_sheet_change(self, e: ft.ControlEvent):
        if self.sheet_selection_updating:
            return
        selected_sheet = e.control.value if e and e.control else None
        if not selected_sheet:
            return

        previous_sheet = self.current_sheet_name
        try:
            with ExcelManager() as manager:
                manager.activate_sheet(selected_sheet)
        except Exception as ex:
            error_message = f"\u30b7\u30fc\u30c8\u306e\u5207\u308a\u66ff\u3048\u306b\u5931\u6557\u3057\u307e\u3057\u305f: {ex}"
            if self.excel_info_label:
                self.excel_info_label.value = error_message
            self.sheet_selection_updating = True
            if self.sheet_selector:
                self.sheet_selector.value = previous_sheet
            self.sheet_selection_updating = False
            self._add_message(ResponseType.ERROR, error_message)
            self._update_ui()
            return

        self.request_queue.put(RequestMessage(RequestType.UPDATE_CONTEXT, {"sheet_name": selected_sheet}))
        self.current_sheet_name = selected_sheet
        if self.current_workbook_name:
            self._save_last_sheet_preference(self.current_workbook_name, selected_sheet)
        workbook = self.current_workbook_name or "Unknown"
        if self.excel_info_label:
            self.excel_info_label.value = f"Workbook: {workbook}\nSheet: {selected_sheet}"
        self._add_message(ResponseType.INFO, f"\u64cd\u4f5c\u5bfe\u8c61\u306e\u30b7\u30fc\u30c8\u3092\u300e{selected_sheet}\u300f\u306b\u8a2d\u5b9a\u3057\u307e\u3057\u305f\u3002")
        self._update_ui()

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

        if response.type is ResponseType.INITIALIZATION_COMPLETE:
            self._set_state(AppState.READY)
            if self.status_label:
                self.status_label.value = response.content or self.status_label.value
        elif response.type is ResponseType.STATUS:
            if self.status_label:
                self.status_label.value = response.content or ""
        elif response.type is ResponseType.ERROR:
            self._set_state(AppState.ERROR)
            if response.content:
                self._add_message(response.type, response.content)
        elif response.type is ResponseType.END_OF_TASK:
            self._set_state(AppState.READY)
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

if __name__ == "__main__":
    ft.app(target=main)
