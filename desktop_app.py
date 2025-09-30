# desktop_app.py

import flet as ft
import threading
import queue
import inspect
import time
import traceback
import os
from dataclasses import dataclass, field
from typing import Dict, Optional, Any, Union
from enum import Enum, auto

from excel_copilot.core.excel_manager import ExcelManager, ExcelConnectionError
from excel_copilot.agent.react_agent import ReActAgent
from excel_copilot.core.browser_copilot_manager import BrowserCopilotManager
from excel_copilot.tools import excel_tools
from excel_copilot.tools.schema_builder import create_tool_schema
from excel_copilot.config import COPILOT_USER_DATA_DIR


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

    def run(self):
        try:
            self._initialize()
            if self.agent and self.browser_manager:
                self._main_loop()
        except Exception as e:
            print(f"Critical error in worker run method: {e}")
            traceback.print_exc()
            self._emit_response(ResponseMessage(ResponseType.ERROR, f"致命的な実行時エラー: {e}"))
        finally:
            self._cleanup()

    def _emit_response(self, message: Union[ResponseMessage, Dict[str, Any]]):
        try:
            self.response_queue.put(ResponseMessage.from_raw(message))
        except Exception as err:
            print(f"Failed to enqueue response: {err}")

    def _initialize(self):
        try:
            print("Worker初期化開始..")
            self._emit_response(ResponseMessage(ResponseType.STATUS, "ブラウザ(Playwright)を起動中..."))
            self.browser_manager = BrowserCopilotManager(user_data_dir=COPILOT_USER_DATA_DIR, headless=False)
            self.browser_manager.start()
            print("BrowserManagerの起動完了")

            self._emit_response(ResponseMessage(ResponseType.STATUS, "AIエージェントを準備中..."))
            tool_functions = [obj for _, obj in inspect.getmembers(excel_tools) if inspect.isfunction(obj)]
            tool_schemas = [create_tool_schema(func) for func in tool_functions]
            self.agent = ReActAgent(tools=tool_functions, tool_schemas=tool_schemas, browser_manager=self.browser_manager, sheet_name=self.sheet_name)
            print("AIエージェントの準備完了")

            self._emit_response(ResponseMessage(ResponseType.INITIALIZATION_COMPLETE, "準備完了。指示をどうぞ。"))
            print("Worker初期化完了")
        except Exception as e:
            print(f"Worker初期化中にエラーが発生しました: {e}")
            traceback.print_exc()
            self._emit_response(ResponseMessage(ResponseType.ERROR, f"初期化エラー: {e}"))

    def _main_loop(self):
        print("メインループを開始します")
        while True:
            raw_request = self.request_queue.get()
            try:
                request = RequestMessage.from_raw(raw_request)
            except ValueError as exc:
                self._emit_response(ResponseMessage(ResponseType.ERROR, f"無効なリクエストを受信しました: {exc}"))
                continue

            if request.type is RequestType.QUIT:
                print("終了リクエスト受信。メインループを終了します")
                break
            if request.type is RequestType.STOP:
                self.stop_event.set()
                continue
            if request.type is RequestType.UPDATE_CONTEXT:
                self._update_context(request.payload)
                continue
            if request.type is RequestType.USER_INPUT:
                if isinstance(request.payload, str):
                    self._execute_task(request.payload)
                else:
                    self._emit_response(ResponseMessage(ResponseType.ERROR, "ユーザー入力が不正です。"))

    def _update_context(self, payload: Optional[Dict[str, Any]]):
        new_sheet_name = payload.get("sheet_name") if isinstance(payload, dict) else None
        self.sheet_name = new_sheet_name
        if self.agent:
            self.agent.sheet_name = new_sheet_name
        sheet_label = new_sheet_name or "未選択"
        self._emit_response(ResponseMessage(ResponseType.INFO, f"操作対象のシートを「{sheet_label}」に変更しました。"))

    def _execute_task(self, user_input: str):
        self.stop_event.clear()
        if not self.agent:
            self._emit_response(ResponseMessage(ResponseType.ERROR, "エージェントが初期化されていません。"))
            return

        try:
            for message_dict in self.agent.run(user_input, self.stop_event):
                self._emit_response(message_dict)
        except ExcelConnectionError as e:
            self._emit_response(ResponseMessage(ResponseType.ERROR, str(e)))
        except Exception as e:
            self._emit_response(ResponseMessage(ResponseType.ERROR, f"タスク実行エラー: {e}"))
        finally:
            if self.stop_event.is_set():
                self._emit_response(ResponseMessage(ResponseType.INFO, "ユーザーの操作により処理を中断しました。"))
            self._emit_response(ResponseMessage(ResponseType.END_OF_TASK))

    def _cleanup(self):
        print("クリーンアップ処理を開始します")
        if self.browser_manager:
            self.browser_manager.close()
        print("ワーカーをクリーンアップしました")


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

        content_controls.append(ft.Text(msg_content, **config.get("text_style", {}), selectable=True))

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

        self.title_label: Optional[ft.Text] = None
        self.status_label: Optional[ft.Text] = None
        self.excel_info_label: Optional[ft.Text] = None
        self.refresh_button: Optional[ft.ElevatedButton] = None
        self.sheet_selector: Optional[ft.Dropdown] = None
        self.chat_list: Optional[ft.ListView] = None
        self.user_input: Optional[ft.TextField] = None
        self.action_button: Optional[ft.Container] = None

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
        self.status_label = ft.Text("初期化中...", size=15, color=ft.Colors.GREY_500, animate_opacity=300, animate_scale=600)

        self.excel_info_label = ft.Text("", size=14, color="#CFD8DC")
        self.refresh_button = ft.ElevatedButton(
            text="更新",
            on_click=self._handle_refresh_click,
            bgcolor=ft.Colors.DEEP_PURPLE_500,
            color=ft.Colors.WHITE,
            scale=1,
            animate_scale=100,
            on_hover=self._handle_button_hover,
        )

        self.sheet_selector = ft.Dropdown(
            options=[],
            width=180,
            on_change=self._on_sheet_change,
            hint_text="シートを選択",
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
            hint_text="A1セルに「こんにちは」と入力して...",
            expand=True,
            on_submit=self._run_copilot,
            border_color="#4A4458",
            focused_border_color="#B39DDB",
            border_radius=10,
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
        return ft.IconButton(icon=ft.Icons.SEND_ROUNDED, on_click=self._run_copilot, icon_color="#B39DDB", tooltip="送信")

    def _make_stop_button(self) -> ft.IconButton:
        return ft.IconButton(icon=ft.Icons.STOP_ROUNDED, on_click=self._stop_task, icon_color="#B39DDB", tooltip="処理を停止")

    def _handle_button_hover(self, e: ft.ControlEvent):
        if e.data == "true":
            e.control.scale = 1.05
        else:
            e.control.scale = 1
        e.control.update()

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
        if self.refresh_button:
            self.refresh_button.disabled = new_state is AppState.INITIALIZING

        if self.status_label:
            self.status_label.opacity = 1
            self.status_label.scale = 1
            if new_state is AppState.INITIALIZING:
                self.status_label.value = "初期化中..."
                self.status_label.color = ft.Colors.GREY_500
            elif is_ready:
                self.status_label.value = "準備完了"
                self.status_label.color = ft.Colors.GREEN_300
            elif is_error:
                self.status_label.value = "エラー発生"
                self.status_label.color = ft.Colors.RED_300
            else:
                self.status_label.color = ft.Colors.GREY_500

        if self.action_button:
            if is_task_in_progress:
                if self.status_label:
                    self.status_label.value = "処理を実行中..."
                    self.status_label.color = ft.Colors.DEEP_PURPLE_300
                    self.status_label.opacity = 0.5
                    self.status_label.scale = 0.95
                self.action_button.content = self._make_stop_button()
                self.action_button.disabled = False
            elif is_stopping:
                if self.status_label:
                    self.status_label.value = "停止処理中..."
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
            print(f"UIの更新に失敗しました: {e}")

    def _add_message(self, msg_type: Union[ResponseType, str], msg_content: str):
        if not self.chat_list:
            return
        msg = ChatMessage(msg_type, msg_content)
        self.chat_list.controls.append(msg)
        self._update_ui()
        time.sleep(0.01)
        msg.appear()

    def _handle_refresh_click(self, e: Optional[ft.ControlEvent]):
        self._refresh_excel_context()

    def _refresh_excel_context(self, is_initial_start: bool = False) -> Optional[str]:
        if not self.sheet_selector or not self.refresh_button:
            return None

        self.sheet_selector.disabled = True
        self.refresh_button.disabled = True
        self.refresh_button.text = "更新中..."
        self._update_ui()

        try:
            with ExcelManager() as manager:
                info_dict = manager.get_active_workbook_and_sheet()
                sheet_names = manager.list_sheet_names()
                self.current_workbook_name = info_dict["workbook_name"]
                self.current_sheet_name = info_dict["sheet_name"]

                self.sheet_selection_updating = True
                info_text = f"ブック: {info_dict['workbook_name']}\nシート: {info_dict['sheet_name']}"
                if self.excel_info_label:
                    self.excel_info_label.value = info_text

                self.sheet_selector.options = [ft.dropdown.Option(name) for name in sheet_names]
                self.sheet_selector.value = info_dict["sheet_name"]
                self.sheet_selector.disabled = False

                if not is_initial_start:
                    self.request_queue.put(RequestMessage(RequestType.UPDATE_CONTEXT, {"sheet_name": info_dict["sheet_name"]}))

                return info_dict["sheet_name"]
        except Exception as ex:
            error_message = f"Excel情報の取得に失敗しました: {ex}"
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
            self.refresh_button.text = "更新"
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
        self._set_state(AppState.STOPPING)
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
            error_message = f"シート切り替えに失敗しました: {ex}"
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
        if self.excel_info_label:
            workbook = self.current_workbook_name or "不明"
            self.excel_info_label.value = f"ブック: {workbook}\nシート: {selected_sheet}"
        self._add_message(ResponseType.INFO, f"操作対象のシートを「{selected_sheet}」に設定しました。")
        self.current_sheet_name = selected_sheet
        self._update_ui()

    def _process_response_queue_loop(self):
        while self.ui_loop_running:
            try:
                raw_message = self.response_queue.get(timeout=0.1)
            except queue.Empty:
                continue
            except Exception as e:
                print(f"レスポンスキューの待機中にエラーが発生しました: {e}")
                continue

            try:
                response = ResponseMessage.from_raw(raw_message)
            except ValueError as exc:
                print(f"レスポンスの解析に失敗しました: {exc}")
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

