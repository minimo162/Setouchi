# desktop_app.py (Flet版 - 最終完成版)

import flet as ft
import threading
import queue
import inspect
import time
import sys
import traceback
import os
from typing import List, Callable, Dict, Optional, Any
from enum import Enum, auto

# --- 既存のコンポーネントをインポート ---
from excel_copilot.core.excel_manager import ExcelManager, ExcelConnectionError
from excel_copilot.agent.react_agent import ReActAgent
from excel_copilot.core.browser_copilot_manager import BrowserCopilotManager
from excel_copilot.tools import excel_tools
from excel_copilot.tools.schema_builder import create_tool_schema
from excel_copilot.config import COPILOT_USER_DATA_DIR

# --- 定数・型定義 ---
class AppState(Enum):
    INITIALIZING = auto()
    READY = auto()
    TASK_IN_PROGRESS = auto()
    STOPPING = auto()
    ERROR = auto()

class Req:
    USER_INPUT = "USER_INPUT"
    STOP = "STOP"
    QUIT = "QUIT"
    UPDATE_CONTEXT = "UPDATE_CONTEXT"

class Res:
    STATUS = "status"
    ERROR = "error"
    AGENT_MESSAGE = "agent_message"
    INFO = "info"
    END_OF_TASK = "end_OF_TASK"
    INITIALIZATION_COMPLETE = "initialization_complete"

# --- バックグラウンド処理 (変更なし) ---
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
            self.response_queue.put({"type": Res.ERROR, "content": f"致命的な実行時エラー: {e}"})
        finally:
            self._cleanup()

    def _initialize(self):
        try:
            print("Worker初期化開始...")
            self.response_queue.put({"type": Res.STATUS, "content": "ブラウザ(Playwright)を起動中..."})
            self.browser_manager = BrowserCopilotManager(user_data_dir=COPILOT_USER_DATA_DIR, headless=False)
            self.browser_manager.start()
            print("BrowserManagerの起動完了。")
            
            self.response_queue.put({"type": Res.STATUS, "content": "AIエージェントを準備中..."})
            tool_functions = [obj for _, obj in inspect.getmembers(excel_tools) if inspect.isfunction(obj)]
            tool_schemas = [create_tool_schema(func) for func in tool_functions]
            self.agent = ReActAgent(tools=tool_functions, tool_schemas=tool_schemas, browser_manager=self.browser_manager, sheet_name=self.sheet_name)
            print("AIエージェントの準備完了。")
            
            self.response_queue.put({"type": Res.INITIALIZATION_COMPLETE, "content": "準備完了。指示をどうぞ。"})
            print("Worker初期化完了。")
        except Exception as e:
            print(f"Worker初期化中にエラーが発生しました: {e}")
            traceback.print_exc()
            self.response_queue.put({"type": Res.ERROR, "content": f"初期化エラー: {e}"})

    def _main_loop(self):
        print("メインループを開始します。")
        while True:
            request = self.request_queue.get()
            req_type = request.get("type")
            payload = request.get("payload")

            if req_type == Req.QUIT:
                print("終了リクエスト受信。メインループを終了します。")
                break
            elif req_type == Req.STOP:
                self.stop_event.set()
                continue
            elif req_type == Req.UPDATE_CONTEXT:
                self._update_context(payload)
            elif req_type == Req.USER_INPUT:
                self._execute_task(payload)

    def _update_context(self, payload: Dict[str, Any]):
        new_sheet_name = payload.get("sheet_name")
        self.sheet_name = new_sheet_name
        if self.agent:
            self.agent.sheet_name = new_sheet_name
        self.response_queue.put({"type": Res.INFO, "content": f"操作対象のシートを「{new_sheet_name or '未選択'}」に変更しました。"})

    def _execute_task(self, user_input: str):
        self.stop_event.clear()
        if not self.agent:
            self.response_queue.put({"type": Res.ERROR, "content": "エージェントが初期化されていません。"})
            return

        try:
            for message_dict in self.agent.run(user_input, self.stop_event):
                self.response_queue.put(message_dict)
        except ExcelConnectionError as e:
            self.response_queue.put({"type": Res.ERROR, "content": str(e)})
        except Exception as e:
            self.response_queue.put({"type": Res.ERROR, "content": f"タスク実行エラー: {e}"})
        finally:
            if self.stop_event.is_set():
                self.response_queue.put({"type": Res.INFO, "content": "ユーザーの指示により処理を中断しました。"})
            self.response_queue.put({"type": Res.END_OF_TASK})

    def _cleanup(self):
        print("クリーンアップ処理を開始します。")
        if self.browser_manager:
            self.browser_manager.close()
        print("ワーカーをクリーンアップしました。")

# --- Flet UI ---
class ChatMessage(ft.ResponsiveRow):
    def __init__(self, msg_type: str, msg_content: str):
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
                "icon_color": "#81C784", # Muted Green
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
                "icon_color": "#E57373", # Muted Red
                "bgcolor": "#5E2A2A",
                "title": "Error",
            },
        }
        config = type_map.get(msg_type, type_map["info"])

        if msg_type in ["info", "status"]:
            self.controls = [
                ft.Column(
                    [ft.Text(msg_content, **config.get("text_style", {}))],
                    col=12,
                    alignment=ft.MainAxisAlignment.CENTER,
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER
                )
            ]
            return

        content_controls = []
        if config.get("title"):
            content_controls.append(ft.Text(config["title"], weight=ft.FontWeight.BOLD, size=12, color=config["icon_color"]))
        
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
                color="#1A000000", # 10% black opacity
                offset=ft.Offset(2, 2),
            )
        )
        
        icon = ft.Icon(name=config["icon"], color=config["icon_color"], size=20)
        icon_container = ft.Container(icon, margin=ft.margin.only(right=8, left=8, top=3))

        bubble_and_icon_row = ft.Row(
            [icon_container, message_bubble] if msg_type != "user" else [message_bubble, icon_container],
            vertical_alignment=ft.CrossAxisAlignment.START,
        )

        if msg_type == "user":
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

def main(page: ft.Page):
    
    request_queue = queue.Queue()
    response_queue = queue.Queue()
    app_state: Optional[AppState] = None
    worker_thread: Optional[threading.Thread] = None
    ui_loop_running = True
    shutdown_requested = False
    window_closed_event = threading.Event()

    page.title = "Excel Co-pilot"
    page.window.width = 1280
    page.window.height = 768
    page.window.min_width = 960
    page.window.min_height = 600
    page.theme_mode = ft.ThemeMode.DARK
    page.bgcolor = "#141218"
    page.window.center()
    page.window.prevent_close = True

    def _set_state(new_state: AppState):
        nonlocal app_state
        if app_state == new_state: return
        app_state = new_state
        
        is_ready = new_state == AppState.READY
        is_task_in_progress = new_state == AppState.TASK_IN_PROGRESS
        is_stopping = new_state == AppState.STOPPING
        can_interact = is_ready or new_state == AppState.ERROR

        user_input.disabled = not can_interact
        refresh_button.disabled = new_state == AppState.INITIALIZING

        if new_state == AppState.INITIALIZING: status_label.value = "初期化中..."
        elif new_state == AppState.READY: status_label.value = "準備完了"
        elif new_state == AppState.ERROR: status_label.value = "エラー発生"
        
        status_label.color = ft.Colors.GREEN_300 if is_ready else ft.Colors.RED_300 if new_state == AppState.ERROR else ft.Colors.GREY_500
        status_label.opacity = 1
        status_label.scale = 1
        
        if is_task_in_progress:
            status_label.value = "処理を実行中..."
            status_label.color = ft.Colors.DEEP_PURPLE_300
            status_label.opacity = 0.5
            status_label.scale = 0.95
            action_button.content = ft.IconButton(icon=ft.Icons.STOP_ROUNDED, icon_color="#B39DDB", on_click=_stop_task, tooltip="処理を停止")
        elif is_stopping:
            status_label.value = "停止処理中..."
            action_button.content = ft.ProgressRing(width=18, height=18, stroke_width=2)
        else:
            action_button.content = ft.IconButton(icon=ft.Icons.SEND_ROUNDED, icon_color="#B39DDB", on_click=_run_copilot, tooltip="送信")
        
        action_button.disabled = not can_interact and not is_task_in_progress

        _update_ui()
        
    def _update_ui():
        try:
            page.update()
        except Exception as e:
            print(f"UIの更新に失敗しました: {e}")

    def _add_message(msg_type: str, msg_content: str):
        msg = ChatMessage(msg_type, msg_content)
        chat_list.controls.append(msg)
        _update_ui()
        time.sleep(0.01)
        msg.appear()
        
    def _refresh_excel_context(e=None, is_initial_start: bool = False):
        refresh_button.disabled = True
        refresh_button.text = "更新中..."
        _update_ui()
        try:
            with ExcelManager() as manager:
                info_dict = manager.get_active_workbook_and_sheet()
                info_text = f"ブック: {info_dict['workbook_name']}\nシート: {info_dict['sheet_name']}"
                excel_info_label.value = info_text
                if not is_initial_start:
                    request_queue.put({"type": Req.UPDATE_CONTEXT, "payload": {"sheet_name": info_dict["sheet_name"]}})
                return info_dict['sheet_name']
        except Exception as ex:
            error_message = f"Excel情報取得失敗: {ex}"
            excel_info_label.value = error_message
            if not is_initial_start:
                _add_message("error", error_message)
            return None
        finally:
            refresh_button.disabled = False
            refresh_button.text = "更新"
            _update_ui()

    def _run_copilot(e):
        user_text = user_input.value
        if not user_text or app_state not in [AppState.READY, AppState.ERROR]:
            return

        _set_state(AppState.TASK_IN_PROGRESS)
        _add_message("user", user_text)
        user_input.value = ""
        request_queue.put({"type": Req.USER_INPUT, "payload": user_text})
        _update_ui()

    def _stop_task(e):
        _set_state(AppState.STOPPING)
        request_queue.put({"type": Req.STOP})

    def _process_response_queue_loop():
        while ui_loop_running:
            try:
                msg = response_queue.get(timeout=0.1)
                msg_type, msg_content = msg.get("type"), msg.get("content", "")

                if msg_type == Res.INITIALIZATION_COMPLETE:
                    _set_state(AppState.READY)
                    status_label.value = msg_content
                elif msg_type == Res.STATUS:
                    status_label.value = msg_content
                elif msg_type == Res.ERROR:
                    _set_state(AppState.ERROR)
                    _add_message("error", msg_content)
                elif msg_type == Res.END_OF_TASK:
                    _set_state(AppState.READY)
                else:
                    _add_message(msg_type, msg_content)
                
                _update_ui()

            except queue.Empty:
                continue
            except Exception as e:
                print(f"レスポンスキューの処理中にエラー: {e}")
                traceback.print_exc()
            
    def _force_exit(reason: str = ''):
        nonlocal ui_loop_running, shutdown_requested
        print(f'Force exit triggered. reason={reason}')
        if shutdown_requested:
            print('Force exit: shutdown already in progress.')
        else:
            shutdown_requested = True
            if not ui_loop_running:
                print('Force exit: UI loop already stopped.')
            else:
                ui_loop_running = False
                try:
                    request_queue.put_nowait({"type": Req.QUIT})
                    print('Force exit: QUIT request posted.')
                except Exception as queue_err:
                    print(f'Force exit: failed to enqueue QUIT: {queue_err}')
                if worker_thread:
                    try:
                        worker_thread.join(timeout=3.0)
                        print('Force exit: worker thread joined or timeout.')
                    except Exception as join_err:
                        print(f'Force exit: worker join error: {join_err}')
            try:
                page.window.prevent_close = False
            except Exception as prevent_err:
                print(f'Force exit: unable to clear prevent_close: {prevent_err}')
            close_requested = False
            try:
                page.window.close()
                close_requested = True
                print('Force exit: window.close() called.')
            except AttributeError:
                try:
                    page.window.destroy()
                    close_requested = True
                    window_closed_event.set()
                    print('Force exit: window.destroy() called.')
                except Exception as destroy_err:
                    print(f'Force exit: window destroy failed: {destroy_err}')
            except Exception as close_err:
                print(f'Force exit: window close failed: {close_err}')
            if close_requested:
                try:
                    page.update()
                except Exception as update_err:
                    print(f'Force exit: page update after close failed: {update_err}')
        if not window_closed_event.is_set():
            try:
                if window_closed_event.wait(timeout=3.0):
                    print('Force exit: window close confirmed.')
                else:
                    print('Force exit: window close wait timed out.')
            except Exception as wait_err:
                print(f'Force exit: waiting for window close failed: {wait_err}')
        os._exit(0)

    def _on_window_event(e):
        event_name = getattr(e, 'event', None)
        data = getattr(e, 'data', None)
        payload_raw = event_name or data or ''
        payload = payload_raw.lower()
        normalized_payload = payload.replace('_', '-')
        print(f'Window event received: event={event_name}, data={data}')
        window_gone_events = {'closed', 'close-completed', 'destroyed'}
        close_request_events = {'close', 'closing', 'close-requested'}
        if normalized_payload in window_gone_events:
            window_closed_event.set()
        if normalized_payload in close_request_events or (normalized_payload in window_gone_events and not shutdown_requested):
            _force_exit(reason=f'window-event:{normalized_payload}')

    def _on_page_disconnect(e):
        print('Page disconnect detected.')
        window_closed_event.set()
        _force_exit(reason='page-disconnect')

    print('Debug: registering window event handlers')
    page.window.prevent_close = True
    page.window.on_event = _on_window_event
    page.on_disconnect = _on_page_disconnect

    # --- UIコンポーネント定義 ---
    title_label = ft.Text("Excel\nCo-pilot", size=26, weight=ft.FontWeight.BOLD, color="#FFFFFF")
    status_label = ft.Text("初期化中...", size=15, color=ft.Colors.GREY_500, animate_opacity=300, animate_scale=600)
    
    def on_hover_button(e):
        e.control.scale = 1.05 if e.data == "true" else 1
        e.control.update()

    excel_info_label = ft.Text("", size=14, color="#CFD8DC")
    refresh_button = ft.ElevatedButton(
        text="更新", on_click=_refresh_excel_context, bgcolor=ft.Colors.DEEP_PURPLE_500, color=ft.Colors.WHITE,
        scale=1, animate_scale=100, on_hover=on_hover_button
    )

    sidebar_content = ft.Column(
        [
            title_label,
            status_label,
            ft.Divider(color="#4A4458"),
            excel_info_label,
            refresh_button,
        ],
        width=200,
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
            color="#1A000000", # 10% black opacity
            offset=ft.Offset(5, 5),
        )
    )

    chat_list = ft.ListView(expand=True, spacing=15, auto_scroll=True, padding=20)
    user_input = ft.TextField(
        hint_text="A1セルに「こんにちは」と入力して...",
        expand=True,
        on_submit=_run_copilot,
        border_color="#4A4458",
        focused_border_color="#B39DDB",
        border_radius=10,
    )
    action_button = ft.Container(
        content=ft.IconButton(icon=ft.Icons.SEND_ROUNDED, on_click=_run_copilot, icon_color="#B39DDB", tooltip="送信"),
        scale=1, animate_scale=100, on_hover=on_hover_button,
        bgcolor="#2C2A3A",
        border_radius=30,
    )

    main_content = ft.Column(
        [
            chat_list,
            ft.Row([user_input, action_button], alignment=ft.MainAxisAlignment.CENTER)
        ],
        expand=True,
    )

    page.add(
        ft.Row(
            [
                sidebar,
                main_content,
            ],
            expand=True,
            spacing=20,
            vertical_alignment=ft.CrossAxisAlignment.STRETCH
        )
    )

    # --- アプリケーション開始 ---
    _set_state(AppState.INITIALIZING)
    _update_ui()
    sheet_name = _refresh_excel_context(is_initial_start=True)
    
    worker = CopilotWorker(request_queue, response_queue, sheet_name)
    worker_thread = threading.Thread(target=worker.run, daemon=True)
    worker_thread.start()
    
    queue_processor = threading.Thread(target=_process_response_queue_loop, daemon=True)
    queue_processor.start()


if __name__ == "__main__":
    ft.app(target=main)
