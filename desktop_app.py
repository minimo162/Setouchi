# desktop_app.py

import argparse
import json
import logging
import os
import queue
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional, Union

import flet as ft

from excel_copilot.agent.prompts import CopilotMode
from excel_copilot.config import (
    COPILOT_USER_DATA_DIR,
    COPILOT_HEADLESS,
    COPILOT_BROWSER_CHANNELS,
    COPILOT_PAGE_GOTO_TIMEOUT_MS,
    COPILOT_SLOW_MO_MS,
)
from excel_copilot.core.excel_manager import ExcelManager, ExcelConnectionError
from excel_copilot.ui.chat import ChatMessage
from excel_copilot.ui.messages import (
    AppState,
    RequestMessage,
    RequestType,
    ResponseMessage,
    ResponseType,
)
from excel_copilot.ui.theme import EXPRESSIVE_PALETTE
from excel_copilot.ui.worker import CopilotWorker

if not logging.getLogger().handlers:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )

FOCUS_WAIT_TIMEOUT_SECONDS = 15.0
PREFERENCE_LAST_WORKBOOK_KEY = "__last_workbook__"
# Container in this Flet build lacks min-height/constraints helpers, so keep a fixed base height.
CHAT_PANEL_BASE_HEIGHT = 600

DEFAULT_AUTOTEST_TIMEOUT_SECONDS = 180.0  # 3-minute auto-test timeout
DEFAULT_AUTOTEST_SOURCE_REFERENCE_URL = (
    "https://ralleti-my.sharepoint.com/:b:/g/personal/yuukikod_ralleti_onmicrosoft_com/"
    "Eao1vzScMVdMpyecR1s8KFEBzr0LMPNp5u5ksA3U3TCMwQ?e=jwmAly"
)
DEFAULT_AUTOTEST_TARGET_REFERENCE_URL = (
    "https://ralleti-my.sharepoint.com/:b:/g/personal/yuukikod_ralleti_onmicrosoft_com/"
    "EdJ586XuxedLsaSArCCve9kB1K79F0BvGqxzuZBhfWWS-w?e=wjR4C2"
)
DEFAULT_AUTOTEST_PROMPT_TEMPLATE = (
    "A1セルを日本語参照: {source_url} 英語参照: {target_url} を使って英訳してください。"
    "B列から連続する列に結果を並べ、B列=翻訳、C列=process_notes_jp、D列以降=参照文ペアとしてください。"
    "参照文は直前に抽出したキーフレーズに関連する文を引用し、citation_output_range は指定しないでください。"
)


def _is_truthy_env(value: Optional[str]) -> bool:
    if value is None:
        return False
    return value.strip().lower() in {"1", "true", "yes", "on"}

def _is_autotest_mode_enabled() -> bool:
    """Check whether the application should run its auto-test scenario."""
    if _is_truthy_env(os.getenv("COPILOT_AUTOTEST_ENABLED")):
        return True
    return bool(os.getenv("COPILOT_AUTOTEST_PROMPT"))

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

        self.status_label: Optional[ft.Text] = None
        self.workbook_selector: Optional[ft.Dropdown] = None
        self.sheet_selector: Optional[ft.Dropdown] = None
        self.chat_list: Optional[ft.ListView] = None
        self.user_input: Optional[ft.TextField] = None
        self.action_button: Optional[ft.Container] = None
        self.save_log_button: Optional[ft.TextButton] = None
        self.workbook_refresh_button: Optional[ft.TextButton] = None
        self.new_chat_button: Optional[ft.TextButton] = None
        self.mode_card_row: Optional[ft.ResponsiveRow] = None
        self._mode_card_map: dict[str, ft.Container] = {}
        self._context_panel: Optional[ft.Container] = None
        self._context_actions: Optional[ft.ResponsiveRow] = None
        self._chat_panel: Optional[ft.Container] = None
        self._composer_panel: Optional[ft.Container] = None
        self._mode_panel_container: Optional[ft.Container] = None
        self._composer_row: Optional[ft.ResponsiveRow] = None
        self._user_input_wrapper: Optional[ft.Container] = None
        self._action_button_wrapper: Optional[ft.Container] = None
        self._content_container: Optional[ft.Container] = None
        self._layout: Optional[ft.ResponsiveRow] = None
        self._main_column: Optional[ft.Column] = None

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

        auto_test_prompt_override = os.getenv("COPILOT_AUTOTEST_PROMPT")
        auto_test_enabled_flag = os.getenv("COPILOT_AUTOTEST_ENABLED")
        self._auto_test_source_url: str = os.getenv(
            "COPILOT_AUTOTEST_SOURCE_URL", DEFAULT_AUTOTEST_SOURCE_REFERENCE_URL
        )
        self._auto_test_target_url: str = os.getenv(
            "COPILOT_AUTOTEST_TARGET_URL", DEFAULT_AUTOTEST_TARGET_REFERENCE_URL
        )
        if auto_test_prompt_override:
            self._auto_test_prompt = auto_test_prompt_override
        elif _is_truthy_env(auto_test_enabled_flag):
            self._auto_test_prompt = DEFAULT_AUTOTEST_PROMPT_TEMPLATE.format(
                source_url=self._auto_test_source_url,
                target_url=self._auto_test_target_url,
            )
        else:
            self._auto_test_prompt = None
        self._auto_test_workbook: Optional[str] = os.getenv("COPILOT_AUTOTEST_WORKBOOK")
        self._auto_test_sheet: Optional[str] = os.getenv("COPILOT_AUTOTEST_SHEET")
        try:
            self._auto_test_delay: float = max(
                0.0, float(os.getenv("COPILOT_AUTOTEST_DELAY", "1.0"))
            )
        except ValueError:
            self._auto_test_delay = 1.0
        try:
            self._auto_test_timeout: float = max(
                0.0,
                float(
                    os.getenv(
                        "COPILOT_AUTOTEST_TIMEOUT",
                        str(int(DEFAULT_AUTOTEST_TIMEOUT_SECONDS)),
                    )
                ),
            )
        except ValueError:
            self._auto_test_timeout = DEFAULT_AUTOTEST_TIMEOUT_SECONDS
        self._auto_test_enabled = bool(self._auto_test_prompt)
        self._auto_test_triggered = False
        self._auto_test_completed = False
        self._auto_test_deadline: Optional[float] = None
        print(
            f"AUTOTEST: enabled={self._auto_test_enabled}, "
            f"workbook={self._auto_test_workbook or '(unchanged)'}, "
            f"sheet={self._auto_test_sheet or '(unchanged)'}",
            flush=True,
        )
        print(
            "AUTOTEST: references",
            {
                "source_url": self._auto_test_source_url,
                "target_url": self._auto_test_target_url,
                "timeout": self._auto_test_timeout,
            },
            flush=True,
        )

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
        self.page.window.min_width = 480
        self.page.window.min_height = 520
        palette = EXPRESSIVE_PALETTE
        self.page.theme = ft.Theme(
            color_scheme_seed=palette["primary"],
            use_material3=True,
            font_family="Yu Gothic UI",
        )
        self.page.theme_mode = ft.ThemeMode.LIGHT
        self.page.bgcolor = palette["background"]
        self.page.window.bgcolor = palette["background"]
        self.page.padding = ft.Padding(0, 0, 0, 0)
        self.page.scroll = ft.ScrollMode.AUTO
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
        palette = EXPRESSIVE_PALETTE

        self.status_label = ft.Text(
            "\u521d\u671f\u5316\u4e2d...",
            size=12,
            color=palette["on_surface_variant"],
            animate_opacity=300,
            animate_scale=600,
        )
        status_chip = ft.Container(
            content=ft.Row(
                [
                    ft.Icon(ft.Icons.CHECK_CIRCLE_OUTLINE, size=18, color=palette["primary"]),
                    self.status_label,
                ],
                spacing=8,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
            ),
            bgcolor=palette["surface_variant"],
            border_radius=18,
            padding=ft.Padding(16, 12, 16, 12),
            border=ft.border.all(1, ft.Colors.with_opacity(0.08, palette["outline"])),
        )

        button_shape = ft.RoundedRectangleBorder(radius=18)
        button_overlay = ft.Colors.with_opacity(0.12, palette["primary"])

        self.new_chat_button = ft.FilledTonalButton(
            text="\u65b0\u3057\u3044\u30c1\u30e3\u30c3\u30c8",
            icon=ft.Icons.CHAT_OUTLINED,
            on_click=self._handle_new_chat_click,
            disabled=True,
            style=ft.ButtonStyle(
                shape=button_shape,
                padding=ft.Padding(18, 12, 18, 12),
                overlay_color=button_overlay,
            ),
        )

        dropdown_style = {
            "border_radius": 18,
            "border_color": palette["outline_variant"],
            "focused_border_color": palette["primary"],
            "fill_color": palette["surface_variant"],
            # Slightly smaller text keeps long workbook/sheet names visible without clipping.
            "text_style": ft.TextStyle(color=palette["on_surface"], size=12),
            "hint_style": ft.TextStyle(color=palette["on_surface_variant"], size=12),
            "disabled": True,
            "filled": True,
            "suffix_icon": ft.Icon(ft.Icons.KEYBOARD_ARROW_DOWN_ROUNDED, color=palette["on_surface_variant"]),
        }

        self.workbook_selector = ft.Dropdown(
            options=[],
            on_change=self._on_workbook_change,
            on_focus=self._on_workbook_dropdown_focus,
            hint_text="\u30d6\u30c3\u30af\u3092\u9078\u629e",
            expand=True,
            **dropdown_style,
        )

        self.workbook_selector_wrapper = ft.GestureDetector(
            content=self.workbook_selector,
            on_tap_down=self._on_workbook_dropdown_tap,
            expand=True,
        )

        self.sheet_selector = ft.Dropdown(
            options=[],
            on_change=self._on_sheet_change,
            on_focus=self._on_sheet_dropdown_focus,
            hint_text="\u30b7\u30fc\u30c8\u3092\u9078\u629e",
            expand=True,
            **dropdown_style,
        )

        self.sheet_selector_wrapper = ft.GestureDetector(
            content=self.sheet_selector,
            on_tap_down=self._on_sheet_dropdown_tap,
            expand=True,
        )

        self.workbook_refresh_button = ft.FilledTonalButton(
            text=self._workbook_refresh_button_default_text,
            icon=ft.Icons.SYNC,
            on_click=self._handle_workbook_refresh_click,
            disabled=True,
            style=ft.ButtonStyle(
                shape=button_shape,
                padding=ft.Padding(18, 12, 18, 12),
                overlay_color=button_overlay,
            ),
        )

        self._context_actions = ft.ResponsiveRow(
            controls=[
                ft.Container(content=self.workbook_refresh_button, col={"xs": 12, "sm": 6}),
                ft.Container(content=self.new_chat_button, col={"xs": 12, "sm": 6}),
            ],
            spacing=12,
            run_spacing=12,
            alignment=ft.MainAxisAlignment.END,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )

        self._context_panel = ft.Container(
            bgcolor=palette["surface"],
            border_radius=24,
            padding=ft.Padding(24, 32, 24, 32),
            border=ft.border.all(1, ft.Colors.with_opacity(0.08, palette["outline"])),
            shadow=ft.BoxShadow(
                spread_radius=0,
                blur_radius=14,
                color=ft.Colors.with_opacity(0.06, "#0F172A"),
                offset=ft.Offset(0, 8),
            ),
            content=ft.Column(
                [
                    status_chip,
                    ft.Column(
                        [
                            ft.Text("\u30d6\u30c3\u30af", size=13, color=palette["on_surface_variant"]),
                            self.workbook_selector_wrapper,
                            ft.Text("\u30b7\u30fc\u30c8", size=13, color=palette["on_surface_variant"]),
                            self.sheet_selector_wrapper,
                        ],
                        spacing=14,
                    ),
                    self._context_actions,
                ],
                spacing=28,
                tight=True,
            ),
        )

        self.chat_list = ft.ListView(
            expand=True,
            spacing=24,
            auto_scroll=True,
            padding=ft.Padding(0, 24, 0, 24),
            clip_behavior=ft.ClipBehavior.HARD_EDGE,
            adaptive=True,
        )

        self.user_input = ft.TextField(
            hint_text="",
            expand=True,
            multiline=True,
            min_lines=3,
            max_lines=5,
            on_submit=self._run_copilot,
            border_radius=18,
            border_color=palette["outline_variant"],
            focused_border_color=palette["primary"],
            cursor_color=palette["primary"],
            selection_color=ft.Colors.with_opacity(0.3, palette["primary"]),
            filled=True,
            fill_color=palette["surface_variant"],
            text_style=ft.TextStyle(color=palette["on_surface"]),
            hint_style=ft.TextStyle(color=palette["on_surface_variant"]),
            content_padding=ft.Padding(18, 16, 18, 16),
        )
        self._apply_mode_to_input_placeholder()

        self.mode_card_row = self._build_mode_cards()

        action_button_content = self._make_send_button()
        self.action_button = ft.Container(
            content=action_button_content,
            width=48,
            height=48,
            bgcolor=palette["primary"],
            border_radius=24,
            alignment=ft.alignment.center,
            ink=True,
            on_hover=self._handle_button_hover,
            animate_scale=100,
            scale=1,
            shadow=ft.BoxShadow(
                spread_radius=0,
                blur_radius=16,
                color=ft.Colors.with_opacity(0.18, palette["primary"]),
                offset=ft.Offset(0, 8),
            ),
            border=ft.border.all(1, ft.Colors.with_opacity(0.1, palette["on_primary"])),
        )

        self._chat_panel = ft.Container(
            expand=True,
            bgcolor=palette["surface_high"],
            border_radius=24,
            padding=ft.Padding(28, 32, 28, 32),
            border=ft.border.all(1, ft.Colors.with_opacity(0.08, palette["outline"])),
            shadow=ft.BoxShadow(
                spread_radius=0,
                blur_radius=18,
                color=ft.Colors.with_opacity(0.06, "#0F172A"),
                offset=ft.Offset(0, 10),
            ),
            clip_behavior=ft.ClipBehavior.HARD_EDGE,
            content=ft.Column(
                controls=[self.chat_list],
                spacing=0,
                expand=True,
            ),
        )

        self.action_button.margin = ft.margin.only(top=4)

        self._user_input_wrapper = ft.Container(
            content=self.user_input,
            col={"xs": 12, "sm": 9, "md": 10, "lg": 11},
        )

        self._action_button_wrapper = ft.Container(
            content=self.action_button,
            col={"xs": 12, "sm": 3, "md": 2, "lg": 1},
            alignment=ft.alignment.center_right,
        )

        self._composer_row = ft.ResponsiveRow(
            controls=[self._user_input_wrapper, self._action_button_wrapper],
            spacing=18,
            run_spacing=12,
            vertical_alignment=ft.CrossAxisAlignment.END,
        )

        self._composer_panel = ft.Container(
            bgcolor=palette["surface_high"],
            border_radius=24,
            padding=ft.Padding(28, 32, 28, 32),
            border=ft.border.all(1, ft.Colors.with_opacity(0.08, palette["outline"])),
            shadow=ft.BoxShadow(
                spread_radius=0,
                blur_radius=18,
                color=ft.Colors.with_opacity(0.06, "#0F172A"),
                offset=ft.Offset(0, 10),
            ),
            clip_behavior=ft.ClipBehavior.NONE,
            content=ft.Column(
                [self._composer_row],
                spacing=24,
            ),
        )

        self._mode_panel_container = ft.Container(
            content=self.mode_card_row,
            bgcolor=palette["surface_variant"],
            border_radius=20,
            padding=ft.Padding(18, 16, 18, 16),
            border=ft.border.all(1, ft.Colors.with_opacity(0.08, palette["outline"])),
            shadow=ft.BoxShadow(
                spread_radius=0,
                blur_radius=12,
                color=ft.Colors.with_opacity(0.04, "#0F172A"),
                offset=ft.Offset(0, 6),
            ),
            margin=ft.margin.only(top=4),
        )

        self._main_column = ft.Column(
            controls=[self._chat_panel, self._composer_panel, self._mode_panel_container],
            expand=True,
            spacing=24,
        )

        self._layout = ft.ResponsiveRow(
            controls=[
                ft.Container(
                    content=ft.Column([self._context_panel], spacing=16),
                    col={"xs": 12, "sm": 12, "md": 4, "lg": 3},
                    expand=True,
                ),
                ft.Container(
                    content=self._main_column,
                    col={"xs": 12, "sm": 12, "md": 8, "lg": 9},
                    expand=True,
                ),
            ],
            spacing=32,
            run_spacing=32,
            expand=True,
        )

        page_body = ft.Column(
            controls=[self._layout],
            spacing=0,
            expand=True,
        )

        self._content_container = ft.Container(
            content=page_body,
            expand=True,
            padding=ft.Padding(40, 48, 40, 48),
        )

        current_width = getattr(self.page, "width", None) or getattr(self.page.window, "width", None)
        current_height = getattr(self.page, "height", None) or getattr(self.page.window, "height", None)
        self._apply_responsive_layout(current_width, current_height)

        self.page.add(self._content_container)

    def _register_window_handlers(self):
        self.page.window.on_event = self._on_window_event
        self.page.on_resize = self._handle_page_resize
        self.page.on_disconnect = self._on_page_disconnect

    def _handle_page_resize(self, e: Optional[ft.ControlEvent]):
        width = getattr(self.page.window, "width", None) or getattr(self.page, "width", None)
        height = getattr(self.page.window, "height", None) or getattr(self.page, "height", None)
        self._apply_responsive_layout(width, height)
        self._update_ui()

    def _apply_responsive_layout(self, width: Optional[Union[int, float]], height: Optional[Union[int, float]]):
        try:
            width_value = float(width or 0)
        except (TypeError, ValueError):
            width_value = 0.0
        if width_value <= 0:
            return
        try:
            height_value = float(height or 0)
        except (TypeError, ValueError):
            height_value = 0.0

        if width_value < 720:
            layout_key = "compact"
        elif width_value < 1180:
            layout_key = "cozy"
        else:
            layout_key = "spacious"

        if layout_key == "compact":
            content_padding = ft.Padding(12, 16, 12, 20)
            panel_padding = ft.Padding(18, 18, 18, 18)
            mode_padding = ft.Padding(12, 10, 12, 10)
            chat_padding = ft.Padding(0, 14, 0, 14)
            composer_spacing = 12
            action_alignment = ft.alignment.center
            action_margin = ft.margin.only(top=12)
            preferred_chat_height = 320
            context_alignment = ft.MainAxisAlignment.START
            main_column_spacing = 18
            list_spacing = 18
        elif layout_key == "cozy":
            content_padding = ft.Padding(20, 26, 20, 30)
            panel_padding = ft.Padding(20, 22, 20, 22)
            mode_padding = ft.Padding(14, 12, 14, 12)
            chat_padding = ft.Padding(0, 18, 0, 18)
            composer_spacing = 14
            action_alignment = ft.alignment.center_right
            action_margin = ft.margin.only(left=10)
            preferred_chat_height = 360
            context_alignment = ft.MainAxisAlignment.END
            main_column_spacing = 20
            list_spacing = 20
        else:
            content_padding = ft.Padding(24, 28, 24, 32)
            panel_padding = ft.Padding(22, 24, 22, 24)
            mode_padding = ft.Padding(14, 12, 14, 12)
            chat_padding = ft.Padding(0, 20, 0, 20)
            composer_spacing = 16
            action_alignment = ft.alignment.center_right
            action_margin = ft.margin.only(left=10)
            preferred_chat_height = 340
            context_alignment = ft.MainAxisAlignment.END
            main_column_spacing = 20
            list_spacing = 22

        if self._content_container:
            self._content_container.padding = content_padding

        for panel in (self._context_panel, self._chat_panel, self._composer_panel):
            if panel:
                panel.padding = panel_padding

        if self._mode_panel_container:
            self._mode_panel_container.padding = mode_padding

        if self._layout:
            if layout_key == "compact":
                spacing_value = 20
            elif layout_key == "cozy":
                spacing_value = 24
            else:
                spacing_value = 28
            self._layout.spacing = spacing_value
            self._layout.run_spacing = spacing_value

        if self._main_column:
            self._main_column.spacing = main_column_spacing

        if self.chat_list:
            self.chat_list.padding = chat_padding
            self.chat_list.spacing = list_spacing

        if self.mode_card_row:
            mode_spacing = 12 if layout_key == "compact" else 18
            self.mode_card_row.spacing = mode_spacing
            self.mode_card_row.run_spacing = mode_spacing

        if self._composer_row:
            self._composer_row.spacing = composer_spacing
            self._composer_row.run_spacing = 12 if layout_key == "compact" else 16

        if self._context_actions:
            self._context_actions.alignment = context_alignment

        available_height = 0.0
        if height_value > 0:
            available_height = max(0.0, height_value - (content_padding.top + content_padding.bottom))

        if self._chat_panel:
            if available_height > 0:
                composer_est = (panel_padding.top + panel_padding.bottom) + 120
                mode_est = (mode_padding.top + mode_padding.bottom) + 110
                spacing_total = max(0, main_column_spacing) * 2
                calculated = available_height - composer_est - mode_est - spacing_total
                if layout_key == "compact":
                    min_chat_height = 200
                elif layout_key == "cozy":
                    min_chat_height = 220
                else:
                    min_chat_height = 240
                max_chat_height = preferred_chat_height
                if calculated <= 0:
                    chat_height = min_chat_height
                else:
                    chat_height = max(min_chat_height, min(max_chat_height, calculated))
            else:
                chat_height = preferred_chat_height
            self._chat_panel.height = chat_height

        if self._action_button_wrapper:
            self._action_button_wrapper.alignment = action_alignment
            self._action_button_wrapper.margin = action_margin

        if self.action_button:
            button_size = 52 if layout_key == "compact" else 48
            self.action_button.width = button_size
            self.action_button.height = button_size

        if self.user_input:
            if layout_key == "compact":
                self.user_input.min_lines = 4
                self.user_input.max_lines = 6
            else:
                self.user_input.min_lines = 3
                self.user_input.max_lines = 5

    def _make_send_button(self) -> ft.IconButton:
        palette = EXPRESSIVE_PALETTE
        return ft.IconButton(
            icon=ft.Icons.SEND_ROUNDED,
            icon_color=palette["on_primary"],
            icon_size=24,
            tooltip="\u9001\u4fe1",
            on_click=self._run_copilot,
            style=ft.ButtonStyle(
                shape=ft.CircleBorder(),
                padding=ft.Padding(0, 0, 0, 0),
                overlay_color=ft.Colors.with_opacity(0.14, palette["on_primary"]),
            ),
        )

    def _make_stop_button(self) -> ft.IconButton:
        palette = EXPRESSIVE_PALETTE
        return ft.IconButton(
            icon=ft.Icons.STOP_ROUNDED,
            icon_color=palette["on_error"],
            icon_size=24,
            tooltip="\u51e6\u7406\u3092\u505c\u6b62",
            on_click=self._stop_task,
            style=ft.ButtonStyle(
                shape=ft.CircleBorder(),
                padding=ft.Padding(0, 0, 0, 0),
                overlay_color=ft.Colors.with_opacity(0.18, palette["error"]),
            ),
        )

    def _handle_button_hover(self, e: ft.ControlEvent):
        if e.data == "true":
            e.control.scale = 1.02
        else:
            e.control.scale = 1
        e.control.update()

    def _build_mode_cards(self) -> ft.ResponsiveRow:
        palette = EXPRESSIVE_PALETTE
        options = [
            {
                "mode": CopilotMode.TRANSLATION_WITH_REFERENCES,
                "title": "\u7ffb\u8a33\uff08\u53c2\u7167\u3042\u308a\uff09",
                "icon": ft.Icons.LINK,
            },
            {
                "mode": CopilotMode.TRANSLATION,
                "title": "\u7ffb\u8a33\uff08\u901a\u5e38\uff09",
                "icon": ft.Icons.SYNC_ALT,
            },
            {
                "mode": CopilotMode.REVIEW,
                "title": "\u7ffb\u8a33\u30c1\u30a7\u30c3\u30af",
                "icon": ft.Icons.SPELLCHECK,
            },
        ]
        self._mode_card_map = {}
        cards: list[ft.Control] = []
        for item in options:
            mode = item["mode"]
            icon_container = ft.Container(
                width=28,
                height=28,
                bgcolor=ft.Colors.with_opacity(0.16, palette["primary"]),
                border_radius=14,
                alignment=ft.alignment.center,
                content=ft.Icon(item["icon"], size=16, color=palette["on_primary"]),
            )
            title_text = ft.Text(
                item["title"],
                size=14,
                weight=ft.FontWeight.BOLD,
                color=palette["on_surface"],
            )
            card_body = ft.Container(
                bgcolor=palette["surface"],
                border_radius=18,
                padding=ft.Padding(18, 16, 18, 16),
                border=ft.border.all(1, ft.Colors.with_opacity(0.1, palette["outline_variant"])),
                content=ft.Column(
                    [
                        ft.Row(
                            [icon_container, title_text],
                            spacing=12,
                            alignment=ft.MainAxisAlignment.START,
                            vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        ),
                    ],
                    spacing=8,
                    tight=True,
                ),
            )
            gesture = ft.GestureDetector(
                content=card_body,
                on_tap=lambda e, value=mode: self._handle_mode_card_select(value),
                mouse_cursor=ft.MouseCursor.CLICK,
            )
            wrapper = ft.Container(content=gesture, col={"xs": 12, "sm": 6, "md": 4, "lg": 4})
            cards.append(wrapper)
            self._mode_card_map[mode.value] = card_body

        row = ft.ResponsiveRow(controls=cards, spacing=18, run_spacing=18)
        self._refresh_mode_cards()
        return row

    def _refresh_mode_cards(self):
        palette = EXPRESSIVE_PALETTE
        for mode_value, card in self._mode_card_map.items():
            is_selected = mode_value == self.mode.value
            card.border = ft.border.all(
                2 if is_selected else 1,
                ft.Colors.with_opacity(0.9, palette["primary"]) if is_selected else ft.Colors.with_opacity(0.1, palette["outline_variant"]),
            )
            card.bgcolor = ft.Colors.with_opacity(0.14, palette["primary"]) if is_selected else palette["surface"]

    def _handle_mode_card_select(self, new_mode: CopilotMode):
        if self.app_state not in {AppState.READY, AppState.ERROR}:
            return
        self._set_mode(new_mode)

    def _set_mode(self, new_mode: CopilotMode):
        if not isinstance(new_mode, CopilotMode):
            return
        if new_mode == self.mode:
            return
        self.mode = new_mode
        self._apply_mode_to_input_placeholder()
        if self.mode_selector:
            self.mode_selector.value = self.mode.value
        self._refresh_mode_cards()
        if self.request_queue:
            self.request_queue.put(RequestMessage(RequestType.UPDATE_CONTEXT, {"mode": self.mode.value}))
        self._update_ui()

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
        self._set_mode(new_mode)

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
            "base": EXPRESSIVE_PALETTE["on_surface_variant"],
            "ready": EXPRESSIVE_PALETTE["primary"],
            "busy": EXPRESSIVE_PALETTE["secondary"],
            "error": EXPRESSIVE_PALETTE["error"],
            "stopping": EXPRESSIVE_PALETTE["secondary"],
            "info": EXPRESSIVE_PALETTE["on_surface_variant"],
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
        if self.new_chat_button:
            if new_state in {AppState.TASK_IN_PROGRESS, AppState.STOPPING}:
                self.new_chat_button.disabled = True
            elif not self._browser_reset_in_progress and can_interact:
                self.new_chat_button.disabled = False

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

    def _handle_new_chat_click(self, e: Optional[ft.ControlEvent]):
        if self._browser_reset_in_progress:
            return
        if self.app_state not in {AppState.READY, AppState.ERROR}:
            return
        if not self.worker or not self.request_queue:
            return

        if self.chat_list:
            self.chat_list.controls.clear()
        with self.history_lock:
            self.chat_history.clear()
        self._update_save_button_state()
        if self.chat_list:
            self.chat_list.update()

        self._browser_reset_in_progress = True
        if self.new_chat_button:
            self.new_chat_button.disabled = True
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

                if self._auto_test_enabled:
                    print(
                        "AUTOTEST: excel context ready",
                        {
                            "workbooks": workbook_names,
                            "active_workbook": active_workbook,
                            "active_sheet": active_sheet,
                        },
                        flush=True,
                    )

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
                if self._auto_test_enabled:
                    print(f"AUTOTEST: excel context error - {error_message}", flush=True)
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
            "base": EXPRESSIVE_PALETTE["on_surface_variant"],
            "info": EXPRESSIVE_PALETTE["primary"],
            "error": EXPRESSIVE_PALETTE["error"],
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
                if self.new_chat_button and self.app_state in {AppState.READY, AppState.ERROR}:
                    self.new_chat_button.disabled = False

        if response.type is ResponseType.INITIALIZATION_COMPLETE:
            self._set_state(AppState.READY)
            if self.status_label:
                self.status_label.value = response.content or self.status_label.value
            self._focus_app_window()
            if self._auto_test_enabled:
                print("AUTOTEST: initialization complete", flush=True)
            self._schedule_auto_test()
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
            if self._auto_test_enabled and status_text:
                print(f"AUTOTEST: status '{status_text}'", flush=True)
        elif response.type is ResponseType.ERROR:
            if self.app_state in {AppState.TASK_IN_PROGRESS, AppState.STOPPING}:
                if self.status_label:
                    self.status_label.value = response.content or "\u51e6\u7406\u4e2d\u306b\u30a8\u30e9\u30fc\u304c\u767a\u751f\u3057\u307e\u3057\u305f"
                    self.status_label.color = status_palette["error"]
                    self.status_label.opacity = 0.9
                if response.content:
                    self._add_message(response.type, response.content)
                if self._auto_test_triggered:
                    print(f"AUTOTEST: error '{response.content}'", flush=True)
            else:
                self._set_state(AppState.ERROR)
                if response.content:
                    self._add_message(response.type, response.content)
                    if self._auto_test_triggered:
                        print(f"AUTOTEST: error '{response.content}'", flush=True)
            if self._browser_reset_in_progress:
                self._browser_reset_in_progress = False
                if self.new_chat_button and self.app_state in {AppState.READY, AppState.ERROR}:
                    self.new_chat_button.disabled = False
            if self._auto_test_triggered:
                self._auto_test_completed = True
        elif response.type is ResponseType.END_OF_TASK:
            self._set_state(AppState.READY)
            if self._auto_test_triggered:
                self._auto_test_completed = True
                print("AUTOTEST: task marked as completed", flush=True)
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
                    if self.new_chat_button and self.app_state in {AppState.READY, AppState.ERROR}:
                        self.new_chat_button.disabled = False
            elif response.content:
                self._add_message(type_value, response.content)
        else:
            if response.content:
                self._add_message(type_value, response.content)

        if response.type is ResponseType.FINAL_ANSWER and self._auto_test_triggered:
            self._auto_test_completed = True
            print("AUTOTEST: received final answer", flush=True)
        elif self._auto_test_enabled and response.type in {ResponseType.OBSERVATION, ResponseType.ACTION, ResponseType.THOUGHT}:
            snippet = (response.content or "").strip()
            if snippet:
                print(f"AUTOTEST: {response.type.value} '{snippet[:120]}'", flush=True)

        self._update_ui()

    def _schedule_auto_test(self) -> None:
        if (
            self._auto_test_triggered
            or not self._auto_test_prompt
            or not self.page
        ):
            return

        self._auto_test_triggered = True
        self._auto_test_completed = False
        print(
            f"AUTOTEST: scheduled (delay={self._auto_test_delay}s, "
            f"workbook={self._auto_test_workbook or '(unchanged)'}, "
            f"sheet={self._auto_test_sheet or '(unchanged)'})",
            flush=True,
        )
        if self._auto_test_timeout:
            self._auto_test_deadline = time.monotonic() + self._auto_test_timeout
            def _timeout_watch():
                try:
                    remaining = self._auto_test_timeout
                    while remaining > 0 and not self._auto_test_completed:
                        time.sleep(min(1.0, remaining))
                        remaining = self._auto_test_deadline - time.monotonic()
                    if not self._auto_test_completed and self._auto_test_enabled:
                        print(
                            f"AUTOTEST: timeout reached after {self._auto_test_timeout}s",
                            flush=True,
                        )
                except Exception:
                    pass
            try:
                self.page.run_thread(_timeout_watch)
            except Exception:
                threading.Thread(target=_timeout_watch, daemon=True).start()

        def _runner():
            try:
                if self._auto_test_delay:
                    time.sleep(self._auto_test_delay)
                self._execute_auto_test()
            except Exception as exc:
                print(f"Auto-test execution failed: {exc}", flush=True)

        try:
            self.page.run_thread(_runner)
        except Exception:
            _runner()

    def _execute_auto_test(self) -> None:
        if self.app_state not in {AppState.READY, AppState.ERROR}:
            return

        def _select_option(dropdown: ft.Dropdown, value: Optional[str], updating_flag: str) -> None:
            if not dropdown or not value:
                return
            options = dropdown.options or []
            option_values = {(option.key or option.text): option for option in options}
            if value not in option_values:
                return
            setattr(self, updating_flag, True)
            dropdown.value = value
            setattr(self, updating_flag, False)
            try:
                dropdown.update()
            except Exception:
                pass

        _select_option(self.workbook_selector, self._auto_test_workbook, "workbook_selection_updating")
        _select_option(self.sheet_selector, self._auto_test_sheet, "sheet_selection_updating")

        print(
            "AUTOTEST: dispatching prompt",
            {
                "workbook": self.workbook_selector.value if self.workbook_selector else None,
                "sheet": self.sheet_selector.value if self.sheet_selector else None,
            },
            flush=True,
        )

        if self.user_input:
            self.user_input.value = self._auto_test_prompt
            self.user_input.update()
        self._update_ui()
        self._run_copilot(None)
        print("AUTOTEST: prompt submitted", flush=True)

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
    autotest_active = _is_autotest_mode_enabled()
    if autotest_active:
        logging.info("Auto test mode detected; forcing Flet view to display.")
        app_kwargs["view"] = ft.AppView.WEB_BROWSER
    elif args.no_browser or COPILOT_HEADLESS:
        app_kwargs["view"] = None
    else:
        app_kwargs["view"] = ft.AppView.WEB_BROWSER
    ft.app(**app_kwargs)
