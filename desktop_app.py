# desktop_app.py

import argparse
import json
import logging
import os
import queue
import platform
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

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
from excel_copilot.ui.theme import EXPRESSIVE_PALETTE, elevated_surface_gradient
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

ASSETS_DIR = Path(__file__).resolve().parent / "assets"
DEFAULT_FONT_PATH = ASSETS_DIR / "fonts" / "NotoSansCJKjp-Regular.otf"
DEFAULT_FONT_RELATIVE_PATH = Path("fonts") / DEFAULT_FONT_PATH.name
PRIMARY_FONT_FAMILY = "NotoSansJP"

DEFAULT_AUTOTEST_TIMEOUT_SECONDS = 180.0  # 3-minute auto-test timeout
DEFAULT_AUTOTEST_SOURCE_REFERENCE_URL = (
    "https://ralleti-my.sharepoint.com/:b:/g/personal/yuukikod_ralleti_onmicrosoft_com/"
    "Eao1vzScMVdMpyecR1s8KFEBzr0LMPNp5u5ksA3U3TCMwQ?e=jwmAly"
)
DEFAULT_AUTOTEST_TARGET_REFERENCE_URL = (
    "https://ralleti-my.sharepoint.com/:b:/g/personal/yuukikod_ralleti_onmicrosoft_com/"
    "EdJ586XuxedLsaSArCCve9kB1K79F0BvGqxzuZBhfWWS-w?e=wjR4C2"
)

MODE_LABELS = {
    CopilotMode.TRANSLATION: "翻訳（通常）",
    CopilotMode.TRANSLATION_WITH_REFERENCES: "翻訳（参照あり）",
    CopilotMode.REVIEW: "翻訳チェック",
}

FORM_FIELD_DEFINITIONS: Dict[CopilotMode, List[Dict[str, Any]]] = {
    CopilotMode.TRANSLATION: [
        {
            "name": "cell_range",
            "label": "ソース範囲",
            "argument": "cell_range",
            "required": True,
            "placeholder": "例: A2:A20",
            "group": "scope",
        },
        {
            "name": "translation_output_range",
            "label": "出力範囲",
            "argument": "translation_output_range",
            "required": True,
            "placeholder": "例: B2:B20",
            "group": "output",
        },
        {
            "name": "target_language",
            "label": "ターゲット言語",
            "argument": "target_language",
            "default": "English",
            "placeholder": "例: English",
            "group": "options",
        },
    ],
    CopilotMode.TRANSLATION_WITH_REFERENCES: [
        {
            "name": "cell_range",
            "label": "ソース範囲",
            "argument": "cell_range",
            "required": True,
            "placeholder": "例: A2:A20",
            "group": "scope",
        },
        {
            "name": "translation_output_range",
            "label": "出力範囲（翻訳・メモ・参照ペア）",
            "argument": "translation_output_range",
            "required": True,
            "placeholder": "例: B2:D20",
            "group": "output",
        },
        {
            "name": "target_language",
            "label": "ターゲット言語",
            "argument": "target_language",
            "default": "English",
            "placeholder": "例: English",
            "group": "options",
        },
        {
            "control": "section",
            "label": "参照資料",
            "description": "必要に応じて HTTP(S) で取得できる参照 URL を記入してください（1 件まで）。",
            "group": "references",
            "expanded": False,
            "children": [
                {
                    "name": "source_reference_urls",
                    "label": "参照URL（原文側）",
                    "argument": "source_reference_urls",
                    "type": "list",
                    "multiline": True,
                    "min_lines": 1,
                    "max_lines": 3,
                    "placeholder": "例: https://example.com/source-guideline",
                    "group": "references",
                },
                {
                    "name": "target_reference_urls",
                    "label": "参照URL（翻訳側）",
                    "argument": "target_reference_urls",
                    "type": "list",
                    "multiline": True,
                    "min_lines": 1,
                    "max_lines": 3,
                    "placeholder": "例: https://example.com/english-reference",
                    "group": "references",
                },
            ],
        },
    ],
    CopilotMode.REVIEW: [
        {
            "name": "source_range",
            "label": "原文範囲",
            "argument": "source_range",
            "required": True,
            "placeholder": "例: B2:B20",
            "group": "scope",
        },
        {
            "name": "translated_range",
            "label": "訳文範囲",
            "argument": "translated_range",
            "required": True,
            "placeholder": "例: C2:C20",
            "group": "scope",
        },
        {
            "name": "status_output_range",
            "label": "ステータス列",
            "argument": "status_output_range",
            "required": True,
            "placeholder": "例: D2:D20",
            "group": "output",
        },
        {
            "name": "issue_output_range",
            "label": "指摘列",
            "argument": "issue_output_range",
            "required": True,
            "placeholder": "例: E2:E20",
            "group": "output",
        },
        {
            "name": "highlight_output_range",
            "label": "ハイライト列",
            "argument": "highlight_output_range",
            "required": True,
            "placeholder": "例: F2:F20",
            "group": "output",
        },
        {
            "name": "corrected_output_range",
            "label": "修正案列（任意）",
            "argument": "corrected_output_range",
            "placeholder": "例: G2:G20",
            "group": "options",
        },
    ],
}
FORM_GROUP_LABELS: Dict[str, str] = {
    "mode": "モード",
    "scope": "対象範囲",
    "output": "出力設定",
    "references": "参考資料",
    "options": "オプション",
}

FORM_GROUP_ORDER: List[str] = ["mode", "scope", "output", "references", "options"]


def _flatten_field_definitions(definitions: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    flat: List[Dict[str, Any]] = []
    for field in definitions:
        if field.get("control") == "section":
            flat.extend(field.get("children", []))
        else:
            flat.append(field)
    return flat


def _iter_mode_field_definitions(mode: CopilotMode) -> List[Dict[str, Any]]:
    return _flatten_field_definitions(FORM_FIELD_DEFINITIONS.get(mode, []))

def _default_autotest_payload(source_url: Optional[str], target_url: Optional[str]) -> str:
    arguments: Dict[str, Any] = {
        "cell_range": "A2:A20",
        "translation_output_range": "B2:D20",
        "target_language": "English",
    }
    if source_url:
        arguments["source_reference_urls"] = [source_url]
    if target_url:
        arguments["target_reference_urls"] = [target_url]
    payload = {
        "mode": CopilotMode.TRANSLATION_WITH_REFERENCES.value,
        "arguments": arguments,
    }
    return json.dumps(payload, ensure_ascii=False)


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

        self._primary_font_family = "Yu Gothic UI"
        self._hint_font_family = "Yu Gothic UI"
        self.status_label: Optional[ft.Text] = None
        self.workbook_selector: Optional[ft.Dropdown] = None
        self.sheet_selector: Optional[ft.Dropdown] = None
        self.chat_list: Optional[ft.ListView] = None
        self.save_log_button: Optional[ft.TextButton] = None
        self.workbook_refresh_button: Optional[ft.TextButton] = None
        self.new_chat_button: Optional[ft.TextButton] = None
        self.mode_card_row: Optional[ft.ResponsiveRow] = None
        self.form_controls: Dict[str, ft.TextField] = {}
        self.form_error_text: Optional[ft.Text] = None
        self._form_submit_button: Optional[ft.Control] = None
        self._form_cancel_button: Optional[ft.Control] = None
        self._form_body_column: Optional[ft.Container] = None
        self._form_tabs: Optional[ft.Tabs] = None
        self._form_progress_indicator: Optional[ft.ProgressRing] = None
        self._form_progress_text: Optional[ft.Text] = None
        self._form_continue_button: Optional[ft.Control] = None
        self._form_panel: Optional[ft.Container] = None
        self._mode_card_map: dict[str, ft.Container] = {}
        self._context_panel: Optional[ft.Container] = None
        self._context_actions: Optional[ft.ResponsiveRow] = None
        self._chat_panel: Optional[ft.Container] = None
        self._mode_panel_container: Optional[ft.Container] = None
        self._content_container: Optional[ft.Container] = None
        self._layout: Optional[ft.ResponsiveRow] = None
        self._main_column: Optional[ft.Column] = None
        self._chat_empty_state: Optional[ft.Container] = None
        self._chat_filter_dropdown: Optional[ft.Dropdown] = None
        self._chat_filter_value: str = "all"
        self._chat_scroll_button: Optional[ft.TextButton] = None
        self._chat_header_subtitle: Optional[ft.Text] = None

        self.chat_history: list[Dict[str, Any]] = []
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
        self._workbook_refresh_button_default_text = "ブック一覧を更新"
        self._status_icon: Optional[ft.Icon] = None
        self._sync_status_text: Optional[ft.Text] = None
        self._context_summary_text: Optional[ft.Text] = None
        self._last_context_refresh_at: Optional[datetime] = None
        self._group_summary_labels: Dict[str, ft.Text] = {}
        self._field_groups: Dict[str, str] = {}
        self._task_recently_completed = False

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
            self._auto_test_prompt = _default_autotest_payload(
                self._auto_test_source_url,
                self._auto_test_target_url,
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
        self._auto_test_shutdown_scheduled = False
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
        font_family = PRIMARY_FONT_FAMILY
        font_path = DEFAULT_FONT_PATH
        if font_path.is_file():
            existing_fonts = dict(getattr(self.page, "fonts", {}) or {})
            existing_fonts[font_family] = DEFAULT_FONT_RELATIVE_PATH.as_posix()
            self.page.fonts = existing_fonts
        else:
            font_family = "Yu Gothic UI"
            logging.warning(
                "Custom font was not found at %s; falling back to system font '%s'.",
                font_path,
                font_family,
            )
        if platform.system() == "Windows":
            hint_family = "Yu Gothic UI"
        else:
            hint_family = font_family
        self._primary_font_family = font_family
        self._hint_font_family = hint_family
        self.page.theme = ft.Theme(
            color_scheme_seed=palette["primary"],
            use_material3=True,
            font_family=font_family,
        )
        self.page.theme_mode = ft.ThemeMode.LIGHT
        self.page.bgcolor = palette["surface_dim"]
        self.page.window.bgcolor = palette["surface_dim"]
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
            print(f"繧｢繝励Μ繧ｦ繧｣繝ｳ繝峨え縺ｮ蜑埼擇陦ｨ遉ｺ縺ｫ螟ｱ謨励＠縺ｾ縺励◆: {focus_err}")

    def _focus_excel_window(self):
        try:
            with ExcelManager(self.current_workbook_name) as manager:
                manager.focus_application_window()
        except Exception as focus_err:
            print(f"Excel繧ｦ繧｣繝ｳ繝峨え縺ｮ蜑埼擇陦ｨ遉ｺ縺ｫ螟ｱ謨励＠縺ｾ縺励◆: {focus_err}")

    def _build_layout(self):
        palette = EXPRESSIVE_PALETTE

        self.status_label = ft.Text(
            "初期化中...",
            size=12,
            color=palette["on_surface_variant"],
            font_family=self._primary_font_family,
            animate_opacity=300,
            animate_scale=600,
        )
        self._status_icon = ft.Icon(ft.Icons.CIRCLE, size=16, color=palette["primary"])
        self._sync_status_text = ft.Text(
            "最終同期: 未実行",
            size=11,
            color=palette["on_surface_variant"],
            font_family=self._hint_font_family,
        )
        status_row = ft.Row(
            controls=[self._status_icon, self.status_label],
            spacing=10,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )
        status_card = ft.Container(
            content=ft.Column(
                [status_row, self._sync_status_text],
                spacing=6,
                tight=True,
            ),
            bgcolor=ft.Colors.with_opacity(0.6, palette["surface_variant"]),
            border_radius=18,
            padding=ft.Padding(18, 14, 18, 16),
            border=ft.border.all(1, ft.Colors.with_opacity(0.08, palette["outline"])),
        )

        button_shape = ft.RoundedRectangleBorder(radius=18)
        button_overlay = ft.Colors.with_opacity(0.12, palette["primary"])

        self.new_chat_button = ft.FilledTonalButton(
            text="新しいチャット",
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
            "text_style": ft.TextStyle(color=palette["on_surface"], size=12, font_family=self._primary_font_family),
            "hint_style": ft.TextStyle(color=palette["on_surface_variant"], size=12, font_family=self._primary_font_family),
            "disabled": True,
            "filled": True,
            "suffix_icon": ft.Icon(ft.Icons.KEYBOARD_ARROW_DOWN_ROUNDED, color=palette["on_surface_variant"]),
        }

        self._context_summary_text = ft.Text(
            "選択中: ブック未選択 / シート未選択",
            size=12,
            color=palette["on_surface_variant"],
            font_family=self._hint_font_family,
        )

        self.workbook_selector = ft.Dropdown(
            options=[],
            on_change=self._on_workbook_change,
            on_focus=self._on_workbook_dropdown_focus,
            hint_text="ブックを選択",
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
            hint_text="シートを選択",
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

        selector_card = ft.Container(
            content=ft.Column(
                [
                    self._context_summary_text,
                    ft.Column(
                        [
                            ft.Text("ブック", size=13, color=palette["on_surface_variant"], font_family=self._primary_font_family),
                            self.workbook_selector_wrapper,
                            ft.Text("シート", size=13, color=palette["on_surface_variant"], font_family=self._primary_font_family),
                            self.sheet_selector_wrapper,
                        ],
                        spacing=14,
                        tight=True,
                    ),
                ],
                spacing=12,
                tight=True,
            ),
            bgcolor=ft.Colors.with_opacity(0.55, palette["surface_variant"]),
            border_radius=18,
            padding=ft.Padding(18, 18, 18, 20),
            border=ft.border.all(1, ft.Colors.with_opacity(0.08, palette["outline"])),
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

        actions_divider = ft.Container(height=1, bgcolor=ft.Colors.with_opacity(0.1, palette["outline"]))

        context_column = ft.Column(
            controls=[status_card, selector_card, actions_divider, self._context_actions],
            spacing=16,
            tight=True,
        )

        self._context_panel = ft.Container(
            bgcolor=palette["surface"],
            border_radius=24,
            padding=ft.Padding(24, 28, 24, 28),
            border=ft.border.all(1, ft.Colors.with_opacity(0.08, palette["outline"])),
            shadow=ft.BoxShadow(
                spread_radius=0,
                blur_radius=14,
                color=ft.Colors.with_opacity(0.06, "#0F172A"),
                offset=ft.Offset(0, 8),
            ),
            content=context_column,
        )

        self._chat_filter_dropdown = ft.Dropdown(
            value=self._chat_filter_value,
            options=[
                ft.dropdown.Option("all", "すべて"),
                ft.dropdown.Option("user", "ユーザー入力"),
                ft.dropdown.Option("ai", "AI応答"),
                ft.dropdown.Option("system", "通知・エラー"),
            ],
            on_change=self._on_chat_filter_change,
            border_radius=18,
            border_color=palette["outline_variant"],
            focused_border_color=palette["primary"],
            fill_color=palette["surface_variant"],
            text_style=ft.TextStyle(color=palette["on_surface"], size=12, font_family=self._primary_font_family),
            hint_text="表示を絞り込む",
            dense=True,
        )
        self._chat_scroll_button = ft.TextButton(
            text="最新の結果へ",
            icon=ft.Icons.ARROW_DOWNWARD,
            on_click=self._scroll_chat_to_latest,
            disabled=True,
        )
        chat_header_title = ft.Text(
            "チャットタイムライン",
            size=15,
            weight=ft.FontWeight.W_600,
            color=palette["on_surface"],
            font_family=self._primary_font_family,
        )
        self._chat_header_subtitle = ft.Text(
            "処理ログと結果が最新順に表示されます。",
            size=12,
            color=palette["on_surface_variant"],
            font_family=self._hint_font_family,
        )
        header_column = ft.Column(
            [chat_header_title, self._chat_header_subtitle],
            spacing=4,
            tight=True,
        )
        header_actions = ft.Row(
            [self._chat_filter_dropdown, self._chat_scroll_button],
            spacing=12,
            alignment=ft.MainAxisAlignment.END,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )
        chat_header_section = ft.Column(
            controls=[
                ft.Row(
                    [header_column, header_actions],
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                    vertical_alignment=ft.CrossAxisAlignment.START,
                ),
                ft.Container(height=1, bgcolor=ft.Colors.with_opacity(0.05, palette["outline"])),
            ],
            spacing=14,
        )

        self._chat_empty_state = ft.Container(
            content=ft.Column(
                [
                    ft.Text(
                        "まだメッセージがありません。",
                        size=13,
                        color=palette["on_surface"],
                        font_family=self._primary_font_family,
                        weight=ft.FontWeight.W_500,
                    ),
                    ft.Text(
                        "フォームを送信すると処理状況と結果がここに表示されます。",
                        size=12,
                        color=palette["on_surface_variant"],
                        font_family=self._hint_font_family,
                    ),
                ],
                spacing=6,
                tight=True,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            ),
            padding=ft.Padding(28, 24, 28, 24),
            border_radius=16,
            border=ft.border.all(1, ft.Colors.with_opacity(0.06, palette["outline_variant"])),
            bgcolor=ft.Colors.with_opacity(0.18, palette["surface_variant"]),
            visible=True,
        )

        self.chat_list = ft.ListView(
            expand=True,
            spacing=24,
            auto_scroll=True,
            padding=ft.Padding(0, 24, 0, 24),
            clip_behavior=ft.ClipBehavior.HARD_EDGE,
            adaptive=True,
        )

        self._chat_panel = ft.Container(
            expand=True,
            bgcolor=palette["surface_high"],
            gradient=elevated_surface_gradient(),
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
                controls=[chat_header_section, self._chat_empty_state, self.chat_list],
                spacing=24,
                expand=True,
            ),
        )

        self._form_panel = self._build_form_panel()

        form_and_chat_row = ft.ResponsiveRow(
            controls=[
                ft.Container(
                    content=self._form_panel,
                    col={"xs": 12, "md": 12, "lg": 6},
                    expand=True,
                ),
                ft.Container(
                    content=self._chat_panel,
                    col={"xs": 12, "md": 12, "lg": 6},
                    expand=True,
                ),
            ],
            spacing=24,
            run_spacing=24,
        )

        self._main_column = ft.Column(
            controls=[form_and_chat_row],
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
            alignment=ft.MainAxisAlignment.CENTER,
            vertical_alignment=ft.CrossAxisAlignment.START,
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
            padding=ft.Padding(28, 36, 28, 36),
            alignment=ft.alignment.top_center,
        )

        self._update_context_summary()
        self._update_chat_empty_state()

        current_width = getattr(self.page, "width", None) or getattr(self.page.window, "width", None)
        current_height = getattr(self.page, "height", None) or getattr(self.page.window, "height", None)
        self._apply_responsive_layout(current_width, current_height)

        self.page.add(self._content_container)

    def _build_form_panel(self) -> ft.Container:
        palette = EXPRESSIVE_PALETTE
        can_interact = self.app_state in {AppState.READY, AppState.ERROR}

        tabs_control, controls_map = self._create_form_controls_for_mode(self.mode)
        self.form_controls = controls_map
        self._form_tabs = tabs_control
        self._form_tabs.expand = True
        self._form_body_column = ft.Container(
            content=self._form_tabs,
            height=420,
            expand=False,
            padding=ft.Padding(4, 0, 4, 0),
        )

        self.form_error_text = ft.Text(
            "",
            color=palette["error"],
            size=12,
            visible=False,
            font_family=self._hint_font_family,
        )

        self._form_submit_button = ft.FilledButton(
            "フォームを送信",
            icon=ft.Icons.CHECK_CIRCLE_OUTLINE,
            on_click=self._submit_form,
            disabled=not can_interact,
        )

        self._form_cancel_button = ft.OutlinedButton(
            "停止",
            icon=ft.Icons.STOP_CIRCLE_OUTLINED,
            on_click=self._stop_task,
            disabled=True,
            visible=False,
        )

        self._form_continue_button = ft.TextButton(
            "続けて実行",
            icon=ft.Icons.REPLAY_OUTLINED,
            on_click=self._handle_form_continue_click,
            visible=False,
        )

        self._form_progress_indicator = ft.ProgressRing(
            width=20,
            height=20,
            stroke_width=2,
            color=palette["primary"],
            visible=False,
        )
        self._form_progress_text = ft.Text(
            "",
            size=12,
            color=palette["on_surface_variant"],
            font_family=self._hint_font_family,
            visible=False,
        )

        progress_cluster = ft.Row(
            controls=[self._form_progress_indicator, self._form_progress_text],
            spacing=8,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )
        action_buttons = ft.Row(
            controls=[self._form_continue_button, self._form_cancel_button, self._form_submit_button],
            alignment=ft.MainAxisAlignment.END,
            spacing=12,
        )
        action_bar = ft.Column(
            controls=[
                ft.Container(height=1, bgcolor=ft.Colors.with_opacity(0.08, palette["outline"])),
                ft.Row(
                    controls=[progress_cluster, action_buttons],
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                ),
            ],
            spacing=12,
            tight=True,
        )

        header = ft.Text(
            "フォーム入力",
            size=16,
            weight=ft.FontWeight.W_600,
            color=palette["primary"],
            font_family=self._primary_font_family,
        )

        content = ft.Column(
            controls=[header, self._form_body_column, self.form_error_text, action_bar],
            spacing=20,
            tight=True,
        )

        panel = ft.Container(
            content=content,
            bgcolor=palette["surface_high"],
            border_radius=24,
            padding=ft.Padding(22, 24, 22, 24),
            border=ft.border.all(1, ft.Colors.with_opacity(0.08, palette["outline"])),
            shadow=ft.BoxShadow(
                spread_radius=0,
                blur_radius=18,
                color=ft.Colors.with_opacity(0.06, "#0F172A"),
                offset=ft.Offset(0, 10),
            ),
            clip_behavior=ft.ClipBehavior.NONE,
        )

        self._update_all_group_summaries()
        return panel

    def _create_form_controls_for_mode(
        self,
        mode: CopilotMode,
        initial_values: Optional[Dict[str, str]] = None,
    ) -> Tuple[ft.Tabs, Dict[str, ft.TextField]]:
        definitions = FORM_FIELD_DEFINITIONS.get(mode, [])
        preserved = initial_values or {}
        palette = EXPRESSIVE_PALETTE

        grouped_controls: Dict[str, List[ft.Control]] = {group: [] for group in FORM_GROUP_ORDER}
        new_controls: Dict[str, ft.TextField] = {}
        self._field_groups = {}
        self._group_summary_labels = {}

        grouped_controls.setdefault("mode", []).append(self._build_mode_selection_control())

        def _build_required_badge() -> ft.Container:
            return ft.Container(
                ft.Text("Required", size=10, weight=ft.FontWeight.W_500, color=palette["primary"]),
                padding=ft.Padding(8, 4, 8, 4),
                bgcolor=ft.Colors.with_opacity(0.12, palette["primary"]),
                border_radius=12,
            )

        def _build_text_field(definition: Dict[str, Any], group: str) -> ft.TextField:
            name = definition["name"]
            value = preserved.get(name, definition.get("default", "")) or ""
            multiline = bool(definition.get("multiline"))

            text_field = ft.TextField(
                label=definition["label"],
                hint_text=definition.get("placeholder", ""),
                value=value,
                expand=True,
                multiline=multiline,
                border_radius=18,
                filled=True,
                fill_color=palette["surface_variant"],
                border_color=palette["outline_variant"],
                focused_border_color=palette["primary"],
                cursor_color=palette["primary"],
                selection_color=ft.Colors.with_opacity(0.2, palette["primary"]),
                text_style=ft.TextStyle(font_family=self._primary_font_family, size=13),
            )
            if multiline and definition.get("min_lines"):
                text_field.min_lines = definition["min_lines"]
            if multiline and definition.get("max_lines"):
                text_field.max_lines = definition["max_lines"]
            if definition.get("type") == "int":
                text_field.keyboard_type = ft.KeyboardType.NUMBER
            if definition.get("required"):
                text_field.suffix = _build_required_badge()
            text_field.on_submit = self._submit_form
            text_field.on_change = lambda e, field_name=name: self._handle_form_value_change(field_name)
            new_controls[name] = text_field
            self._field_groups[name] = group
            return text_field

        
        for field in definitions:
            group_key = field.get("group", "scope")
            if field.get("control") == "section":
                child_controls: List[ft.Control] = []
                for child_definition in field.get("children", []):
                    child_group = child_definition.get("group", group_key)
                    child_field = _build_text_field(child_definition, child_group)
                    child_controls.append(child_field)
                if not child_controls:
                    continue
                section_elements: List[ft.Control] = [
                    ft.Text(
                        field.get("label", ""),
                        size=14,
                        weight=ft.FontWeight.W_500,
                        color=palette["primary"],
                        font_family=self._primary_font_family,
                    )
                ]
                description_text = field.get("description")
                if description_text:
                    section_elements.append(
                        ft.Text(
                            description_text,
                            size=12,
                            color=palette["on_surface_variant"],
                            font_family=self._hint_font_family,
                        )
                    )
                section_elements.extend(child_controls)
                grouped_controls.setdefault(group_key, []).append(
                    ft.Container(
                        content=ft.Column(section_elements, spacing=12, tight=True),
                        padding=ft.Padding(20, 18, 20, 20),
                        border_radius=18,
                        bgcolor=ft.Colors.with_opacity(0.08, palette["surface_variant"]),
                    )
                )
            else:
                text_field = _build_text_field(field, group_key)
                grouped_controls.setdefault(group_key, []).append(text_field)

        tabs: List[ft.Tab] = []
        for group_key in FORM_GROUP_ORDER:
            controls_for_group = grouped_controls.get(group_key) or []
            if not controls_for_group:
                continue
            summary_label = ft.Text(
                "未入力",
                size=12,
                color=palette["on_surface_variant"],
                font_family=self._hint_font_family,
            )
            self._group_summary_labels[group_key] = summary_label

            if group_key == "mode":
                summary_label.value = f"現在: {MODE_LABELS.get(self.mode, self.mode.value)}"
                header_row = ft.Row(
                    controls=[summary_label],
                    alignment=ft.MainAxisAlignment.END,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                )
                tab_body_controls: List[ft.Control] = [header_row] + controls_for_group
                body_spacing = 12
            else:
                header_row = ft.Row(
                    controls=[
                        ft.Text(
                            FORM_GROUP_LABELS.get(group_key, group_key.title()),
                            size=15,
                            weight=ft.FontWeight.W_600,
                            color=palette["primary"],
                            font_family=self._primary_font_family,
                        ),
                        summary_label,
                    ],
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                )
                tab_body_controls = [header_row, ft.Column(controls_for_group, spacing=12, tight=True)]
                body_spacing = 16

            tab_content = ft.Column(
                controls=tab_body_controls,
                spacing=body_spacing,
                tight=True,
            )
            tabs.append(
                ft.Tab(
                    text=FORM_GROUP_LABELS.get(group_key, group_key.title()),
                    content=ft.Container(tab_content, padding=ft.Padding(4, 6, 4, 6)),
                )
            )

        tabs_control = ft.Tabs(
            tabs=tabs,
            animation_duration=200,
            expand=True,
            divider_color=ft.Colors.with_opacity(0.08, palette["outline"]),
            indicator_color=palette["primary"],
        )

        return tabs_control, new_controls

    def _refresh_form_panel(self) -> None:
        if not self._form_body_column:
            return
        preserved_values = {name: ctrl.value for name, ctrl in self.form_controls.items()}
        tabs_control, controls_map = self._create_form_controls_for_mode(self.mode, preserved_values)
        self.form_controls = controls_map
        self._form_tabs = tabs_control
        self._form_body_column.content = tabs_control
        self._set_form_error("")
        if self._form_submit_button:
            self._form_submit_button.disabled = self.app_state not in {AppState.READY, AppState.ERROR}
        if self._form_continue_button:
            self._form_continue_button.visible = False
        self._update_all_group_summaries()
        self._update_ui()

    def _set_form_error(self, message: str) -> None:
        if not self.form_error_text:
            return
        self.form_error_text.value = message
        self.form_error_text.visible = bool(message)

    def _handle_form_continue_click(self, e: Optional[ft.ControlEvent]) -> None:
        if self._form_continue_button:
            self._form_continue_button.visible = False
        first_control = next(iter(self.form_controls.values()), None)
        if first_control:
            try:
                first_control.focus()
            except Exception:
                pass
        self._update_ui()

    def _handle_form_value_change(self, field_name: str) -> None:
        group_key = self._field_groups.get(field_name)
        if not group_key:
            return
        self._update_group_summary(group_key)

    def _update_all_group_summaries(self) -> None:
        for group_key in FORM_GROUP_ORDER:
            self._update_group_summary(group_key)

    def _update_group_summary(self, group_key: str) -> None:
        label = self._group_summary_labels.get(group_key)
        if not label:
            return
        summary_value = self._compute_group_summary(group_key)
        if group_key == "mode":
            label.value = f"現在: {summary_value}"
        else:
            label.value = summary_value
        try:
            label.update()
        except Exception:
            pass

    def _compute_group_summary(self, group_key: str) -> str:
        if group_key == "mode":
            return MODE_LABELS.get(self.mode, self.mode.value)
        if group_key == "references":
            total_urls = 0
            for name, assigned_group in self._field_groups.items():
                if assigned_group != "references":
                    continue
                control = self.form_controls.get(name)
                if not control:
                    continue
                lines = [
                    line.strip()
                    for line in (control.value or "").replace("\r\n", "\n").split("\n")
                    if line.strip()
                ]
                total_urls += len(lines)
            return f"{total_urls} 件のURL" if total_urls else "未登録"

        values: List[str] = []
        for definition in _iter_mode_field_definitions(self.mode):
            name = definition["name"]
            if self._field_groups.get(name) != group_key:
                continue
            control = self.form_controls.get(name)
            if not control:
                continue
            value = (control.value or "").strip()
            if not value:
                continue
            values.append(value)
        if not values:
            return "未入力"
        if group_key == "output" and len(values) > 1:
            return f"{values[0]} ほか{len(values) - 1} 件"
        if group_key == "options" and len(values) > 1:
            return f"{values[0]} 他 {len(values) - 1} 項目"
        return values[0]

    def _split_list_values(self, raw_text: str) -> List[str]:
        tokens: List[str] = []
        for chunk in raw_text.replace(",", "\n").splitlines():
            item = chunk.strip()
            if item:
                tokens.append(item)
        return tokens

    def _collect_form_payload(self) -> Tuple[Optional[Dict[str, Any]], Optional[str], Optional[Dict[str, Any]]]:
        definitions = _iter_mode_field_definitions(self.mode)
        arguments: Dict[str, Any] = {}
        errors: List[str] = []

        for field in definitions:
            name = field["name"]
            ctrl = self.form_controls.get(name)
            raw_value = ""
            if ctrl and isinstance(ctrl.value, str):
                raw_value = ctrl.value.strip()
            if not raw_value and field.get("default"):
                raw_value = str(field["default"])
                if ctrl:
                    ctrl.value = raw_value

            if field.get("required") and not raw_value:
                errors.append(f"{field['label']}を入力してください。")
                continue

            if not raw_value:
                continue

            field_type = field.get("type", "str")
            argument_key = field["argument"]

            if field_type == "int":
                try:
                    value = int(raw_value)
                except ValueError:
                    errors.append(f"{field['label']}は整数で入力してください。")
                    continue
                if field.get("min") and value < field["min"]:
                    errors.append(f"{field['label']}は{field['min']}以上で入力してください。")
                    continue
                arguments[argument_key] = value
            elif field_type == "list":
                items = self._split_list_values(raw_value)
                if field.get("required") and not items:
                    errors.append(f"{field['label']}を入力してください。")
                    continue
                if items:
                    arguments[argument_key] = items
            else:
                arguments[argument_key] = raw_value

        if errors:
            return None, "\n".join(errors), None

        tool_name = FORM_TOOL_NAMES.get(self.mode)
        if not tool_name:
            return None, "現在のモードで使用できるツールが見つかりません。", None

        if self.mode in {CopilotMode.TRANSLATION, CopilotMode.TRANSLATION_WITH_REFERENCES}:
            arguments.setdefault("target_language", "English")

        if tool_name == "translate_range_with_references":
            has_reference = any(arguments.get(key) for key in ("source_reference_urls", "target_reference_urls"))
            if not has_reference:
                return None, "参照URLを1件以上入力してください。", None

        payload: Dict[str, Any] = {
            "mode": self.mode.value,
            "tool_name": tool_name,
            "arguments": arguments,
        }
        if self.current_workbook_name:
            payload["workbook_name"] = self.current_workbook_name
        if self.current_sheet_name:
            payload["sheet_name"] = self.current_sheet_name

        return payload, None, arguments

    def _format_form_summary(self, arguments: Dict[str, Any]) -> str:
        mode_label = MODE_LABELS.get(self.mode, self.mode.value)
        lines = [f"繝輔か繝ｼ繝騾∽ｿ｡ ({mode_label})"]
        for field in _iter_mode_field_definitions(self.mode):
            key = field["argument"]
            value = arguments.get(key)
            if value is None or value == "":
                continue
            if isinstance(value, list):
                display_value = ", ".join(value)
            else:
                display_value = str(value)
            lines.append(f"- {field['label']}: {display_value}")
        return "\n".join(lines)

    def _submit_form(self, e: Optional[ft.ControlEvent]):
        if self.app_state not in {AppState.READY, AppState.ERROR}:
            return

        payload, error_message, arguments = self._collect_form_payload()
        if error_message:
            self._set_form_error(error_message)
            self._update_ui()
            return

        assert payload is not None and arguments is not None
        self._set_form_error("")
        summary_message = self._format_form_summary(arguments)
        metadata: Dict[str, Any] = {"mode": self.mode.value, "mode_label": MODE_LABELS.get(self.mode, self.mode.value)}
        if self.current_workbook_name:
            metadata["workbook"] = self.current_workbook_name
        if self.current_sheet_name:
            metadata["sheet"] = self.current_sheet_name
        if self._form_continue_button:
            self._form_continue_button.visible = False

        self._set_state(AppState.TASK_IN_PROGRESS)
        self._add_message("user", summary_message, metadata)
        self.request_queue.put(RequestMessage(RequestType.USER_INPUT, payload))
        self._update_ui()

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
            preferred_chat_height = 360
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
            preferred_chat_height = 420
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
            preferred_chat_height = 520
            context_alignment = ft.MainAxisAlignment.END
            main_column_spacing = 20
            list_spacing = 22

        if self._content_container:
            self._content_container.padding = content_padding

        for panel in (self._context_panel, self._chat_panel):
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
                    min_chat_height = 240
                elif layout_key == "cozy":
                    min_chat_height = 280
                else:
                    min_chat_height = 320
                if calculated <= 0:
                    chat_height = min_chat_height
                else:
                    max_chat_height = max(preferred_chat_height, calculated)
                    chat_height = max(min_chat_height, min(max_chat_height, calculated))
            else:
                chat_height = preferred_chat_height
            self._chat_panel.height = chat_height

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
                font_family=self._primary_font_family,
            )
            card_body = ft.Container(
                bgcolor=ft.Colors.with_opacity(0.1, palette["surface_variant"]),
                border_radius=12,
                padding=ft.Padding(16, 14, 16, 14),
                border=ft.border.all(1, ft.Colors.with_opacity(0.08, palette["outline_variant"])),
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

    def _build_mode_selection_control(self) -> ft.Container:
        palette = EXPRESSIVE_PALETTE
        mode_row = self._build_mode_cards()
        self.mode_card_row = mode_row
        instruction = ft.Text(
            "実行する処理を選択してください",
            size=13,
            color=palette["on_surface_variant"],
            font_family=self._hint_font_family,
        )

        return ft.Column(
            [instruction, mode_row],
            spacing=10,
            tight=True,
        )

    def _refresh_mode_cards(self):
        palette = EXPRESSIVE_PALETTE
        for mode_value, card in self._mode_card_map.items():
            is_selected = mode_value == self.mode.value
            card.border = ft.border.all(
                2 if is_selected else 1,
                ft.Colors.with_opacity(0.9, palette["primary"]) if is_selected else ft.Colors.with_opacity(0.1, palette["outline_variant"]),
            )
            card.bgcolor = (
                ft.Colors.with_opacity(0.18, palette["primary"])
                if is_selected
                else ft.Colors.with_opacity(0.1, palette["surface_variant"])
            )

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
        if self.mode_selector:
            self.mode_selector.value = self.mode.value
        self._refresh_mode_cards()
        self._refresh_form_panel()
        if self.request_queue:
            self.request_queue.put(RequestMessage(RequestType.UPDATE_CONTEXT, {"mode": self.mode.value}))
        self._update_ui()

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
        previous_state = self.app_state
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

        if self.form_controls:
            for control in self.form_controls.values():
                control.disabled = not can_interact
        if self._form_submit_button:
            self._form_submit_button.disabled = not can_interact
        if self._form_cancel_button:
            if is_task_in_progress:
                self._form_cancel_button.visible = True
                self._form_cancel_button.disabled = False
            elif is_stopping:
                self._form_cancel_button.visible = True
                self._form_cancel_button.disabled = True
            else:
                self._form_cancel_button.visible = False
                self._form_cancel_button.disabled = True
        if self._form_continue_button:
            if new_state in {AppState.TASK_IN_PROGRESS, AppState.STOPPING}:
                self._form_continue_button.visible = False
                self._form_continue_button.disabled = True
                self._task_recently_completed = False
            elif is_ready and previous_state in {AppState.TASK_IN_PROGRESS, AppState.STOPPING}:
                self._task_recently_completed = True
                self._form_continue_button.visible = True
                self._form_continue_button.disabled = False
            elif is_error:
                self._form_continue_button.visible = False
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

        icon_config = {
            AppState.READY: (ft.Icons.CHECK_CIRCLE_OUTLINE, status_palette["ready"]),
            AppState.TASK_IN_PROGRESS: (ft.Icons.AUTORENEW, status_palette["busy"]),
            AppState.STOPPING: (ft.Icons.PAUSE_CIRCLE_OUTLINE, status_palette["stopping"]),
            AppState.ERROR: (ft.Icons.ERROR_OUTLINE, status_palette["error"]),
            AppState.INITIALIZING: (ft.Icons.HOURGLASS_TOP, status_palette["base"]),
        }
        if self._status_icon:
            icon_name, icon_color = icon_config.get(new_state, (ft.Icons.CIRCLE, status_palette["base"]))
            self._status_icon.name = icon_name
            self._status_icon.color = icon_color

        if self.status_label:
            self.status_label.opacity = 1
            self.status_label.scale = 1
            if new_state is AppState.INITIALIZING:
                self._status_message_override = None
                self._status_color_override = None
                self.status_label.value = "初期化中..."
                self.status_label.color = status_palette["base"]
            elif is_ready:
                if self._status_message_override:
                    self.status_label.value = self._status_message_override
                    self.status_label.color = self._status_color_override or status_palette["ready"]
                else:
                    self.status_label.value = "待機中"
                    self.status_label.color = status_palette["ready"]
            elif is_error:
                self._status_message_override = None
                self._status_color_override = None
                self.status_label.value = "エラー"
                self.status_label.color = status_palette["error"]
            elif is_task_in_progress:
                self._status_message_override = None
                self._status_color_override = None
                self.status_label.value = "処理を実行中..."
                self.status_label.color = status_palette["busy"]
                self.status_label.opacity = 0.5
                self.status_label.scale = 0.95
            elif is_stopping:
                self._status_message_override = None
                self._status_color_override = None
                self.status_label.value = "処理を停止しています..."
                self.status_label.color = status_palette["stopping"]
            else:
                self._status_message_override = None
                self._status_color_override = None
                self.status_label.color = status_palette["base"]

        self._update_form_progress_message()
        self._update_ui()

    def _update_form_progress_message(self, message: Optional[str] = None) -> None:
        if not self._form_progress_indicator or not self._form_progress_text:
            return
        if self.app_state not in {AppState.TASK_IN_PROGRESS, AppState.STOPPING}:
            self._form_progress_indicator.visible = False
            self._form_progress_text.visible = False
            self._form_progress_text.value = ""
            return
        progress_text = message or (self.status_label.value if self.status_label else "処理を実行中...")
        self._form_progress_indicator.visible = True
        self._form_progress_text.visible = True
        self._form_progress_text.value = progress_text
        try:
            self._form_progress_text.update()
        except Exception:
            pass

    def _update_context_summary(self) -> None:
        workbook = self.current_workbook_name or "未選択"
        sheet = self.current_sheet_name or "未選択"
        if self._context_summary_text:
            self._context_summary_text.value = f"選択中: {workbook} / {sheet}"
            try:
                self._context_summary_text.update()
            except Exception:
                pass
        if self._chat_header_subtitle:
            self._chat_header_subtitle.value = f"{workbook} / {sheet} のアクティビティを表示中"
            try:
                self._chat_header_subtitle.update()
            except Exception:
                pass
        if self._sync_status_text:
            if self._last_context_refresh_at:
                timestamp_str = self._last_context_refresh_at.strftime("%H:%M:%S")
                self._sync_status_text.value = f"最終同期: {timestamp_str}"
            else:
                self._sync_status_text.value = "最終同期: 未実行"
            try:
                self._sync_status_text.update()
            except Exception:
                pass

    def _update_ui(self):
        try:
            self.page.update()
        except Exception as e:
            print(f"UI\u306e\u66f4\u65b0\u306b\u5931\u6557\u3057\u307e\u3057\u305f: {e}")

    def _add_message(
        self,
        msg_type: Union[ResponseType, str],
        msg_content: str,
        metadata: Optional[Dict[str, Any]] = None,
    ) -> None:
        if not msg_content and not metadata:
            return

        timestamp = datetime.now()
        msg_type_value = msg_type.value if isinstance(msg_type, ResponseType) else str(msg_type)
        metadata_payload: Dict[str, Any] = dict(metadata or {})
        metadata_payload.setdefault("timestamp", timestamp.isoformat(timespec="seconds"))
        metadata_payload.setdefault("display_time", timestamp.strftime("%H:%M"))

        self._append_history(msg_type_value, msg_content, metadata_payload)
        self._update_save_button_state()

        should_display = self._should_display_message_type(msg_type_value)
        if not should_display or not self.chat_list:
            self._update_chat_empty_state()
            return

        msg = ChatMessage(msg_type, msg_content, metadata=metadata_payload)
        self.chat_list.controls.append(msg)
        self._update_ui()
        time.sleep(0.01)
        msg.appear()
        self._update_chat_empty_state()

    def _append_history(self, msg_type: str, msg_content: str, metadata: Optional[Dict[str, Any]] = None):
        entry = {
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "type": msg_type,
            "content": (msg_content or "").replace("\r\n", "\n"),
            "metadata": dict(metadata or {}),
        }
        with self.history_lock:
            self.chat_history.append(entry)

    def _update_save_button_state(self):
        if not self.save_log_button:
            return
        with self.history_lock:
            has_history = bool(self.chat_history)
        self.save_log_button.disabled = not has_history
        self._update_chat_empty_state()

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

    def _on_chat_filter_change(self, e: Optional[ft.ControlEvent]):
        selected = (e.control.value if e and e.control else None) or "all"
        self._chat_filter_value = selected
        self._refresh_chat_view_from_history()

    def _should_display_message_type(self, msg_type: str) -> bool:
        filter_value = self._chat_filter_value or "all"
        if filter_value == "all":
            return True
        ai_categories = {"final_answer", "observation", "thought", "action"}
        system_categories = {"info", "status", "error"}
        if filter_value == "user":
            return msg_type == "user"
        if filter_value == "ai":
            return msg_type in ai_categories
        if filter_value == "system":
            return msg_type in system_categories
        return True

    def _refresh_chat_view_from_history(self):
        if not self.chat_list:
            return
        self.chat_list.controls.clear()
        with self.history_lock:
            entries = list(self.chat_history)
        for entry in entries:
            if not self._should_display_message_type(entry["type"]):
                continue
            metadata = entry.get("metadata", {})
            chat_msg = ChatMessage(entry["type"], entry["content"], metadata=metadata, animate=False)
            chat_msg.opacity = 1
            chat_msg.offset = ft.Offset(0, 0)
            self.chat_list.controls.append(chat_msg)
        self._update_chat_empty_state()
        try:
            self.chat_list.update()
        except Exception:
            pass

    def _update_chat_empty_state(self):
        has_visible_entries = False
        with self.history_lock:
            for entry in self.chat_history:
                if self._should_display_message_type(entry["type"]):
                    has_visible_entries = True
                    break
        if self._chat_empty_state:
            self._chat_empty_state.visible = not has_visible_entries
            try:
                self._chat_empty_state.update()
            except Exception:
                pass
        if self._chat_scroll_button:
            with self.history_lock:
                has_history = bool(self.chat_history)
            self._chat_scroll_button.disabled = not has_history
            try:
                self._chat_scroll_button.update()
            except Exception:
                pass

    def _scroll_chat_to_latest(self, e: Optional[ft.ControlEvent]):
        if not self.chat_list or not self.chat_list.controls:
            return
        try:
            self.chat_list.scroll_to(index=len(self.chat_list.controls) - 1, duration=300)
        except Exception:
            pass


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
        self._refresh_chat_view_from_history()

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
                        raise ExcelConnectionError("開いている Excel ブックが見つかりません。")

                    active_workbook: Optional[str] = None
                    active_sheet: Optional[str] = None

                    def _fetch_active_context() -> Tuple[Optional[str], Optional[str]]:
                        try:
                            context = manager.get_active_workbook_and_sheet()
                        except ExcelConnectionError as context_err:
                            print(f"Failed to obtain active workbook and sheet: {context_err}")
                            return None, None
                        return (
                            context.get("workbook_name"),
                            context.get("sheet_name"),
                        )

                    should_activate_target = (
                        target_workbook
                        and target_workbook in workbook_names
                        and not auto_triggered
                    )

                    if should_activate_target:
                        try:
                            active_workbook = manager.activate_workbook(target_workbook)
                        except ExcelConnectionError as activate_err:
                            print(f"Failed to activate requested workbook '{target_workbook}': {activate_err}")
                            self._add_message(
                                ResponseType.INFO,
                                f"ブック『{target_workbook}』をアクティブにできませんでした: {activate_err}",
                            )
                            active_workbook, active_sheet = _fetch_active_context()
                        else:
                            active_workbook, active_sheet = _fetch_active_context()
                    else:
                        active_workbook, active_sheet = _fetch_active_context()

                    if not active_workbook and workbook_names:
                        active_workbook = workbook_names[0]

                    try:
                        sheet_names = manager.list_sheet_names()
                    except ExcelConnectionError as sheet_err:
                        print(f"Failed to fetch sheet names: {sheet_err}")
                        sheet_names = []

                    preferred_sheet = self._load_last_sheet_preference(active_workbook)
                    if (
                        preferred_sheet
                        and preferred_sheet in sheet_names
                        and preferred_sheet != active_sheet
                        and not auto_triggered
                    ):
                        try:
                            active_sheet = manager.activate_sheet(preferred_sheet)
                        except ExcelConnectionError as activate_err:
                            print(f"保存済みシート '{preferred_sheet}' の復元に失敗しました: {activate_err}")
                            self._add_message(
                                ResponseType.INFO,
                                f"保存済みシート『{preferred_sheet}』を開けませんでした: {activate_err}",
                            )
                        else:
                            try:
                                sheet_names = manager.list_sheet_names()
                            except ExcelConnectionError as sheet_err:
                                print(f"Failed to fetch sheet names after activation: {sheet_err}")
                                sheet_names = []

                    if not active_sheet and sheet_names:
                        active_sheet = sheet_names[0]

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

                self._last_context_refresh_at = datetime.now()
                self._update_context_summary()
                return active_sheet

            except Exception as ex:
                error_message = f"Excel の状態更新中にエラーが発生しました: {ex}"
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
                self._last_context_refresh_at = None
                self._update_context_summary()
                if not auto_triggered and not is_initial_start:
                    self._add_message(ResponseType.ERROR, error_message, {"source": "excel_refresh"})
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
        self._update_context_summary()
        self._update_ui()

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
            self._update_form_progress_message(status_text or None)
        elif response.type is ResponseType.ERROR:
            if self.app_state in {AppState.TASK_IN_PROGRESS, AppState.STOPPING}:
                if self.status_label:
                    self.status_label.value = response.content or "\u51e6\u7406\u4e2d\u306b\u30a8\u30e9\u30fc\u304c\u767a\u751f\u3057\u307e\u3057\u305f"
                    self.status_label.color = status_palette["error"]
                    self.status_label.opacity = 0.9
                if response.content:
                    self._add_message(response.type, response.content, response.metadata)
                    self._update_form_progress_message(response.content)
                if self._auto_test_triggered:
                    print(f"AUTOTEST: error '{response.content}'", flush=True)
            else:
                self._set_state(AppState.ERROR)
                if response.content:
                    self._add_message(response.type, response.content, response.metadata)
                    self._update_form_progress_message(response.content)
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
                self._add_message(type_value, response.content, response.metadata)
        else:
            if response.content:
                self._add_message(type_value, response.content, response.metadata)

        if response.type is ResponseType.FINAL_ANSWER:
            if self._auto_test_triggered:
                final_text = (response.content or "").strip()
                self._auto_test_completed = True
                print(f"AUTOTEST: final answer '{final_text}'", flush=True)
                self._schedule_autotest_shutdown()
            elif response.content:
                print(f"Final answer received outside autotest: {(response.content or '').strip()}", flush=True)
        elif self._auto_test_enabled and response.type in {ResponseType.OBSERVATION, ResponseType.ACTION, ResponseType.THOUGHT}:
            snippet = (response.content or "").strip()
            if snippet:
                print(f"AUTOTEST: {response.type.value} '{snippet[:120]}'", flush=True)

        self._update_ui()

    def _schedule_autotest_shutdown(self) -> None:
        if self._auto_test_shutdown_scheduled:
            return
        self._auto_test_shutdown_scheduled = True

        def _shutdown():
            try:
                time.sleep(1.0)
            except Exception:
                pass
            print("AUTOTEST: shutting down after final answer", flush=True)
            try:
                self.page.window.prevent_close = False
            except Exception:
                pass
            self._force_exit(reason="autotest-final-answer")

        threading.Thread(target=_shutdown, daemon=True).start()

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

    def _load_autotest_override(self) -> Dict[str, Any]:
        prompt_text = (self._auto_test_prompt or "").strip()
        if not prompt_text:
            return {}
        try:
            parsed = json.loads(prompt_text)
        except json.JSONDecodeError:
            return {}
        return parsed if isinstance(parsed, dict) else {}

    def _build_autotest_form_values(self, override_payload: Optional[Dict[str, Any]]) -> Dict[str, str]:
        definitions = FORM_FIELD_DEFINITIONS.get(self.mode, [])
        flat_definitions = _flatten_field_definitions(definitions)
        defaults_by_mode: Dict[CopilotMode, Dict[str, Any]] = {
            CopilotMode.TRANSLATION_WITH_REFERENCES: {
                "cell_range": "A2:A20",
                "translation_output_range": "B2:D20",
                "target_language": "English",
                "source_reference_urls": [self._auto_test_source_url] if self._auto_test_source_url else [],
                "target_reference_urls": [self._auto_test_target_url] if self._auto_test_target_url else [],
            },
            CopilotMode.TRANSLATION: {
                "cell_range": "A2:A20",
                "translation_output_range": "B2:B20",
                "target_language": "English",
            },
            CopilotMode.REVIEW: {
                "source_range": "B2:B20",
                "translated_range": "C2:C20",
                "status_output_range": "D2:D20",
                "issue_output_range": "E2:E20",
                "highlight_output_range": "F2:F20",
            },
        }
        defaults = defaults_by_mode.get(self.mode, {})

        override_candidates: List[Dict[str, Any]] = []
        if isinstance(override_payload, dict):
            override_candidates.append(override_payload)
            arguments_override = override_payload.get("arguments")
            if isinstance(arguments_override, dict):
                override_candidates.insert(0, arguments_override)

        values: Dict[str, str] = {}
        for field in flat_definitions:
            name = field["name"]
            argument_key = field["argument"]
            field_type = field.get("type", "str")

            raw_value: Any = None
            for candidate in override_candidates:
                if name in candidate:
                    raw_value = candidate[name]
                    break
                if argument_key in candidate:
                    raw_value = candidate[argument_key]
                    break
            if raw_value is None and name in defaults:
                raw_value = defaults[name]
            if raw_value is None:
                continue

            if field_type == "list":
                if isinstance(raw_value, str):
                    text_value = raw_value
                elif isinstance(raw_value, list):
                    text_value = "\n".join(str(item) for item in raw_value if str(item).strip())
                else:
                    text_value = str(raw_value)
            else:
                text_value = str(raw_value)
            values[name] = text_value
        return values

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

        override_payload = self._load_autotest_override()

        desired_mode = CopilotMode.TRANSLATION_WITH_REFERENCES
        if isinstance(override_payload, dict):
            override_mode_value = override_payload.get("mode")
            if isinstance(override_mode_value, str):
                try:
                    desired_mode = CopilotMode(override_mode_value)
                except ValueError:
                    pass
            else:
                tool_name = override_payload.get("tool_name")
                if isinstance(tool_name, str):
                    for mode_candidate, tool_identifier in FORM_TOOL_NAMES.items():
                        if tool_name == tool_identifier:
                            desired_mode = mode_candidate
                            break

        if self.mode != desired_mode:
            self._set_mode(desired_mode)

        _select_option(self.workbook_selector, self._auto_test_workbook, "workbook_selection_updating")
        _select_option(self.sheet_selector, self._auto_test_sheet, "sheet_selection_updating")

        print(
            "AUTOTEST: dispatching form submission",
            {
                "workbook": self.workbook_selector.value if self.workbook_selector else None,
                "sheet": self.sheet_selector.value if self.sheet_selector else None,
                "mode": self.mode.value,
            },
            flush=True,
        )

        form_values = self._build_autotest_form_values(override_payload)
        if not form_values:
            print("AUTOTEST: no form values produced; skipping submission", flush=True)
            return

        for name, text_value in form_values.items():
            control = self.form_controls.get(name)
            if not control:
                continue
            control.value = text_value
            self._handle_form_value_change(name)
            try:
                control.update()
            except Exception:
                pass

        self._set_form_error("")
        self._update_ui()
        self._submit_form(None)
        print("AUTOTEST: form submitted", flush=True)

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
    app_kwargs["assets_dir"] = str(ASSETS_DIR)
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
