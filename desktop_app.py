# desktop_app.py

import argparse
import json
import logging
import math
import os
import queue
import random
import platform
import re
import sys
import threading
import time
from datetime import datetime
from itertools import zip_longest
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

if os.environ.get("PYTHONUNBUFFERED") != "1":
    os.environ["PYTHONUNBUFFERED"] = "1"
    os.execv(sys.executable, [sys.executable, "-u", *sys.argv])

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
from excel_copilot.tools.actions import ExcelActions
from excel_copilot.ui.chat import ChatMessage
from excel_copilot.ui.messages import (
    AppState,
    RequestMessage,
    RequestType,
    ResponseMessage,
    ResponseType,
)
from excel_copilot.ui.theme import (
    EXPRESSIVE_PALETTE,
    TYPE_SCALE,
    RUNWAY_NOISE_TOKEN,
    RUNWAY_PARTICLE_TOKEN,
    accent_glow_gradient,
    depth_shadow,
    elevated_surface_gradient,
    floating_shadow,
    glass_border,
    glass_surface,
    metallic_bloom_gradient,
    motion_token,
    primary_surface_gradient,
    prism_card_gradient,
)
from excel_copilot.ui.worker import CopilotWorker

HAS_MATERIAL_STATE = hasattr(ft, "MaterialState")


def _material_state_value(default_value: Any, hovered_value: Optional[Any] = None) -> Any:
    """Return state-aware style values when supported by the current Flet build."""
    if HAS_MATERIAL_STATE:
        values = {ft.MaterialState.DEFAULT: default_value}
        if hovered_value is not None:
            values[ft.MaterialState.HOVERED] = hovered_value
        return values
    return default_value

def _ensure_console_logging() -> None:
    """Ensure root logger always streams to console so exceptions are visible."""
    root_logger = logging.getLogger()
    if not root_logger.handlers:
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        )
    formatter = logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s")
    stream_handler_exists = False
    for handler in root_logger.handlers:
        if isinstance(handler, logging.StreamHandler):
            handler.setLevel(logging.INFO)
            if handler.formatter is None:
                handler.setFormatter(formatter)
            stream_handler_exists = True
            break
    if not stream_handler_exists:
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(formatter)
        root_logger.addHandler(console_handler)
    root_logger.setLevel(logging.INFO)


_ensure_console_logging()

FOCUS_WAIT_TIMEOUT_SECONDS = 15.0
PREFERENCE_LAST_WORKBOOK_KEY = "__last_workbook__"
PREFERENCE_FORM_VALUES_KEY = "__form_values__"
FORM_VALUE_SAVE_DEBOUNCE_SECONDS = 0.75
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

WORKER_INIT_TIMEOUT_DEFAULT = 45.0
EXCEL_CONNECT_TIMEOUT_DEFAULT = 5.0

MODE_LABELS = {
    CopilotMode.TRANSLATION: "翻訳（通常）",
    CopilotMode.TRANSLATION_WITH_REFERENCES: "翻訳（参照あり）",
    CopilotMode.REVIEW: "翻訳チェック",
}
MISSING_CONTEXT_ERROR_MESSAGE = "対象ブックとシートを選択してから送信してください。"

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
            "name": "review_output_range",
            "label": "出力範囲（ステータス／指摘／ハイライト／修正案）",
            "argument": "review_output_range",
            "required": True,
            "placeholder": "例: D2:G20",
            "group": "output",
        },
    ],
}

FORM_TOOL_NAMES: Dict[CopilotMode, str] = {
    CopilotMode.TRANSLATION: "translate_range_without_references",
    CopilotMode.TRANSLATION_WITH_REFERENCES: "translate_range_with_references",
    CopilotMode.REVIEW: "check_translation_quality",
}
FORM_GROUP_LABELS: Dict[str, str] = {
    "mode": "モード",
    "scope": "対象範囲",
    "output": "出力設定",
    "references": "参考資料",
    "options": "オプション",
}

FORM_GROUP_ORDER: List[str] = ["mode", "scope", "output", "options", "references"]


def _flatten_field_definitions(definitions: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    flat: List[Dict[str, Any]] = []
    for field in definitions:
        if field.get("control") == "section":
            flat.extend(field.get("children", []))
        else:
            flat.append(field)
    return flat


def _split_sheet_reference(range_ref: Optional[str], default_sheet: Optional[str]) -> Tuple[Optional[str], str]:
    """Split a range like "Sheet1!A2:B5" into (sheet, range) pairs."""

    if not isinstance(range_ref, str):
        return default_sheet, ""

    cleaned = range_ref.strip()
    if not cleaned:
        return default_sheet, ""

    if "!" not in cleaned:
        return default_sheet, cleaned

    sheet_part, cell_part = cleaned.split("!", 1)
    sheet_part = sheet_part.strip()
    if sheet_part.startswith("[") and "]" in sheet_part:
        sheet_part = sheet_part.split("]", 1)[1]
    if sheet_part.startswith("'") and sheet_part.endswith("'"):
        sheet_part = sheet_part[1:-1].replace("''", "'")
    elif sheet_part.startswith('"') and sheet_part.endswith('"'):
        sheet_part = sheet_part[1:-1]

    cell_part = cell_part.strip()
    if cell_part.startswith("'") and cell_part.endswith("'"):
        cell_part = cell_part[1:-1].replace("''", "'")
    elif cell_part.startswith('"') and cell_part.endswith('"'):
        cell_part = cell_part[1:-1]

    return (sheet_part or default_sheet), cell_part


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
        logging.info("CopilotApp.__init__ started")
        self.page = page
        self.request_queue: "queue.Queue[RequestMessage]" = queue.Queue()
        self.response_queue: "queue.Queue[ResponseMessage]" = queue.Queue()
        self.worker_thread: Optional[threading.Thread] = None
        self.queue_thread: Optional[threading.Thread] = None
        self.worker: Optional[CopilotWorker] = None
        timeout_env = os.getenv("COPILOT_WORKER_INIT_TIMEOUT")
        try:
            timeout_value = float(timeout_env) if timeout_env else WORKER_INIT_TIMEOUT_DEFAULT
        except (TypeError, ValueError):
            timeout_value = WORKER_INIT_TIMEOUT_DEFAULT
        self._worker_init_timeout_seconds = max(5.0, timeout_value)
        excel_timeout_env = os.getenv("COPILOT_EXCEL_CONNECT_TIMEOUT")
        try:
            excel_timeout_value = float(excel_timeout_env) if excel_timeout_env else EXCEL_CONNECT_TIMEOUT_DEFAULT
        except (TypeError, ValueError):
            excel_timeout_value = EXCEL_CONNECT_TIMEOUT_DEFAULT
        self._excel_connect_timeout_seconds = max(1.0, excel_timeout_value)
        self._worker_started_at: Optional[float] = None
        self._worker_init_last_message_time: Optional[float] = None
        self._worker_init_timed_out = False
        self._excel_connection_failed = False
        self.app_state: Optional[AppState] = None
        self.ui_loop_running = True
        self._shutdown_lock = threading.Lock()
        self.shutdown_requested = False
        self.shutdown_finalized = False
        self.worker_shutdown_event = threading.Event()
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
        self._mode_segment_row: Optional[ft.Row] = None
        self.form_controls: Dict[str, ft.TextField] = {}
        self.form_error_text: Optional[ft.Text] = None
        self._form_submit_button: Optional[ft.Control] = None
        self._form_cancel_button: Optional[ft.Control] = None
        self._form_body_column: Optional[ft.Container] = None
        self._form_sections: Optional[ft.Control] = None
        self._group_status_indicators: Dict[str, Dict[str, ft.Control]] = {}
        self._form_progress_indicator: Optional[ft.ProgressRing] = None
        self._form_progress_text: Optional[ft.Text] = None
        self._process_timeline_step_refs: Dict[str, Dict[str, ft.Control]] = {}
        self._form_panel: Optional[ft.Container] = None
        self._mode_segment_map: dict[str, ft.Container] = {}
        self._runway_context_capsule: Optional[ft.Container] = None
        self._runway_context_host: Optional[ft.Container] = None
        self._drawer_context_host: Optional[ft.Container] = None
        self._context_capsule_parent: Optional[ft.Container] = None
        self._context_actions: Optional[ft.ResponsiveRow] = None
        self._chat_panel: Optional[ft.Container] = None
        self._chat_floating_shell: Optional[ft.Container] = None
        self._command_dock_container: Optional[ft.Container] = None
        self._timeline_shell: Optional[ft.Container] = None
        self._body_stack: Optional[ft.Stack] = None
        self._context_drawer: Optional[ft.AnimatedContainer] = None
        self._drawer_scrim: Optional[ft.Container] = None
        self._drawer_scrim_gesture: Optional[ft.GestureDetector] = None
        self._context_drawer_visible = False
        self._drawer_toggle_button: Optional[ft.FilledButton] = None
        self._hero_state_value: Optional[ft.Text] = None
        self._hero_mode_value: Optional[ft.Text] = None
        self._hero_workbook_value: Optional[ft.Text] = None
        self._hero_sheet_value: Optional[ft.Text] = None
        self._hero_stat_cards: Dict[str, Dict[str, Any]] = {}
        self._hero_metric_history: Dict[str, Any] = {}
        self._hero_completed_jobs = 0
        self._hero_rows_processed = 0
        self._hero_banner_container: Optional[ft.Container] = None
        self._hero_foreground_layer: Optional[ft.Container] = None
        self._hero_aurora_layer: Optional[ft.Container] = None
        self._hero_particle_layer: Optional[ft.Container] = None
        self._hero_context_pill_values: Dict[str, ft.Text] = {}
        self._hero_title_switcher: Optional[ft.AnimatedSwitcher] = None
        self._hero_title_variants: List[str] = [
            "CELL-TO-COSMOS RUNWAY",
            "REFERENCE HYPERDRIVE",
            "TRANSLATION CONTINUUM",
        ]
        self._hero_title_phrase_index = 0
        self._hero_title_value: str = ""
        self._hero_tagline_richtext: Optional[ft.RichText] = None
        self._hero_tagline_dynamic_span: Optional[ft.TextSpan] = None
        self._hero_parallax_offset = 0.0
        self._hero_breathing_timer: Optional[threading.Timer] = None
        self._hero_breathing_toggle = False
        self._hero_breathing_active = False
        self._command_palette_dialog: Optional[ft.AlertDialog] = None
        self._mode_panel_container: Optional[ft.Container] = None
        self._content_container: Optional[ft.Container] = None
        self._layout: Optional[ft.ResponsiveRow] = None
        self._main_column: Optional[ft.Column] = None
        self._chat_empty_state: Optional[ft.Container] = None
        self._chat_header_subtitle: Optional[ft.Text] = None

        self.chat_history: list[Dict[str, Any]] = []
        self.history_lock = threading.Lock()
        self.log_dir = Path(COPILOT_USER_DATA_DIR) / "setouchi_logs"
        self.preference_file = Path(COPILOT_USER_DATA_DIR) / "setouchi_state.json"
        self.preference_lock = threading.Lock()
        self._persisted_form_values: Dict[str, Dict[str, str]] = self._load_last_form_values()
        self._persisted_form_values.setdefault(self.mode.value, {})
        self._form_value_save_timer: Optional[threading.Timer] = None
        self._pending_form_seed: Optional[Dict[str, str]] = None

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
        self._latest_translation_job: Optional[Dict[str, Any]] = None
        self._latest_translation_summary_text: str = ""

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
        logging.info("CopilotApp.__init__ completed")

    def mount(self):
        logging.info("CopilotApp.mount invoked")
        self._set_state(AppState.INITIALIZING)
        self._update_ui()
        sheet_name = self._refresh_excel_context(is_initial_start=True, allow_fail=True)

        logging.info(
            "Starting Copilot worker thread (timeout %.1fs)...",
            self._worker_init_timeout_seconds,
        )
        self.worker = CopilotWorker(
            self.request_queue,
            self.response_queue,
            sheet_name,
            self.current_workbook_name,
        )
        self.worker_thread = threading.Thread(target=self.worker.run, daemon=True)
        self.worker_thread.start()
        self._worker_started_at = time.monotonic()
        self._worker_init_last_message_time = self._worker_started_at
        logging.info(
            "Copilot worker thread launched (alive=%s)",
            self.worker_thread.is_alive() if self.worker_thread else False,
        )

        logging.info("Starting response queue processing thread...")
        self.queue_thread = threading.Thread(target=self._process_response_queue_loop, daemon=True)
        self.queue_thread.start()
        logging.info(
            "Response queue thread launched (alive=%s)",
            self.queue_thread.is_alive() if self.queue_thread else False,
        )

        self.request_queue.put(RequestMessage(RequestType.UPDATE_CONTEXT, {"mode": self.mode.value}))

    def _configure_page(self):
        self.page.title = "Setouchi Excel Copilot"
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
        status_cluster = ft.Column(
            controls=[status_row, self._sync_status_text],
            spacing=4,
            tight=True,
        )

        button_shape = ft.RoundedRectangleBorder(radius=22)
        button_overlay = ft.Colors.with_opacity(0.08, palette["primary"])
        button_bg = _material_state_value(
            ft.Colors.with_opacity(0.14, palette["primary"]),
            ft.Colors.with_opacity(0.22, palette["primary"]),
        )

        self.new_chat_button = ft.FilledTonalButton(
            text="新しいチャット",
            icon=ft.Icons.CHAT_OUTLINED,
            on_click=self._handle_new_chat_click,
            disabled=True,
            style=ft.ButtonStyle(
                shape=button_shape,
                padding=ft.Padding(20, 12, 20, 12),
                bgcolor=button_bg,
                color=palette["on_primary"],
                overlay_color=button_overlay,
            ),
        )

        dropdown_style = {
            "border_radius": 22,
            "border_color": ft.Colors.with_opacity(0.24, palette["outline"]),
            "focused_border_color": palette["primary"],
            "fill_color": glass_surface(0.52),
            "text_style": ft.TextStyle(color=palette["on_surface"], size=12, font_family=self._primary_font_family),
            "hint_style": ft.TextStyle(color=palette["on_surface_variant"], size=12, font_family=self._primary_font_family),
            "disabled": True,
            "filled": True,
            "suffix_icon": ft.Icon(ft.Icons.KEYBOARD_ARROW_DOWN_ROUNDED, color=palette["on_surface_variant"]),
        }

        subtitle_scale = TYPE_SCALE["subtitle"]
        self._context_summary_text = ft.Text(
            "選択中: ブック未選択 / シート未選択",
            size=subtitle_scale["size"],
            weight=subtitle_scale["weight"],
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
                padding=ft.Padding(20, 12, 20, 12),
                bgcolor=button_bg,
                color=palette["on_primary"],
                overlay_color=button_overlay,
            ),
        )

        selector_labels = ft.Column(
            controls=[
                ft.Text("ブック", size=13, color=palette["on_surface_variant"], font_family=self._primary_font_family),
                self.workbook_selector_wrapper,
                ft.Text("シート", size=13, color=palette["on_surface_variant"], font_family=self._primary_font_family),
                self.sheet_selector_wrapper,
            ],
            spacing=10,
            tight=True,
        )

        selector_row = ft.ResponsiveRow(
            controls=[
                ft.Container(content=selector_labels, col={"xs": 12, "sm": 12, "md": 6}),
                ft.Container(
                    content=ft.Column(
                        [
                            ft.Text(
                                "現在の選択",
                                size=12,
                                color=palette["on_surface_variant"],
                                font_family=self._hint_font_family,
                            ),
                            self._context_summary_text,
                        ],
                        spacing=6,
                        tight=True,
                    ),
                    col={"xs": 12, "sm": 12, "md": 6},
                ),
            ],
            spacing=14,
            run_spacing=14,
            alignment=ft.MainAxisAlignment.START,
        )

        self._context_actions = ft.ResponsiveRow(
            controls=[ft.Container(content=self.workbook_refresh_button, col={"xs": 12, "sm": 6, "md": 4})],
            spacing=12,
            run_spacing=12,
            alignment=ft.MainAxisAlignment.END,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )

        self._runway_context_capsule = ft.Container(
            content=ft.Column(
                [
                    status_cluster,
                    ft.Container(height=1, bgcolor=ft.Colors.with_opacity(0.08, palette["outline"])),
                    selector_row,
                    self._context_actions,
                ],
                spacing=14,
                tight=True,
            ),
            bgcolor=glass_surface(0.82),
            gradient=elevated_surface_gradient(),
            border_radius=28,
            padding=ft.Padding(24, 22, 24, 24),
            border=glass_border(0.4),
            shadow=depth_shadow("md"),
        )

        self._runway_context_host = ft.Container(content=self._runway_context_capsule, expand=True)
        self._drawer_context_host = ft.Container(expand=True)
        self._context_capsule_parent = None
        self._mount_context_capsule(self._runway_context_host)

        caption_scale = TYPE_SCALE["caption"]
        self._chat_header_subtitle = ft.Text(
            "処理ログと結果が最新順に表示されます。",
            size=caption_scale["size"],
            weight=caption_scale["weight"],
            color=palette["on_surface_variant"],
            font_family=self._hint_font_family,
        )
        chat_header_section = ft.Column(
            controls=[
                ft.Row(
                    [
                        ft.Container(
                            content=self._chat_header_subtitle,
                            expand=True,
                            alignment=ft.alignment.center_left,
                        ),
                    ],
                    alignment=ft.MainAxisAlignment.START,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                ),
                ft.Container(height=1, bgcolor=ft.Colors.with_opacity(0.05, palette["outline"])),
            ],
            spacing=12,
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
            border=glass_border(0.24),
            bgcolor=glass_surface(0.58),
            visible=True,
            alignment=ft.alignment.center,
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
            bgcolor=glass_surface(0.95),
            gradient=elevated_surface_gradient(),
            border_radius=30,
            padding=ft.Padding(30, 34, 30, 34),
            border=glass_border(0.4),
            shadow=floating_shadow("lg"),
            clip_behavior=ft.ClipBehavior.HARD_EDGE,
            content=ft.Column(
                controls=[chat_header_section, self._chat_empty_state, self.chat_list],
                spacing=24,
                expand=True,
            ),
        )

        self._form_panel = self._build_form_panel()

        hero_title_scale = TYPE_SCALE["hero"]
        body_scale = TYPE_SCALE["body"]
        initial_metrics = self._collect_hero_metrics(suppress_delta=True)
        hero_stat_controls: List[ft.Control] = []
        for metric in initial_metrics:
            hero_stat_controls.append(
                ft.Container(
                    col={"xs": 12, "sm": 6, "md": 4, "lg": 4},
                    padding=ft.Padding(2, 6, 2, 6),
                    content=self._build_hero_stat_card(metric, body_scale, caption_scale),
                )
            )

        hero_badge = ft.Container(
            padding=ft.Padding(18, 8, 22, 8),
            border_radius=999,
            gradient=accent_glow_gradient(),
            border=ft.border.all(1, ft.Colors.with_opacity(0.32, palette["inverse_on_surface"])),
            content=ft.Row(
                [
                    ft.Icon(ft.Icons.AUTO_AWESOME_ROUNDED, size=18, color=palette["inverse_on_surface"]),
                    ft.Column(
                        [
                            ft.Text(
                                "AURORA RUNWAY",
                                size=TYPE_SCALE["eyebrow"]["size"],
                                weight=TYPE_SCALE["eyebrow"]["weight"],
                                color=palette["inverse_on_surface"],
                                font_family=self._primary_font_family,
                            ),
                            ft.Text(
                                "世界をとるフライトプラン",
                                size=caption_scale["size"],
                                color=ft.Colors.with_opacity(0.9, palette["inverse_on_surface"]),
                                font_family=self._hint_font_family,
                            ),
                        ],
                        spacing=2,
                        tight=True,
                    ),
                ],
                alignment=ft.MainAxisAlignment.CENTER,
                spacing=12,
            ),
        )
        hero_badge_wrapper = ft.Stack(
            controls=[
                ft.Container(
                    width=320,
                    height=60,
                    border_radius=999,
                    gradient=accent_glow_gradient(),
                    opacity=0.32,
                    shadow=floating_shadow("sm"),
                ),
                ft.Container(content=hero_badge, alignment=ft.alignment.center),
            ],
            height=60,
        )

        hero_context_pills, hero_pill_map = self._build_hero_context_pills()
        self._hero_context_pill_values = hero_pill_map

        drawer_button_style = ft.ButtonStyle(
            shape=ft.RoundedRectangleBorder(radius=26),
            padding=ft.Padding(24, 12, 24, 12),
            bgcolor=_material_state_value(
                ft.Colors.with_opacity(0.22, palette["inverse_on_surface"]),
                ft.Colors.with_opacity(0.32, palette["inverse_on_surface"]),
            ),
            color=palette["inverse_on_surface"],
            overlay_color=ft.Colors.with_opacity(0.12, palette["inverse_on_surface"]),
        )
        self._drawer_toggle_button = ft.FilledButton(
            text="コンテキストを開く",
            icon=ft.Icons.TUNE,
            on_click=self._toggle_context_drawer,
            style=drawer_button_style,
        )

        command_palette_button = self._build_command_palette_button()
        hero_actions = ft.ResponsiveRow(
            controls=[
                ft.Container(content=self._drawer_toggle_button, col={"xs": 12, "sm": 6, "md": 4}),
                ft.Container(content=self.new_chat_button, col={"xs": 12, "sm": 6, "md": 4}),
                ft.Container(content=command_palette_button, col={"xs": 12, "sm": 12, "md": 4}),
            ],
            spacing=12,
            run_spacing=12,
            alignment=ft.MainAxisAlignment.START,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )

        primary_title = self._compose_hero_title_value()
        hero_title_control = ft.Text(
            primary_title,
            size=hero_title_scale["size"],
            weight=hero_title_scale["weight"],
            color=palette["inverse_on_surface"],
            font_family=self._primary_font_family,
        )
        self._hero_title_switcher = ft.AnimatedSwitcher(
            hero_title_control,
            transition=ft.AnimatedSwitcherTransition.SCALE,
            duration=350,
        )
        self._hero_title_value = primary_title

        tagline_intro = ft.TextSpan(
            "セルから銀河まで翻訳の軌道をつなぐ、",
            style=ft.TextStyle(
                size=body_scale["size"],
                color=ft.Colors.with_opacity(0.88, palette["inverse_on_surface"]),
                font_family=self._hint_font_family,
            ),
        )
        tagline_accent = ft.TextSpan(
            "Setouchi Runway",
            style=ft.TextStyle(
                size=body_scale["size"],
                weight=ft.FontWeight.W_600,
                color=palette["inverse_on_surface"],
                font_family=self._primary_font_family,
            ),
        )
        tagline_dynamic = ft.TextSpan(
            self._resolve_mode_tagline(),
            style=ft.TextStyle(
                size=body_scale["size"],
                color=ft.Colors.with_opacity(0.88, palette["inverse_on_surface"]),
                font_family=self._hint_font_family,
            ),
        )
        hero_tagline = ft.RichText(
            spans=[tagline_intro, tagline_accent, ft.TextSpan(" "), tagline_dynamic],
            width=620,
        )
        self._hero_tagline_richtext = hero_tagline
        self._hero_tagline_dynamic_span = tagline_dynamic

        hero_stat_row = ft.ResponsiveRow(hero_stat_controls, spacing=12, run_spacing=12)

        hero_foreground = ft.Container(
            padding=ft.Padding(32, 32, 32, 36),
            border_radius=40,
            gradient=elevated_surface_gradient(),
            border=glass_border(0.18),
            content=ft.Column(
                [
                    hero_badge_wrapper,
                    hero_context_pills,
                    self._hero_title_switcher,
                    hero_tagline,
                    self._runway_context_host,
                    hero_stat_row,
                    hero_actions,
                ],
                spacing=20,
                tight=True,
            ),
        )
        hero_foreground.offset = ft.Offset(0, 0)
        self._hero_foreground_layer = hero_foreground

        hero_background = ft.Container(
            gradient=metallic_bloom_gradient(),
            border_radius=40,
            expand=True,
        )
        noise_spec = RUNWAY_NOISE_TOKEN
        hero_noise_layer = ft.Container(
            border_radius=40,
            expand=True,
            bgcolor=ft.Colors.with_opacity(noise_spec.get("opacity", 0.18), palette["inverse_on_surface"]),
            blur=ft.Blur(noise_spec.get("blur", 48), noise_spec.get("blur", 48)),
            opacity=noise_spec.get("opacity", 0.18),
        )
        blend_mode = noise_spec.get("blend_mode")
        if blend_mode:
            hero_noise_layer.blend_mode = blend_mode
        hero_aurora_layer = self._build_hero_aurora_layer()
        hero_aurora_layer.offset = ft.Offset(0, 0)
        self._hero_aurora_layer = hero_aurora_layer
        hero_overlay = ft.Container(
            border_radius=40,
            gradient=ft.RadialGradient(
                center=ft.alignment.center_right,
                radius=1.15,
                colors=[
                    ft.Colors.with_opacity(0.48, palette["secondary"]),
                    ft.Colors.with_opacity(0.05, palette["surface"]),
                ],
            ),
            expand=True,
            opacity=0.9,
        )
        particle_layer = self._build_hero_particle_layer()
        particle_layer.offset = ft.Offset(0, 0)
        self._hero_particle_layer = particle_layer

        hero_stack = ft.Stack(
            controls=[hero_background, hero_noise_layer, hero_aurora_layer, hero_overlay, particle_layer, hero_foreground],
            expand=True,
            clip_behavior=ft.ClipBehavior.ANTI_ALIAS,
        )

        hero_banner = ft.Container(
            content=hero_stack,
            border_radius=40,
            shadow=floating_shadow("lg"),
            clip_behavior=ft.ClipBehavior.ANTI_ALIAS,
        )
        hero_banner.animate_scale = motion_token("long")
        hero_banner.scale = ft.transform.Scale(1.0, 1.0, 1.0)
        self._hero_banner_container = hero_banner

        self._layout = ft.ResponsiveRow(
            controls=[
                ft.Container(
                    content=self._form_panel,
                    col={"xs": 12, "sm": 11, "md": 9, "lg": 8},
                    expand=True,
                )
            ],
            spacing=28,
            run_spacing=28,
            alignment=ft.MainAxisAlignment.CENTER,
            vertical_alignment=ft.CrossAxisAlignment.START,
            expand=True,
        )

        self._chat_floating_shell = ft.Container(
            content=ft.ResponsiveRow(
                controls=[
                    ft.Container(
                        content=self._chat_panel,
                        col={"xs": 12, "sm": 12, "md": 10, "lg": 8},
                        expand=True,
                    )
                ],
                spacing=0,
                run_spacing=0,
                alignment=ft.MainAxisAlignment.CENTER,
                vertical_alignment=ft.CrossAxisAlignment.START,
            ),
            expand=False,
            alignment=ft.alignment.bottom_center,
        )

        page_body = ft.Column(
            controls=[hero_banner, self._layout, self._chat_floating_shell],
            spacing=32,
            expand=True,
        )

        self._content_container = ft.Container(
            content=page_body,
            expand=True,
            padding=ft.Padding(32, 42, 32, 44),
            alignment=ft.alignment.top_center,
            bgcolor=palette["background"],
            gradient=ft.LinearGradient(
                begin=ft.alignment.top_center,
                end=ft.alignment.bottom_center,
                colors=[
                    ft.Colors.with_opacity(0.95, palette["surface_dim"]),
                    palette["background"],
                ],
            ),
        )

        scrim_color = ft.Colors.with_opacity(0.35, palette["on_surface"])
        self._drawer_scrim = ft.Container(
            expand=True,
            bgcolor=scrim_color,
            visible=False,
            opacity=0.0,
            animate_opacity=300,
        )
        self._drawer_scrim_gesture = ft.GestureDetector(
            content=self._drawer_scrim,
            on_tap=self._toggle_context_drawer,
            visible=False,
        )
        drawer_panel = ft.Container(
            bgcolor=glass_surface(0.9),
            gradient=elevated_surface_gradient(),
            border_radius=28,
            padding=ft.Padding(26, 30, 26, 30),
            border=glass_border(0.42),
            shadow=depth_shadow("lg"),
            content=self._drawer_context_host,
        )
        self._context_drawer = ft.AnimatedContainer(
            content=drawer_panel,
            width=420,
            alignment=ft.alignment.center_right,
            offset=ft.Offset(1.1, 0),
            animate_offset=350,
            visible=False,
        )
        drawer_wrapper = ft.Container(
            content=self._context_drawer,
            alignment=ft.alignment.center_right,
            expand=True,
            padding=ft.Padding(16, 32, 16, 32),
        )

        self._body_stack = ft.Stack(
            controls=[self._content_container, self._drawer_scrim_gesture, drawer_wrapper],
            expand=True,
        )

        self._update_context_summary()
        self._update_chat_empty_state()
        self._update_context_action_button()
        self._update_hero_overview()

        current_width = getattr(self.page, "width", None) or getattr(self.page.window, "width", None)
        current_height = getattr(self.page, "height", None) or getattr(self.page.window, "height", None)
        self._apply_responsive_layout(current_width, current_height)

        self.page.add(self._body_stack)

    def _build_form_panel(self) -> ft.Container:
        palette = EXPRESSIVE_PALETTE
        can_interact = self.app_state in {AppState.READY, AppState.ERROR}

        sections_control, controls_map = self._create_form_controls_for_mode(self.mode)
        self.form_controls = controls_map
        self._form_sections = sections_control
        self._form_body_column = ft.Container(
            content=sections_control,
            expand=True,
            padding=ft.Padding(4, 0, 4, 0),
        )

        self.form_error_text = ft.Text(
            "",
            color=palette["error"],
            size=12,
            visible=False,
            font_family=self._hint_font_family,
        )

        pill_shape = ft.StadiumBorder()
        submit_style = ft.ButtonStyle(
            shape=pill_shape,
            padding=ft.Padding(28, 12, 28, 12),
            bgcolor=_material_state_value(
                palette["primary"],
                ft.Colors.with_opacity(0.9, palette["primary"]),
            ),
            color=palette["on_primary"],
            overlay_color=ft.Colors.with_opacity(0.08, palette["on_primary"]),
        )
        cancel_style = ft.ButtonStyle(
            shape=pill_shape,
            padding=ft.Padding(24, 12, 24, 12),
            side=ft.BorderSide(1, ft.Colors.with_opacity(0.4, palette["outline"])),
            color=palette["on_surface"],
            overlay_color=ft.Colors.with_opacity(0.06, palette["on_surface"]),
        )

        self._form_submit_button = ft.FilledButton(
            "フォームを送信",
            icon=ft.Icons.CHECK_CIRCLE_OUTLINE,
            on_click=self._submit_form,
            disabled=not can_interact,
            style=submit_style,
        )

        self._form_cancel_button = ft.OutlinedButton(
            "停止",
            icon=ft.Icons.STOP_CIRCLE_OUTLINED,
            on_click=self._stop_task,
            disabled=True,
            visible=False,
            style=cancel_style,
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
            controls=[self._form_cancel_button, self._form_submit_button],
            alignment=ft.MainAxisAlignment.END,
            spacing=12,
        )
        action_bar = ft.Column(
            controls=[
                ft.Container(height=1, bgcolor=ft.Colors.with_opacity(0.06, palette["outline"])),
                ft.Row(
                    controls=[progress_cluster, action_buttons],
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                ),
            ],
            spacing=12,
            tight=True,
        )

        timeline_block = self._build_process_timeline()

        content = ft.Column(
            controls=[self._form_body_column, timeline_block, self.form_error_text, action_bar],
            spacing=24,
            tight=True,
        )

        panel = ft.Container(
            content=content,
            bgcolor=glass_surface(0.94),
            gradient=elevated_surface_gradient(),
            border_radius=30,
            padding=ft.Padding(26, 28, 26, 30),
            border=glass_border(0.4),
            shadow=depth_shadow("lg"),
            clip_behavior=ft.ClipBehavior.NONE,
        )

        self._mode_panel_container = panel
        self._command_dock_container = panel
        self._update_all_group_summaries()
        self._update_process_timeline_state()
        return panel

    def _create_form_controls_for_mode(
        self,
        mode: CopilotMode,
        initial_values: Optional[Dict[str, str]] = None,
    ) -> Tuple[ft.Control, Dict[str, ft.TextField]]:
        definitions = FORM_FIELD_DEFINITIONS.get(mode, [])
        seed_values = self._build_seed_form_values(mode, initial_values)
        palette = EXPRESSIVE_PALETTE

        grouped_controls: Dict[str, List[ft.Control]] = {group: [] for group in FORM_GROUP_ORDER}
        new_controls: Dict[str, ft.TextField] = {}
        self._field_groups = {}
        self._group_summary_labels = {}
        self._group_status_indicators = {}

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
            value = seed_values.get(name, definition.get("default", "")) or ""
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
            helper_text = definition.get("helper")
            if helper_text:
                text_field.helper_text = helper_text
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

        section_blocks: List[ft.Control] = []
        grid_cards: List[ft.Control] = []
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
            summary_chip = ft.Container(
                content=summary_label,
                padding=ft.Padding(12, 4, 12, 4),
                border_radius=999,
                bgcolor=ft.Colors.with_opacity(0.16, palette["surface_variant"]),
            )
            status_text = ft.Text(
                "Need Input",
                size=11,
                weight=ft.FontWeight.W_600,
                color=palette["warning"],
                font_family=self._primary_font_family,
            )
            status_chip = ft.Container(
                content=status_text,
                padding=ft.Padding(12, 4, 12, 4),
                border_radius=999,
                bgcolor=ft.Colors.with_opacity(0.16, palette["warning"]),
            )
            self._group_status_indicators[group_key] = {"label": status_text, "chip": status_chip}

            if group_key == "mode":
                summary_label.value = f"現在: {MODE_LABELS.get(self.mode, self.mode.value)}"
                section_blocks.append(
                    ft.Container(
                        content=ft.Column(
                            [
                                ft.Row(
                                    [
                                        ft.Text(
                                            "MODE",
                                            size=TYPE_SCALE["eyebrow"]["size"],
                                            weight=TYPE_SCALE["eyebrow"]["weight"],
                                            color=palette["primary"],
                                            font_family=self._primary_font_family,
                                        ),
                                        ft.Row(
                                            [status_chip, summary_chip],
                                            spacing=8,
                                            alignment=ft.MainAxisAlignment.END,
                                            vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                        ),
                                    ],
                                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                ),
                                ft.Column(controls_for_group, spacing=10, tight=True),
                            ],
                            spacing=10,
                            tight=True,
                        ),
                        border_radius=28,
                        padding=ft.Padding(22, 18, 22, 20),
                        bgcolor=glass_surface(0.92),
                        border=glass_border(0.32),
                        shadow=floating_shadow("sm"),
                    )
                )
            else:
                grid_cards.append(
                    ft.Container(
                        content=ft.Column(
                            [
                                ft.Row(
                                    [
                                        ft.Container(
                                            content=ft.Text(
                                                FORM_GROUP_LABELS.get(group_key, group_key.title()),
                                                size=15,
                                                weight=ft.FontWeight.W_600,
                                                color=palette["primary"],
                                                font_family=self._primary_font_family,
                                            ),
                                            expand=True,
                                        ),
                                        ft.Column(
                                            [status_chip, summary_chip],
                                            spacing=6,
                                            alignment=ft.MainAxisAlignment.END,
                                            horizontal_alignment=ft.CrossAxisAlignment.END,
                                        ),
                                    ],
                                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                ),
                                ft.Column(controls_for_group, spacing=10, tight=True),
                            ],
                            spacing=12,
                            tight=True,
                        ),
                        border_radius=26,
                        padding=ft.Padding(20, 20, 22, 22),
                        gradient=prism_card_gradient(),
                        border=glass_border(0.28),
                        shadow=depth_shadow("sm"),
                        col={"xs": 12, "md": 6},
                    )
                )

        if grid_cards:
            section_blocks.append(
                ft.ResponsiveRow(
                    controls=grid_cards,
                    spacing=18,
                    run_spacing=18,
                    alignment=ft.MainAxisAlignment.START,
                )
            )

        sections_control = ft.Column(section_blocks, spacing=18, tight=True)

        return sections_control, new_controls

    def _build_process_timeline(self) -> ft.Container:
        palette = EXPRESSIVE_PALETTE
        caption_scale = TYPE_SCALE["caption"]
        body_scale = TYPE_SCALE["body"]
        steps = [
            {
                "key": "context",
                "title": "Excel コンテキスト",
                "subtitle": "ブック/シート同期",
                "icon": ft.Icons.SYNC_ALT,
            },
            {
                "key": "copilot",
                "title": "Copilot 対話",
                "subtitle": "ブラウザオーケストレーション",
                "icon": ft.Icons.SMART_TOY,
            },
            {
                "key": "excel",
                "title": "Excel 反映",
                "subtitle": "セル書き込みと検証",
                "icon": ft.Icons.GRID_ON,
            },
        ]

        self._process_timeline_step_refs = {}
        step_rows: List[ft.Control] = []
        total_steps = len(steps)

        timeline_title = ft.Text(
            "Nebula Timeline",
            size=TYPE_SCALE["title"]["size"],
            weight=TYPE_SCALE["title"]["weight"],
            color=palette["primary"],
            font_family=self._primary_font_family,
        )
        timeline_caption = ft.Text(
            "翻訳航路の現在地を色と光で可視化します。",
            size=caption_scale["size"],
            color=palette["on_surface_variant"],
            font_family=self._hint_font_family,
        )

        for idx, spec in enumerate(steps):
            icon_control = ft.Icon(spec["icon"], size=20, color=palette["on_surface_variant"])
            halo = ft.Container(
                width=52,
                height=52,
                border_radius=26,
                alignment=ft.alignment.center,
                bgcolor=ft.Colors.with_opacity(0.18, palette["surface_variant"]),
                content=icon_control,
            )
            glow = ft.Container(
                width=82,
                height=82,
                border_radius=999,
                gradient=accent_glow_gradient(),
                opacity=0.35,
            )
            axis_stack = ft.Stack([glow, halo], width=82, height=82)
            connector = ft.Container(
                width=4,
                height=60,
                bgcolor=ft.Colors.with_opacity(0.12, palette["outline"]),
                visible=idx < total_steps - 1,
            )
            axis_column = ft.Column(
                controls=[axis_stack, connector],
                spacing=4,
                alignment=ft.MainAxisAlignment.START,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            )

            title_text = ft.Text(
                spec["title"],
                size=body_scale["size"],
                weight=ft.FontWeight.W_600,
                color=palette["on_surface"],
                font_family=self._primary_font_family,
            )
            subtitle_text = ft.Text(
                spec["subtitle"],
                size=caption_scale["size"],
                color=palette["on_surface_variant"],
                font_family=self._hint_font_family,
                max_lines=2,
                overflow=ft.TextOverflow.ELLIPSIS,
            )

            card = ft.Container(
                content=ft.Column(
                    [
                        title_text,
                        subtitle_text,
                    ],
                    spacing=6,
                    tight=True,
                ),
                padding=ft.Padding(22, 18, 24, 20),
                border_radius=28,
                gradient=prism_card_gradient(),
                border=glass_border(0.26),
                shadow=depth_shadow("micro"),
                expand=True,
            )

            step_row = ft.Row(
                [
                    ft.Container(content=axis_column, width=96),
                    card,
                ],
                spacing=18,
                vertical_alignment=ft.CrossAxisAlignment.START,
            )
            step_rows.append(step_row)

            self._process_timeline_step_refs[spec["key"]] = {
                "halo": halo,
                "glow": glow,
                "icon": icon_control,
                "title": title_text,
                "subtitle": subtitle_text,
                "card": card,
                "connector": connector if idx < total_steps - 1 else None,
            }

        column = ft.Column([timeline_title, timeline_caption, *step_rows], spacing=16, tight=True)
        shell = ft.Container(
            content=column,
            border_radius=32,
            padding=ft.Padding(26, 24, 26, 28),
            bgcolor=glass_surface(0.95),
            border=glass_border(0.32),
            shadow=depth_shadow("md"),
        )
        self._timeline_shell = shell
        return shell

    def _update_process_timeline_state(self) -> None:
        if not self._process_timeline_step_refs:
            return
        palette = EXPRESSIVE_PALETTE
        state = self.app_state or AppState.INITIALIZING
        if state is AppState.INITIALIZING:
            step_states = {"context": "active", "copilot": "idle", "excel": "idle"}
        elif state is AppState.READY:
            step_states = {"context": "complete", "copilot": "idle", "excel": "idle"}
        elif state is AppState.TASK_IN_PROGRESS:
            step_states = {"context": "complete", "copilot": "active", "excel": "idle"}
        elif state is AppState.STOPPING:
            step_states = {"context": "complete", "copilot": "complete", "excel": "active"}
        elif state is AppState.ERROR:
            step_states = {"context": "complete", "copilot": "error", "excel": "idle"}
        else:
            step_states = {"context": "idle", "copilot": "idle", "excel": "idle"}

        variants = {
            "idle": {
                "halo_bg": ft.Colors.with_opacity(0.18, palette["surface_variant"]),
                "halo_gradient": None,
                "icon_color": palette["on_surface_variant"],
                "title_color": palette["on_surface"],
                "subtitle_color": ft.Colors.with_opacity(0.85, palette["on_surface_variant"]),
                "glow_opacity": 0.3,
                "card_border_color": ft.Colors.with_opacity(0.22, palette["outline"]),
                "shadow_level": "micro",
                "connector_color": ft.Colors.with_opacity(0.12, palette["outline"]),
            },
            "active": {
                "halo_bg": ft.Colors.with_opacity(0.4, palette["primary"]),
                "halo_gradient": metallic_bloom_gradient(),
                "icon_color": palette["on_primary"],
                "title_color": palette["on_surface"],
                "subtitle_color": ft.Colors.with_opacity(0.95, palette["on_surface"]),
                "glow_opacity": 0.65,
                "card_border_color": ft.Colors.with_opacity(0.46, palette["primary"]),
                "shadow_level": "sm",
                "connector_gradient": ft.LinearGradient(
                    begin=ft.alignment.top_center,
                    end=ft.alignment.bottom_center,
                    colors=[palette["primary"], palette["secondary"]],
                ),
            },
            "complete": {
                "halo_bg": ft.Colors.with_opacity(0.32, palette["tertiary"]),
                "halo_gradient": metallic_bloom_gradient(True),
                "icon_color": palette["on_tertiary"],
                "title_color": palette["on_surface"],
                "subtitle_color": ft.Colors.with_opacity(0.95, palette["on_surface"]),
                "glow_opacity": 0.55,
                "card_border_color": ft.Colors.with_opacity(0.38, palette["tertiary"]),
                "shadow_level": "sm",
                "connector_gradient": ft.LinearGradient(
                    begin=ft.alignment.top_center,
                    end=ft.alignment.bottom_center,
                    colors=[palette["tertiary"], palette["primary"]],
                ),
            },
            "error": {
                "halo_bg": ft.Colors.with_opacity(0.3, palette["error"]),
                "halo_gradient": None,
                "icon_color": palette["on_error"],
                "title_color": palette["error"],
                "subtitle_color": palette["on_error"],
                "glow_opacity": 0.5,
                "card_border_color": ft.Colors.with_opacity(0.5, palette["error"]),
                "shadow_level": "md",
                "connector_gradient": ft.LinearGradient(
                    begin=ft.alignment.top_center,
                    end=ft.alignment.bottom_center,
                    colors=[palette["error"], ft.Colors.with_opacity(0.4, palette["error"])],
                ),
            },
        }

        for key, state_name in step_states.items():
            refs = self._process_timeline_step_refs.get(key)
            if not refs:
                continue
            variant = variants.get(state_name, variants["idle"])
            halo = refs.get("halo")
            glow = refs.get("glow")
            icon_chip = refs.get("icon")
            title_text = refs.get("title")
            subtitle_text = refs.get("subtitle")
            card = refs.get("card")
            connector = refs.get("connector")
            if isinstance(halo, ft.Container):
                halo.bgcolor = variant.get("halo_bg")
                halo.gradient = variant.get("halo_gradient")
                self._safe_update_control(halo)
            if isinstance(glow, ft.Container):
                glow.opacity = variant.get("glow_opacity", glow.opacity or 0.3)
                self._safe_update_control(glow)
            if isinstance(icon_chip, ft.Icon):
                icon_chip.color = variant.get("icon_color", icon_chip.color)
                self._safe_update_control(icon_chip)
            if title_text:
                title_text.color = variant["title_color"]
                try:
                    title_text.update()
                except Exception:
                    pass
            if subtitle_text:
                subtitle_text.color = variant["subtitle_color"]
                try:
                    subtitle_text.update()
                except Exception:
                    pass
            if isinstance(card, ft.Container):
                border_color = variant.get("card_border_color")
                if border_color:
                    card.border = ft.border.all(1, border_color)
                shadow_level = variant.get("shadow_level")
                if shadow_level:
                    card.shadow = depth_shadow(shadow_level)
                self._safe_update_control(card)
            if isinstance(connector, ft.Container):
                gradient = variant.get("connector_gradient")
                if gradient:
                    connector.gradient = gradient
                    connector.bgcolor = None
                else:
                    connector.gradient = None
                    connector.bgcolor = variant.get("connector_color")
                self._safe_update_control(connector)

    def _refresh_form_panel(self) -> None:
        if not self._form_body_column:
            return
        if self._pending_form_seed is not None:
            preserved_values = dict(self._pending_form_seed)
            self._pending_form_seed = None
        else:
            preserved_values: Dict[str, str] = {}
            for name, ctrl in self.form_controls.items():
                value = getattr(ctrl, "value", None)
                if isinstance(value, str):
                    preserved_values[name] = value
                elif value is not None:
                    preserved_values[name] = str(value)
        sections_control, controls_map = self._create_form_controls_for_mode(self.mode, preserved_values)
        self.form_controls = controls_map
        self._form_sections = sections_control
        if self._form_body_column:
            self._form_body_column.content = sections_control
        self._set_form_error("")
        self._update_submit_button_state()
        self._update_all_group_summaries()
        self._update_ui()

    def _set_form_error(self, message: str) -> None:
        if not self.form_error_text:
            return
        self.form_error_text.value = message
        self.form_error_text.visible = bool(message)

    def _handle_form_value_change(self, field_name: str) -> None:
        control = self.form_controls.get(field_name)
        mode_key = self.mode.value if isinstance(self.mode, CopilotMode) else str(self.mode)
        if control is not None:
            value = getattr(control, "value", "")
            if value is None:
                normalized_value = ""
            elif isinstance(value, str):
                normalized_value = value
            else:
                normalized_value = str(value)
            mode_store = self._persisted_form_values.setdefault(mode_key, {})
            mode_store[field_name] = normalized_value
            self._schedule_form_value_save()
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
        status_state = self._resolve_group_status_state(group_key, summary_value)
        self._update_group_status_chip(group_key, status_state)

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
        return values[0]

    def _resolve_group_status_state(self, group_key: str, summary_value: str) -> str:
        normalized = (summary_value or "").strip()
        if group_key == "mode":
            return "ready"
        if not normalized or normalized in {"未入力", "未登録"}:
            return "pending"
        if normalized.startswith("0 件"):
            return "pending"
        return "ready"

    def _update_group_status_chip(self, group_key: str, status_name: str) -> None:
        indicators = self._group_status_indicators.get(group_key)
        if not indicators:
            return
        palette = EXPRESSIVE_PALETTE
        styles = {
            "pending": {
                "label": "Need Input",
                "text_color": palette["warning"],
                "bgcolor": ft.Colors.with_opacity(0.18, palette["warning"]),
            },
            "ready": {
                "label": "Ready",
                "text_color": palette["success"],
                "bgcolor": ft.Colors.with_opacity(0.18, palette["success"]),
            },
        }
        style = styles.get(status_name, styles["pending"])
        status_label = indicators.get("label")
        if isinstance(status_label, ft.Text):
            status_label.value = style["label"]
            status_label.color = style["text_color"]
            self._safe_update_control(status_label)
        chip_control = indicators.get("chip")
        if isinstance(chip_control, ft.Container):
            chip_control.bgcolor = style["bgcolor"]
            self._safe_update_control(chip_control)

    def _split_list_values(self, raw_text: str) -> List[str]:
        tokens: List[str] = []
        for chunk in raw_text.replace(",", "\n").splitlines():
            item = chunk.strip()
            if item:
                tokens.append(item)
        return tokens

    @staticmethod
    def _split_sheet_and_range(value: str) -> Tuple[Optional[str], str]:
        if "!" not in value:
            return None, value
        sheet_name, inner = value.split("!", 1)
        return sheet_name, inner

    @staticmethod
    def _column_label_to_index(label: str) -> int:
        result = 0
        for ch in label:
            if not ("A" <= ch <= "Z"):
                raise ValueError(f"無効な列指定です: {label}")
            result = result * 26 + (ord(ch) - ord("A") + 1)
        return result

    @staticmethod
    def _column_index_to_label(index: int) -> str:
        if index <= 0:
            raise ValueError(f"列番号は1以上で指定してください（指定値: {index}）")
        label = ""
        while index > 0:
            index, remainder = divmod(index - 1, 26)
            label = chr(ord("A") + remainder) + label
        return label

    @staticmethod
    def _parse_cell_reference(cell: str) -> Tuple[str, int]:
        match = re.fullmatch(r"([A-Za-z]+)(\d+)", cell)
        if not match:
            raise ValueError(f"セル参照の形式が正しくありません: {cell}")
        column = match.group(1).upper()
        row = int(match.group(2))
        return column, row

    def _derive_review_output_ranges(self, combined_range: str) -> Dict[str, str]:
        if not combined_range:
            raise ValueError("出力範囲を入力してください。")

        sheet_name, inner_range = self._split_sheet_and_range(combined_range.strip())
        inner_range = inner_range.replace("$", "")
        if ":" not in inner_range:
            raise ValueError("出力範囲は開始セルと終了セルを「:」で区切って指定してください。")

        start_cell, end_cell = [part.strip() for part in inner_range.split(":", 1)]
        start_col_label, start_row = self._parse_cell_reference(start_cell)
        end_col_label, end_row = self._parse_cell_reference(end_cell)

        if end_row < start_row:
            raise ValueError("出力範囲の終了行は開始行以上になるように指定してください。")

        start_col_index = self._column_label_to_index(start_col_label)
        end_col_index = self._column_label_to_index(end_col_label)

        if end_col_index < start_col_index:
            raise ValueError("出力範囲の列指定が逆転しています。左端から右端に向かって指定してください。")

        column_count = end_col_index - start_col_index + 1
        if column_count < 3:
            raise ValueError("出力範囲には少なくとも3列（ステータス／指摘／ハイライト）が必要です。")
        if column_count > 4:
            raise ValueError("出力範囲は最大4列（ステータス／指摘／ハイライト／修正案）までにしてください。")

        prefix = f"{sheet_name}!" if sheet_name else ""

        def build_range(offset: int) -> str:
            column_label = self._column_index_to_label(start_col_index + offset)
            return f"{prefix}{column_label}{start_row}:{column_label}{end_row}"

        result: Dict[str, str] = {
            "status_output_range": build_range(0),
            "issue_output_range": build_range(1),
            "highlight_output_range": build_range(2),
        }
        if column_count >= 4:
            result["corrected_output_range"] = build_range(3)
        return result

    def _collect_form_payload(self) -> Tuple[Optional[Dict[str, Any]], Optional[str], Optional[Dict[str, Any]]]:
        definitions = _iter_mode_field_definitions(self.mode)
        arguments: Dict[str, Any] = {}
        summary_arguments: Dict[str, Any] = {}
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
                summary_arguments[argument_key] = value
            elif field_type == "float":
                try:
                    value = float(raw_value)
                except ValueError:
                    errors.append(f"{field['label']}は数値で入力してください。")
                    continue
                if math.isnan(value) or math.isinf(value):
                    errors.append(f"{field['label']}は有限の数値で入力してください。")
                    continue
                min_value = field.get("min")
                if min_value is not None and value < min_value:
                    errors.append(f"{field['label']}は{min_value}以上で入力してください。")
                    continue
                arguments[argument_key] = value
                summary_arguments[argument_key] = value
            elif field_type == "list":
                items = self._split_list_values(raw_value)
                if field.get("required") and not items:
                    errors.append(f"{field['label']}を入力してください。")
                    continue
                if items:
                    arguments[argument_key] = items
                    summary_arguments[argument_key] = items
            else:
                arguments[argument_key] = raw_value
                summary_arguments[argument_key] = raw_value

        if errors:
            return None, "\n".join(errors), None

        tool_name = FORM_TOOL_NAMES.get(self.mode)
        if not tool_name:
            logging.error("No tool mapping found for mode '%s'.", getattr(self.mode, "value", self.mode))
            return None, "現在のモードで使用できるツールが見つかりません。", None

        if self.mode in {CopilotMode.TRANSLATION, CopilotMode.TRANSLATION_WITH_REFERENCES}:
            arguments.setdefault("target_language", "English")

        if tool_name == "translate_range_with_references":
            has_reference = any(arguments.get(key) for key in ("source_reference_urls", "target_reference_urls"))
            if not has_reference:
                return None, "参照URLを1件以上入力してください。", None

        if self.mode is CopilotMode.REVIEW:
            combined_range = arguments.pop("review_output_range", None)
            try:
                derived_ranges = self._derive_review_output_ranges(combined_range or "")
            except ValueError as exc:
                return None, str(exc), None
            arguments.update(derived_ranges)

        payload: Dict[str, Any] = {
            "mode": self.mode.value,
            "tool_name": tool_name,
            "arguments": arguments,
        }
        if self.current_workbook_name:
            payload["workbook_name"] = self.current_workbook_name
        if self.current_sheet_name:
            payload["sheet_name"] = self.current_sheet_name

        return payload, None, summary_arguments

    def _format_form_summary(self, arguments: Dict[str, Any]) -> str:
        mode_label = MODE_LABELS.get(self.mode, self.mode.value)
        lines = [f"フォーム送信 ({mode_label})"]
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
        if not self._has_active_excel_context():
            self._set_form_error(MISSING_CONTEXT_ERROR_MESSAGE)
            self._update_ui()
            return

        payload, error_message, arguments = self._collect_form_payload()
        if error_message:
            self._set_form_error(error_message)
            self._update_ui()
            return

        assert payload is not None and arguments is not None
        tool_arguments = dict(payload.get("arguments", {}))
        job_context: Optional[Dict[str, Any]] = None
        if self.mode in {CopilotMode.TRANSLATION, CopilotMode.TRANSLATION_WITH_REFERENCES}:
            source_texts = self._capture_source_texts(tool_arguments)
            job_context = {
                "mode": self.mode,
                "workbook": self.current_workbook_name,
                "sheet": tool_arguments.get("sheet_name") or self.current_sheet_name,
                "source_range": tool_arguments.get("cell_range"),
                "output_range": tool_arguments.get("translation_output_range"),
                "overwrite_source": bool(tool_arguments.get("overwrite_source")),
                "source_texts": source_texts,
                "arguments": tool_arguments,
            }
            self._latest_translation_job = job_context
            self._latest_translation_summary_text = ""
        else:
            self._latest_translation_job = None
            self._latest_translation_summary_text = ""
        self._set_form_error("")
        summary_message = self._format_form_summary(arguments)
        metadata: Dict[str, Any] = {"mode": self.mode.value, "mode_label": MODE_LABELS.get(self.mode, self.mode.value)}
        if self.current_workbook_name:
            metadata["workbook"] = self.current_workbook_name
        if self.current_sheet_name:
            metadata["sheet"] = self.current_sheet_name
        self._set_state(AppState.TASK_IN_PROGRESS)
        if job_context is None:
            self._add_message("user", summary_message, metadata)
        self.request_queue.put(RequestMessage(RequestType.USER_INPUT, payload))
        self._flush_pending_form_value_save()
        self._update_ui()

    def _register_window_handlers(self):
        self.page.window.on_event = self._on_window_event
        self.page.on_resize = self._handle_page_resize
        self.page.on_disconnect = self._on_page_disconnect
        self.page.on_scroll = self._handle_page_scroll

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

        dock_panels = [self._chat_panel, self._form_panel]
        for panel in dock_panels:
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

        if self._mode_segment_row:
            mode_spacing = 8 if layout_key == "compact" else 14
            self._mode_segment_row.spacing = mode_spacing

        if self._context_actions:
            self._context_actions.alignment = context_alignment

        if layout_key == "compact":
            self._mount_context_capsule(self._drawer_context_host)
            if self._drawer_toggle_button:
                self._drawer_toggle_button.visible = True
                try:
                    self._drawer_toggle_button.update()
                except Exception:
                    pass
            if not self._context_drawer_visible:
                self._set_context_drawer_visibility(False)
        else:
            self._mount_context_capsule(self._runway_context_host)
            if self._drawer_toggle_button:
                self._drawer_toggle_button.visible = False
                try:
                    self._drawer_toggle_button.update()
                except Exception:
                    pass
            if self._context_drawer_visible:
                self._set_context_drawer_visibility(False)

        available_height = 0.0
        if height_value > 0:
            available_height = max(0.0, height_value - (content_padding.top + content_padding.bottom))

        if self._chat_panel:
            if available_height > 0:
                composer_est = (panel_padding.top + panel_padding.bottom) + 120
                mode_est = (mode_padding.top + mode_padding.bottom) + 110
                if self._main_column:
                    spacing_total = max(0, main_column_spacing) * 2
                else:
                    layout_spacing = getattr(self._layout, "spacing", 0) if self._layout else 0
                    spacing_total = max(0, layout_spacing)
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

    def _build_hero_aurora_layer(self) -> ft.Container:
        palette = EXPRESSIVE_PALETTE
        ribbon_specs = [
            {
                "left": -80,
                "top": 12,
                "width": 420,
                "height": 220,
                "angle": -10,
                "opacity": 0.55,
                "colors": [palette["tertiary"], ft.Colors.with_opacity(0.4, palette["secondary"]), ft.Colors.with_opacity(0.1, palette["primary"])],
                "blur": 95,
            },
            {
                "left": 180,
                "top": 40,
                "width": 360,
                "height": 180,
                "angle": 12,
                "opacity": 0.45,
                "colors": [ft.Colors.with_opacity(0.6, palette["secondary"]), ft.Colors.with_opacity(0.2, palette["primary"])],
                "blur": 85,
            },
            {
                "left": 60,
                "top": 140,
                "width": 520,
                "height": 200,
                "angle": -4,
                "opacity": 0.38,
                "colors": [ft.Colors.with_opacity(0.55, palette["primary_container"]), ft.Colors.with_opacity(0.18, palette["secondary_container"])],
                "blur": 75,
            },
        ]
        ribbons: List[ft.Control] = []
        for spec in ribbon_specs:
            ribbon = ft.Container(
                width=spec["width"],
                height=spec["height"],
                border_radius=360,
                gradient=ft.LinearGradient(
                    begin=ft.alignment.top_left,
                    end=ft.alignment.bottom_right,
                    colors=spec["colors"],
                ),
                opacity=spec["opacity"],
                blur=ft.Blur(sigma_x=spec.get("blur", 80), sigma_y=spec.get("blur", 80)),
            )
            angle = math.radians(spec.get("angle", 0))
            ribbon.rotate = ft.transform.Rotate(angle, alignment=ft.alignment.center)
            ribbons.append(
                ft.Positioned(
                    left=spec["left"],
                    top=spec["top"],
                    child=ribbon,
                )
            )
        return ft.Container(
            content=ft.Stack(controls=ribbons, expand=True),
            border_radius=40,
            expand=True,
            clip_behavior=ft.ClipBehavior.ANTI_ALIAS,
        )

    def _build_hero_particle_layer(self) -> ft.Container:
        palette = EXPRESSIVE_PALETTE
        token = RUNWAY_PARTICLE_TOKEN
        particle_count = int(token.get("count", 14))
        min_size = float(token.get("min_size", 4))
        max_size = float(token.get("max_size", 14))
        base_opacity = float(token.get("opacity", 0.8))
        rand = random.Random(0x5E70C1)
        width = 560
        height = 220

        def _particle_color() -> str:
            return rand.choice([
                palette["inverse_on_surface"],
                palette["tertiary"],
                palette["secondary"],
                palette["primary_container"],
            ])

        particles: List[ft.Control] = []
        for _ in range(particle_count):
            size = rand.uniform(min_size, max_size)
            opacity = rand.uniform(base_opacity * 0.35, base_opacity)
            hue = _particle_color()
            glow = ft.Container(
                width=size * 3.2,
                height=size * 3.2,
                border_radius=999,
                gradient=ft.RadialGradient(
                    center=ft.Alignment(0, 0),
                    radius=1.1,
                    colors=[ft.Colors.with_opacity(opacity * 0.6, hue), ft.Colors.with_opacity(0, hue)],
                ),
                blur=ft.Blur(size * 1.4, size * 1.6),
                opacity=opacity,
            )
            dot = ft.Container(
                width=size,
                height=size,
                border_radius=999,
                bgcolor=ft.Colors.with_opacity(opacity, hue),
            )
            particles.append(
                ft.Positioned(
                    left=rand.uniform(-40, width),
                    top=rand.uniform(0, height),
                    child=ft.Stack([glow, dot]),
                )
            )
        return ft.Container(
            content=ft.Stack(controls=particles, expand=True),
            border_radius=40,
            expand=True,
            padding=ft.Padding(0, 0, 0, 0),
        )

    def _build_hero_context_pills(self) -> Tuple[ft.ResponsiveRow, Dict[str, ft.Text]]:
        palette = EXPRESSIVE_PALETTE
        caption_scale = TYPE_SCALE["caption"]
        body_scale = TYPE_SCALE["body"]
        pill_specs = [
            {
                "key": "mode",
                "label": "モード",
                "icon": ft.Icons.AUTO_FIX_HIGH_ROUNDED,
                "colors": [ft.Colors.with_opacity(0.4, palette["secondary"]), ft.Colors.with_opacity(0.14, palette["surface_high"])],
            },
            {
                "key": "workbook",
                "label": "ブック",
                "icon": ft.Icons.INSERT_DRIVE_FILE_ROUNDED,
                "colors": [ft.Colors.with_opacity(0.38, palette["primary"]), ft.Colors.with_opacity(0.12, palette["surface_high"])],
            },
            {
                "key": "sheet",
                "label": "シート",
                "icon": ft.Icons.GRID_VIEW_ROUNDED,
                "colors": [ft.Colors.with_opacity(0.35, palette["tertiary"]), ft.Colors.with_opacity(0.12, palette["surface_high"])],
            },
        ]
        pill_values: Dict[str, ft.Text] = {}
        controls: List[ft.Control] = []

        def _default_value(key: str) -> str:
            if key == "mode":
                return MODE_LABELS.get(self.mode, self.mode.value)
            if key == "workbook":
                return self.current_workbook_name or "未選択"
            if key == "sheet":
                return self.current_sheet_name or "未選択"
            return ""

        for spec in pill_specs:
            value_text = ft.Text(
                _default_value(spec["key"]),
                size=body_scale["size"],
                weight=ft.FontWeight.W_600,
                color=palette["inverse_on_surface"],
                font_family=self._primary_font_family,
                no_wrap=True,
            )
            pill_values[spec["key"]] = value_text
            label_text = ft.Text(
                spec["label"],
                size=caption_scale["size"],
                color=ft.Colors.with_opacity(0.82, palette["inverse_on_surface"]),
                font_family=self._hint_font_family,
            )
            icon_container = ft.Container(
                width=28,
                height=28,
                alignment=ft.alignment.center,
                border_radius=14,
                bgcolor=ft.Colors.with_opacity(0.18, palette["inverse_on_surface"]),
                content=ft.Icon(spec["icon"], size=16, color=palette["inverse_on_surface"]),
            )
            pill = ft.Container(
                col={"xs": 12, "sm": 4, "md": 4, "lg": 4},
                padding=ft.Padding(18, 12, 18, 12),
                border_radius=28,
                bgcolor=glass_surface(0.48),
                gradient=ft.LinearGradient(
                    begin=ft.alignment.top_left,
                    end=ft.alignment.bottom_right,
                    colors=spec["colors"],
                ),
                border=glass_border(0.26),
                shadow=floating_shadow("sm"),
                content=ft.Column(
                    [
                        ft.Row([icon_container, label_text], spacing=8, alignment=ft.MainAxisAlignment.START, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                        value_text,
                    ],
                    spacing=6,
                    tight=True,
                ),
            )
            controls.append(pill)

        row = ft.ResponsiveRow(controls=controls, spacing=12, run_spacing=12)
        return row, pill_values

    def _update_hero_context_pills(self) -> None:
        if not self._hero_context_pill_values:
            return
        values = {
            "mode": MODE_LABELS.get(self.mode, self.mode.value),
            "workbook": self.current_workbook_name or "未選択",
            "sheet": self.current_sheet_name or "未選択",
        }
        for key, control in self._hero_context_pill_values.items():
            if not control:
                continue
            new_value = values.get(key)
            if new_value is None or control.value == new_value:
                continue
            control.value = new_value
            try:
                control.update()
            except Exception:
                pass

    def _build_command_palette_button(self) -> ft.Control:
        palette = EXPRESSIVE_PALETTE
        button = ft.FilledTonalButton(
            text="コマンドパレット",
            icon=ft.Icons.KEYBOARD_COMMAND_KEY,
            on_click=self._open_command_palette,
            tooltip="⌘K / Ctrl+K",
            style=ft.ButtonStyle(
                padding=ft.Padding(22, 12, 22, 12),
                bgcolor=_material_state_value(
                    ft.Colors.with_opacity(0.16, palette["inverse_on_surface"]),
                    ft.Colors.with_opacity(0.28, palette["inverse_on_surface"]),
                ),
                color=palette["inverse_on_surface"],
                overlay_color=ft.Colors.with_opacity(0.12, palette["inverse_on_surface"]),
            ),
        )
        self._command_palette_button = button
        return button

    def _open_command_palette(self, e: Optional[ft.ControlEvent] = None):
        palette = EXPRESSIVE_PALETTE
        shortcuts = [
            ("⌘ + Enter / Ctrl + Enter", "フォームを即時送信"),
            ("⌘ + R / Ctrl + R", "ブック情報を最新化"),
            ("⌘ + B / Ctrl + B", "コンテキストドロワー切替"),
            ("Esc", "処理の中断要求"),
        ]
        rows: List[ft.Control] = []
        for combo, desc in shortcuts:
            rows.append(
                ft.Row(
                    controls=[
                        ft.Text(
                            combo,
                            size=13,
                            weight=ft.FontWeight.W_600,
                            color=palette["on_surface"],
                            font_family=self._primary_font_family,
                        ),
                        ft.Text(
                            desc,
                            size=13,
                            color=palette["on_surface_variant"],
                            font_family=self._hint_font_family,
                        ),
                    ],
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                )
            )

        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("Command Palette", font_family=self._primary_font_family, size=18),
            content=ft.Column(rows, spacing=8, width=440, tight=True),
            actions=[ft.TextButton("閉じる", on_click=self._close_command_palette)],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        self._command_palette_dialog = dialog
        self.page.dialog = dialog
        dialog.open = True
        try:
            self.page.update()
        except Exception:
            pass

    def _close_command_palette(self, e: Optional[ft.ControlEvent] = None):
        if not self._command_palette_dialog:
            return
        self._command_palette_dialog.open = False
        try:
            self._command_palette_dialog.update()
        except Exception:
            pass
        self.page.dialog = None
        self._command_palette_dialog = None

    def _read_queue_size(self) -> int:
        if not self.request_queue:
            return 0
        try:
            return self.request_queue.qsize()
        except Exception:
            return 0

    def _collect_hero_metrics(self, suppress_delta: bool = False) -> List[Dict[str, Any]]:
        hero_state_label = {
            AppState.INITIALIZING: "初期化中",
            AppState.READY: "READY",
            AppState.TASK_IN_PROGRESS: "実行中",
            AppState.STOPPING: "停止要求中",
            AppState.ERROR: "エラー",
        }.get(self.app_state, "初期化中")
        workbook = self.current_workbook_name or "未選択"
        sheet = self.current_sheet_name or "未選択"
        queue_size = self._read_queue_size()
        metrics: List[Dict[str, Any]] = [
            {"id": "state", "label": "状態", "value": hero_state_label, "raw": hero_state_label},
            {"id": "mode", "label": "モード", "value": MODE_LABELS.get(self.mode, self.mode.value), "raw": self.mode.value},
            {"id": "workbook", "label": "ブック", "value": workbook, "raw": workbook},
            {"id": "sheet", "label": "シート", "value": sheet, "raw": sheet},
            {"id": "queue", "label": "待ちジョブ", "value": str(queue_size), "raw": queue_size, "unit": "件"},
            {
                "id": "jobs",
                "label": "完了ジョブ",
                "value": str(self._hero_completed_jobs),
                "raw": self._hero_completed_jobs,
                "unit": "件",
            },
            {
                "id": "rows",
                "label": "処理セル",
                "value": str(self._hero_rows_processed),
                "raw": self._hero_rows_processed,
                "unit": "行",
            },
        ]

        enriched: List[Dict[str, Any]] = []
        for metric in metrics:
            metric_id = metric["id"]
            value_raw = metric.get("raw", metric.get("value"))
            prev_value = self._hero_metric_history.get(metric_id)
            delta_text: Optional[str] = None
            trend_icon: Optional[str] = None
            if isinstance(value_raw, (int, float)) and not suppress_delta and prev_value is not None:
                diff = value_raw - prev_value
                if diff > 0:
                    delta_text = f"+{diff}"
                    trend_icon = ft.Icons.TRENDING_UP
                elif diff < 0:
                    delta_text = str(diff)
                    trend_icon = ft.Icons.TRENDING_DOWN
                else:
                    delta_text = "安定"
                    trend_icon = ft.Icons.DRAG_HANDLE
            metric["delta_text"] = delta_text
            metric["trend_icon"] = trend_icon
            enriched.append(metric)
            self._hero_metric_history[metric_id] = value_raw
        return enriched

    def _advance_hero_title_phrase(self) -> None:
        if not self._hero_title_variants:
            return
        self._hero_title_phrase_index = (self._hero_title_phrase_index + 1) % len(self._hero_title_variants)

    def _compose_hero_title_value(self) -> str:
        mode_label = MODE_LABELS.get(self.mode, self.mode.value)
        if not self._hero_title_variants:
            return f"Setouchi Excel Copilot · {mode_label}"
        phrase = self._hero_title_variants[self._hero_title_phrase_index]
        return f"{phrase} · {mode_label}"

    def _build_hero_value_text(self, value: Optional[str], body_scale: Dict[str, Any]) -> ft.Text:
        palette = EXPRESSIVE_PALETTE
        display_value = value if value is not None and str(value).strip() else "—"
        return ft.Text(
            display_value,
            size=body_scale["size"] + 2,
            weight=ft.FontWeight.W_600,
            color=palette["inverse_on_surface"],
            font_family=self._primary_font_family,
        )

    def _build_hero_stat_card(
        self,
        metric: Dict[str, Any],
        body_scale: Dict[str, Any],
        caption_scale: Dict[str, Any],
    ) -> ft.Control:
        palette = EXPRESSIVE_PALETTE
        metric_id = metric.get("id", "")
        value_text = self._build_hero_value_text(metric.get("value"), body_scale)
        value_switcher = ft.AnimatedSwitcher(value_text, duration=260)
        unit_label = ft.Text(
            metric.get("unit", ""),
            size=12,
            color=ft.Colors.with_opacity(0.78, palette["inverse_on_surface"]),
            visible=bool(metric.get("unit")),
            font_family=self._hint_font_family,
        )
        delta_text = ft.Text(
            metric.get("delta_text") or "",
            size=12,
            color=ft.Colors.with_opacity(0.9, palette["inverse_on_surface"]),
            visible=bool(metric.get("delta_text")),
            font_family=self._hint_font_family,
        )
        trend_icon = ft.Icon(
            metric.get("trend_icon") or ft.Icons.LENS,
            size=14,
            color=palette["inverse_on_surface"],
            visible=bool(metric.get("trend_icon")),
        )
        delta_row = ft.Row(
            controls=[trend_icon, delta_text],
            spacing=6,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
            visible=bool(metric.get("delta_text")),
        )
        card = ft.AnimatedContainer(
            content=ft.Column(
                [
                    ft.Row(
                        [
                            ft.Text(
                                metric.get("label", "-"),
                                size=caption_scale["size"],
                                weight=caption_scale["weight"],
                                color=ft.Colors.with_opacity(0.82, palette["inverse_on_surface"]),
                                font_family=self._hint_font_family,
                            ),
                            unit_label,
                        ],
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                    ),
                    value_switcher,
                    delta_row,
                ],
                spacing=10,
                tight=True,
            ),
            border_radius=26,
            gradient=ft.LinearGradient(
                begin=ft.alignment.top_left,
                end=ft.alignment.bottom_right,
                colors=[
                    ft.Colors.with_opacity(0.22, palette["inverse_on_surface"]),
                    ft.Colors.with_opacity(0.08, palette["surface_variant"]),
                ],
            ),
            border=glass_border(0.18),
            padding=ft.Padding(20, 18, 20, 18),
            shadow=floating_shadow("md"),
            on_hover=lambda e, key=metric_id: self._handle_hero_card_hover(key, e),
        )
        card.scale = ft.transform.Scale(1.0, 1.0, 1.0)
        self._hero_stat_cards[metric_id] = {
            "container": card,
            "value_switcher": value_switcher,
            "unit_label": unit_label,
            "delta_text": delta_text,
            "trend_icon": trend_icon,
            "delta_row": delta_row,
        }
        return card

    def _handle_hero_card_hover(self, metric_id: str, e: Optional[ft.ControlEvent]):
        card_state = self._hero_stat_cards.get(metric_id)
        if not card_state:
            return
        container = card_state.get("container")
        if not container:
            return
        is_hovered = str(getattr(e, "data", "")).lower() == "true"
        container.animate_scale = ft.animation.Animation(220, "easeOut")
        container.scale = ft.transform.Scale(1.02 if is_hovered else 1.0, 1.02 if is_hovered else 1.0, 1.0)
        self._safe_update_control(container)

    def _update_metric_card(self, metric: Dict[str, Any]):
        metric_id = metric.get("id")
        card_state = self._hero_stat_cards.get(metric_id)
        if not card_state:
            return
        value_switcher: Optional[ft.AnimatedSwitcher] = card_state.get("value_switcher")
        if value_switcher:
            new_value = self._build_hero_value_text(metric.get("value"), TYPE_SCALE["body"])
            value_switcher.content = new_value
            self._safe_update_control(value_switcher)
        unit_label: Optional[ft.Text] = card_state.get("unit_label")
        if unit_label:
            unit_value = metric.get("unit", "")
            unit_label.value = unit_value
            unit_label.visible = bool(unit_value)
            self._safe_update_control(unit_label)
        delta_text: Optional[ft.Text] = card_state.get("delta_text")
        trend_icon: Optional[ft.Icon] = card_state.get("trend_icon")
        delta_row: Optional[ft.Row] = card_state.get("delta_row")
        has_delta = bool(metric.get("delta_text"))
        if delta_text:
            delta_text.value = metric.get("delta_text") or ""
            delta_text.visible = has_delta
            self._safe_update_control(delta_text)
        if trend_icon:
            icon_name = metric.get("trend_icon") or ft.Icons.LENS
            trend_icon.name = icon_name
            trend_icon.visible = has_delta and metric.get("trend_icon") is not None
            self._safe_update_control(trend_icon)
        if delta_row:
            delta_row.visible = has_delta
            self._safe_update_control(delta_row)

    def _resolve_mode_tagline(self) -> str:
        taglines = {
            CopilotMode.TRANSLATION: "翻訳をエレガントに量産する集中モード。",
            CopilotMode.TRANSLATION_WITH_REFERENCES: "参照と調和しながら精度を極める演算。",
            CopilotMode.REVIEW: "ニュアンスチェックと差分検証を光速で。",
        }
        return taglines.get(self.mode, "静かな熱量でオペレーションを底上げします。")

    def _update_hero_title_text(self, new_value: str) -> None:
        if not self._hero_title_switcher or not new_value:
            return
        if new_value == self._hero_title_value:
            return
        hero_scale = TYPE_SCALE["hero"]
        updated_title = ft.Text(
            new_value,
            size=hero_scale["size"],
            weight=hero_scale["weight"],
            color=EXPRESSIVE_PALETTE["inverse_on_surface"],
            font_family=self._primary_font_family,
        )
        self._hero_title_switcher.content = updated_title
        self._hero_title_value = new_value
        self._safe_update_control(self._hero_title_switcher)

    def _update_hero_tagline(self) -> None:
        if not self._hero_tagline_dynamic_span or not self._hero_tagline_richtext:
            return
        new_text = self._resolve_mode_tagline()
        if self._hero_tagline_dynamic_span.text == new_text:
            return
        self._hero_tagline_dynamic_span.text = new_text
        self._safe_update_control(self._hero_tagline_richtext)

    def _handle_page_scroll(self, e: Optional[ft.ControlEvent]):
        if not e:
            return
        delta_value = 0.0
        raw = getattr(e, "data", None)
        if isinstance(raw, str) and raw.strip():
            try:
                payload = json.loads(raw)
                delta_value = float(payload.get("pixels_y") or payload.get("offset") or 0.0)
            except (json.JSONDecodeError, TypeError, ValueError):
                try:
                    delta_value = float(raw)
                except (TypeError, ValueError):
                    delta_value = 0.0
        elif isinstance(raw, (int, float)):
            delta_value = float(raw)
        self._update_hero_parallax(delta_value)

    def _update_hero_parallax(self, delta_y: float) -> None:
        try:
            delta_value = float(delta_y or 0.0)
        except (TypeError, ValueError):
            delta_value = 0.0
        normalized = max(-1.0, min(1.0, delta_value / 900.0))
        self._hero_parallax_offset = normalized
        if self._hero_foreground_layer:
            self._hero_foreground_layer.offset = ft.Offset(0, normalized * 0.6)
            self._safe_update_control(self._hero_foreground_layer)
        if self._hero_particle_layer:
            self._hero_particle_layer.offset = ft.Offset(0, normalized * 0.9)
            self._safe_update_control(self._hero_particle_layer)

    def _toggle_hero_breathing(self, enabled: bool) -> None:
        if enabled:
            if self._hero_breathing_active:
                return
            self._hero_breathing_active = True
            self._schedule_hero_breath()
            return
        self._hero_breathing_active = False
        self._stop_hero_breathing_timer()
        if self._hero_banner_container:
            self._hero_banner_container.scale = ft.transform.Scale(1.0, 1.0, 1.0)
            self._safe_update_control(self._hero_banner_container)

    def _schedule_hero_breath(self):
        if not self._hero_breathing_active:
            return
        self._hero_breathing_toggle = not self._hero_breathing_toggle
        target_scale = 1.015 if self._hero_breathing_toggle else 0.99
        if self._hero_banner_container:
            self._hero_banner_container.animate_scale = ft.animation.Animation(600, "easeInOut")
            self._hero_banner_container.scale = ft.transform.Scale(target_scale, target_scale, 1.0)
            self._safe_update_control(self._hero_banner_container)
        self._stop_hero_breathing_timer()
        self._hero_breathing_timer = threading.Timer(0.6, self._schedule_hero_breath)
        self._hero_breathing_timer.daemon = True
        self._hero_breathing_timer.start()

    def _stop_hero_breathing_timer(self):
        if self._hero_breathing_timer:
            self._hero_breathing_timer.cancel()
            self._hero_breathing_timer = None

    def _safe_update_control(self, control: Optional[ft.Control]) -> None:
        if not control:
            return
        try:
            control.update()
        except Exception:
            pass

    def _mount_context_capsule(self, host: Optional[ft.Container]) -> None:
        if not self._runway_context_capsule:
            return
        if host is self._context_capsule_parent:
            return
        if self._context_capsule_parent:
            self._context_capsule_parent.content = None
            try:
                self._context_capsule_parent.update()
            except Exception:
                pass
        self._context_capsule_parent = host
        if host:
            host.content = self._runway_context_capsule
            try:
                host.update()
            except Exception:
                pass
        for candidate in (self._runway_context_host, self._drawer_context_host):
            if not candidate:
                continue
            candidate.visible = candidate is host
            try:
                candidate.update()
            except Exception:
                pass

    def _toggle_context_drawer(self, e: Optional[ft.ControlEvent] = None):
        self._set_context_drawer_visibility(not self._context_drawer_visible)

    def _set_context_drawer_visibility(self, visible: bool) -> None:
        self._context_drawer_visible = visible
        if self._context_drawer:
            self._context_drawer.visible = visible
            self._context_drawer.offset = ft.Offset(0, 0) if visible else ft.Offset(1.1, 0)
            try:
                self._context_drawer.update()
            except Exception:
                pass
        if self._drawer_scrim:
            self._drawer_scrim.visible = visible
            self._drawer_scrim.opacity = 1.0 if visible else 0.0
            try:
                self._drawer_scrim.update()
            except Exception:
                pass
        if self._drawer_scrim_gesture:
            self._drawer_scrim_gesture.visible = visible
        self._update_context_action_button()
        self._update_ui()

    def _update_context_action_button(self) -> None:
        if not self._drawer_toggle_button:
            return
        if self._context_drawer_visible:
            self._drawer_toggle_button.text = "コンテキストを閉じる"
            self._drawer_toggle_button.icon = ft.Icons.CLOSE
        else:
            self._drawer_toggle_button.text = "コンテキストを開く"
            self._drawer_toggle_button.icon = ft.Icons.TUNE
        try:
            self._drawer_toggle_button.update()
        except Exception:
            pass

    def _update_hero_overview(self) -> None:
        metrics = self._collect_hero_metrics()
        for metric in metrics:
            self._update_metric_card(metric)
        self._advance_hero_title_phrase()
        self._update_hero_title_text(self._compose_hero_title_value())
        self._update_hero_tagline()
        self._update_hero_context_pills()

    def _build_mode_segmented_control(self) -> ft.Row:
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
        self._mode_segment_map = {}
        segments: List[ft.Control] = []
        for spec in options:
            mode = spec["mode"]
            segment = ft.Container(
                content=ft.Row(
                    [
                        ft.Icon(spec["icon"], size=16, color=palette["primary"]),
                        ft.Text(
                            spec["title"],
                            size=13,
                            weight=ft.FontWeight.W_600,
                            font_family=self._primary_font_family,
                            color=palette["on_surface"],
                        ),
                    ],
                    spacing=8,
                    alignment=ft.MainAxisAlignment.CENTER,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                ),
                padding=ft.Padding(18, 10, 18, 10),
                border_radius=999,
                bgcolor=glass_surface(0.6),
                border=ft.border.all(1, ft.Colors.with_opacity(0.18, palette["outline_variant"])),
                on_click=lambda e, value=mode: self._handle_mode_card_select(value),
                mouse_cursor=ft.MouseCursor.CLICK,
            )
            segments.append(segment)
            self._mode_segment_map[mode.value] = segment
        row = ft.Row(
            controls=segments,
            spacing=10,
            wrap=True,
            alignment=ft.MainAxisAlignment.CENTER,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )
        self._mode_segment_row = row
        self._refresh_mode_segments()
        return row

    def _build_mode_selection_control(self) -> ft.Container:
        palette = EXPRESSIVE_PALETTE
        mode_row = self._build_mode_segmented_control()
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

    def _refresh_mode_segments(self):
        palette = EXPRESSIVE_PALETTE
        for mode_value, segment in self._mode_segment_map.items():
            is_selected = mode_value == self.mode.value
            segment.bgcolor = (
                ft.Colors.with_opacity(0.28, palette["primary"]) if is_selected else glass_surface(0.6)
            )
            segment.border = ft.border.all(
                2 if is_selected else 1,
                ft.Colors.with_opacity(0.9, palette["primary"])
                if is_selected
                else ft.Colors.with_opacity(0.18, palette["outline_variant"]),
            )
            if isinstance(segment.content, ft.Row):
                for inner in segment.content.controls:
                    if isinstance(inner, ft.Text):
                        inner.color = palette["on_primary"] if is_selected else palette["on_surface"]
                    if isinstance(inner, ft.Icon):
                        inner.color = palette["on_primary"] if is_selected else palette["primary"]

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
        self._persisted_form_values.setdefault(new_mode.value, {})
        self._pending_form_seed = self._build_seed_form_values(new_mode)
        if self.mode_selector:
            self.mode_selector.value = self.mode.value
        self._refresh_mode_segments()
        self._refresh_form_panel()
        if self.request_queue:
            self.request_queue.put(RequestMessage(RequestType.UPDATE_CONTEXT, {"mode": self.mode.value}))
        self._update_hero_overview()
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

        if not can_interact and self._context_drawer_visible:
            self._set_context_drawer_visibility(False)

        if self.form_controls:
            for control in self.form_controls.values():
                control.disabled = not can_interact
        self._update_submit_button_state()
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
        if self._drawer_toggle_button:
            self._drawer_toggle_button.disabled = not can_interact
        if self._command_palette_button:
            self._command_palette_button.disabled = not can_interact

        if self.workbook_refresh_button:
            if self._manual_refresh_in_progress:
                self.workbook_refresh_button.disabled = True
            else:
                self.workbook_refresh_button.disabled = not can_interact
            if not self._manual_refresh_in_progress and can_interact:
                self.workbook_refresh_button.text = self._workbook_refresh_button_default_text
            try:
                self.workbook_refresh_button.update()
            except Exception:
                pass

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

        self._toggle_hero_breathing(is_task_in_progress)
        self._update_process_timeline_state()
        self._update_form_progress_message()
        self._update_hero_overview()
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
        self._update_hero_overview()

    def _update_ui(self):
        try:
            self.page.update()
        except Exception as e:
            print(f"UI\u306e\u66f4\u65b0\u306b\u5931\u6557\u3057\u307e\u3057\u305f: {e}")

    _CHAT_VISIBLE_TYPES = {
        "user",
        ResponseType.FINAL_ANSWER.value,
        ResponseType.ERROR.value,
    }

    def _add_message(
        self,
        msg_type: Union[ResponseType, str],
        msg_content: str,
        metadata: Optional[Dict[str, Any]] = None,
    ) -> None:
        msg_type_value = msg_type.value if isinstance(msg_type, ResponseType) else str(msg_type)
        if msg_type_value not in self._CHAT_VISIBLE_TYPES:
            return

        if not msg_content and not metadata:
            return

        timestamp = datetime.now()
        metadata_payload: Dict[str, Any] = dict(metadata or {})
        metadata_payload.setdefault("timestamp", timestamp.isoformat(timespec="seconds"))
        metadata_payload.setdefault("display_time", timestamp.strftime("%H:%M"))

        self._append_history(msg_type_value, msg_content, metadata_payload)
        self._update_save_button_state()

        if not self.chat_list:
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

    def _normalize_display_text(self, value: Any) -> str:
        if value is None:
            return ""
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value).strip()

    def _flatten_range_values(self, values: Any) -> List[str]:
        if values is None:
            return []
        flattened: List[Any] = []
        if isinstance(values, list):
            for row in values:
                if isinstance(row, list):
                    flattened.extend(row)
                else:
                    flattened.append(row)
        else:
            flattened.append(values)
        return [self._normalize_display_text(item) for item in flattened]

    def _read_range_matrix(self, workbook: Optional[str], sheet: Optional[str], cell_range: Optional[str]) -> List[List[Any]]:
        if not workbook or not cell_range:
            return []
        resolved_sheet = sheet
        resolved_range = cell_range.strip()
        if "!" in resolved_range:
            override_sheet, normalized = _split_sheet_reference(resolved_range, sheet)
            resolved_sheet = override_sheet
            resolved_range = normalized
        if not resolved_range:
            return []
        try:
            with ExcelManager(workbook) as manager:
                actions = ExcelActions(manager)
                raw_values = actions.read_range(resolved_range, resolved_sheet)
        except Exception as exc:
            print(f"Failed to read range '{cell_range}': {exc}")
            return []
        if raw_values is None:
            return []
        if isinstance(raw_values, list):
            if raw_values and isinstance(raw_values[0], list):
                return raw_values
            return [[item] for item in raw_values]
        return [[raw_values]]

    def _capture_source_texts(self, tool_arguments: Dict[str, Any]) -> List[str]:
        workbook = self.current_workbook_name
        sheet_name = tool_arguments.get("sheet_name") or self.current_sheet_name
        cell_range = tool_arguments.get("cell_range")
        matrix = self._read_range_matrix(workbook, sheet_name, cell_range)
        return self._flatten_range_values(matrix)

    def _capture_translated_texts(self, job: Dict[str, Any]) -> List[str]:
        workbook = job.get("workbook")
        sheet_name = job.get("sheet")
        target_range = job.get("output_range") or job.get("source_range")
        matrix = self._read_range_matrix(workbook, sheet_name, target_range)
        if not matrix:
            return []
        translations: List[str] = []
        for row in matrix:
            if isinstance(row, list) and row:
                translations.append(self._normalize_display_text(row[0]))
            else:
                translations.append(self._normalize_display_text(row))
        return translations

    def _parse_translation_texts(self, final_text: str) -> Optional[List[str]]:
        if not final_text:
            return None
        try:
            payload = json.loads(final_text)
        except json.JSONDecodeError:
            return None
        if not isinstance(payload, list):
            return None
        translations: List[str] = []
        for item in payload:
            if isinstance(item, dict):
                candidate = item.get("translated_text")
                if candidate is None:
                    continue
                translations.append(self._normalize_display_text(candidate))
            else:
                translations.append(self._normalize_display_text(item))
        return translations or None

    def _build_translation_pairs(self, final_text: str, job: Dict[str, Any]) -> List[Tuple[str, str]]:
        source_texts: List[str] = list(job.get("source_texts") or [])
        if not source_texts:
            arguments = job.get("arguments") or {}
            source_texts = self._capture_source_texts(arguments if isinstance(arguments, dict) else {})
            job["source_texts"] = source_texts
        translations = self._parse_translation_texts(final_text)
        if translations is None:
            translations = self._capture_translated_texts(job)
        if not source_texts and not translations:
            return []
        pairs: List[Tuple[str, str]] = []
        for src, tgt in zip_longest(source_texts, translations, fillvalue=""):
            src_text = src if isinstance(src, str) else self._normalize_display_text(src)
            tgt_text = tgt if isinstance(tgt, str) else self._normalize_display_text(tgt)
            pairs.append((src_text, tgt_text))
        return pairs

    def _display_translation_summary(
        self,
        pairs: List[Tuple[str, str]],
        fallback_text: str,
        job: Dict[str, Any],
        metadata: Optional[Dict[str, Any]],
    ) -> None:
        palette = EXPRESSIVE_PALETTE
        summary_lines: List[str] = []
        for index, (source_text, translated_text) in enumerate(pairs, start=1):
            if not source_text and not translated_text:
                continue
            summary_lines.append(f"{index}. 原文: {source_text}")
            summary_lines.append(f"   訳文: {translated_text}")
        fallback_clean = fallback_text.strip() if fallback_text else ""
        text_to_copy = "\n".join(summary_lines).strip() if summary_lines else fallback_clean
        self._latest_translation_summary_text = text_to_copy

        if self.chat_list:
            self.chat_list.controls.clear()

            header = ft.Text(
                "翻訳結果",
                size=14,
                weight=ft.FontWeight.W_600,
                color=palette["on_surface"],
                font_family=self._primary_font_family,
            )
            copy_button = ft.OutlinedButton(
                "全件コピー",
                icon=ft.Icons.CONTENT_COPY,
                on_click=lambda e, text=text_to_copy: self._copy_translation_summary(text),
                disabled=not text_to_copy,
            )
            header_row = ft.Row(
                controls=[header, copy_button],
                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
            )
            display_text = text_to_copy or "翻訳結果が確認できませんでした。"
            summary_body = ft.Text(
                display_text,
                size=13,
                color=palette["on_surface"],
                font_family=self._primary_font_family,
                selectable=True,
            )
            summary_container = ft.Container(
                content=summary_body,
                padding=ft.Padding(16, 14, 16, 14),
                border_radius=12,
                bgcolor=ft.Colors.with_opacity(0.08, palette["surface_variant"]),
                border=ft.border.all(1, ft.Colors.with_opacity(0.06, palette["outline_variant"])),
            )
            summary_column = ft.Column(
                controls=[header_row, summary_container],
                spacing=16,
                tight=True,
            )
            self.chat_list.controls.append(summary_column)
            self._update_chat_empty_state()
            try:
                self.chat_list.update()
            except Exception:
                pass

        with self.history_lock:
            self.chat_history.clear()
        history_metadata = {
            "mode": job.get("mode").value if isinstance(job.get("mode"), CopilotMode) else job.get("mode"),
            "workbook": job.get("workbook"),
            "sheet": job.get("sheet"),
            "source_range": job.get("source_range"),
            "translation_output_range": job.get("output_range"),
        }
        if metadata:
            history_metadata.update(metadata)
        history_content = text_to_copy or fallback_clean
        self._append_history("translation_summary", history_content, history_metadata)
        self._update_save_button_state()

    def _copy_translation_summary(self, text: str) -> None:
        if not text:
            return
        setter = getattr(self.page, "set_clipboard", None)
        if callable(setter):
            try:
                setter(text)
            except Exception as copy_err:
                print(f"Failed to copy translation summary: {copy_err}")
        if self.status_label:
            self.status_label.value = "翻訳結果をコピーしました。"
            self.status_label.color = EXPRESSIVE_PALETTE["primary"]
            try:
                self.status_label.update()
            except Exception:
                pass

    def _has_active_excel_context(self) -> bool:
        return bool(self.current_workbook_name) and bool(self.current_sheet_name)

    def _update_submit_button_state(self) -> None:
        if not self._form_submit_button:
            return
        can_submit = (
            self.app_state in {AppState.READY, AppState.ERROR}
            and self._has_active_excel_context()
        )
        self._form_submit_button.disabled = not can_submit
        try:
            self._form_submit_button.update()
        except Exception:
            pass

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

    def _refresh_chat_view_from_history(self):
        if not self.chat_list:
            return
        self.chat_list.controls.clear()
        with self.history_lock:
            entries = list(self.chat_history)
        for entry in entries:
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
        with self.history_lock:
            has_history = bool(self.chat_history)
        if self._chat_empty_state:
            self._chat_empty_state.visible = not has_history
            try:
                self._chat_empty_state.update()
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

    def _load_last_form_values(self) -> Dict[str, Dict[str, str]]:
        with self.preference_lock:
            if not self.preference_file.exists():
                return {}
            try:
                raw_text = self.preference_file.read_text(encoding="utf-8")
                data = json.loads(raw_text) if raw_text else {}
            except (OSError, json.JSONDecodeError) as err:
                print(f"Failed to load form preference: {err}")
                return {}
        stored = data.get(PREFERENCE_FORM_VALUES_KEY)
        if not isinstance(stored, dict):
            return {}
        result: Dict[str, Dict[str, str]] = {}
        for mode_key, values in stored.items():
            if not isinstance(mode_key, str) or not isinstance(values, dict):
                continue
            cleaned: Dict[str, str] = {}
            for field_name, field_value in values.items():
                if isinstance(field_name, str) and isinstance(field_value, str):
                    cleaned[field_name] = field_value
            if cleaned:
                result[mode_key] = cleaned
        return result

    def _save_last_form_values(self) -> None:
        with self.preference_lock:
            try:
                if self.preference_file.exists():
                    raw_text = self.preference_file.read_text(encoding="utf-8")
                    loaded = json.loads(raw_text) if raw_text else {}
                    preferences = dict(loaded) if isinstance(loaded, dict) else {}
                else:
                    preferences = {}
            except (OSError, json.JSONDecodeError) as err:
                print(f"Failed to read form preference: {err}")
                preferences = {}

            payload: Dict[str, Dict[str, str]] = {}
            for mode_key, values in self._persisted_form_values.items():
                if not isinstance(mode_key, str) or not isinstance(values, dict):
                    continue
                filtered = {name: value for name, value in values.items() if isinstance(name, str) and isinstance(value, str)}
                if filtered:
                    payload[mode_key] = filtered

            if payload:
                preferences[PREFERENCE_FORM_VALUES_KEY] = payload
            else:
                preferences.pop(PREFERENCE_FORM_VALUES_KEY, None)

            try:
                self.preference_file.parent.mkdir(parents=True, exist_ok=True)
                self.preference_file.write_text(
                    json.dumps(preferences, ensure_ascii=False, indent=2),
                    encoding="utf-8",
                )
            except OSError as err:
                print(f"Failed to write form preference: {err}")

    def _handle_form_value_save_timer(self) -> None:
        self._form_value_save_timer = None
        self._save_last_form_values()

    def _schedule_form_value_save(self) -> None:
        existing = self._form_value_save_timer
        if existing:
            self._form_value_save_timer = None
            try:
                existing.cancel()
            except Exception:
                pass
        timer = threading.Timer(
            FORM_VALUE_SAVE_DEBOUNCE_SECONDS,
            self._handle_form_value_save_timer,
        )
        timer.daemon = True
        self._form_value_save_timer = timer
        try:
            timer.start()
        except Exception:
            self._form_value_save_timer = None
            self._save_last_form_values()

    def _cancel_pending_form_value_save(self) -> None:
        timer = self._form_value_save_timer
        if not timer:
            return
        self._form_value_save_timer = None
        try:
            timer.cancel()
        except Exception:
            pass

    def _flush_pending_form_value_save(self) -> None:
        timer = self._form_value_save_timer
        if timer:
            self._form_value_save_timer = None
            try:
                timer.cancel()
            except Exception:
                pass
        self._save_last_form_values()

    def _build_seed_form_values(
        self,
        mode: CopilotMode,
        overrides: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, str]:
        base = dict(self._persisted_form_values.get(mode.value, {}))
        if overrides:
            for key, value in overrides.items():
                if not isinstance(key, str):
                    continue
                if isinstance(value, str):
                    base[key] = value
                elif value is not None:
                    base[key] = str(value)
        return base

    def _refresh_excel_context_once(
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
                self._update_submit_button_state()
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

                self.current_workbook_name = None
                self.current_sheet_name = None
                self._last_excel_snapshot = {}
                self._last_context_refresh_at = None
                self._update_context_summary()
                self._update_submit_button_state()
                if not auto_triggered and not is_initial_start:
                    self._add_message(ResponseType.ERROR, error_message, {"source": "excel_refresh"})
                self._update_ui()
                return None

    def _refresh_excel_context(
        self,
        is_initial_start: bool = False,
        desired_workbook: Optional[str] = None,
        auto_triggered: bool = False,
        allow_fail: bool = False,
    ) -> Optional[str]:
        try:
            result = self._refresh_excel_context_once(
                is_initial_start=is_initial_start,
                desired_workbook=desired_workbook,
                auto_triggered=auto_triggered,
            )
        except ExcelConnectionError as exc:
            if allow_fail:
                logging.warning("Excel context refresh failed: %s", exc)
                self._handle_excel_connection_failure()
                return None
            raise
        except Exception as exc:
            if allow_fail:
                logging.warning("Excel context refresh raised an unexpected error: %s", exc)
                self._handle_excel_connection_failure()
                return None
            raise

        if result is None:
            if allow_fail:
                logging.warning("Excel context refresh returned no active sheet.")
                self._handle_excel_connection_failure()
            return None

        self._clear_excel_connection_warning()
        return result

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
        logging.info("Response queue loop started.")
        while self.ui_loop_running:
            try:
                raw_message = self.response_queue.get(timeout=0.1)
            except queue.Empty:
                self._maybe_handle_worker_init_timeout()
                continue
            except Exception as e:
                print(f"\u30ec\u30b9\u30dd\u30f3\u30b9\u30ad\u30e5\u30fc\u51e6\u7406\u4e2d\u306b\u30a8\u30e9\u30fc\u304c\u767a\u751f\u3057\u307e\u3057\u305f: {e}")
                continue

            try:
                response = ResponseMessage.from_raw(raw_message)
            except ValueError as exc:
                print(f"\u30ec\u30b9\u30dd\u30f3\u30b9\u306e\u89e3\u6790\u306b\u5931\u6557\u3057\u307e\u3057\u305f: {exc}")
                continue

            self._worker_init_last_message_time = time.monotonic()
            self._display_response(response)

    def _display_response(self, response: ResponseMessage):
        type_value = response.metadata.get("source_type", response.type.value)
        status_palette = {
            "base": EXPRESSIVE_PALETTE["on_surface_variant"],
            "info": EXPRESSIVE_PALETTE["primary"],
            "error": EXPRESSIVE_PALETTE["error"],
        }

        if response.type is ResponseType.SHUTDOWN_COMPLETE:
            print("Shutdown: worker reported cleanup complete.")
            self.worker_shutdown_event.set()
            if self.shutdown_requested:
                self._finalize_shutdown(reason="worker-signal")
            return

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
            self._worker_init_timed_out = False
            self._worker_started_at = None
            self._worker_init_last_message_time = None
            self._set_state(AppState.READY)
            if self.status_label:
                self.status_label.value = response.content or self.status_label.value
                self._status_color_override = None
                self.status_label.color = EXPRESSIVE_PALETTE["primary"]
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
            self._latest_translation_job = None
            if response.content:
                print(f"ERROR: {response.content}", flush=True)
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
        elif response.type is ResponseType.FINAL_ANSWER:
            self._handle_final_answer(response)
        else:
            if response.content:
                self._add_message(type_value, response.content, response.metadata)

        if self._auto_test_enabled and response.type in {ResponseType.OBSERVATION, ResponseType.ACTION, ResponseType.THOUGHT}:
            snippet = (response.content or "").strip()
            if snippet:
                print(f"AUTOTEST: {response.type.value} '{snippet[:120]}'", flush=True)

        self._update_ui()

    def _handle_final_answer(self, response: ResponseMessage) -> None:
        final_text = (response.content or "").strip()
        if self._auto_test_triggered:
            self._latest_translation_job = None
            self._auto_test_completed = True
            print(f"AUTOTEST: final answer '{final_text}'", flush=True)
            self._schedule_autotest_shutdown()
            return
        if final_text:
            print(f"Final answer received outside autotest: {final_text}", flush=True)

        job = self._latest_translation_job
        mode = job.get("mode") if isinstance(job, dict) else None
        if isinstance(mode, CopilotMode):
            is_translation_mode = mode in {CopilotMode.TRANSLATION, CopilotMode.TRANSLATION_WITH_REFERENCES}
        else:
            is_translation_mode = str(mode) in {CopilotMode.TRANSLATION.value, CopilotMode.TRANSLATION_WITH_REFERENCES.value}

        if job and is_translation_mode:
            pairs = self._build_translation_pairs(final_text, job)
            processed_pairs = len(pairs)
            self._hero_completed_jobs += 1
            if processed_pairs:
                self._hero_rows_processed += processed_pairs
            self._update_hero_overview()
            self._display_translation_summary(pairs, final_text, job, response.metadata)
            self._latest_translation_job = None
        elif final_text:
            self._add_message(response.type, final_text, response.metadata)

    def _maybe_handle_worker_init_timeout(self) -> None:
        if self._worker_init_timed_out:
            return
        if self.app_state != AppState.INITIALIZING:
            return
        if self._worker_started_at is None:
            return
        last_event = self._worker_init_last_message_time or self._worker_started_at
        if time.monotonic() - last_event < self._worker_init_timeout_seconds:
            return
        self._worker_init_timed_out = True
        logging.error(
            "Copilot worker failed to finish initialization within %.1f seconds.",
            self._worker_init_timeout_seconds,
        )
        timeout_message = "Copilot 初期化がタイムアウトしました。Playwright の起動状況を確認してください。"
        if self.status_label:
            self.status_label.value = timeout_message
            self.status_label.color = EXPRESSIVE_PALETTE["error"]
        self._set_state(AppState.ERROR)
        self._update_ui()

    def _handle_excel_connection_failure(self) -> None:
        if not self._excel_connection_failed:
            logging.error("Excel connection unavailable; continuing without workbook context.")
        self._excel_connection_failed = True
        message = "Excel に接続できませんでした。"
        self._status_message_override = message
        self._status_color_override = EXPRESSIVE_PALETTE["error"]
        if self.status_label:
            self.status_label.value = message
            self.status_label.color = EXPRESSIVE_PALETTE["error"]
        self._update_ui()

    def _clear_excel_connection_warning(self) -> None:
        if not self._excel_connection_failed:
            return
        self._excel_connection_failed = False
        self._status_message_override = None
        self._status_color_override = None
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
        normalized_reason = reason or "unspecified"
        if self.shutdown_requested:
            print("Force exit: shutdown already in progress.")
            self._finalize_shutdown(reason=normalized_reason)
            return

        print(f"Force exit triggered. reason={normalized_reason}")
        self.shutdown_requested = True
        self._stop_background_excel_polling()

        try:
            self.request_queue.put_nowait(RequestMessage(RequestType.QUIT))
            print("Force exit: QUIT request posted.")
        except Exception as queue_err:
            print(f"Force exit: failed to enqueue QUIT: {queue_err}")

        if not self.worker_thread or not self.worker_thread.is_alive():
            self.worker_shutdown_event.set()

        try:
            self.page.window.prevent_close = False
        except Exception as prevent_err:
            print(f"Force exit: unable to clear prevent_close: {prevent_err}")

        self._finalize_shutdown(reason=normalized_reason)

    def _finalize_shutdown(self, reason: str = ""):
        normalized_reason = reason or "unspecified"
        with self._shutdown_lock:
            if self.shutdown_finalized:
                print(f"Shutdown already finalized. reason={normalized_reason}")
                return
            self.shutdown_finalized = True

        self._flush_pending_form_value_save()
        self._hero_breathing_active = False
        self._stop_hero_breathing_timer()
        print(f"Shutdown finalization started. reason={normalized_reason}")

        worker_active = self.worker_thread is not None and self.worker_thread.is_alive()
        if worker_active:
            if self.worker_shutdown_event.wait(timeout=5.0):
                print("Shutdown: worker shutdown signal received.")
            else:
                print("Shutdown: worker shutdown wait timed out.")
        else:
            self.worker_shutdown_event.set()

        self.ui_loop_running = False
        current_thread = threading.current_thread()

        if worker_active and self.worker_thread is not current_thread:
            try:
                self.worker_thread.join(timeout=2.0)
                print("Shutdown: worker thread joined or timeout.")
            except Exception as join_err:
                print(f"Shutdown: worker thread join failed: {join_err}")

        if (
            self.queue_thread
            and self.queue_thread.is_alive()
            and self.queue_thread is not current_thread
        ):
            try:
                self.queue_thread.join(timeout=2.0)
                print("Shutdown: queue thread joined or timeout.")
            except Exception as join_err:
                print(f"Shutdown: queue thread join failed: {join_err}")

        window = getattr(self.page, "window", None)
        close_requested = False
        if window:
            try:
                window.close()
                close_requested = True
                print("Shutdown: window.close() called.")
            except AttributeError:
                try:
                    window.destroy()
                    close_requested = True
                    self.window_closed_event.set()
                    print("Shutdown: window.destroy() called.")
                except Exception as destroy_err:
                    print(f"Shutdown: window destroy failed: {destroy_err}")
            except Exception as close_err:
                print(f"Shutdown: window close failed: {close_err}")
        else:
            print("Shutdown: page window not available.")

        if not close_requested:
            try:
                self.page.update()
            except Exception as update_err:
                print(f"Shutdown: page update after close failed: {update_err}")

        if not self.window_closed_event.is_set():
            try:
                if self.window_closed_event.wait(timeout=3.0):
                    print("Shutdown: window close confirmed.")
                else:
                    print("Shutdown: window close wait timed out.")
            except Exception as wait_err:
                print(f"Shutdown: waiting for window close failed: {wait_err}")
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
