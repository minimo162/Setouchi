# excel_copilot/config.py

from __future__ import annotations

from pathlib import Path
from typing import List
import os
import platform


def _to_bool(value: str | None, default: bool = False) -> bool:
    if value is None:
        return default
    normalized = value.strip().lower()
    if not normalized:
        return default
    if normalized in {"1", "true", "t", "yes", "y", "on"}:
        return True
    if normalized in {"0", "false", "f", "no", "n", "off"}:
        return False
    return default


def _ensure_directory(path: Path) -> Path:
    try:
        path.mkdir(parents=True, exist_ok=True)
    except Exception:
        # The consumer code handles error reporting if the directory truly cannot be created.
        pass
    return path

class AppSettings:
    """アプリケーション全体の設定を管理するクラス"""

    def __init__(self):
        # --- ReAct Agent Settings ---
        # エージェントの最大反復回数
        self.MAX_ITERATIONS: int = self._get_int("COPILOT_MAX_ITERATIONS", 10, minimum=1)
        # AIに渡す会話履歴の最大メッセージ数（システムプロンプトを除く）
        self.HISTORY_MAX_MESSAGES: int = self._get_int("COPILOT_HISTORY_MAX_MESSAGES", 10, minimum=1)

        # --- BrowserCopilotManager Settings ---
        # Chrome/Edgeのユーザーデータディレクトリ
        # このパスは環境に合わせて設定してください
        self.USER_DATA_DIR: Path = self._resolve_user_data_dir()
        self.PREFERRED_BROWSER_CHANNELS: List[str] = self._get_browser_channels()
        self.PLAYWRIGHT_HEADLESS: bool = _to_bool(os.getenv("COPILOT_HEADLESS"), False)
        self.PLAYWRIGHT_SLOW_MO_MS: int = self._get_int("COPILOT_SLOW_MO_MS", 50, minimum=0)
        self.PLAYWRIGHT_PAGE_GOTO_TIMEOUT_MS: int = self._get_int(
            "COPILOT_PAGE_GOTO_TIMEOUT_MS", 90000, minimum=1000
        )
        self.PLAYWRIGHT_SUPPRESS_FOCUS: bool = _to_bool(os.getenv("COPILOT_SUPPRESS_BROWSER_FOCUS"), True)

    def _get_int(self, env_key: str, default: int, minimum: int | None = None, maximum: int | None = None) -> int:
        raw_value = os.getenv(env_key)
        if raw_value is None:
            return default
        try:
            parsed = int(raw_value)
        except (TypeError, ValueError):
            return default
        if minimum is not None and parsed < minimum:
            return minimum
        if maximum is not None and parsed > maximum:
            return maximum
        return parsed

    def _resolve_user_data_dir(self) -> Path:
        env_override = os.getenv("COPILOT_USER_DATA_DIR")
        if env_override:
            expanded = Path(env_override).expanduser()
            return _ensure_directory(expanded)
        return self._get_default_user_data_dir()

    def _get_default_user_data_dir(self) -> Path:
        """OSに応じてデフォルトのブラウザユーザーデータディレクトリパスを取得する"""
        system = platform.system()
        if system == "Windows":
            base_path = Path(os.getenv('LOCALAPPDATA', ''))
            # 一般的なEdgeのパス
            edge_path = base_path / 'Microsoft' / 'Edge' / 'User Data'
            if edge_path.exists():
                return edge_path / 'Default'
            # 一般的なChromeのパス
            chrome_path = base_path / 'Google' / 'Chrome' / 'User Data'
            if chrome_path.exists():
                return chrome_path / 'Default'
        elif system == 'Darwin':  # macOS
            base_path = Path.home() / 'Library' / 'Application Support'
            # 一般的なEdgeのパス
            edge_path = base_path / 'Microsoft Edge'
            if edge_path.exists():
                return edge_path / 'Default'
            # 一般的なChromeのパス
            chrome_path = base_path / 'Google' / 'Chrome'
            if chrome_path.exists():
                return chrome_path / 'Default'

        # デフォルトが見つからない場合は、カレントディレクトリに作成
        fallback_path = _ensure_directory(Path.cwd() / 'browser_user_data')
        return fallback_path

    def _get_browser_channels(self) -> List[str]:
        raw = os.getenv("COPILOT_BROWSER_CHANNELS")
        if not raw:
            return ["msedge", "chrome"]
        channels: List[str] = []
        for candidate in raw.split(","):
            normalized = candidate.strip()
            if normalized:
                channels.append(normalized)
        return channels or ["msedge", "chrome"]

# 設定クラスのインスタンスを作成
settings = AppSettings()

# グローバル変数としてエクスポート
MAX_ITERATIONS = settings.MAX_ITERATIONS
HISTORY_MAX_MESSAGES = settings.HISTORY_MAX_MESSAGES
COPILOT_USER_DATA_DIR = str(settings.USER_DATA_DIR)
COPILOT_BROWSER_CHANNELS = settings.PREFERRED_BROWSER_CHANNELS
COPILOT_HEADLESS = settings.PLAYWRIGHT_HEADLESS
COPILOT_SLOW_MO_MS = settings.PLAYWRIGHT_SLOW_MO_MS
COPILOT_PAGE_GOTO_TIMEOUT_MS = settings.PLAYWRIGHT_PAGE_GOTO_TIMEOUT_MS
COPILOT_SUPPRESS_BROWSER_FOCUS = settings.PLAYWRIGHT_SUPPRESS_FOCUS
