# excel_copilot/config.py

from pathlib import Path
import os

class AppSettings:
    """アプリケーション全体の設定を管理するクラス"""

    def __init__(self):
        # --- ReAct Agent Settings ---
        # エージェントの最大反復回数
        self.MAX_ITERATIONS: int = 10
        # AIに渡す会話履歴の最大メッセージ数（システムプロンプトを除く）
        self.HISTORY_MAX_MESSAGES: int = 10

        # --- BrowserCopilotManager Settings ---
        # Chrome/Edgeのユーザーデータディレクトリ
        # このパスは環境に合わせて設定してください
        self.USER_DATA_DIR: Path = self._get_default_user_data_dir()

    def _get_default_user_data_dir(self) -> Path:
        """OSに応じてデフォルトのブラウザユーザーデータディレクトリパスを取得する"""
        if os.name == 'nt':  # Windows
            base_path = Path(os.getenv('LOCALAPPDATA', ''))
            # 一般的なEdgeのパス
            edge_path = base_path / 'Microsoft' / 'Edge' / 'User Data'
            if edge_path.exists():
                return edge_path / 'Default'
            # 一般的なChromeのパス
            chrome_path = base_path / 'Google' / 'Chrome' / 'User Data'
            if chrome_path.exists():
                return chrome_path / 'Default'
        elif os.uname().sysname == 'Darwin':  # macOS
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
        fallback_path = Path.cwd() / 'browser_user_data'
        fallback_path.mkdir(exist_ok=True)
        return fallback_path

# 設定クラスのインスタンスを作成
settings = AppSettings()

# グローバル変数としてエクスポート
MAX_ITERATIONS = settings.MAX_ITERATIONS
HISTORY_MAX_MESSAGES = settings.HISTORY_MAX_MESSAGES
COPILOT_USER_DATA_DIR = str(settings.USER_DATA_DIR)
