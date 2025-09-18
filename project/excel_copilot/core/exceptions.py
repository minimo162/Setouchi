# excel_copilot/core/exceptions.py

class ExcelCoPilotError(Exception):
    """アプリケーション固有のすべてのエラーの基底クラス。"""
    pass

class ExcelConnectionError(ExcelCoPilotError):
    """ワークブックへの接続に関する問題が発生した際に送出される。"""
    pass

class ToolExecutionError(ExcelCoPilotError):
    """ツールの実行中にエラーが発生した際に送出される。"""
    pass

class LLMResponseError(ExcelCoPilotError):
    """LLMが不正または予期しない応答を返した際に送出される。"""
    pass

class SchemaGenerationError(ExcelCoPilotError):
    """ツールのJSONスキーマ生成に失敗した際に送出される。"""
    pass