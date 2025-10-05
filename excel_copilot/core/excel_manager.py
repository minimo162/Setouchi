# excel_copilot/core/excel_manager.py

import xlwings as xw
import time
from .exceptions import ExcelConnectionError

class ExcelManager:
    """
    Excelアプリケーションへの接続とライフサイクルを管理するクラス。
    """
    def __init__(self):
        self.app: xw.App | None = None
        self.book: xw.Book | None = None

    def __enter__(self):
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.book = None
        self.app = None

    def connect(self, retries: int = 5, delay: int = 2):
        """
        実行中のExcelアプリケーションに接続し、アクティブなワークブックを取得する。
        """
        last_exception = None
        for i in range(retries):
            try:
                if not xw.apps:
                    raise ExcelConnectionError("実行中のExcelアプリケーションが見つかりません。")
                
                self.app = xw.apps.active
                if self.app is None:
                    raise ExcelConnectionError("アクティブなExcelプロセスが見つかりませんでした。")
                
                self.app.visible = True
                self.app.screen_updating = True
                
                if not self.app.books:
                     raise ExcelConnectionError("開かれているExcelワークブックが見つかりません。")
                self.book = self.app.books.active
                if self.book is None:
                     raise ExcelConnectionError("アクティブなワークブックが見つかりませんでした。")
                
                _ = self.book.name
                print(f"Excelに接続しました。アクティブなブック: {self.book.name}")
                return

            except Exception as e:
                last_exception = e
                print(f"Excelへの接続に失敗 (試行 {i + 1}/{retries}): {e}")
                time.sleep(delay)
        
        raise ExcelConnectionError(f"{retries}回リトライしましたが、Excelへの接続に失敗しました。") from last_exception

    def get_active_workbook(self) -> xw.Book:
        """現在アクティブなワークブックオブジェクトを返す"""
        if not self.book:
            raise ExcelConnectionError("ワークブックに接続されていません。")
        return self.book

    def focus_application_window(self) -> None:
        """Excelアプリケーションとアクティブブックを前面に表示する"""

        app = self.app
        book = self.book

        if app:
            try:
                app.visible = True
            except Exception:
                pass
            try:
                app.activate(steal_focus=True)
            except TypeError:
                try:
                    app.activate()
                except Exception:
                    pass
            except Exception:
                try:
                    app.activate()
                except Exception:
                    pass

        if book:
            try:
                book.activate()
            except Exception:
                pass

    def get_active_workbook_and_sheet(self) -> dict[str, str]:
        """現在アクティブなExcelブックとシート名を取得し、辞書として返す。"""
        try:
            book = self.get_active_workbook()
            book_name = book.name
            sheet_name = book.sheets.active.name
            return {"workbook_name": book_name, "sheet_name": sheet_name}
        except Exception as e:
            raise ExcelConnectionError(f"ブックとシートの取得中にエラーが発生しました: {e}")

    def list_sheet_names(self) -> list[str]:
        """アクティブなワークブック内のシート名一覧を返す。"""
        try:
            book = self.get_active_workbook()
            return [sheet.name for sheet in book.sheets]
        except Exception as e:
            raise ExcelConnectionError(f"シート一覧の取得中にエラーが発生しました: {e}")

    def activate_sheet(self, sheet_name: str) -> str:
        """指定したシートをアクティブにする。"""
        try:
            book = self.get_active_workbook()
            sheet = book.sheets[sheet_name]
            sheet.activate()
            return sheet.name
        except Exception as e:
            raise ExcelConnectionError(f"シート '{sheet_name}' のアクティブ化に失敗しました: {e}")
