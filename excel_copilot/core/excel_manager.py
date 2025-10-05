# excel_copilot/core/excel_manager.py

import xlwings as xw
import time
from typing import Optional, List

from .exceptions import ExcelConnectionError

class ExcelManager:
    """
    Excelアプリケーションへの接続とライフサイクルを管理するクラス。
    """
    def __init__(self, workbook_name: Optional[str] = None):
        self.app: xw.App | None = None
        self.book: xw.Book | None = None
        self._requested_workbook_name: Optional[str] = workbook_name

    def __enter__(self):
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.book = None
        self.app = None

    def connect(self, workbook_name: Optional[str] = None, retries: int = 5, delay: int = 2):
        """
        実行中のExcelアプリケーションに接続し、アクティブなワークブックを取得する。
        """
        last_exception = None
        requested_name = workbook_name or self._requested_workbook_name
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

                books: List[xw.Book] = list(self.app.books)
                target_book: Optional[xw.Book] = None

                if requested_name:
                    for candidate in books:
                        if candidate.name == requested_name:
                            target_book = candidate
                            break

                if target_book is None:
                    try:
                        target_book = self.app.books.active
                    except Exception:
                        target_book = None

                if target_book is None and books:
                    target_book = books[0]

                if target_book is None:
                    raise ExcelConnectionError("対象のワークブックを特定できませんでした。")

                self.book = target_book
                self._requested_workbook_name = target_book.name

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

    def list_workbook_names(self) -> list[str]:
        """開いているすべてのワークブック名を返す"""
        if not self.app:
            raise ExcelConnectionError("Excelアプリケーションに接続されていません。")
        try:
            return [book.name for book in self.app.books]
        except Exception as exc:
            raise ExcelConnectionError(f"ワークブック一覧の取得に失敗しました: {exc}") from exc

    def activate_workbook(self, workbook_name: str) -> str:
        """指定したワークブックをアクティブにして返す"""
        if not self.app:
            raise ExcelConnectionError("Excelアプリケーションに接続されていません。")

        for book in self.app.books:
            if book.name == workbook_name:
                try:
                    book.activate()
                except Exception:
                    pass
                self.book = book
                self._requested_workbook_name = workbook_name
                return book.name

        raise ExcelConnectionError(f"ワークブック '{workbook_name}' をアクティブにできませんでした。")

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
