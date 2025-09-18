import re
from typing import List, Any, Optional, Dict

from excel_copilot.core.browser_copilot_manager import BrowserCopilotManager
from excel_copilot.core.exceptions import ToolExecutionError

from .actions import ExcelActions

def writetocell(actions: ExcelActions, cell: str, value: Any, sheetname: Optional[str] = None) -> str:
    """
    Excelシートの特定のセルに値を書き込みます。
    """
    return actions.write_to_cell(cell, value, sheetname)

def readcellvalue(actions: ExcelActions, cell: str, sheetname: Optional[str] = None) -> Any:
    """
    Excelシートの特定のセルの値を読み取ります。
    """
    return actions.read_cell_value(cell, sheetname)

def getallsheetnames(actions: ExcelActions) -> str:
    """
    現在開いているExcelワークブック内のすべてのシート名を取得します。
    """
    names = actions.get_sheet_names()
    return f"利用可能なシートは次の通りです: {', '.join(names)}"

def copyrange(actions: ExcelActions, sourcerange: str, destinationrange: str, sheetname: Optional[str] = None) -> str:
    """
    指定した範囲を別の場所にコピーします。
    """
    return actions.copy_range(sourcerange, destinationrange, sheetname)

def executeexcelformula(actions: ExcelActions, cell: str, formula: str, sheetname: Optional[str] = None) -> str:
    """
    指定したセルにExcelの数式を設定します。
    """
    return actions.set_formula(cell, formula, sheetname)

def readrangevalues(actions: ExcelActions, cellrange: str, sheetname: Optional[str] = None) -> str:
    """
    指定した範囲のセルから値を読み取ります。1セルでも範囲として指定可能です。
    """
    values = actions.read_range(cellrange, sheetname)
    return f"範囲 '{cellrange}' の値は次の通りです: {values}"

def writerangevalues(actions: ExcelActions, cellrange: str, data: List[List[Any]], sheetname: Optional[str] = None) -> str:
    """
    指定した範囲に2次元リストのデータを書き込みます。1セルでも対応可能です。
    """
    return actions.write_range(cellrange, data, sheetname)

def getactiveworkbookandsheet(actions: ExcelActions) -> str:
    """
    現在アクティブなExcelブックとシート名を取得します。
    """
    info_dict = actions.get_active_workbook_and_sheet()
    return f"ブック: {info_dict['workbook_name']}, シート: {info_dict['sheet_name']}"

def formatrange(actions: ExcelActions,
                 cellrange: str,
                 sheetname: Optional[str] = None,
                 fontname: Optional[str] = None,
                 fontsize: Optional[float] = None,
                 fontcolorhex: Optional[str] = None,
                 bold: Optional[bool] = None,
                 italic: Optional[bool] = None,
                 fillcolorhex: Optional[str] = None,
                 columnwidth: Optional[float] = None,
                 rowheight: Optional[float] = None,
                 horizontalalignment: Optional[str] = None,
                 borderstyle: Optional[Dict[str, Any]] = None) -> str:
    """
    指定した範囲に書式設定を適用します。
    """
    return actions.format_range(
        cell_range=cellrange,
        sheet_name=sheetname,
        font_name=fontname,
        font_size=fontsize,
        font_color_hex=fontcolorhex,
        bold=bold,
        italic=italic,
        fill_color_hex=fillcolorhex,
        column_width=columnwidth,
        row_height=rowheight,
        horizontal_alignment=horizontalalignment,
        border_style=borderstyle
    )

import json
def translate_range_contents(
    actions: ExcelActions,
    browser_manager: BrowserCopilotManager,
    cell_range: str,
    target_language: str = "English",
    sheet_name: Optional[str] = None
) -> str:
    """
    指定された範囲のセルを読み込み、テキスト部分のみをAIで翻訳し、同じ範囲に書き戻します。
    数値や空白セルは変更されません。
    """
    try:
        # 1. データの読み取り
        original_data = actions.read_range(cell_range, sheet_name)
        if not isinstance(original_data, list):
            original_data = [[original_data]]
        elif original_data and not isinstance(original_data[0], list):
            original_data = [original_data]

        texts_to_translate = []
        text_positions = []
        for r, row in enumerate(original_data):
            for c, cell in enumerate(row):
                if isinstance(cell, str) and re.search(r'[ぁ-んァ-ン一-龯]', cell):
                    texts_to_translate.append(cell)
                    text_positions.append((r, c))

        if not texts_to_translate:
            return f"範囲 '{cell_range}' 内に翻訳対象のテキストが見つかりませんでした。"

        # 2. 翻訳の実行（JSON形式を要求）
        translation_prompt = (
            f"以下のJSONリストに格納された日本語の各テキストを、それぞれ{target_language}に翻訳し、"
            f"翻訳後のテキストを格納したJSONリスト形式で返してください。リストの順序と要素数は変えないでください。"
            f"応答はJSONのみとし、前後に説明やコードブロックのマークアップを含めないでください。\n\n"
            f"{json.dumps(texts_to_translate, ensure_ascii=False)}"
        )
        response = browser_manager.ask(translation_prompt)

        try:
            # 応答がコードブロックで囲まれている場合を考慮してJSONを抽出
            match = re.search(r'\{.*\}|\[.*\]', response, re.DOTALL)
            if match:
                json_str = match.group(0)
                translated_texts = json.loads(json_str)
            else:
                # コードブロックがない場合は、そのまま解析を試みる
                translated_texts = json.loads(response)
        except json.JSONDecodeError:
            raise ToolExecutionError(f"AIからの翻訳結果をJSONとして解析できませんでした。応答: {response}")

        if not isinstance(translated_texts, list) or len(translated_texts) != len(texts_to_translate):
            raise ToolExecutionError("翻訳前と翻訳後でテキストの数や形式が一致しません。")

        # 3. 元のデータ構造に翻訳結果を反映
        new_data = [row[:] for row in original_data]
        for i, (r, c) in enumerate(text_positions):
            new_data[r][c] = translated_texts[i]

        # 4. Excelへの書き込み
        return actions.write_range(cell_range, new_data, sheet_name)

    except Exception as e:
        raise ToolExecutionError(f"範囲 '{cell_range}' の翻訳中にエラーが発生しました: {e}")

def insert_shape(actions: ExcelActions,
                 cell_range: str,
                 shape_type: str,
                 sheet_name: Optional[str] = None,
                 fill_color_hex: Optional[str] = None,
                 line_color_hex: Optional[str] = None) -> str:
    """
    指定したセル範囲に、指定した書式で図形を挿入します。
    :param cell_range: 図形を挿入する範囲 (例: "A1:C5")
    :param shape_type: 挿入する図形の種類 (例: "四角形", "楕円")
    :param sheet_name: 対象シート名（省略可）
    :param fill_color_hex: 塗りつぶしの色 (16進数, 例: "#FF0000")
    :param line_color_hex: 枠線の色 (16進数, 例: "#0000FF")
    """
    return actions.insert_shape_in_range(cell_range, shape_type, sheet_name, fill_color_hex, line_color_hex)

def format_shape(actions: ExcelActions, fill_color_hex: Optional[str] = None, line_color_hex: Optional[str] = None, sheet_name: Optional[str] = None) -> str:
    """
    [非推奨] この関数は使わないでください。代わりに insert_shape 関数の引数で色を指定してください。
    """
    return actions.format_last_shape(fill_color_hex, line_color_hex, sheet_name)