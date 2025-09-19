import re
from typing import List, Any, Optional, Dict, Tuple

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

def check_translation_quality(
    actions: ExcelActions,
    browser_manager: BrowserCopilotManager,
    source_range: str,
    translated_range: str,
    status_output_range: str,
    issue_output_range: str,
    sheet_name: Optional[str] = None,
    batch_size: int = 3,
) -> str:
    """Compare source and translated ranges, then record review results."""
    try:
        cell_ref_re = re.compile(r"([A-Za-z]+)(\d+)")

        def _parse_range_dimensions(range_ref: str) -> Tuple[int, int]:
            ref = range_ref.split("!")[-1].replace("$", "").strip()
            if not ref:
                raise ToolExecutionError("Range string is empty.")
            if ":" not in ref:
                return 1, 1
            start_ref, end_ref = ref.split(":", 1)
            start_match = cell_ref_re.fullmatch(start_ref)
            end_match = cell_ref_re.fullmatch(end_ref)
            if not start_match or not end_match:
                raise ToolExecutionError("Range format is invalid.")

            def _col_to_index(col: str) -> int:
                result = 0
                for ch in col.upper():
                    if not ("A" <= ch <= "Z"):
                        raise ToolExecutionError("Range format is invalid.")
                    result = result * 26 + (ord(ch) - ord("A") + 1)
                return result

            start_col = _col_to_index(start_match.group(1))
            end_col = _col_to_index(end_match.group(1))
            start_row = int(start_match.group(2))
            end_row = int(end_match.group(2))
            rows = abs(end_row - start_row) + 1
            cols = abs(end_col - start_col) + 1
            return rows, cols

        def _reshape_to_dimensions(data: Any, rows: int, cols: int) -> List[List[Any]]:
            if isinstance(data, list) and data and all(isinstance(row, list) for row in data):
                if len(data) == rows and all(len(row) == cols for row in data):
                    return [row[:] for row in data]

            flattened: List[Any] = []
            if isinstance(data, list):
                for item in data:
                    if isinstance(item, list):
                        flattened.extend(item)
                    else:
                        flattened.append(item)
            elif data is None:
                flattened.append("")
            else:
                flattened.append(data)

            expected = rows * cols
            if len(flattened) != expected:
                raise ToolExecutionError(
                    f"Expected {expected} values for range but got {len(flattened)}."
                )

            reshaped: List[List[Any]] = []
            for r in range(rows):
                start_index = r * cols
                reshaped.append(list(flattened[start_index:start_index + cols]))
            return reshaped

        src_rows, src_cols = _parse_range_dimensions(source_range)
        trans_rows, trans_cols = _parse_range_dimensions(translated_range)
        status_rows, status_cols = _parse_range_dimensions(status_output_range)
        issue_rows, issue_cols = _parse_range_dimensions(issue_output_range)

        if (src_rows, src_cols) != (trans_rows, trans_cols):
            raise ToolExecutionError("Source range and translated range sizes do not match.")
        if (src_rows, src_cols) != (status_rows, status_cols) or (src_rows, src_cols) != (issue_rows, issue_cols):
            raise ToolExecutionError("Output ranges must match the source range size.")

        source_data = _reshape_to_dimensions(actions.read_range(source_range, sheet_name), src_rows, src_cols)
        translated_data = _reshape_to_dimensions(actions.read_range(translated_range, sheet_name), src_rows, src_cols)

        status_matrix = [["" for _ in range(src_cols)] for _ in range(src_rows)]
        issue_matrix = [["" for _ in range(src_cols)] for _ in range(src_rows)]

        review_entries: List[Dict[str, Any]] = []
        id_to_position: Dict[str, Tuple[int, int]] = {}
        needs_revision_count = 0

        for r in range(src_rows):
            for c in range(src_cols):
                original_text = source_data[r][c]
                translated_text = translated_data[r][c]
                if isinstance(original_text, str) and original_text.strip():
                    if isinstance(translated_text, str) and translated_text.strip():
                        entry_id = f"{r}:{c}"
                        review_entries.append(
                            {
                                "id": entry_id,
                                "original_text": original_text,
                                "translated_text": translated_text,
                            }
                        )
                        id_to_position[entry_id] = (r, c)
                    else:
                        status_matrix[r][c] = "要修正"
                        issue_matrix[r][c] = "英訳セルが空、または無効です。"
                        needs_revision_count += 1
                else:
                    status_matrix[r][c] = ""
                    issue_matrix[r][c] = ""

        if not review_entries:
            actions.write_range(status_output_range, status_matrix, sheet_name)
            actions.write_range(issue_output_range, issue_matrix, sheet_name)
            return "翻訳チェックの対象となる文字列が存在しなかったため、結果列を初期化しました。"

        normalized_batch_size = max(1, min(batch_size or 1, 5))
        normalized_batch_size = min(normalized_batch_size, len(review_entries))

        for batch in (review_entries[i:i + normalized_batch_size] for i in range(0, len(review_entries), normalized_batch_size)):
            payload = json.dumps(batch, ensure_ascii=False)
            analysis_prompt = (
                "あなたは英訳の品質を評価するレビュアーです。各項目について、英訳が原文の意味・ニュアンス・文法・スペル・主語述語の対応として適切かを確認してください。"
                "各項目には 'original_text'（日本語原文）と 'translated_text'（英訳）が含まれています。"
                "翻訳に問題がなければ status は 'OK' とし、notes は空文字または簡潔な補足にしてください。"
                "修正が必要な場合は status を 'REVISE' とし、notes には日本語で『Issue: ... / Suggestion: ...』の形式で問題点と修正案を記述してください。"
                "AIは JSON のみを返し、余計な文章やマークアップを付けないでください。"
                f"\n\n{payload}\n"
            )

            response = browser_manager.ask(analysis_prompt)
            match = re.search(r"\[[\s\S]*\]|\{[\s\S]*\}", response)
            json_candidates = []
            if match:
                json_candidates.append(match.group(0))
            json_candidates.extend(re.findall(r"\{[\s\S]*?\}", response))
            if not json_candidates:
                json_candidates.append(response)

            batch_results = None
            for candidate in json_candidates:
                candidate = candidate.strip()
                try:
                    batch_results = json.loads(candidate)
                    break
                except json.JSONDecodeError:
                    continue

            if batch_results is None:
                raise ToolExecutionError(f"翻訳チェックの結果をJSONとして解析できませんでした。応答: {response}")

            if not isinstance(batch_results, list):
                raise ToolExecutionError("翻訳チェックの応答形式が不正です。JSON配列を返してください。")

            for item in batch_results:
                if not isinstance(item, dict):
                    raise ToolExecutionError("翻訳チェックの応答に不正な要素が含まれています。")
                item_id = item.get("id")
                if item_id not in id_to_position:
                    raise ToolExecutionError("翻訳チェックの応答に未知のIDが含まれています。")
                status_value = str(item.get("status", "")).strip().upper()
                notes_value = str(item.get("notes", "")).strip()

                row_idx, col_idx = id_to_position[item_id]
                if status_value in {"OK", "PASS", "GOOD"}:
                    status_matrix[row_idx][col_idx] = "OK"
                    issue_matrix[row_idx][col_idx] = notes_value or ""
                elif status_value in {"REVISE", "NG", "FAIL", "ISSUE"}:
                    status_matrix[row_idx][col_idx] = "要修正"
                    issue_matrix[row_idx][col_idx] = notes_value or "修正内容を追記してください。"
                    needs_revision_count += 1
                else:
                    status_matrix[row_idx][col_idx] = status_value or "要確認"
                    issue_matrix[row_idx][col_idx] = notes_value or "ステータスが解釈できませんでした。"
                    needs_revision_count += 1

        actions.write_range(status_output_range, status_matrix, sheet_name)
        actions.write_range(issue_output_range, issue_matrix, sheet_name)

        processed_items = len(review_entries)
        return (
            f"翻訳チェックを完了しました。対象 {processed_items} 件中、要修正 {needs_revision_count} 件の結果を"
            f" '{status_output_range}' と '{issue_output_range}' に書き込みました。"
        )

    except ToolExecutionError:
        raise
    except Exception as e:
        raise ToolExecutionError(f"翻訳チェック中にエラーが発生しました: {e}") from e


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

