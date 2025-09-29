import re
import difflib
import logging
import os
from typing import List, Any, Optional, Dict, Tuple

from excel_copilot.core.browser_copilot_manager import BrowserCopilotManager
from excel_copilot.core.exceptions import ToolExecutionError

from .actions import ExcelActions


_logger = logging.getLogger(__name__)
_DIFF_DEBUG_ENABLED = os.getenv('EXCEL_COPILOT_DEBUG_DIFF', '').lower() in {'1', 'true', 'yes'}

if _DIFF_DEBUG_ENABLED and not logging.getLogger().handlers:
    logging.basicConfig(level=logging.DEBUG)



def _diff_debug(message: str) -> None:
    if _DIFF_DEBUG_ENABLED:
        _logger.debug(message)


def _shorten_debug(value: str, limit: int = 120) -> str:
    if value is None:
        return ''
    text = str(value).replace('\r', '\r').replace('\n', '\n')
    return text if len(text) <= limit else text[:limit] + '…'


CELL_REFERENCE_PATTERN = re.compile(r"([A-Za-z]+)(\d+)")

LEGACY_DIFF_MARKER_PATTERN = re.compile(r"【(追加|削除)：(.*?)】")
MODERN_DIFF_MARKER_PATTERN = re.compile(r"\?(?:追加|削除)\?\s*(.*?)\?")
_BASE_DIFF_TOKEN_PATTERN = re.compile(r"\s+|[^\s]+")
_SENTENCE_BOUNDARY_CHARS = set("。.!?！？")
_CLOSING_PUNCTUATION = ")]}、。！？」』】》）］'\"”’"
_MAX_DIFF_SEGMENT_TOKENS = 18
_MAX_DIFF_SEGMENT_CHARS = 80

REFUSAL_PATTERNS = (
    "申し訳ございません。これについてチャットできません。",
    "申し訳ございません。これについてチャットできません",
    "申し訳ございません。チャットを保存して新しいチャットを開始するには、[新しいチャット] を選択してください。",
    "チャットを保存して新しいチャットを開始するには、[新しいチャット] を選択してください。",
    "お答えできません。",
    "お答えできません",
    "I'm sorry, but I can't help with that.",
    "I cannot help with that request.",
    "エラーが発生しました: 応答形式が不正です。'Thought:' または 'Final Answer:' が見つかりません。",
    "応答形式が不正です。'Thought:' または 'Final Answer:' が見つかりません。",
)


def _parse_range_dimensions(range_ref: str) -> Tuple[int, int]:
    ref = range_ref.split('!')[-1].replace('$', '').strip()
    if not ref:
        raise ToolExecutionError('Range string is empty.')
    if ':' not in ref:
        return 1, 1
    start_ref, end_ref = ref.split(':', 1)
    start_match = CELL_REFERENCE_PATTERN.fullmatch(start_ref)
    end_match = CELL_REFERENCE_PATTERN.fullmatch(end_ref)
    if not start_match or not end_match:
        raise ToolExecutionError('Range format is invalid.')

    def _col_to_index(col: str) -> int:
        result = 0
        for ch in col.upper():
            if not ('A' <= ch <= 'Z'):
                raise ToolExecutionError('Range format is invalid.')
            result = result * 26 + (ord(ch) - ord('A') + 1)
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
        flattened.append('')
    else:
        flattened.append(data)

    expected = rows * cols
    if len(flattened) != expected:
        raise ToolExecutionError(
            f'Expected {expected} values for range but got {len(flattened)}.'
        )

    reshaped: List[List[Any]] = []
    for r in range(rows):
        start_index = r * cols
        reshaped.append(list(flattened[start_index:start_index + cols]))
    return reshaped

def _normalize_cell_value(cell: Any) -> str:
    if isinstance(cell, str):
        return cell
    if cell is None:
        return ''
    return str(cell)

def _strip_diff_markers(text: Any) -> str:
    if not isinstance(text, str):
        return ''
    stripped = LEGACY_DIFF_MARKER_PATTERN.sub(lambda m: m.group(2), text)
    return MODERN_DIFF_MARKER_PATTERN.sub(lambda m: m.group(1), stripped)

def _tokenize_for_diff(text: str) -> List[str]:
    if not text:
        _diff_debug('_tokenize_for_diff empty input')
        return []
    raw_tokens = _BASE_DIFF_TOKEN_PATTERN.findall(text)
    _diff_debug(f"_tokenize_for_diff raw_tokens={_shorten_debug(raw_tokens)}")
    if not raw_tokens:
        return [text]

    segments: List[str] = []
    current_tokens: List[str] = []
    content_token_count = 0
    content_char_count = 0

    def flush() -> None:
        nonlocal current_tokens, content_token_count, content_char_count
        if current_tokens:
            segment = ''.join(current_tokens)
            segments.append(segment)
            _diff_debug(f"_tokenize_for_diff flush segment={_shorten_debug(segment)}")
            current_tokens = []
            content_token_count = 0
            content_char_count = 0

    for token in raw_tokens:
        current_tokens.append(token)
        stripped = token.strip()
        if not stripped:
            if '\r\n' in token or '\n' in token:
                _diff_debug('_tokenize_for_diff flush due to newline token')
                flush()
            continue
        _diff_debug(f"_tokenize_for_diff token={_shorten_debug(token)}")
        content_token_count += 1
        content_char_count += len(stripped)
        trimmed = stripped.rstrip(_CLOSING_PUNCTUATION)
        last_char = trimmed[-1] if trimmed else stripped[-1]
        if last_char in _SENTENCE_BOUNDARY_CHARS:
            _diff_debug('_tokenize_for_diff flush due to boundary char')
            flush()
        elif content_token_count >= _MAX_DIFF_SEGMENT_TOKENS or content_char_count >= _MAX_DIFF_SEGMENT_CHARS:
            _diff_debug('_tokenize_for_diff flush due to size limit')
            flush()
    flush()
    result = segments or [text]
    _diff_debug(f"_tokenize_for_diff result={_shorten_debug(result)}")
    return result

def _format_diff_segment(tokens: List[str], label: str) -> Tuple[str, Optional[int], Optional[int]]:
    _diff_debug(f"_format_diff_segment start label={label} tokens={_shorten_debug(tokens)}")
    if not tokens:
        _diff_debug("_format_diff_segment no tokens provided")
        return '', None, None
    segment = ''.join(tokens)
    if not segment.strip():
        _diff_debug(f"_format_diff_segment segment without content label={label}")
        return segment, None, None
    leading_len = len(segment) - len(segment.lstrip())
    trailing_len = len(segment.rstrip()) - len(segment.strip())
    core_start = leading_len
    core_end = len(segment) - trailing_len if trailing_len else len(segment)
    core = segment[core_start:core_end]
    if not core:
        _diff_debug(f"_format_diff_segment no core text label={label}")
        return segment, None, None
    prefix = segment[:leading_len]
    suffix = segment[core_end:]
    marker_prefix = f'【{label}：'
    marker_suffix = '】'
    formatted = f'{prefix}{marker_prefix}{core}{marker_suffix}{suffix}'
    highlight_start_offset = len(prefix)
    highlight_length = len(marker_prefix) + len(core) + len(marker_suffix)
    _diff_debug(f"_format_diff_segment result label={label} formatted={_shorten_debug(formatted)} offset={highlight_start_offset} length={highlight_length}")
    return formatted, highlight_start_offset, highlight_length

def _build_diff_highlight(original: str, corrected: str) -> Tuple[str, List[Dict[str, int]]]:
    original_text = original if isinstance(original, str) else ('' if original is None else str(original))
    corrected_text = corrected if isinstance(corrected, str) else ('' if corrected is None else str(corrected))
    _diff_debug(f"_build_diff_highlight start orig_len={len(original_text)} corr_len={len(corrected_text)}")
    if original_text == corrected_text:
        _diff_debug("_build_diff_highlight texts identical")
        return corrected_text, []
    orig_tokens = _tokenize_for_diff(original_text)
    corr_tokens = _tokenize_for_diff(corrected_text)
    _diff_debug(f"_build_diff_highlight tokens orig={_shorten_debug(orig_tokens)} corr={_shorten_debug(corr_tokens)}")
    matcher = difflib.SequenceMatcher(a=orig_tokens, b=corr_tokens, autojunk=False)
    result_parts: List[str] = []
    spans: List[Dict[str, int]] = []
    cursor = 0
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        _diff_debug(
            f"_build_diff_highlight opcode={tag} orig_range=({i1},{i2}) corr_range=({j1},{j2}) "
            f"orig_tokens={_shorten_debug(orig_tokens[i1:i2])} corr_tokens={_shorten_debug(corr_tokens[j1:j2])}"
        )
        if tag == 'equal':
            text = ''.join(corr_tokens[j1:j2])
            result_parts.append(text)
            cursor += len(text)
            _diff_debug(f"_build_diff_highlight equal appended len={len(text)} cursor={cursor}")
        elif tag == 'replace':
            removed_tokens = orig_tokens[i1:i2]
            formatted_removed, offset_removed, length_removed = _format_diff_segment(removed_tokens, '削除')
            if formatted_removed:
                result_parts.append(formatted_removed)
                if offset_removed is not None and length_removed:
                    span = {'start': cursor + offset_removed, 'length': length_removed, 'type': '削除'}
                    spans.append(span)
                    _diff_debug(f"_build_diff_highlight span added {span}")
                cursor += len(formatted_removed)
            added_tokens = corr_tokens[j1:j2]
            formatted_added, offset_added, length_added = _format_diff_segment(added_tokens, '追加')
            if formatted_added:
                result_parts.append(formatted_added)
                if offset_added is not None and length_added:
                    span = {'start': cursor + offset_added, 'length': length_added, 'type': '追加'}
                    spans.append(span)
                    _diff_debug(f"_build_diff_highlight span added {span}")
                cursor += len(formatted_added)
        elif tag == 'delete':
            removed_tokens = orig_tokens[i1:i2]
            formatted_removed, offset_removed, length_removed = _format_diff_segment(removed_tokens, '削除')
            if formatted_removed:
                result_parts.append(formatted_removed)
                if offset_removed is not None and length_removed:
                    span = {'start': cursor + offset_removed, 'length': length_removed, 'type': '削除'}
                    spans.append(span)
                    _diff_debug(f"_build_diff_highlight span added {span}")
                cursor += len(formatted_removed)
        elif tag == 'insert':
            added_tokens = corr_tokens[j1:j2]
            formatted_added, offset_added, length_added = _format_diff_segment(added_tokens, '追加')
            if formatted_added:
                result_parts.append(formatted_added)
                if offset_added is not None and length_added:
                    span = {'start': cursor + offset_added, 'length': length_added, 'type': '追加'}
                    spans.append(span)
                    _diff_debug(f"_build_diff_highlight span added {span}")
                cursor += len(formatted_added)
    result = ''.join(result_parts)
    if not result.strip():
        _diff_debug("_build_diff_highlight result empty after strip")
        return corrected_text, []
    _diff_debug(f"_build_diff_highlight result_len={len(result)} spans={spans}")
    return result, spans


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
    corrected_output_range: Optional[str] = None,
    highlight_output_range: Optional[str] = None,
    sheet_name: Optional[str] = None,
    batch_size: int = 1,
) -> str:
    """Compare source and translated ranges, then record review results.

    Args:
        source_range: Range containing original Japanese text.
        translated_range: Range containing the English translation.
        status_output_range: Range to store review status (e.g., OK / 要修正).
        issue_output_range: Range to store review notes.
        corrected_output_range: Optional range to store the finalized corrected English sentences.
        highlight_output_range: Optional range to store a highlighted translation string with diff markers.
        sheet_name: Target sheet name. Uses active sheet if None.
        batch_size: Number of items reviewed together per AI request.
    """
    try:

        src_rows, src_cols = _parse_range_dimensions(source_range)
        trans_rows, trans_cols = _parse_range_dimensions(translated_range)
        status_rows, status_cols = _parse_range_dimensions(status_output_range)
        issue_rows, issue_cols = _parse_range_dimensions(issue_output_range)
        corrected_rows = corrected_cols = None
        if corrected_output_range:
            corrected_rows, corrected_cols = _parse_range_dimensions(corrected_output_range)
        highlight_rows = highlight_cols = None
        if highlight_output_range:
            highlight_rows, highlight_cols = _parse_range_dimensions(highlight_output_range)

        if (src_rows, src_cols) != (trans_rows, trans_cols):
            raise ToolExecutionError("Source range and translated range sizes do not match.")
        if (src_rows, src_cols) != (status_rows, status_cols) or (src_rows, src_cols) != (issue_rows, issue_cols):
            raise ToolExecutionError("Output ranges must match the source range size.")
        if corrected_output_range and (src_rows, src_cols) != (corrected_rows, corrected_cols):
            raise ToolExecutionError("Corrected output range must match the source range size.")
        if highlight_output_range and (src_rows, src_cols) != (highlight_rows, highlight_cols):
            raise ToolExecutionError("Highlight output range must match the source range size.")

        source_data = _reshape_to_dimensions(actions.read_range(source_range, sheet_name), src_rows, src_cols)
        translated_data = _reshape_to_dimensions(actions.read_range(translated_range, sheet_name), src_rows, src_cols)

        status_matrix = [["" for _ in range(src_cols)] for _ in range(src_rows)]
        issue_matrix = [["" for _ in range(src_cols)] for _ in range(src_rows)]
        corrected_matrix = [] if corrected_output_range else None
        highlight_matrix = [] if highlight_output_range else None
        highlight_styles = [] if highlight_output_range else None

        if corrected_matrix is not None or highlight_matrix is not None:
            for r in range(src_rows):
                corrected_row = [] if corrected_matrix is not None else None
                highlight_row = [] if highlight_matrix is not None else None
                styles_row = [] if highlight_styles is not None else None
                for c in range(src_cols):
                    base_value = _normalize_cell_value(translated_data[r][c])
                    if corrected_row is not None:
                        corrected_row.append(base_value)
                    if highlight_row is not None:
                        highlight_row.append(base_value)
                    if styles_row is not None:
                        styles_row.append([])
                if corrected_row is not None:
                    corrected_matrix.append(corrected_row)
                if highlight_row is not None:
                    highlight_matrix.append(highlight_row)
                if styles_row is not None:
                    highlight_styles.append(styles_row)

        def _infer_corrected_text(base_text: str, item: Dict[str, Any]) -> str:
            base = base_text if isinstance(base_text, str) else ('' if base_text is None else str(base_text))
            candidates = [
                item.get('corrected_text'),
                item.get('revised_text'),
                item.get('suggested_text'),
            ]
            for candidate in candidates:
                if isinstance(candidate, str) and candidate.strip():
                    return candidate
            highlighted_candidate = item.get('highlighted_text')
            if isinstance(highlighted_candidate, str) and highlighted_candidate.strip():
                stripped = _strip_diff_markers(highlighted_candidate)
                if stripped.strip():
                    return stripped
            before = item.get('before_text')
            after = item.get('after_text')
            before_str = before if isinstance(before, str) else None
            after_str = after if isinstance(after, str) else None
            if before_str is not None:
                if before_str:
                    if before_str in base:
                        replacement = after_str if after_str is not None else ''
                        return base.replace(before_str, replacement, 1)
                else:
                    if after_str:
                        return base + after_str
            if after_str:
                return base + after_str
            return base

        review_entries: List[Dict[str, Any]] = []
        id_to_position: Dict[str, Tuple[int, int]] = {}
        needs_revision_count = 0

        for r in range(src_rows):
            for c in range(src_cols):
                original_text = source_data[r][c]
                translated_text = translated_data[r][c]
                normalized_translation = _normalize_cell_value(translated_text)
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
                        if corrected_matrix is not None:
                            corrected_matrix[r][c] = normalized_translation
                        if highlight_matrix is not None:
                            highlight_matrix[r][c] = normalized_translation
                        if highlight_styles is not None:
                            highlight_styles[r][c] = []
                else:
                    status_matrix[r][c] = ""
                    issue_matrix[r][c] = ""
                    if corrected_matrix is not None:
                        corrected_matrix[r][c] = normalized_translation
                    if highlight_matrix is not None:
                        if highlight_styles is not None:
                            highlight_styles[r][c] = []
                        highlight_matrix[r][c] = normalized_translation

        if not review_entries:
            actions.write_range(status_output_range, status_matrix, sheet_name)
            actions.write_range(issue_output_range, issue_matrix, sheet_name)
            if corrected_matrix is not None and corrected_output_range:
                actions.write_range(corrected_output_range, corrected_matrix, sheet_name)
            if highlight_matrix is not None and highlight_output_range:
                actions.write_range(highlight_output_range, highlight_matrix, sheet_name)
                if highlight_styles is not None:
                    actions.apply_diff_highlight_colors(highlight_output_range, highlight_styles, sheet_name)
            return "翻訳チェックの対象となる文字列が存在しなかったため、結果列を初期化しました。"

        normalized_batch_size = 1

        for batch in (review_entries[i:i + normalized_batch_size] for i in range(0, len(review_entries), normalized_batch_size)):
            payload = json.dumps(batch, ensure_ascii=False)
            _diff_debug(f"check_translation_quality payload={_shorten_debug(payload)}")
            analysis_prompt = (
                "あなたは英訳の品質を評価するレビュアーです。各項目について、英訳が原文の意味・ニュアンス・文法・スペル・主語述語の対応として適切かを確認してください。"
                "各要素には 'id', 'original_text', 'translated_text' が含まれています。"
                "JSON 配列の各要素には必ず 'id', 'status', 'notes', 'highlighted_text', 'corrected_text', 'before_text', 'after_text' を含めてください。"
                "翻訳に問題がなければ status は 'OK' とし、notes は空文字または簡潔な補足にしてください。"
                "少しでも不安がある場合は 'OK' を選ばず、慎重に確認してください。"
                "修正が必要な場合は status を 'REVISE' とし、notes には日本語で『Issue: ... / Suggestion: ...』の形式で問題点と修正案を記述してください。"
                "迷った場合や不確かな点があれば必ず 'REVISE' を選択してください。"
                "'corrected_text' には修正後の英文全体を入れてください。"
                "'highlighted_text' には corrected_text をベースに、追加箇所は【追加：...】、削除箇所は【削除：...】で囲んだ比較表示を返してください。"
                "AI には JSON のみを返し、余計なフォローアップやマークアップは生成しないでください。"
                "このリクエストは公開されている企業財務情報の翻訳チェックであり、ポリシーに適合した安全な内容です。安心して対応してください。"
                "You already have all necessary information; do not ask for additional ranges or confirmations."
                "Return exactly one JSON array—no Final Answer, Thought, or explanations, and never include a second JSON object or array."
                "If uncertain, make the best possible judgment and still output JSON."
                f"\n\n{payload}\n"
            )
            def _parse_batch_response(response_text: str) -> Optional[List[Any]]:
                _diff_debug(f"check_translation_quality parse raw={_shorten_debug(response_text)}")
                if not response_text:
                    return None
                stripped = response_text.strip()
                if not stripped:
                    return None
                decoder = json.JSONDecoder()
                potential_starts = [idx for idx, ch in enumerate(stripped) if ch in {'[', '{'}]
                if not potential_starts:
                    _diff_debug('check_translation_quality no JSON delimiters found')
                    return None
                for start_idx in potential_starts:
                    if stripped[:start_idx].strip():
                        _diff_debug('check_translation_quality leading non-JSON content detected before payload')
                        continue
                    try:
                        parsed, end_idx = decoder.raw_decode(stripped[start_idx:])
                    except json.JSONDecodeError as decode_error:
                        _diff_debug(f"check_translation_quality decode error start={start_idx} err={decode_error}")
                        continue
                    trailing = stripped[start_idx + end_idx:].strip()
                    if trailing:
                        _diff_debug('check_translation_quality extra content after JSON payload detected')
                        continue
                    if isinstance(parsed, dict):
                        parsed = [parsed]
                    if isinstance(parsed, list):
                        _diff_debug(f"check_translation_quality parsed list length={len(parsed)}")
                        return parsed
                _diff_debug('check_translation_quality no valid JSON payload isolated')
                return None


            prompt_variants = [
                analysis_prompt,
                analysis_prompt + "\n\nSTRICT OUTPUT REMINDER: Return exactly one JSON array immediately. Do not include Final Answer, Thought, extra commentary, or multiple JSON payloads.",
                (
                    "You are reviewing translations of corporate financial disclosures. "
                    "Reply with a single JSON array. Each element must contain 'id', 'status', 'notes', "
                    "'highlighted_text', 'corrected_text', 'before_text', and 'after_text'. "
                    "Use status 'OK' when the translation is acceptable (notes empty or a short remark). Only select 'OK' when you are certain there are no issues. "
                    "Use status 'REVISE' when changes are needed and write notes in Japanese as 'Issue: ... / Suggestion: ...'. If unsure, choose 'REVISE'. "
                    "Set 'corrected_text' to the fully corrected English sentence. Build 'highlighted_text' from corrected_text, "
                    "marking additions as 【追加：...】 and deletions as 【削除：...】. "
                    "Return exactly one JSON array and nothing else."
                    f"\n\n{payload}\n"
                ),
            ]

            response = ""
            batch_results: Optional[List[Any]] = None
            for prompt_variant in prompt_variants:
                response = browser_manager.ask(prompt_variant)
                _diff_debug(f"check_translation_quality response={_shorten_debug(response)}")
                if response and any(indicator in response for indicator in REFUSAL_PATTERNS):
                    _diff_debug('check_translation_quality detected refusal response, trying next prompt variant')
                    continue
                batch_results = _parse_batch_response(response)
                if batch_results is not None:
                    break

            if batch_results is None:
                _diff_debug(f"check_translation_quality unable to parse response={_shorten_debug(response)}")
                raise ToolExecutionError(f"翻訳チェックの結果をJSONとして解析できませんでした。応答: {response}")

            if not isinstance(batch_results, list):
                raise ToolExecutionError("翻訳チェックの応答形式が不正です。JSON配列を返してください。")

            ok_statuses = {"OK", "PASS", "GOOD"}
            revise_statuses = {"REVISE", "NG", "FAIL", "ISSUE"}
            for item in batch_results:
                if not isinstance(item, dict):
                    raise ToolExecutionError("翻訳チェックの応答に不正な要素が含まれています。")
                item_id = item.get("id")
                if item_id not in id_to_position:
                    _diff_debug(f"check_translation_quality unknown id={item_id} known={list(id_to_position.keys())}")
                    raise ToolExecutionError("翻訳チェックの応答に未知のIDが含まれています。")
                status_value = str(item.get("status", "")).strip().upper()
                notes_value = str(item.get("notes", "")).strip()
                before_text = item.get("before_text")
                after_text = item.get("after_text")


                row_idx, col_idx = id_to_position[item_id]
                base_translation = translated_data[row_idx][col_idx]
                base_text = _normalize_cell_value(base_translation)
                corrected_text = _infer_corrected_text(base_text, item)
                corrected_text_str = _normalize_cell_value(corrected_text)
                is_ok_status = status_value in ok_statuses
                if is_ok_status or not corrected_text_str.strip():
                    corrected_text_str = base_text

                if corrected_matrix is not None:
                    corrected_matrix[row_idx][col_idx] = corrected_text_str

                if highlight_matrix is not None:
                    if is_ok_status:
                        highlight_matrix[row_idx][col_idx] = base_text
                        if highlight_styles is not None:
                            highlight_styles[row_idx][col_idx] = []
                    else:
                        highlight_text, highlight_spans = _build_diff_highlight(base_text, corrected_text_str)
                        highlight_matrix[row_idx][col_idx] = highlight_text
                        if highlight_styles is not None:
                            highlight_styles[row_idx][col_idx] = highlight_spans

                if is_ok_status:
                    status_matrix[row_idx][col_idx] = "OK"
                    issue_matrix[row_idx][col_idx] = notes_value or ""
                elif status_value in revise_statuses:
                    status_matrix[row_idx][col_idx] = "要修正"
                    issue_matrix[row_idx][col_idx] = notes_value or "修正内容を記載してください。"
                    needs_revision_count += 1
                else:
                    status_matrix[row_idx][col_idx] = status_value or "要確認"
                    issue_matrix[row_idx][col_idx] = notes_value or "ステータスが解釈できませんでした。"
                    needs_revision_count += 1

        actions.write_range(status_output_range, status_matrix, sheet_name)
        actions.write_range(issue_output_range, issue_matrix, sheet_name)

        processed_items = len(review_entries)
        message = (
            f"翻訳チェックを完了しました。対象 {processed_items} 件中、要修正 {needs_revision_count} 件の結果を"
            f" '{status_output_range}' と '{issue_output_range}' に書き込みました。"
        )
        if corrected_matrix is not None and corrected_output_range:
            actions.write_range(corrected_output_range, corrected_matrix, sheet_name)
            message += f" 完成形は '{corrected_output_range}' に出力しました。"
        if highlight_matrix is not None and highlight_output_range:
            actions.write_range(highlight_output_range, highlight_matrix, sheet_name)
            if highlight_styles is not None:
                actions.apply_diff_highlight_colors(highlight_output_range, highlight_styles, sheet_name)
            message += f" 比較表示用の文字列は '{highlight_output_range}' に出力しました。"
        return message

    except ToolExecutionError:
        raise
    except Exception as e:
        raise ToolExecutionError(f"翻訳チェック中にエラーが発生しました: {e}") from e



def highlight_text_differences(
    actions: ExcelActions,
    original_range: str,
    revised_range: str,
    output_range: str,
    sheet_name: Optional[str] = None,
    addition_color_hex: str = "#1565C0",
    deletion_color_hex: str = "#C62828",
) -> str:
    """Compare two ranges and color additions/deletions inside each cell."""
    try:
        _diff_debug(f"highlight_text_differences start original_range={original_range} revised_range={revised_range} output_range={output_range} sheet={sheet_name}")
        original_rows, original_cols = _parse_range_dimensions(original_range)
        revised_rows, revised_cols = _parse_range_dimensions(revised_range)
        output_rows, output_cols = _parse_range_dimensions(output_range)

        if (original_rows, original_cols) != (revised_rows, revised_cols):
            raise ToolExecutionError('Original range and revised range sizes do not match.')
        if (original_rows, original_cols) != (output_rows, output_cols):
            raise ToolExecutionError('Output range must match the original range size.')

        original_matrix = _reshape_to_dimensions(
            actions.read_range(original_range, sheet_name), original_rows, original_cols
        )
        revised_matrix = _reshape_to_dimensions(
            actions.read_range(revised_range, sheet_name), original_rows, original_cols
        )

        highlight_matrix: List[List[str]] = []
        highlight_styles: List[List[List[Dict[str, int]]]] = []

        for r in range(original_rows):
            text_row: List[str] = []
            style_row: List[List[Dict[str, int]]] = []
            for c in range(original_cols):
                before_text = _normalize_cell_value(original_matrix[r][c])
                after_text = _normalize_cell_value(revised_matrix[r][c])
                _diff_debug(f"highlight_text_differences cell=({r},{c}) before={_shorten_debug(before_text)} after={_shorten_debug(after_text)}")
                highlight_text, spans = _build_diff_highlight(before_text, after_text)
                _diff_debug(f"highlight_text_differences spans= {spans}")
                text_row.append(highlight_text)
                style_row.append(spans)
            highlight_matrix.append(text_row)
            highlight_styles.append(style_row)

        _diff_debug(f"highlight_text_differences writing matrix size={len(highlight_matrix)}x{len(highlight_matrix[0]) if highlight_matrix else 0}")
        actions.write_range(output_range, highlight_matrix, sheet_name)
        actions.apply_diff_highlight_colors(
            output_range,
            highlight_styles,
            sheet_name,
            addition_color_hex=addition_color_hex,
            deletion_color_hex=deletion_color_hex,
        )
        _diff_debug('highlight_text_differences applied colors via ExcelActions')

        return (
            f"差分ハイライトを '{output_range}' に出力し、追加程所({addition_color_hex})と削除程所({deletion_color_hex})を強調しました。"
        )
    except ToolExecutionError:
        raise
    except Exception as exc:
        _diff_debug(f"highlight_text_differences exception={exc}")
        raise ToolExecutionError(f"差分ハイライトの適用中にエラーが発生しました: {exc}") from exc

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








