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



def _split_sheet_and_range(range_ref: str, default_sheet: Optional[str]) -> Tuple[Optional[str], str]:
    cleaned = (range_ref or "").strip()
    if not cleaned:
        raise ToolExecutionError("Range string is empty.")
    if "!" not in cleaned:
        return default_sheet, cleaned
    sheet_part, cell_part = cleaned.split("!", 1)
    sheet_part = sheet_part.strip()
    if sheet_part.startswith("'") and sheet_part.endswith("'"):
        sheet_part = sheet_part[1:-1].replace("''", "'")
    elif sheet_part.startswith('"') and sheet_part.endswith('"'):
        sheet_part = sheet_part[1:-1]
    cell_part = cell_part.strip()
    if not cell_part:
        raise ToolExecutionError("Range string is empty.")
    return sheet_part or default_sheet, cell_part


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
    sheet_name: Optional[str] = None,
    reference_ranges: Optional[List[str]] = None,
    citation_output_range: Optional[str] = None,
    reference_urls: Optional[List[str]] = None,
    translation_output_range: Optional[str] = None,
    overwrite_source: bool = False,
    rows_per_batch: Optional[int] = None,
) -> str:
    """Translate text ranges with optional references and controlled output."""
    try:
        target_sheet, normalized_range = _split_sheet_and_range(cell_range, sheet_name)
        source_rows, source_cols = _parse_range_dimensions(normalized_range)

        raw_original = actions.read_range(normalized_range, target_sheet)
        original_data = _reshape_to_dimensions(raw_original, source_rows, source_cols)

        if source_rows == 0 or source_cols == 0:
            return f"範囲 '{cell_range}' に翻訳対象のテキストが見つかりませんでした。"

        writing_to_source_directly = translation_output_range is None
        if writing_to_source_directly:
            if not overwrite_source:
                raise ToolExecutionError(
                    "翻訳結果の出力先が指定されていません。translation_output_range を指定するか"
                    " overwrite_source を True にしてください。"
                )
            output_sheet = target_sheet
            output_range = normalized_range
            output_matrix = [row[:] for row in original_data]
        else:
            output_sheet, output_range = _split_sheet_and_range(translation_output_range, target_sheet)
            out_rows, out_cols = _parse_range_dimensions(output_range)
            if (out_rows, out_cols) != (source_rows, source_cols):
                raise ToolExecutionError(
                    "translation_output_range のサイズは翻訳対象範囲と一致させてください。"
                )
            raw_output = actions.read_range(output_range, output_sheet)
            try:
                output_matrix = _reshape_to_dimensions(raw_output, out_rows, out_cols)
            except ToolExecutionError:
                output_matrix = [["" for _ in range(out_cols)] for _ in range(out_rows)]

        reference_entries: List[Dict[str, Any]] = []
        if reference_ranges:
            range_list = [reference_ranges] if isinstance(reference_ranges, str) else list(reference_ranges)
            for raw_range in range_list:
                ref_sheet, ref_range = _split_sheet_and_range(raw_range, target_sheet)
                try:
                    ref_data = actions.read_range(ref_range, ref_sheet)
                except ToolExecutionError as exc:
                    raise ToolExecutionError(f"参照文献範囲 '{raw_range}' の読み取りに失敗しました: {exc}") from exc

                reference_lines: List[str] = []
                if isinstance(ref_data, list):
                    for ref_row in ref_data:
                        if isinstance(ref_row, list):
                            row_text = []
                            for value in ref_row:
                                text_value = _normalize_cell_value(value).strip()
                                if text_value:
                                    row_text.append(text_value)
                            if row_text:
                                reference_lines.append(" ".join(row_text))
                        else:
                            text_value = _normalize_cell_value(ref_row).strip()
                            if text_value:
                                reference_lines.append(text_value)
                else:
                    text_value = _normalize_cell_value(ref_data).strip()
                    if text_value:
                        reference_lines.append(text_value)

                if reference_lines:
                    entry: Dict[str, Any] = {
                        "id": f"R{len(reference_entries) + 1}",
                        "source_range": raw_range,
                        "content": reference_lines,
                    }
                    if ref_sheet:
                        entry["sheet"] = ref_sheet
                    reference_entries.append(entry)

            if reference_ranges and not reference_entries:
                raise ToolExecutionError("指定された参照文献範囲から利用可能なテキストを取得できませんでした。")

        reference_url_entries: List[Dict[str, str]] = []
        if reference_urls:
            url_list = [reference_urls] if isinstance(reference_urls, str) else list(reference_urls)
            for raw_url in url_list:
                if not isinstance(raw_url, str):
                    raise ToolExecutionError("reference_urls の各要素は文字列で指定してください。")
                url = raw_url.strip()
                if not url:
                    continue
                reference_url_entries.append({
                    "id": f"U{len(reference_url_entries) + 1}",
                    "url": url,
                })

        use_references = bool(reference_entries or reference_url_entries)

        prompt_parts: List[str]
        if use_references:
            prompt_parts = [
                f"以下のJSONリストに格納された日本語テキストを、それぞれ{target_language}に翻訳してください。\n",
                "翻訳では必ず提供された参照文献やURLの文章を根拠として使用し、引用した表現を活かした自然な訳文を作成してください。\n",
                "各翻訳文には対応する参照ID（例: [R1], [U2]）を本文内に含めてください。\n",
                "入力テキストの順序は維持してください。\n",
            ]
            if reference_entries:
                prompt_parts.append(f"参照文献リスト:\n{json.dumps(reference_entries, ensure_ascii=False)}\n")
            if reference_url_entries:
                prompt_parts.append(f"参照可能なURLリスト:\n{json.dumps(reference_url_entries, ensure_ascii=False)}\n")
            prompt_parts.append(
                "応答はJSON配列のみとし、各要素は必ず次のキーを含めてください:\n"
                "- \"translated_text\": 翻訳結果の文字列（必要な参照IDを含む）\n"
                "- \"evidence\": 翻訳に使用した参照文やURLの文章を複数含む文字列配列\n"
                "前後に説明文やコードブロックを含めないでください。\n"
            )
            prompt_preamble = "".join(prompt_parts)
        else:
            prompt_preamble = (
                f"以下のJSONリストに格納された日本語テキストを、それぞれ{target_language}に翻訳し、"
                "翻訳後のテキストを格納したJSONリスト形式で返してください。\n"
                "入力テキストの順序は維持してください。\n"
                "応答はJSONのみとし、前後に説明やコードブロックのマークアップを含めないでください。\n"
            )

        batch_size = rows_per_batch if rows_per_batch is not None else 1
        if batch_size < 1:
            raise ToolExecutionError("rows_per_batch は 1 以上で指定してください。")

        translations: Dict[Tuple[int, int], str] = {}
        evidences: Dict[Tuple[int, int], str] = {}

        for row_start in range(0, source_rows, batch_size):
            row_end = min(row_start + batch_size, source_rows)
            chunk_texts: List[str] = []
            chunk_positions: List[Tuple[int, int]] = []

            for row_idx in range(row_start, row_end):
                for col_idx in range(source_cols):
                    cell_value = original_data[row_idx][col_idx]
                    if isinstance(cell_value, str) and re.search(r"[ぁ-んァ-ン一-龯]", cell_value):
                        chunk_texts.append(cell_value)
                        chunk_positions.append((row_idx, col_idx))

            if not chunk_texts:
                continue

            texts_json = json.dumps(chunk_texts, ensure_ascii=False)
            prompt = f"{prompt_preamble}{texts_json}"
            response = browser_manager.ask(prompt)

            try:
                match = re.search(r"{.*}|\[.*\]", response, re.DOTALL)
                json_payload = match.group(0) if match else response
                parsed_payload = json.loads(json_payload)
            except json.JSONDecodeError as exc:
                raise ToolExecutionError(
                    f"AIからの翻訳結果をJSONとして解析できませんでした。応答: {response}"
                ) from exc

            if use_references:
                if not isinstance(parsed_payload, list) or len(parsed_payload) != len(chunk_texts):
                    raise ToolExecutionError("翻訳前と翻訳後でテキストの件数が一致しません。")
                for item, position in zip(parsed_payload, chunk_positions):
                    if not isinstance(item, dict):
                        raise ToolExecutionError(
                            "参照文献やURLを利用する場合、翻訳結果はオブジェクトのリストで返してください。"
                        )
                    translation_value = item.get("translated_text") or item.get("translation") or item.get("output")
                    if not isinstance(translation_value, str):
                        raise ToolExecutionError("翻訳結果のJSONに 'translated_text' が含まれていません。")
                    translations[position] = translation_value

                    evidence_value = item.get("evidence") or item.get("justification") or item.get("support")
                    if isinstance(evidence_value, list):
                        collected = [str(v).strip() for v in evidence_value if isinstance(v, (str, int, float))]
                        evidences[position] = "\n\n".join(filter(None, collected))
                    elif isinstance(evidence_value, str):
                        evidences[position] = evidence_value.strip()
                    elif evidence_value is None:
                        evidences[position] = ""
                    else:
                        evidences[position] = str(evidence_value)
            else:
                if not isinstance(parsed_payload, list) or len(parsed_payload) != len(chunk_texts):
                    raise ToolExecutionError("翻訳前と翻訳後でテキストの件数が一致しません。")
                if not all(isinstance(item, str) for item in parsed_payload):
                    raise ToolExecutionError("翻訳結果は文字列のリストで返してください。")
                for translation_value, position in zip(parsed_payload, chunk_positions):
                    translations[position] = translation_value

        if not translations:
            return f"範囲 '{cell_range}' に翻訳対象のテキストが見つかりませんでした。"

        for (row_idx, col_idx), translated_text in translations.items():
            output_matrix[row_idx][col_idx] = translated_text

        messages: List[str] = []
        messages.append(actions.write_range(output_range, output_matrix, output_sheet))

        if not writing_to_source_directly and overwrite_source:
            source_matrix = [row[:] for row in original_data]
            for (row_idx, col_idx), translated_text in translations.items():
                source_matrix[row_idx][col_idx] = translated_text
            messages.append(actions.write_range(normalized_range, source_matrix, target_sheet))

        if use_references:
            if not citation_output_range:
                raise ToolExecutionError(
                    "参照文献が指定された場合は、根拠を書き込む範囲 (citation_output_range) を指定してください。"
                )

            citation_sheet, citation_range = _split_sheet_and_range(citation_output_range, target_sheet)
            cite_rows, cite_cols = _parse_range_dimensions(citation_range)

            if cite_rows != source_rows:
                raise ToolExecutionError("citation_output_range の行数は翻訳対象範囲と一致させてください。")
            if cite_cols not in {1, source_cols}:
                raise ToolExecutionError(
                    "citation_output_range の列数は1列または翻訳対象範囲と同じ列数にしてください。"
                )

            existing_citation = actions.read_range(citation_range, citation_sheet)
            try:
                citation_matrix = _reshape_to_dimensions(existing_citation, cite_rows, cite_cols)
            except ToolExecutionError:
                citation_matrix = [["" for _ in range(cite_cols)] for _ in range(cite_rows)]

            if cite_cols == source_cols:
                for row_idx in range(cite_rows):
                    for col_idx in range(cite_cols):
                        citation_matrix[row_idx][col_idx] = ""
                for (row_idx, col_idx), evidence_text in evidences.items():
                    if row_idx < cite_rows and col_idx < cite_cols:
                        citation_matrix[row_idx][col_idx] = evidence_text
            else:
                for row_idx in range(cite_rows):
                    entries = [evidences[pos] for pos in evidences if pos[0] == row_idx and evidences[pos]]
                    citation_matrix[row_idx][0] = "\n\n".join(entries)

            messages.append(actions.write_range(citation_range, citation_matrix, citation_sheet))

        return "\n".join(messages)

    except ToolExecutionError:
        raise
    except Exception as e:
        raise ToolExecutionError(f"範囲 '{cell_range}' の翻訳中にエラーが発生しました: {e}") from e

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








