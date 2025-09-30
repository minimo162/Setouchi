import re
import difflib
import logging
import os
import string
from typing import List, Any, Optional, Dict, Tuple, Set

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
    return text if len(text) <= limit else text[:limit] + '窶ｦ'



_MIN_QUOTE_TOKEN_COVERAGE = 0.5
_PUNCT_TRANSLATION_TABLE = {ord(ch): ' ' for ch in string.punctuation}
_PUNCT_TRANSLATION_TABLE.update({
    0x2010: ' ',
    0x2011: ' ',
    0x2012: ' ',
    0x2013: ' ',
    0x2014: ' ',
    0x2212: ' ',
})

def _normalize_for_match(text: str) -> str:
    if not isinstance(text, str):
        return ''
    replacements = {
        chr(0x201c): '"',
        chr(0x201d): '"',
        chr(0x2019): "'",
        chr(0x2018): "'",
    }
    normalized = text
    for src, dst in replacements.items():
        normalized = normalized.replace(src, dst)
    normalized = normalized.translate(_PUNCT_TRANSLATION_TABLE)
    normalized = re.sub(r'\s+', ' ', normalized)
    return normalized.strip().lower()



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

def _column_label_to_index(label: str) -> int:
    result = 0
    for ch in label.upper():
        if not ('A' <= ch <= 'Z'):
            raise ToolExecutionError('Range format is invalid.')
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result


def _index_to_column_label(index: int) -> str:
    if index <= 0:
        raise ToolExecutionError('Column index must be positive.')
    label_chars: List[str] = []
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        label_chars.append(chr(ord('A') + remainder))
    return ''.join(reversed(label_chars))


def _parse_range_bounds(range_ref: str) -> Tuple[int, int, int, int]:
    ref = range_ref.split('!')[-1].replace('$', '').strip()
    if not ref:
        raise ToolExecutionError('Range string is empty.')
    if ':' in ref:
        start_ref, end_ref = ref.split(':', 1)
    else:
        start_ref = end_ref = ref
    start_match = CELL_REFERENCE_PATTERN.fullmatch(start_ref)
    end_match = CELL_REFERENCE_PATTERN.fullmatch(end_ref)
    if not start_match or not end_match:
        raise ToolExecutionError('Range format is invalid.')
    start_col = _column_label_to_index(start_match.group(1)) - 1
    start_row = int(start_match.group(2)) - 1
    end_col = _column_label_to_index(end_match.group(1)) - 1
    end_row = int(end_match.group(2)) - 1
    if start_row > end_row:
        start_row, end_row = end_row, start_row
    if start_col > end_col:
        start_col, end_col = end_col, start_col
    return start_row, start_col, end_row, end_col


def _build_range_reference(start_row: int, end_row: int, start_col: int, end_col: int) -> str:
    start_cell = f"{_index_to_column_label(start_col + 1)}{start_row + 1}"
    if start_row == end_row and start_col == end_col:
        return start_cell
    end_cell = f"{_index_to_column_label(end_col + 1)}{end_row + 1}"
    return f"{start_cell}:{end_cell}"



CELL_REFERENCE_PATTERN = re.compile(r"([A-Za-z]+)(\d+)")

LEGACY_DIFF_MARKER_PATTERN = re.compile(r"【(追加|削除)：(.*?)】")
MODERN_DIFF_MARKER_PATTERN = re.compile(r"\?(?:追加|削除)\?\s*(.*?)\?")
_BASE_DIFF_TOKEN_PATTERN = re.compile(r"\s+|[^\s]+")
_SENTENCE_BOUNDARY_CHARS = set("!.?。！？")
_CLOSING_PUNCTUATION = ")]}、。！？」』】》）］'\"”’
_MAX_DIFF_SEGMENT_TOKENS = 18
_MAX_DIFF_SEGMENT_CHARS = 80

REFUSAL_PATTERNS = (
    "逕ｳ縺苓ｨｳ縺斐＊縺・EE縺帙ｓ縲ゅ％繧後↓縺､縺・EE繝√Ε繝・EE縺ｧ縺阪∪縺帙ｓ縲・,
    "逕ｳ縺苓ｨｳ縺斐＊縺・EE縺帙ｓ縲ゅ％繧後↓縺､縺・EE繝√Ε繝・EE縺ｧ縺阪∪縺帙ｓ",
    "逕ｳ縺苓ｨｳ縺斐＊縺・EE縺帙ｓ縲ゅメ繝｣繝・EE繧剃ｿ晏ｭ倥＠縺ｦ譁ｰ縺励＞繝√Ε繝・EE繧帝幕蟋九☆繧九↓縺ｯ縲ー譁ｰ縺励＞繝√Ε繝・EE] 繧帝∈謚槭＠縺ｦ縺上□縺輔＞縲・,
    "繝√Ε繝・EE繧剃ｿ晏ｭ倥＠縺ｦ譁ｰ縺励＞繝√Ε繝・EE繧帝幕蟋九☆繧九↓縺ｯ縲ー譁ｰ縺励＞繝√Ε繝・EE] 繧帝∈謚槭＠縺ｦ縺上□縺輔＞縲・,
    "縺顔ｭ斐∴縺ｧ縺阪∪縺帙ｓ縲・,
    "縺顔ｭ斐∴縺ｧ縺阪∪縺帙ｓ",
    "I'm sorry, but I can't help with that.",
    "I cannot help with that request.",
    "繧ｨ繝ｩ繝ｼ縺檎匱逕溘＠縺ｾ縺励◆: 蠢懃ｭ泌ｽ｢蠑上′荳肴ｭ｣縺ｧ縺吶・Thought:' 縺ｾ縺滂ｿｽE 'Final Answer:' 縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲・,
    "蠢懃ｭ泌ｽ｢蠑上′荳肴ｭ｣縺ｧ縺吶・Thought:' 縺ｾ縺滂ｿｽE 'Final Answer:' 縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲・,
)

JAPANESE_CHAR_PATTERN = re.compile(r'[縺-繝ｿ荳-鯀ｿ]')



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

    start_col = _column_label_to_index(start_match.group(1))
    end_col = _column_label_to_index(end_match.group(1))
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
    marker_prefix = f'縲須label}EEEE
    marker_suffix = '縲・
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
            formatted_removed, offset_removed, length_removed = _format_diff_segment(removed_tokens, '蜑企勁')
            if formatted_removed:
                result_parts.append(formatted_removed)
                if offset_removed is not None and length_removed:
                    span = {'start': cursor + offset_removed, 'length': length_removed, 'type': '蜑企勁'}
                    spans.append(span)
                    _diff_debug(f"_build_diff_highlight span added {span}")
                cursor += len(formatted_removed)
            added_tokens = corr_tokens[j1:j2]
            formatted_added, offset_added, length_added = _format_diff_segment(added_tokens, '霑ｽ蜉')
            if formatted_added:
                result_parts.append(formatted_added)
                if offset_added is not None and length_added:
                    span = {'start': cursor + offset_added, 'length': length_added, 'type': '霑ｽ蜉'}
                    spans.append(span)
                    _diff_debug(f"_build_diff_highlight span added {span}")
                cursor += len(formatted_added)
        elif tag == 'delete':
            removed_tokens = orig_tokens[i1:i2]
            formatted_removed, offset_removed, length_removed = _format_diff_segment(removed_tokens, '蜑企勁')
            if formatted_removed:
                result_parts.append(formatted_removed)
                if offset_removed is not None and length_removed:
                    span = {'start': cursor + offset_removed, 'length': length_removed, 'type': '蜑企勁'}
                    spans.append(span)
                    _diff_debug(f"_build_diff_highlight span added {span}")
                cursor += len(formatted_removed)
        elif tag == 'insert':
            added_tokens = corr_tokens[j1:j2]
            formatted_added, offset_added, length_added = _format_diff_segment(added_tokens, '霑ｽ蜉')
            if formatted_added:
                result_parts.append(formatted_added)
                if offset_added is not None and length_added:
                    span = {'start': cursor + offset_added, 'length': length_added, 'type': '霑ｽ蜉'}
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
    Excel繧ｷ繝ｼ繝茨ｿｽE迚ｹ螳夲ｿｽE繧ｻ繝ｫ縺ｫ蛟､繧呈嶌縺崎ｾｼ縺ｿ縺ｾ縺吶・
    """
    return actions.write_to_cell(cell, value, sheetname)

def readcellvalue(actions: ExcelActions, cell: str, sheetname: Optional[str] = None) -> Any:
    """
    Excel繧ｷ繝ｼ繝茨ｿｽE迚ｹ螳夲ｿｽE繧ｻ繝ｫ縺ｮ蛟､繧定ｪｭ縺ｿ蜿悶ｊ縺ｾ縺吶・
    """
    return actions.read_cell_value(cell, sheetname)

def getallsheetnames(actions: ExcelActions) -> str:
    """
    迴ｾ蝨ｨ髢九＞縺ｦ縺・EEExcel繝ｯ繝ｼ繧ｯ繝悶ャ繧ｯ蜀・EE縺吶∋縺ｦ縺ｮ繧ｷ繝ｼ繝亥錐繧貞叙蠕励＠縺ｾ縺吶・
    """
    names = actions.get_sheet_names()
    return f"蛻ｩ逕ｨ蜿ｯ閭ｽ縺ｪ繧ｷ繝ｼ繝茨ｿｽE谺｡縺ｮ騾壹ｊ縺ｧ縺・ {', '.join(names)}"

def copyrange(actions: ExcelActions, sourcerange: str, destinationrange: str, sheetname: Optional[str] = None) -> str:
    """
    謖・EE縺励◆遽・EE繧貞挨縺ｮ蝣ｴ謇縺ｫ繧ｳ繝費ｿｽE縺励∪縺吶・
    """
    return actions.copy_range(sourcerange, destinationrange, sheetname)

def executeexcelformula(actions: ExcelActions, cell: str, formula: str, sheetname: Optional[str] = None) -> str:
    """
    謖・EE縺励◆繧ｻ繝ｫ縺ｫExcel縺ｮ謨ｰ蠑上ｒ險ｭ螳壹＠縺ｾ縺吶・
    """
    return actions.set_formula(cell, formula, sheetname)

def readrangevalues(actions: ExcelActions, cellrange: str, sheetname: Optional[str] = None) -> str:
    """
    謖・EE縺励◆遽・EE縺ｮ繧ｻ繝ｫ縺九ｉ蛟､繧定ｪｭ縺ｿ蜿悶ｊ縺ｾ縺吶・繧ｻ繝ｫ縺ｧ繧らｯ・EE縺ｨ縺励※謖・EE蜿ｯ閭ｽ縺ｧ縺吶・
    """
    values = actions.read_range(cellrange, sheetname)
    return f"遽・EE '{cellrange}' 縺ｮ蛟､縺ｯ谺｡縺ｮ騾壹ｊ縺ｧ縺・ {values}"

def writerangevalues(actions: ExcelActions, cellrange: str, data: List[List[Any]], sheetname: Optional[str] = None) -> str:
    """
    謖・EE縺励◆遽・EE縺ｫ2谺｡蜈・EE繧ｹ繝茨ｿｽE繝・EE繧ｿ繧呈嶌縺崎ｾｼ縺ｿ縺ｾ縺吶・繧ｻ繝ｫ縺ｧ繧ょｯｾ蠢懷庄閭ｽ縺ｧ縺吶・
    """
    return actions.write_range(cellrange, data, sheetname)

def getactiveworkbookandsheet(actions: ExcelActions) -> str:
    """
    迴ｾ蝨ｨ繧｢繧ｯ繝・EE繝悶↑Excel繝悶ャ繧ｯ縺ｨ繧ｷ繝ｼ繝亥錐繧貞叙蠕励＠縺ｾ縺吶・
    """
    info_dict = actions.get_active_workbook_and_sheet()
    return f"繝悶ャ繧ｯ: {info_dict['workbook_name']}, 繧ｷ繝ｼ繝・ {info_dict['sheet_name']}"

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
    謖・EE縺励◆遽・EE縺ｫ譖ｸ蠑剰ｨｭ螳壹ｒ驕ｩ逕ｨ縺励∪縺吶・
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
    '''Translate Japanese text for a range with optional references and controlled output.'''
    try:
        target_sheet, normalized_range = _split_sheet_and_range(cell_range, sheet_name)
        source_rows, source_cols = _parse_range_dimensions(normalized_range)

        raw_original = actions.read_range(normalized_range, target_sheet)
        original_data = _reshape_to_dimensions(raw_original, source_rows, source_cols)

        if source_rows == 0 or source_cols == 0:
            return f"遽・EE '{cell_range}' 縺ｫ鄙ｻ險ｳ蟇ｾ雎｡縺ｮ繝・EE繧ｹ繝医′隕九▽縺九ｊ縺ｾ縺帙ｓ縺ｧ縺励◆縲・

        source_matrix = [row[:] for row in original_data]
        writing_to_source_directly = translation_output_range is None
        if writing_to_source_directly and not overwrite_source:
            raise ToolExecutionError(
                "鄙ｻ險ｳ邨先棡縺ｮ蜃ｺ蜉幢ｿｽE縺梧欠螳壹＆繧後※縺・EE縺帙ｓ縲Ｕranslation_output_range 繧呈欠螳壹☆繧九° overwrite_source 繧・True 縺ｫ縺励※縺上□縺輔＞縲・
            )
        if writing_to_source_directly:
            output_sheet = target_sheet
            output_range = normalized_range
            output_matrix = source_matrix
        else:
            output_sheet, output_range = _split_sheet_and_range(translation_output_range, target_sheet)
            out_rows, out_cols = _parse_range_dimensions(output_range)
            if (out_rows, out_cols) != (source_rows, source_cols):
                raise ToolExecutionError("translation_output_range 縺ｮ繧ｵ繧､繧ｺ縺ｯ鄙ｻ險ｳ蟇ｾ雎｡遽・EE縺ｨ荳閾ｴ縺輔○縺ｦ縺上□縺輔＞縲・)
            raw_output = actions.read_range(output_range, output_sheet)
            try:
                output_matrix = _reshape_to_dimensions(raw_output, out_rows, out_cols)
            except ToolExecutionError:
                output_matrix = [["" for _ in range(out_cols)] for _ in range(out_rows)]

        reference_entries: List[Dict[str, Any]] = []
        reference_text_pool: List[str] = []
        if reference_ranges:
            range_list = [reference_ranges] if isinstance(reference_ranges, str) else list(reference_ranges)
            for raw_range in range_list:
                ref_sheet, ref_range = _split_sheet_and_range(raw_range, target_sheet)
                try:
                    ref_data = actions.read_range(ref_range, ref_sheet)
                except ToolExecutionError as exc:
                    raise ToolExecutionError(f"蜿ゑｿｽE譁・EE遽・EE '{raw_range}' 縺ｮ隱ｭ縺ｿ蜿悶ｊ縺ｫ螟ｱ謨励＠縺ｾ縺励◆: {exc}") from exc

                reference_lines: List[str] = []
                if isinstance(ref_data, list):
                    for ref_row in ref_data:
                        if isinstance(ref_row, list):
                            row_text: List[str] = []
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
                    reference_text_pool.extend(reference_lines)

            if reference_ranges and not reference_entries:
                raise ToolExecutionError("謖・EE縺輔ｌ縺溷盾EE譁・EE遽・EE縺九ｉ蛻ｩ逕ｨ蜿ｯ閭ｽ縺ｪ繝・EE繧ｹ繝医ｒ蜿門ｾ励〒縺阪∪縺帙ｓ縺ｧ縺励◆縲・)

        reference_url_entries: List[Dict[str, str]] = []
        if reference_urls:
            url_list = [reference_urls] if isinstance(reference_urls, str) else list(reference_urls)
            for raw_url in url_list:
                if not isinstance(raw_url, str):
                    raise ToolExecutionError("reference_urls 縺ｮ蜷・EE邏縺ｯ譁・EEEE縺ｧ謖・EE縺励※縺上□縺輔＞縲・)
                url = raw_url.strip()
                if not url:
                    continue
                reference_url_entries.append({
                    "id": f"U{len(reference_url_entries) + 1}",
                    "url": url,
                })

        use_references = bool(reference_entries or reference_url_entries)
        reference_text_pool = [text for text in reference_text_pool if text]
        normalized_reference_text_pool: List[str] = []
        for _ref_text in reference_text_pool:
            normalized = _normalize_for_match(_ref_text)
            if normalized:
                normalized_reference_text_pool.append(normalized)

        def _sanitize_evidence_value(value: str) -> str:
            cleaned = value.strip()
            if cleaned.lower().startswith("source:"):
                cleaned = cleaned.split(":", 1)[1].strip()
            return cleaned

        def _expand_keyword_variants(keywords: List[str]) -> List[str]:
            seen: Set[str] = set()
            expanded: List[str] = []

            def _add_candidate(candidate: str) -> None:
                value = candidate.strip()
                if not value:
                    return
                lowered = value.lower()
                if lowered in seen:
                    return
                seen.add(lowered)
                expanded.append(value)

            dash_variants = ['-'] + [chr(code) for code in (0x2010, 0x2011, 0x2012, 0x2013, 0x2014)]
            fullwidth_space = chr(0x3000)

            for keyword in keywords:
                base = keyword.strip()
                if not base:
                    continue
                _add_candidate(base)

                normalized = base.replace(fullwidth_space, ' ').strip()
                for alt_dash in dash_variants[1:]:
                    normalized = normalized.replace(alt_dash, '-')

                if '-' in normalized:
                    _add_candidate(normalized.replace('-', ' '))
                if ' ' in normalized:
                    _add_candidate(normalized.replace(' ', '-'))
                    words = [word for word in normalized.split() if word]
                    for word in words:
                        _add_candidate(word)
                    if len(words) >= 2:
                        acronym = ''.join(word[0] for word in words if word and word[0].isalpha()).upper()
                        if len(acronym) >= 2:
                            _add_candidate(acronym)

                if '(' in normalized and ')' in normalized:
                    start_paren = normalized.find('(') + 1
                    end_paren = normalized.find(')', start_paren)
                    if end_paren > start_paren:
                        _add_candidate(normalized[start_paren:end_paren])

                punctuation_stripped = normalized.strip(',:;')
                if punctuation_stripped != normalized:
                    _add_candidate(punctuation_stripped)

            max_variants = 6
            return expanded[:max_variants]

        def _expand_keyword_variants(keywords: List[str]) -> List[str]:
            seen: Set[str] = set()
            expanded: List[str] = []

            def _add_candidate(candidate: str) -> None:
                value = candidate.strip()
                if not value:
                    return
                lowered = value.lower()
                if lowered in seen:
                    return
                seen.add(lowered)
                expanded.append(value)

            dash_variants = ['-'] + [chr(code) for code in (0x2010, 0x2011, 0x2012, 0x2013, 0x2014)]

            for keyword in keywords:
                base = keyword.strip()
                if not base:
                    continue
                _add_candidate(base)

                normalized = base.replace('縲', ' ').strip()
                for alt_dash in dash_variants[1:]:
                    normalized = normalized.replace(alt_dash, '-')

                if '-' in normalized:
                    _add_candidate(normalized.replace('-', ' '))
                if ' ' in normalized:
                    _add_candidate(normalized.replace(' ', '-'))
                    words = [word for word in normalized.split() if word]
                    for word in words:
                        _add_candidate(word)
                    if len(words) >= 2:
                        acronym = ''.join(word[0] for word in words if word and word[0].isalpha()).upper()
                        if len(acronym) >= 2:
                            _add_candidate(acronym)

                if '(' in normalized and ')' in normalized:
                    start_paren = normalized.find('(') + 1
                    end_paren = normalized.find(')', start_paren)
                    if end_paren > start_paren:
                        _add_candidate(normalized[start_paren:end_paren])

                punctuation_stripped = normalized.strip(',:;')
                if punctuation_stripped != normalized:
                    _add_candidate(punctuation_stripped)

            max_variants = 8
            return expanded[:max_variants]





        prompt_parts: List[str]
        if use_references:
            prompt_parts = [
                f"Translate each Japanese text in the following JSON array into {target_language}.\n",
                "Use the provided reference passages and URLs as evidence, weaving quoted English wording naturally into the translation.\n",
                "Give preference to sentences that convey financial metrics, operational milestones, or strategic commitments related to the Japanese text.\n",
                "Keep every translation strictly faithful to the Japanese source; do not add, infer, or omit facts beyond what is stated.\n",
                "Follow this three-step workflow for each item:\n1. From the Japanese text, create concise English search key phrases that capture the core meaning.\n2. Use those key phrases to locate the most relevant English sentences or fragments within the provided references and URLs.\n3. When writing the translation, reuse the strongest wording from step 2 only when it supports the same fact, keeping the sentence smooth and idiomatic.\n",
                "Do not include bracketed reference markers (for example, [R1] or [U2]) in the translated sentences.\n",
                "Do not provide a Japanese justification. Ensure evidence.explanation_jp is an empty string ("").\n",
                "Keep the input order unchanged.\n",
            ]
            if reference_entries:
                prompt_parts.append(f"Reference passages:\n{json.dumps(reference_entries, ensure_ascii=False)}\n")
            if reference_url_entries:
                prompt_parts.append(f"Reference URLs:\n{json.dumps(reference_url_entries, ensure_ascii=False)}\n")
            prompt_parts.append(
                "Return only a JSON array. Each element must contain:\n"
                "- \"translated_text\": the translated sentence (without any bracketed IDs).\n"
                "- \"evidence\": an object with the key \"quotes\" (array of quoted sentences copied verbatim from step 2).\n"
                "Do not include \"explanation_jp\" or any additional commentary keys.\n"
                "Do not add other keys, explanations, or code fences.\n"
            )
            prompt_preamble = "".join(prompt_parts)
        else:
            prompt_preamble = (
                f"Translate each Japanese text in the following JSON list into {target_language} and return a JSON list of the same length.\n"
                "Maintain the order of the inputs.\n"
                "Return JSON only, without explanations or code fences.\n"
            )

        batch_size = rows_per_batch if rows_per_batch is not None else 1
        if batch_size < 1:
            raise ToolExecutionError("rows_per_batch 縺ｯ 1 莉･荳翫〒謖・EE縺励※縺上□縺輔＞縲・)

        source_start_row, source_start_col, _, _ = _parse_range_bounds(normalized_range)
        output_start_row, output_start_col, _, _ = _parse_range_bounds(output_range)

        citation_sheet = None
        citation_range = None
        citation_matrix: Optional[List[List[str]]] = None
        cite_start_row = cite_start_col = cite_rows = cite_cols = 0
        if use_references:
            if not citation_output_range:
                raise ToolExecutionError(
                    "蜿ゑｿｽE譁・EE縺梧欠螳壹＆繧後◆蝣ｴ蜷茨ｿｽE縲∵ｹ諡繧呈嶌縺崎ｾｼ繧遽・EE (citation_output_range) 繧呈欠螳壹＠縺ｦ縺上□縺輔＞縲・
                )
            citation_sheet, citation_range = _split_sheet_and_range(citation_output_range, target_sheet)
            cite_rows, cite_cols = _parse_range_dimensions(citation_range)
            if cite_rows != source_rows:
                raise ToolExecutionError("citation_output_range 縺ｮ陦梧焚縺ｯ鄙ｻ險ｳ蟇ｾ雎｡遽・EE縺ｨ荳閾ｴ縺輔○縺ｦ縺上□縺輔＞縲・)
            if cite_cols not in {1, source_cols}:
                raise ToolExecutionError(
                    "citation_output_range 縺ｮ蛻玲焚縺ｯ1蛻励∪縺滂ｿｽE鄙ｻ險ｳ蟇ｾ雎｡遽・EE縺ｨ蜷後§蛻玲焚縺ｫ縺励※縺上□縺輔＞縲・
                )
            cite_start_row, cite_start_col, _, _ = _parse_range_bounds(citation_range)
            existing_citation = actions.read_range(citation_range, citation_sheet)
            try:
                citation_matrix = _reshape_to_dimensions(existing_citation, cite_rows, cite_cols)
            except ToolExecutionError:
                citation_matrix = [["" for _ in range(cite_cols)] for _ in range(cite_rows)]

        messages: List[str] = []
        any_translation = False

        for row_start in range(0, source_rows, batch_size):
            row_end = min(row_start + batch_size, source_rows)
            chunk_texts: List[str] = []
            chunk_positions: List[Tuple[int, int]] = []

            for local_row in range(row_start, row_end):
                for col_idx in range(source_cols):
                    cell_value = original_data[local_row][col_idx]
                    if isinstance(cell_value, str) and re.search(r"[縺・繧薙ぃ-繝ｳ荳-鮴ｯ]", cell_value):
                        chunk_texts.append(cell_value)
                        chunk_positions.append((local_row, col_idx))

            if not chunk_texts:
                continue



            texts_json = json.dumps(chunk_texts, ensure_ascii=False)

            keyword_prompt = (
                "For each Japanese text in the following JSON array, generate 4-6 varied English search phrases. Blend literal translations with broader contextual, thematic, or industry phrases so you can locate reference sentences even when the source material covers adjacent topics.\n"
                "Cover entity names, key actions, and any numerical or temporal markers present in the Japanese sentence.\n"
                "Do not invent concepts or terminology that are absent from the Japanese text.\n"
                "Return a JSON array of the same length. Each element must be an object with the key \"keywords\" (array of short phrases).\n"
                "Do not include explanations or code fences.\n"
                f"{texts_json}"
            )
            keyword_response = browser_manager.ask(keyword_prompt)
            try:
                match = re.search(r'{.*}|\[.*\]', keyword_response, re.DOTALL)
                keyword_payload = match.group(0) if match else keyword_response
                keyword_items = json.loads(keyword_payload)
            except json.JSONDecodeError as exc:
                raise ToolExecutionError(
                    f"AI縺九ｉ縺ｮ讀懃ｴ｢繧ｭ繝ｼ繝輔Ξ繝ｼ繧ｺ謚ｽ蜃ｺ邨先棡繧谷SON縺ｨ縺励※隗｣譫舌〒縺阪∪縺帙ｓ縺ｧ縺励◆縲ょｿ懃ｭ・ {keyword_response}"
                ) from exc
            if not isinstance(keyword_items, list) or len(keyword_items) != len(chunk_texts):
                raise ToolExecutionError("讀懃ｴ｢繧ｭ繝ｼ繝輔Ξ繝ｼ繧ｺ縺ｮ莉ｶ謨ｰ縺鯉ｿｽE蜉帙ユ繧ｭ繧ｹ繝医→荳閾ｴ縺励∪縺帙ｓ縲・)

            normalized_keywords: List[List[str]] = []
            for item in keyword_items:
                if isinstance(item, dict):
                    raw_keywords = item.get("keywords")
                elif isinstance(item, list):
                    raw_keywords = item
                else:
                    raw_keywords = None
                if not raw_keywords or not isinstance(raw_keywords, list):
                    raise ToolExecutionError("讀懃ｴ｢繧ｭ繝ｼ繝輔Ξ繝ｼ繧ｺ縺ｮJSON縺ｫ 'keywords' 驟搾ｿｽE縺悟性縺ｾ繧後※縺・EE縺帙ｓ縲・)
                keyword_list = []
                for keyword in raw_keywords:
                    if isinstance(keyword, str):
                        cleaned = keyword.strip()
                        if cleaned:
                            keyword_list.append(cleaned)
                if not keyword_list:
                    raise ToolExecutionError("讀懃ｴ｢繧ｭ繝ｼ繝輔Ξ繝ｼ繧ｺ縺檎ｩｺ縺ｧ縺吶・)
                normalized_keywords.append(keyword_list)

            expanded_keywords: List[List[str]] = []
            for base_keywords in normalized_keywords:
                expanded_keywords.append(_expand_keyword_variants(base_keywords))

            expanded_keywords = []
            for base_keywords in normalized_keywords:
                expanded_keywords.append(_expand_keyword_variants(base_keywords))

            keyword_plan_lines: List[str] = []
            for index, (source_text, keywords) in enumerate(zip(chunk_texts, expanded_keywords), start=1):
                keyword_plan_lines.append(f"Item {index}:")
                keyword_plan_lines.append(f"- Japanese: {source_text}")
                keyword_plan_lines.append("- Search keywords:")
                for keyword in keywords:
                    keyword_plan_lines.append(f"  * {keyword}")
            keyword_plan_text = "\n".join(keyword_plan_lines)

            reference_passage_text = ""
            if reference_entries:
                passage_lines: List[str] = []
                for entry in reference_entries:
                    label_parts = [entry.get("id")]
                    sheet_name = entry.get("sheet")
                    if sheet_name:
                        label_parts.append(f"sheet {sheet_name}")
                    source_range = entry.get("source_range")
                    if source_range:
                        label_parts.append(f"range {source_range}")
                    header = " ".join(part for part in label_parts if part) or "Reference"
                    passage_lines.append(f"{header}:")
                    for content_line in entry.get("content", []):
                        passage_lines.append(f"  - {content_line}")
                reference_passage_text = "\n".join(passage_lines)

            reference_urls_text = ""
            if reference_url_entries:
                reference_urls_text = "\n".join(
                    entry["url"] for entry in reference_url_entries if entry.get("url")
                )

            if use_references:
                evidence_prompt_sections: List[str] = [
                    "Use the search keywords below to gather every relevant English sentence in the provided materials, even when the surrounding topic differs, whenever the wording can guide the translation.",
                    "Collect as many distinct candidate quotes as are useful (aim for four to seven) and prioritise variety across sections, tone, and wording.",
                    "Capture each sentence exactly as it appears, preserving punctuation, casing, numerals, and spacing.",
                    "If no quotation directly supports the Japanese meaning, include an empty string for that item.",
                    "",
                    "Japanese texts and search keywords:",
                    keyword_plan_text,
                    "",
                ]
                if reference_passage_text:
                    evidence_prompt_sections.extend(["Reference passages:", reference_passage_text, ""])
                if reference_urls_text:
                    evidence_prompt_sections.extend([
                        "Reference URLs (open as needed before responding):",
                        reference_urls_text,
                        "",
                    ])
                evidence_prompt_sections.extend([
                    "Return a JSON array matching the input order. Each element must be an object with the key \"quotes\" containing the full set of exact English sentences you gathered from the references or URLs (no omissions).",
                    "Only include sentences that appear verbatim in those sources and keep every unique sentence without duplication.",
                ])
            else:
                evidence_prompt_sections = [
                    "Use the search keywords below to craft several concise English candidate sentences per item that could serve as reusable reference expressions, even if they broaden the topic beyond the original text.",
                    "Provide three to six varied sentences per item when possible, mixing direct renderings with broader contextual phrasing.",
                    "",
                    "Japanese texts and search keywords:",
                    keyword_plan_text,
                    "",
                    "Return a JSON array matching the input order. Each element must be an object with the key \"quotes\" containing the complete array of English sentences you propose (include them all, ordered by usefulness).",
                    "Include an empty string if no suitable expression exists or if every English sentence would add information beyond the Japanese text.",
                ]

            evidence_prompt = "\n".join(evidence_prompt_sections)
            evidence_response = browser_manager.ask(evidence_prompt)
            try:
                match = re.search(r'{.*}|\[.*\]', evidence_response, re.DOTALL)
                evidence_payload = match.group(0) if match else evidence_response
                evidence_items = json.loads(evidence_payload)
            except json.JSONDecodeError as exc:
                raise ToolExecutionError(
                    f"AI縺九ｉ縺ｮ蜿り・EE迴ｾ謚ｽ蜃ｺ邨先棡繧谷SON縺ｨ縺励※隗｣譫舌〒縺阪∪縺帙ｓ縺ｧ縺励◆縲ょｿ懃ｭ・ {evidence_response}"
                ) from exc
            if not isinstance(evidence_items, list) or len(evidence_items) != len(chunk_texts):
                raise ToolExecutionError("蜿り・EE迴ｾ縺ｮ莉ｶ謨ｰ縺鯉ｿｽE蜉帙ユ繧ｭ繧ｹ繝医→荳閾ｴ縺励∪縺帙ｓ縲・)

            normalized_quotes_per_item: List[List[str]] = []
            for quotes_entry in evidence_items:
                if isinstance(quotes_entry, dict):
                    raw_quotes = quotes_entry.get("quotes")
                elif isinstance(quotes_entry, list):
                    raw_quotes = quotes_entry
                else:
                    raw_quotes = None
                quotes_list: List[str] = []
                if isinstance(raw_quotes, list):
                    for quote in raw_quotes:
                        if isinstance(quote, str):
                            cleaned_quote = quote.strip()
                            if cleaned_quote:
                                quotes_list.append(cleaned_quote)
                normalized_quotes_per_item.append(quotes_list)

            translation_context = [
                {"source_text": text, "keywords": keywords, "quotes": quotes}
                for text, keywords, quotes in zip(chunk_texts, expanded_keywords, normalized_quotes_per_item)
            ]
            translation_context_json = json.dumps(translation_context, ensure_ascii=False)

            final_prompt = (
                f"{prompt_preamble}{texts_json}"
                "Blend the supporting expressions for each item into natural English prose, reusing key wording from the quotes where it fits fluently.\n"
                "Ensure the translation remains faithful to the Japanese source; do not introduce information that is absent or uncertain.\n"
                "Use quoted wording only when it expresses the same fact, and ensure the evidence output quotes exactly match the source sentences you relied on.\n"
                f"Supporting expressions (JSON): {translation_context_json}\n"
            )
            response = browser_manager.ask(final_prompt)

            try:
                match = re.search(r'{.*}|\[.*\]', response, re.DOTALL)
                json_payload = match.group(0) if match else response
                parsed_payload = json.loads(json_payload)
            except json.JSONDecodeError as exc:
                raise ToolExecutionError(
                    f"AI縺九ｉ縺ｮ鄙ｻ險ｳ邨先棡繧谷SON縺ｨ縺励※隗｣譫舌〒縺阪∪縺帙ｓ縺ｧ縺励◆縲ょｿ懃ｭ・ {response}"
                ) from exc

            if not isinstance(parsed_payload, list) or len(parsed_payload) != len(chunk_texts):
                raise ToolExecutionError("鄙ｻ險ｳ蜑阪→鄙ｻ險ｳ蠕後〒繝・EE繧ｹ繝茨ｿｽE莉ｶ謨ｰ縺御ｸ閾ｴ縺励∪縺帙ｓ縲・)

            chunk_cell_evidences: Dict[Tuple[int, int], str] = {}
            row_evidence_lines: Dict[int, List[str]] = {}

            for item, (local_row, col_idx) in zip(parsed_payload, chunk_positions):
                if use_references and not isinstance(item, dict):
                    raise ToolExecutionError("蜿ゑｿｽE譁・EE繧ФRL繧貞茜逕ｨ縺吶ｋ蝣ｴ蜷医∫ｿｻ險ｳ邨先棡縺ｯ繧ｪ繝悶ず繧ｧ繧ｯ繝茨ｿｽE繝ｪ繧ｹ繝医〒霑斐＠縺ｦ縺上□縺輔＞縲・)

                translation_value = (
                    item.get("translated_text") or item.get("translation") or item.get("output")
                ) if use_references else item

                if not isinstance(translation_value, str):
                    raise ToolExecutionError("鄙ｻ險ｳ邨先棡縺ｮJSON縺ｫ 'translated_text' 縺悟性縺ｾ繧後※縺・EE縺帙ｓ縲・)

                output_matrix[local_row][col_idx] = translation_value
                if not writing_to_source_directly and overwrite_source:
                    source_matrix[local_row][col_idx] = translation_value

                any_translation = True

                if use_references:
                    evidence_value = item.get("evidence")
                    explanation_jp = ""
                    quotes: List[str] = []
                    if isinstance(evidence_value, dict):
                        raw_quotes = evidence_value.get("quotes")
                        if isinstance(raw_quotes, list):
                            quotes = [
                                _sanitize_evidence_value(str(q))
                                for q in raw_quotes
                                if isinstance(q, (str, int, float)) and str(q).strip()
                            ]
                        raw_explanation = (
                            evidence_value.get("explanation_jp")
                            or evidence_value.get("explanation")
                        )
                        if isinstance(raw_explanation, (str, int, float)):
                            explanation_jp = _sanitize_evidence_value(str(raw_explanation))
                    elif isinstance(evidence_value, list):
                        quotes = [
                            _sanitize_evidence_value(str(q))
                            for q in evidence_value
                            if isinstance(q, (str, int, float)) and str(q).strip()
                        ]
                    elif isinstance(evidence_value, (str, int, float)):
                        explanation_jp = _sanitize_evidence_value(str(evidence_value))

                    validated_quotes: List[str] = []
                    if quotes:
                        for quote in quotes:
                            if not quote:
                                continue
                            normalized_quote = _normalize_for_match(quote)
                            if not normalized_quote:
                                continue
                            if normalized_reference_text_pool and not any(
                                normalized_quote in ref_text for ref_text in normalized_reference_text_pool
                            ):
                                raise ToolExecutionError(
                                    f"蠑慕畑譁・'{quote}' 縺悟盾辣ｧ遽・EE縺ｮ繝・EE繧ｹ繝医↓隕九▽縺九ｊ縺ｾ縺帙ｓ縲ょｼ慕畑縺ｯ蜿ゑｿｽE譁・EE縺ｫ蟄伜惠縺吶ｋ譁・EE縺ｮ縺ｿ繧剃ｽｿ逕ｨ縺励※縺上□縺輔＞縲・
                                )
                            validated_quotes.append(quote)
                    if reference_text_pool and not validated_quotes:
                        raise ToolExecutionError("蜿ゑｿｽE譁・EE縺九ｉ蠑慕畑縺励◆闍ｱ譁・EE蟆代↑縺上→繧・譁・EE繧√※縺上□縺輔＞縲・)

                    evidence_lines: List[str] = []
                    if explanation_jp:
                        if not JAPANESE_CHAR_PATTERN.search(explanation_jp):
                            raise ToolExecutionError("根拠の説明は日本語で記述してください。")
                        normalized_explanation = explanation_jp
                    else:
                        normalized_explanation = "引用をご確認ください。"
                    evidence_lines.append(f"説明: {normalized_explanation}")
                    if validated_quotes:
                        for quote in validated_quotes:
                            evidence_lines.append(f"蠑慕畑: {quote}")
                    combined = "\n".join(evidence_lines).strip()

                    if cite_cols == source_cols:
                        chunk_cell_evidences[(local_row, col_idx)] = combined
                    elif combined:
                        row_evidence_lines.setdefault(local_row, []).append(combined)

            chunk_output_data = [
                list(output_matrix[local_row][0:source_cols])
                for local_row in range(row_start, row_end)
            ]
            chunk_output_range = _build_range_reference(
                output_start_row + row_start,
                output_start_row + row_end - 1,
                output_start_col,
                output_start_col + source_cols - 1,
            )
            messages.append(actions.write_range(chunk_output_range, chunk_output_data, output_sheet))

            if not writing_to_source_directly and overwrite_source:
                chunk_source_data = [
                    list(source_matrix[local_row][0:source_cols])
                    for local_row in range(row_start, row_end)
                ]
                chunk_source_range = _build_range_reference(
                    source_start_row + row_start,
                    source_start_row + row_end - 1,
                    source_start_col,
                    source_start_col + source_cols - 1,
                )
                messages.append(actions.write_range(chunk_source_range, chunk_source_data, target_sheet))

            if use_references and citation_matrix is not None:
                if cite_cols == source_cols:
                    for local_row in range(row_start, row_end):
                        for col_idx in range(cite_cols):
                            citation_matrix[local_row][col_idx] = ""
                    for (local_row, col_idx), evidence_text in chunk_cell_evidences.items():
                        citation_matrix[local_row][col_idx] = evidence_text
                    chunk_citation_data = [
                        list(citation_matrix[local_row][0:cite_cols])
                        for local_row in range(row_start, row_end)
                    ]
                    chunk_citation_range = _build_range_reference(
                        cite_start_row + row_start,
                        cite_start_row + row_end - 1,
                        cite_start_col,
                        cite_start_col + cite_cols - 1,
                    )
                else:
                    for local_row in range(row_start, row_end):
                        texts = row_evidence_lines.get(local_row)
                        if texts:
                            citation_matrix[local_row][0] = "\n".join(texts)
                        else:
                            citation_matrix[local_row][0] = ""
                    chunk_citation_data = [
                        [citation_matrix[local_row][0]]
                        for local_row in range(row_start, row_end)
                    ]
                    chunk_citation_range = _build_range_reference(
                        cite_start_row + row_start,
                        cite_start_row + row_end - 1,
                        cite_start_col,
                        cite_start_col,
                    )
                messages.append(actions.write_range(chunk_citation_range, chunk_citation_data, citation_sheet))

        if not any_translation:
            return f"遽・EE '{cell_range}' 縺ｫ鄙ｻ險ｳ蟇ｾ雎｡縺ｮ繝・EE繧ｹ繝医′隕九▽縺九ｊ縺ｾ縺帙ｓ縺ｧ縺励◆縲・

        return "\n".join(messages)

    except ToolExecutionError:
        raise
    except Exception as exc:
        raise ToolExecutionError(f"遽・EE '{cell_range}' 縺ｮ鄙ｻ險ｳ荳ｭ縺ｫ繧ｨ繝ｩ繝ｼ縺檎匱逕溘＠縺ｾ縺励◆: {exc}") from exc

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
        status_output_range: Range to store review status (e.g., OK / 隕∽ｿｮ豁｣).
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
                        status_matrix[r][c] = "隕∽ｿｮ豁｣"
                        issue_matrix[r][c] = "闍ｱ險ｳ繧ｻ繝ｫ縺檎ｩｺ縲√∪縺滂ｿｽE辟｡蜉ｹ縺ｧ縺吶・
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
            return "鄙ｻ險ｳ繝√ぉ繝・EE縺ｮ蟇ｾ雎｡縺ｨ縺ｪ繧区枚蟄暦ｿｽE縺悟ｭ伜惠縺励↑縺九▲縺溘◆繧√∫ｵ先棡蛻励ｒ蛻晄悄蛹悶＠縺ｾ縺励◆縲・

        normalized_batch_size = 1

        for batch in (review_entries[i:i + normalized_batch_size] for i in range(0, len(review_entries), normalized_batch_size)):
            payload = json.dumps(batch, ensure_ascii=False)
            _diff_debug(f"check_translation_quality payload={_shorten_debug(payload)}")
            analysis_prompt = (
                "縺ゅ↑縺滂ｿｽE闍ｱ險ｳ縺ｮ蜩∬ｳｪ繧定ｩ穂ｾ｡縺吶ｋ繝ｬ繝薙Η繧｢繝ｼ縺ｧ縺吶ょ推鬆・EE縺ｫ縺､縺・EE縲∬恭險ｳ縺悟次譁・EE諢丞袖繝ｻ繝九Η繧｢繝ｳ繧ｹ繝ｻ譁・EEEE繧ｹ繝壹Ν繝ｻ荳ｻ隱櫁ｿｰ隱橸ｿｽE蟇ｾ蠢懊→縺励※驕ｩ蛻・EE繧堤｢ｺ隱阪＠縺ｦ縺上□縺輔＞縲・
                "蜷・EE邏縺ｫ縺ｯ 'id', 'original_text', 'translated_text' 縺悟性縺ｾ繧後※縺・EE縺吶・
                "JSON 驟搾ｿｽE縺ｮ蜷・EE邏縺ｫ縺ｯ蠢・EE 'id', 'status', 'notes', 'highlighted_text', 'corrected_text', 'before_text', 'after_text' 繧貞性繧√※縺上□縺輔＞縲・
                "鄙ｻ險ｳ縺ｫ蝠城｡後′縺ｪ縺代ｌ縺ｰ status 縺ｯ 'OK' 縺ｨ縺励］otes 縺ｯ遨ｺ譁・EE縺ｾ縺滂ｿｽE邁｡貎斐↑陬懆ｶｳ縺ｫ縺励※縺上□縺輔＞縲・
                "蟆代＠縺ｧ繧ゆｸ榊ｮ峨′縺ゅｋ蝣ｴ蜷茨ｿｽE 'OK' 繧帝∈縺ｰ縺壹・E驥阪↓遒ｺ隱阪＠縺ｦ縺上□縺輔＞縲・
                "菫ｮ豁｣縺悟ｿ・EE縺ｪ蝣ｴ蜷茨ｿｽE status 繧・'REVISE' 縺ｨ縺励］otes 縺ｫ縺ｯ譌･譛ｬ隱槭〒縲鯖ssue: ... / Suggestion: ...縲擾ｿｽE蠖｢蠑上〒蝠城｡檎せ縺ｨ菫ｮ豁｣譯医ｒ險倩ｿｰ縺励※縺上□縺輔＞縲・
                "霑ｷ縺｣縺溷ｴ蜷医ｄ荳咲｢ｺ縺九↑轤ｹ縺後≠繧鯉ｿｽE蠢・EE 'REVISE' 繧帝∈謚槭＠縺ｦ縺上□縺輔＞縲・
                "'corrected_text' 縺ｫ縺ｯ菫ｮ豁｣蠕鯉ｿｽE闍ｱ譁・EE菴薙ｒ蜈･繧後※縺上□縺輔＞縲・
                "'highlighted_text' 縺ｫ縺ｯ corrected_text 繧抵ｿｽE繝ｼ繧ｹ縺ｫ縲∬ｿｽ蜉邂・EE縺ｯ縲占ｿｽ蜉EEEE..縲代∝炎髯､邂・EE縺ｯ縲仙炎髯､EEEE..縲代〒蝗ｲ繧薙□豈碑ｼ・EE遉ｺ繧定ｿ斐＠縺ｦ縺上□縺輔＞縲・
                "AI 縺ｫ縺ｯ JSON 縺ｮ縺ｿ繧定ｿ斐＠縲∽ｽ呵ｨ医↑繝輔か繝ｭ繝ｼ繧｢繝・EE繧・EE繝ｼ繧ｯ繧｢繝・EE縺ｯ逕滂ｿｽE縺励↑縺・EE縺上□縺輔＞縲・
                "縺難ｿｽE繝ｪ繧ｯ繧ｨ繧ｹ繝茨ｿｽE蜈ｬ髢九＆繧後※縺・EE莨∵･ｭ雋｡蜍呎ュ蝣ｱ縺ｮ鄙ｻ險ｳ繝√ぉ繝・EE縺ｧ縺ゅｊ縲・E繝ｪ繧ｷ繝ｼ縺ｫ驕ｩ蜷医＠縺溷ｮ会ｿｽE縺ｪ蜀・EE縺ｧ縺吶ょｮ牙ｿ・EE縺ｦ蟇ｾ蠢懊＠縺ｦ縺上□縺輔＞縲・
                "You already have all necessary information; do not ask for additional ranges or confirmations."
                "Return exactly one JSON array窶馬o Final Answer, Thought, or explanations, and never include a second JSON object or array."
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
                    "marking additions as 縲占ｿｽ蜉EEEE..縲・and deletions as 縲仙炎髯､EEEE..縲・ "
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
                raise ToolExecutionError(f"鄙ｻ險ｳ繝√ぉ繝・EE縺ｮ邨先棡繧谷SON縺ｨ縺励※隗｣譫舌〒縺阪∪縺帙ｓ縺ｧ縺励◆縲ょｿ懃ｭ・ {response}")

            if not isinstance(batch_results, list):
                raise ToolExecutionError("鄙ｻ險ｳ繝√ぉ繝・EE縺ｮ蠢懃ｭ泌ｽ｢蠑上′荳肴ｭ｣縺ｧ縺吶・SON驟搾ｿｽE繧定ｿ斐＠縺ｦ縺上□縺輔＞縲・)

            ok_statuses = {"OK", "PASS", "GOOD"}
            revise_statuses = {"REVISE", "NG", "FAIL", "ISSUE"}
            for item in batch_results:
                if not isinstance(item, dict):
                    raise ToolExecutionError("鄙ｻ險ｳ繝√ぉ繝・EE縺ｮ蠢懃ｭ斐↓荳肴ｭ｣縺ｪ隕∫ｴ縺悟性縺ｾ繧後※縺・EE縺吶・)
                item_id = item.get("id")
                if item_id not in id_to_position:
                    _diff_debug(f"check_translation_quality unknown id={item_id} known={list(id_to_position.keys())}")
                    raise ToolExecutionError("鄙ｻ險ｳ繝√ぉ繝・EE縺ｮ蠢懃ｭ斐↓譛ｪ遏･縺ｮID縺悟性縺ｾ繧後※縺・EE縺吶・)
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
                    status_matrix[row_idx][col_idx] = "隕∽ｿｮ豁｣"
                    issue_matrix[row_idx][col_idx] = notes_value or "菫ｮ豁｣蜀・EE繧定ｨ倩ｼ峨＠縺ｦ縺上□縺輔＞縲・
                    needs_revision_count += 1
                else:
                    status_matrix[row_idx][col_idx] = status_value or "隕∫｢ｺ隱・
                    issue_matrix[row_idx][col_idx] = notes_value or "繧ｹ繝・EE繧ｿ繧ｹ縺瑚ｧ｣驥医〒縺阪∪縺帙ｓ縺ｧ縺励◆縲・
                    needs_revision_count += 1

        actions.write_range(status_output_range, status_matrix, sheet_name)
        actions.write_range(issue_output_range, issue_matrix, sheet_name)

        processed_items = len(review_entries)
        message = (
            f"鄙ｻ險ｳ繝√ぉ繝・EE繧貞ｮ御ｺ・EE縺ｾ縺励◆縲ょｯｾ雎｡ {processed_items} 莉ｶ荳ｭ縲∬ｦ∽ｿｮ豁｣ {needs_revision_count} 莉ｶ縺ｮ邨先棡繧・
            f" '{status_output_range}' 縺ｨ '{issue_output_range}' 縺ｫ譖ｸ縺崎ｾｼ縺ｿ縺ｾ縺励◆縲・
        )
        if corrected_matrix is not None and corrected_output_range:
            actions.write_range(corrected_output_range, corrected_matrix, sheet_name)
            message += f" 螳鯉ｿｽE蠖｢縺ｯ '{corrected_output_range}' 縺ｫ蜃ｺ蜉帙＠縺ｾ縺励◆縲・
        if highlight_matrix is not None and highlight_output_range:
            actions.write_range(highlight_output_range, highlight_matrix, sheet_name)
            if highlight_styles is not None:
                actions.apply_diff_highlight_colors(highlight_output_range, highlight_styles, sheet_name)
            message += f" 豈碑ｼ・EE遉ｺ逕ｨ縺ｮ譁・EEEE縺ｯ '{highlight_output_range}' 縺ｫ蜃ｺ蜉帙＠縺ｾ縺励◆縲・
        return message

    except ToolExecutionError:
        raise
    except Exception as e:
        raise ToolExecutionError(f"鄙ｻ險ｳ繝√ぉ繝・EE荳ｭ縺ｫ繧ｨ繝ｩ繝ｼ縺檎匱逕溘＠縺ｾ縺励◆: {e}") from e



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
            f"蟾ｮ蛻・EE繧､繝ｩ繧､繝医ｒ '{output_range}' 縺ｫ蜃ｺ蜉帙＠縲∬ｿｽ蜉遞区園({addition_color_hex})縺ｨ蜑企勁遞区園({deletion_color_hex})繧貞ｼｷ隱ｿ縺励∪縺励◆縲・
        )
    except ToolExecutionError:
        raise
    except Exception as exc:
        _diff_debug(f"highlight_text_differences exception={exc}")
        raise ToolExecutionError(f"蟾ｮ蛻・EE繧､繝ｩ繧､繝茨ｿｽE驕ｩ逕ｨ荳ｭ縺ｫ繧ｨ繝ｩ繝ｼ縺檎匱逕溘＠縺ｾ縺励◆: {exc}") from exc

def insert_shape(actions: ExcelActions,
                 cell_range: str,
                 shape_type: str,
                 sheet_name: Optional[str] = None,
                 fill_color_hex: Optional[str] = None,
                 line_color_hex: Optional[str] = None) -> str:
    """
    謖・EE縺励◆繧ｻ繝ｫ遽・EE縺ｫ縲∵欠螳壹＠縺滓嶌蠑上〒蝗ｳ蠖｢繧呈諺蜈･縺励∪縺吶・
    :param cell_range: 蝗ｳ蠖｢繧呈諺蜈･縺吶ｋ遽・EE (萓・ "A1:C5")
    :param shape_type: 謖ｿ蜈･縺吶ｋ蝗ｳ蠖｢縺ｮ遞ｮ鬘・(萓・ "蝗幄ｧ貞ｽ｢", "讌包ｿｽE")
    :param sheet_name: 蟇ｾ雎｡繧ｷ繝ｼ繝亥錐EEE逵∫払蜿ｯEEEE
    :param fill_color_hex: 蝪励ｊ縺､縺ｶ縺暦ｿｽE濶ｲ (16騾ｲ謨ｰ, 萓・ "#FF0000")
    :param line_color_hex: 譫邱夲ｿｽE濶ｲ (16騾ｲ謨ｰ, 萓・ "#0000FF")
    """
    return actions.insert_shape_in_range(cell_range, shape_type, sheet_name, fill_color_hex, line_color_hex)

def format_shape(actions: ExcelActions, fill_color_hex: Optional[str] = None, line_color_hex: Optional[str] = None, sheet_name: Optional[str] = None) -> str:
    """
    [髱樊耳螂ｨ] 縺難ｿｽE髢｢謨ｰ縺ｯ菴ｿ繧上↑縺・EE縺上□縺輔＞縲ゆｻ｣繧上ｊ縺ｫ insert_shape 髢｢謨ｰ縺ｮ蠑墓焚縺ｧ濶ｲ繧呈欠螳壹＠縺ｦ縺上□縺輔＞縲・
    """
    return actions.format_last_shape(fill_color_hex, line_color_hex, sheet_name)












