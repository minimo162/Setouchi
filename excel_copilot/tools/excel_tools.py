import html
import ast
import json
import re
import difflib
import logging
import math
import os
import string
import copy
from threading import Event
from typing import List, Any, Optional, Dict, Tuple, Set, Mapping
from pathlib import Path
from urllib.parse import urlparse

from excel_copilot.core.browser_copilot_manager import BrowserCopilotManager
from excel_copilot.core.exceptions import ToolExecutionError, UserStopRequested

from .actions import ExcelActions


_logger = logging.getLogger(__name__)
_DIFF_DEBUG_ENABLED = os.getenv('EXCEL_COPILOT_DEBUG_DIFF', '').lower() in {'1', 'true', 'yes'}
_REVIEW_DEBUG_ENABLED = os.getenv('EXCEL_COPILOT_DEBUG_REVIEW', '').lower() in {'1', 'true', 'yes', 'on'}



def _review_debug(message: str) -> None:
    if _REVIEW_DEBUG_ENABLED:
        print(f"[review-debug] {message}")



_NO_QUOTES_PLACEHOLDER = "引用なし"
_HTML_ENTITY_PATTERN = re.compile(r"&(?:[A-Za-z][A-Za-z0-9]{1,31}|#[0-9]{1,7}|#x[0-9A-Fa-f]{1,6});")
DEFAULT_REFERENCE_PAIR_COLUMNS = 10
_MIN_CONTEXT_BLOCK_WIDTH = 2 + DEFAULT_REFERENCE_PAIR_COLUMNS
if _DIFF_DEBUG_ENABLED and not logging.getLogger().handlers:
    logging.basicConfig(level=logging.DEBUG)



try:
    _ITEMS_PER_TRANSLATION_REQUEST = max(
        1, int(os.getenv('EXCEL_COPILOT_TRANSLATION_ITEMS_PER_REQUEST', '1'))
    )
except ValueError:
    _ITEMS_PER_TRANSLATION_REQUEST = 1

try:
    _TRANSLATION_UTF16_BUDGET = int(os.getenv("EXCEL_COPILOT_TRANSLATION_UTF16_BUDGET", "4000"))
except ValueError:
    _TRANSLATION_UTF16_BUDGET = 4000
if _TRANSLATION_UTF16_BUDGET <= 0:
    _TRANSLATION_UTF16_BUDGET = None  # type: ignore[assignment]



def _diff_debug(message: str) -> None:
    if _DIFF_DEBUG_ENABLED:
        _logger.debug(message)


def _shorten_debug(value: str, limit: int = 120) -> str:
    if value is None:
        return ''
    text = str(value).replace('\r', '\r').replace('\n', '\n')
    return text if len(text) <= limit else text[:limit] + '窶ｦ'



_MOJIBAKE_MARKERS: Set[str] = frozenset('縺繧繝邨蜑螟蠖蛻蝣鬟荳菫譁蜿遘遉遞迚邱遽驟髣髴髢霑蜉')
_BRACKETED_URL_PATTERN = re.compile(r"\[(\d+)\]\((?:https?|ftp)://[^)]+\)")
_JAPANESE_CHAR_PATTERN = re.compile(r"[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufa6d]")
JAPANESE_CHAR_PATTERN = _JAPANESE_CHAR_PATTERN  # 後方互換用のエイリアス

def _mojibake_penalty(text: str) -> int:
    penalty = 0
    for ch in text:
        code_point = ord(ch)
        if ch in _MOJIBAKE_MARKERS:
            penalty += 2
        elif 0xE000 <= code_point <= 0xF8FF:
            penalty += 3
        elif ch == '\uFFFD':
            penalty += 4
    return penalty


def _contains_japanese(text: str) -> bool:
    if not text:
        return False
    return bool(_JAPANESE_CHAR_PATTERN.search(text))


def _count_japanese_characters(text: str) -> int:
    count = 0
    for ch in text:
        code_point = ord(ch)
        if (
            0x3040 <= code_point <= 0x30FF
            or 0x3400 <= code_point <= 0x4DBF
            or 0x4E00 <= code_point <= 0x9FFF
            or 0xF900 <= code_point <= 0xFAFF
            or 0xFF66 <= code_point <= 0xFF9D
        ):
            count += 1
    return count


def _maybe_fix_mojibake(text: str) -> str:
    if not isinstance(text, str) or not text:
        return text

    original_penalty = _mojibake_penalty(text)
    original_japanese = _count_japanese_characters(text)

    candidates: List[str] = []

    def _register_candidate(converter: Any) -> None:
        try:
            candidate = converter()
        except (UnicodeError, ValueError):
            return
        if not isinstance(candidate, str):
            return
        if not candidate or candidate == text:
            return
        if candidate not in candidates:
            candidates.append(candidate)

    _register_candidate(lambda: text.encode('cp932', errors='strict').decode('utf-8', errors='strict'))
    _register_candidate(lambda: text.encode('utf-8', errors='strict').decode('cp932', errors='strict'))
    _register_candidate(lambda: text.encode('latin-1', errors='strict').decode('cp932', errors='strict'))
    _register_candidate(lambda: text.encode('cp932', errors='strict').decode('latin-1', errors='strict'))
    _register_candidate(lambda: text.encode('latin-1', errors='strict').decode('utf-8', errors='strict'))
    _register_candidate(lambda: text.encode('utf-8', errors='strict').decode('latin-1', errors='strict'))

    if not candidates:
        return text

    best_text = text
    best_score = (original_penalty, -original_japanese)

    for candidate in candidates:
        candidate_penalty = _mojibake_penalty(candidate)
        candidate_japanese = _count_japanese_characters(candidate)
        candidate_score = (candidate_penalty, -candidate_japanese)
        if candidate_score < best_score:
            best_text = candidate
            best_score = candidate_score

    return best_text


def _maybe_unescape_html_entities(text: str) -> str:
    if not text or '&' not in text:
        return text
    if not _HTML_ENTITY_PATTERN.search(text):
        return text
    try:
        unescaped = html.unescape(text)
    except Exception:
        return text
    return unescaped


def _measure_utf16_length(value: Optional[str]) -> int:
    """Return the number of UTF-16 code units in the provided string."""
    if value is None:
        return 0
    if not isinstance(value, str):
        value = str(value)
    if not value:
        return 0
    try:
        return len(value.encode("utf-16-le")) // 2
    except Exception:
        # Fall back to Python's code point count if encoding fails unexpectedly.
        return len(value)


def _unescape_matrix_values(matrix: List[List[Any]]) -> List[List[Any]]:
    return [
        [
            _maybe_unescape_html_entities(cell) if isinstance(cell, str) else cell
            for cell in row
        ]
        for row in matrix
    ]

def _generate_keyword_variants(base: str) -> List[str]:
    """Produce diverse keyword variants to widen reference searches."""
    variants: List[str] = []
    seen: Set[str] = set()

    def _add(candidate: str) -> None:
        cleaned = (candidate or '').strip()
        if not cleaned:
            return
        lowered = cleaned.lower()
        if lowered in seen:
            return
        seen.add(lowered)
        variants.append(cleaned)

    candidate_base = (base or '').replace('\u3000', ' ').strip()
    if not candidate_base:
        return []

    dash_normalised = (
        candidate_base
        .replace('\u2010', '-')
        .replace('\u2011', '-')
        .replace('\u2012', '-')
        .replace('\u2013', '-')
        .replace('\u2014', '-')
        .replace('\u2212', '-')
    )
    _add(candidate_base)
    _add(dash_normalised)
    _add(dash_normalised.lower())
    _add(dash_normalised.upper())

    space_normalised = ' '.join(dash_normalised.split())
    _add(space_normalised)
    _add(space_normalised.replace(' ', '-'))
    _add(dash_normalised.replace('-', ' '))

    raw_tokens = [tok for tok in re.split(r'[\s/&+\u30fb\uFF65\u301C\uFF5E~]+', space_normalised) if tok]

    def _add_word_forms(token: str) -> None:
        if not token:
            return
        _add(token)
        lower_tok = token.lower()
        _add(lower_tok)
        title_tok = token.title()
        if title_tok != token:
            _add(title_tok)
        if lower_tok.endswith('ies') and len(lower_tok) > 3:
            _add(lower_tok[:-3] + 'y')
        if lower_tok.endswith('ing') and len(lower_tok) > 4:
            stem = lower_tok[:-3]
            _add(stem)
            if not stem.endswith('e'):
                _add(stem + 'e')
        if lower_tok.endswith('ed') and len(lower_tok) > 3:
            stem = lower_tok[:-2]
            _add(stem)
            if not stem.endswith('e'):
                _add(stem + 'e')
        if lower_tok.endswith('s') and len(lower_tok) > 3:
            _add(lower_tok[:-1])
        if lower_tok.endswith('es') and len(lower_tok) > 4:
            _add(lower_tok[:-2])

    for token in raw_tokens:
        _add_word_forms(token)

    if len(raw_tokens) >= 2:
        for i in range(len(raw_tokens) - 1):
            pair = f"{raw_tokens[i]} {raw_tokens[i + 1]}"
            _add(pair)
            _add(pair.replace(' ', '-'))
    if len(raw_tokens) >= 3:
        trio = ' '.join(raw_tokens[:3])
        _add(trio)
        _add(trio.replace(' ', '-'))

    punctuation_sanitised = re.sub(r'[,:;]', ' ', space_normalised)
    if punctuation_sanitised != space_normalised:
        _add(' '.join(punctuation_sanitised.split()))

    return variants
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
    if cell_part.startswith("'") and cell_part.endswith("'"):
        cell_part = cell_part[1:-1].replace("''", "'")
    elif cell_part.startswith('"') and cell_part.endswith('"'):
        cell_part = cell_part[1:-1]
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
_SENTENCE_BOUNDARY_CHARS = set("!.?。？！")
_CLOSING_PUNCTUATION = ")]},。、？！「」『』'\"》】〕〉"
_MAX_DIFF_SEGMENT_TOKENS = 10
_MAX_DIFF_SEGMENT_CHARS = 48

# ハイライトは色分けで伝えるためラベル文字列は空のままとする。
_HIGHLIGHT_LABELS = {
    "DEL": "",
    "ADD": "",
}

REFUSAL_PATTERNS = (
    "申し訳ございません。これには対応できません。",
    "申し訳ございません。これには対応できません",
    "申し訳ございません。チャットを保存して新しいチャットを開始するには、[新しいチャット] を選択してください。",
    "チャットを保存して新しいチャットを開始するには、[新しいチャット] を選択してください。",
    "お答えできません。",
    "お答えできません",
    "I'm sorry, but I can't help with that.",
    "I cannot help with that request.",
    "エラーが発生しました: 応答形式が不正です。'Thought:' または 'Final Answer:' が見つかりません。",
    "応答形式が不正です。'Thought:' または 'Final Answer:' が見つかりません。",
)

# JAPANESE_CHAR_PATTERN はファイル先頭で定義した _JAPANESE_CHAR_PATTERN のエイリアスを使用します。



def _parse_range_dimensions(range_ref: str) -> Tuple[int, int]:
    ref = range_ref.split('!')[-1].replace('$', '').strip()
    if not ref:
        raise ToolExecutionError('Range string is empty.')
    if ':' not in ref:
        if not CELL_REFERENCE_PATTERN.fullmatch(ref):
            raise ToolExecutionError(
                f"Range '{range_ref}' is not a valid Excel reference. Use A1-style addresses such as 'A1' or 'A1:B5'."
            )
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
        return _maybe_unescape_html_entities(cell)
    if cell is None:
        return ''
    return str(cell)


def _format_issue_notes(notes_value: Optional[str]) -> str:
    normalized = _normalize_cell_value(notes_value).strip()
    if not normalized:
        return "課題: / 提案: "

    replacements = (
        (r"(?i)issue\s*[:：]", "課題: "),
        (r"(?i)suggestion\s*[:：]", "提案: "),
        (r"(?i)note\s*[:：]", "メモ: "),
    )
    result = normalized
    for pattern, replacement in replacements:
        result = re.sub(pattern, replacement, result)

    result = result.replace(" / ", " ／ ")

    if not _contains_japanese(result):
        return "課題: 内容を日本語で記入してください。／ 提案: 内容を日本語で記入してください。"

    return result


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
    trailing_len = len(segment) - len(segment.rstrip())
    core_start = leading_len
    core_end = len(segment) - trailing_len if trailing_len else len(segment)
    core = segment[core_start:core_end]
    if not core:
        _diff_debug(f"_format_diff_segment no core text label={label}")
        return segment, None, None
    prefix = segment[:leading_len]
    suffix = segment[core_end:]
    marker_prefix = f"[{label}]"
    marker_suffix = f"[{label}]"
    formatted = f'{prefix}{marker_prefix}{core}{marker_suffix}{suffix}'
    highlight_start_offset = len(prefix)
    highlight_length = len(marker_prefix) + len(core)
    _diff_debug(f"_format_diff_segment result label={label} formatted={_shorten_debug(formatted)} offset={highlight_start_offset} length={highlight_length}")
    return formatted, highlight_start_offset, highlight_length

def _split_shared_context(original_segment: str, corrected_segment: str) -> Tuple[str, str, str, str]:
    if not original_segment and not corrected_segment:
        return '', '', '', ''

    prefix_len = 0
    max_prefix = min(len(original_segment), len(corrected_segment))
    while prefix_len < max_prefix and original_segment[prefix_len] == corrected_segment[prefix_len]:
        prefix_len += 1

    suffix_len = 0
    max_suffix_original = len(original_segment) - prefix_len
    max_suffix_corrected = len(corrected_segment) - prefix_len
    while (
        suffix_len < max_suffix_original
        and suffix_len < max_suffix_corrected
        and original_segment[len(original_segment) - suffix_len - 1] == corrected_segment[len(corrected_segment) - suffix_len - 1]
    ):
        suffix_len += 1

    trimmed_end_original = len(original_segment) - suffix_len if suffix_len else len(original_segment)
    trimmed_end_corrected = len(corrected_segment) - suffix_len if suffix_len else len(corrected_segment)

    common_prefix = original_segment[:prefix_len]
    trimmed_original = original_segment[prefix_len:trimmed_end_original]
    trimmed_corrected = corrected_segment[prefix_len:trimmed_end_corrected]
    common_suffix = corrected_segment[trimmed_end_corrected:] if suffix_len else ''

    return common_prefix, trimmed_original, trimmed_corrected, common_suffix

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
            added_tokens = corr_tokens[j1:j2]
            removed_segment = ''.join(removed_tokens)
            added_segment = ''.join(added_tokens)
            prefix, trimmed_removed, trimmed_added, suffix = _split_shared_context(removed_segment, added_segment)

            if not trimmed_removed and not trimmed_added and (removed_segment or added_segment):
                trimmed_removed = removed_segment
                trimmed_added = added_segment
                prefix = ''
                suffix = ''

            if trimmed_removed and not trimmed_added:
                _diff_debug("_build_diff_highlight replace segment has removal without addition; highlighting deletion only")

            if prefix:
                result_parts.append(prefix)
                cursor += len(prefix)

            if trimmed_removed:
                formatted_removed, offset_removed, length_removed = _format_diff_segment([trimmed_removed], 'DEL')
                if formatted_removed:
                    result_parts.append(formatted_removed)
                    if offset_removed is not None and length_removed:
                        span = {'start': cursor + offset_removed, 'length': length_removed, 'type': 'DEL'}
                        spans.append(span)
                        _diff_debug(f"_build_diff_highlight span added {span}")
                    cursor += len(formatted_removed)

            if trimmed_added:
                formatted_added, offset_added, length_added = _format_diff_segment([trimmed_added], 'ADD')
                if formatted_added:
                    result_parts.append(formatted_added)
                    if offset_added is not None and length_added:
                        span = {'start': cursor + offset_added, 'length': length_added, 'type': 'ADD'}
                        spans.append(span)
                        _diff_debug(f"_build_diff_highlight span added {span}")
                    cursor += len(formatted_added)

            if suffix:
                result_parts.append(suffix)
                cursor += len(suffix)
        elif tag == 'delete':
            removed_tokens = orig_tokens[i1:i2]
            formatted_removed, offset_removed, length_removed = _format_diff_segment(removed_tokens, 'DEL')
            if formatted_removed:
                result_parts.append(formatted_removed)
                if offset_removed is not None and length_removed:
                    span = {'start': cursor + offset_removed, 'length': length_removed, 'type': 'DEL'}
                    spans.append(span)
                    _diff_debug(f"_build_diff_highlight span added {span}")
                cursor += len(formatted_removed)
        elif tag == 'insert':
            added_tokens = corr_tokens[j1:j2]
            formatted_added, offset_added, length_added = _format_diff_segment(added_tokens, 'ADD')
            if formatted_added:
                result_parts.append(formatted_added)
                if offset_added is not None and length_added:
                    span = {'start': cursor + offset_added, 'length': length_added, 'type': 'ADD'}
                    spans.append(span)
                    _diff_debug(f"_build_diff_highlight span added {span}")
                cursor += len(formatted_added)
    result = ''.join(result_parts)
    if not result.strip():
        _diff_debug("_build_diff_highlight result empty after strip")
        return corrected_text, []
    clean_text, marker_spans = _parse_highlight_markup(result)
    if marker_spans:
        _diff_debug(f"_build_diff_highlight parsed spans count={len(marker_spans)}")
        return clean_text, marker_spans
    _diff_debug(f"_build_diff_highlight result_len={len(result)} spans={spans}")
    return clean_text, spans


def _parse_highlight_markup(raw_text: str) -> Tuple[str, List[Dict[str, int]]]:
    if not isinstance(raw_text, str) or not raw_text:
        return "" if raw_text is None else str(raw_text), []

    pattern = re.compile(r"\[(DEL|ADD)\](.*?)\[(DEL|ADD)\]", re.DOTALL)
    output_segments: List[str] = []
    spans: List[Dict[str, int]] = []
    cursor = 0
    current_length = 0

    for match in pattern.finditer(raw_text):
        open_type = match.group(1)
        segment_text = match.group(2)
        close_type = match.group(3)
        if open_type != close_type:
            continue

        leading_text = raw_text[cursor:match.start()]
        if leading_text:
            output_segments.append(leading_text)
            current_length += len(leading_text)

        if segment_text:
            prefix_ws_len = len(segment_text) - len(segment_text.lstrip())
            suffix_ws_len = len(segment_text) - len(segment_text.rstrip())

            if prefix_ws_len:
                prefix_ws = segment_text[:prefix_ws_len]
                output_segments.append(prefix_ws)
                current_length += len(prefix_ws)

            core_text = segment_text[prefix_ws_len: len(segment_text) - suffix_ws_len if suffix_ws_len else len(segment_text)]
            if core_text:
                span_start = current_length
                output_segments.append(core_text)
                span_length = len(core_text)
                current_length += span_length
                spans.append({"start": span_start, "length": span_length, "type": open_type.upper()})

            if suffix_ws_len:
                suffix_ws = segment_text[len(segment_text) - suffix_ws_len:]
                output_segments.append(suffix_ws)
                current_length += len(suffix_ws)

        cursor = match.end()

    if cursor < len(raw_text):
        trailing_text = raw_text[cursor:]
        output_segments.append(trailing_text)

    clean_text = "".join(output_segments)
    return clean_text, spans


def _attach_highlight_labels(text: str, spans: List[Dict[str, int]]) -> Tuple[str, List[Dict[str, int]]]:
    if not spans:
        return text, spans

    sorted_spans = sorted(spans, key=lambda span: span.get("start", 0))
    offset = 0
    modified_text = text
    updated_spans: List[Dict[str, int]] = []

    for span in sorted_spans:
        span_type = (span.get("type") or "").upper()
        label = _HIGHLIGHT_LABELS.get(span_type, "")
        label_len = len(label)
        original_start = int(span.get("start", 0))
        original_length = int(span.get("length", 0))
        insert_position = original_start + offset

        if label_len:
            modified_text = modified_text[:insert_position] + label + modified_text[insert_position:]
            offset += label_len
            new_span = dict(span)
            new_span["start"] = original_start + (offset - label_len)
            new_span["length"] = original_length + label_len
            updated_spans.append(new_span)
        else:
            new_span = dict(span)
            new_span["start"] = original_start + offset
            updated_spans.append(new_span)

    return modified_text, updated_spans


def _render_textual_diff_markup(
    text: str,
    spans: List[Dict[str, int]],
    addition_marker: str = "[ADD]",
    deletion_marker: str = "[DEL]",
) -> str:
    """
    Wrap diff spans with textual markers when rich text coloring is unavailable.
    """
    if not isinstance(text, str):
        text = "" if text is None else str(text)
    if not spans:
        return text

    text_len = len(text)
    normalized_spans: List[Tuple[int, int, str]] = []

    for span in spans:
        if not isinstance(span, dict):
            continue
        try:
            start = int(span.get("start", 0))
            length = int(span.get("length", 0))
        except Exception:
            continue
        if length <= 0:
            continue
        span_type = str(span.get("type", "")).strip().upper()
        if span_type in {"ADD", "ADDITION", "INSERT", "INSERTED", "追加"}:
            marker = addition_marker
        elif span_type in {"DEL", "DELETION", "DELETE", "REMOVED", "削除"}:
            marker = deletion_marker
        else:
            continue
        start = max(start, 0)
        end = start + length
        if start >= text_len:
            continue
        if end > text_len:
            end = text_len
        if end <= start:
            continue
        normalized_spans.append((start, end, marker))

    if not normalized_spans:
        return text

    normalized_spans.sort(key=lambda item: item[0])
    pieces: List[str] = []
    cursor = 0

    for start, end, marker in normalized_spans:
        if cursor < start:
            pieces.append(text[cursor:start])
            segment_start = start
        else:
            segment_start = cursor
        if segment_start >= end:
            continue
        pieces.append(marker)
        pieces.append(text[segment_start:end])
        pieces.append(marker)
        cursor = max(cursor, end)

    if cursor < text_len:
        pieces.append(text[cursor:])

    return ''.join(pieces)


def writetocell(actions: ExcelActions, cell: str, value: Any, sheetname: Optional[str] = None) -> str:
    """Write a value into a single Excel cell.

    

    Args:

        actions: Excel automation helper injected by the agent runtime.

        cell: A1-style reference for the destination cell.

        value: Data to write into the cell.

        sheetname: Optional sheet override; defaults to the active sheet.

    """
    return actions.write_to_cell(cell, value, sheetname)

def readcellvalue(actions: ExcelActions, cell: str, sheetname: Optional[str] = None) -> Any:
    """Read the value stored in a single cell.

    

    Args:

        actions: Excel automation helper injected by the agent runtime.

        cell: A1-style reference to read.

        sheetname: Optional sheet override; defaults to the active sheet.

    """
    return actions.read_cell_value(cell, sheetname)

def getallsheetnames(actions: ExcelActions) -> str:
    """Return all sheet names from the active workbook.

    

    Args:

        actions: Excel automation helper injected by the agent runtime.

    """
    names = actions.get_sheet_names()
    return f"利用可能なシートは次の通りです: {', '.join(names)}"

def copyrange(actions: ExcelActions, sourcerange: str, destinationrange: str, sheetname: Optional[str] = None) -> str:
    """Copy values and formatting from a source range into a destination range.

    

    Args:

        actions: Excel automation helper injected by the agent runtime.

        sourcerange: A1-style range to copy from.

        destinationrange: A1-style range to copy into.

        sheetname: Optional sheet override; defaults to the active sheet.

    """
    return actions.copy_range(sourcerange, destinationrange, sheetname)

def executeexcelformula(actions: ExcelActions, cell: str, formula: str, sheetname: Optional[str] = None) -> str:
    """Set or replace an Excel formula on a cell.

    

    Args:

        actions: Excel automation helper injected by the agent runtime.

        cell: A1-style cell reference where the formula is applied.

        formula: Excel formula text.

        sheetname: Optional sheet override; defaults to the active sheet.

    """
    return actions.set_formula(cell, formula, sheetname)

def readrangevalues(actions: ExcelActions, cellrange: str, sheetname: Optional[str] = None) -> str:
    """Read values from a range and summarise the result.

    

    Args:

        actions: Excel automation helper injected by the agent runtime.

        cellrange: A1-style range to read.

        sheetname: Optional sheet override; defaults to the active sheet.

    """
    values = actions.read_range(cellrange, sheetname)
    return f"範囲 '{cellrange}' の値は次の通りです: {values}"

def writerangevalues(actions: ExcelActions, cellrange: str, data: List[List[Any]], sheetname: Optional[str] = None) -> str:
    """Write a 2D list of values into a range, validating the shape first.

    

    Args:

        actions: Excel automation helper injected by the agent runtime.

        cellrange: A1-style range that must match the data shape.

        data: Two-dimensional list of values to write.

        sheetname: Optional sheet override; defaults to the active sheet.

    """
    return actions.write_range(cellrange, data, sheetname)

def getactiveworkbookandsheet(actions: ExcelActions) -> str:
    """Report the currently active workbook and sheet names.

    

    Args:

        actions: Excel automation helper injected by the agent runtime.

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
    """Apply the provided formatting properties to a range.

    

    Args:

        actions: Excel automation helper injected by the agent runtime.

        cellrange: A1-style range to format.

        sheetname: Optional sheet override; defaults to the active sheet.

        fontname: Optional font family to apply.

        fontsize: Optional font size in points.

        fontcolorhex: Optional font colour specified as #RRGGBB.

        bold: Optional flag to toggle bold text.

        italic: Optional flag to toggle italic text.

        fillcolorhex: Optional fill colour specified as #RRGGBB.

        columnwidth: Optional column width in Excel units.

        rowheight: Optional row height in Excel units.

        horizontalalignment: Optional horizontal alignment keyword.

        borderstyle: Optional mapping describing border configuration.

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

def translate_range_contents(
    actions: ExcelActions,
    browser_manager: BrowserCopilotManager,
    cell_range: str,
    target_language: str = "English",
    sheet_name: Optional[str] = None,
    citation_output_range: Optional[str] = None,
    reference_urls: Optional[List[str]] = None,
    source_reference_urls: Optional[List[str]] = None,
    target_reference_urls: Optional[List[str]] = None,
    translation_output_range: Optional[str] = None,
    overwrite_source: bool = False,
    length_ratio_limit: Optional[float] = None,
    length_ratio_min: Optional[float] = None,
    rows_per_batch: Optional[int] = None,
    stop_event: Optional[Event] = None,
    output_mode: str = "translation_with_context",
) -> str:
    """Translate Japanese text in Excel while orchestrating reference-guided translation.

    Args:
        actions: Excel automation helper injected by the agent runtime.
        browser_manager: Shared browser manager used for LLM interactions.
        cell_range: Range containing the source Japanese text.
        target_language: Target language for the translation (e.g., \"English\").
        sheet_name: Optional sheet override; defaults to the active sheet.
        citation_output_range: Optional range used to write citation evidence.
        reference_urls: Legacy list of reference URLs (treated as source-language material).
        source_reference_urls: URLs to the original-language reference documents.
        target_reference_urls: URLs to the target-language reference documents used for pairing.
        translation_output_range: Range where translated content, process notes, and reference pairs are written.
        overwrite_source: Whether to overwrite the source range directly.
        length_ratio_limit: Optional upper bound for the ratio (translated UTF-16 length / source UTF-16 length)
            enforced in translation-only mode. When None, no upper limit is applied.
        length_ratio_min: Optional lower bound for the ratio (translated UTF-16 length / source UTF-16 length)
            enforced in translation-only mode. When None, no lower limit is applied.
        rows_per_batch: Optional cap for batch size when chunking the translation work.
        stop_event: Optional cancellation event set when the user interrupts the operation.
        output_mode: Controls whether contextual columns (process notes / reference pairs) are emitted.
    """

    try:
        def _ensure_not_stopped() -> None:
            if stop_event and stop_event.is_set():
                raise UserStopRequested("ユーザーによって処理が中断されました。")

        _ensure_not_stopped()

        def _strip_enclosing_quotes(text: str) -> str:
            if not isinstance(text, str):
                return text
            trimmed = text.strip()
            if len(trimmed) >= 2 and trimmed[0] == trimmed[-1] and trimmed[0] in {"'", '"'}:
                core = trimmed[1:-1]
                if trimmed[0] == "'":
                    return core.replace("''", "'")
                return core
            return trimmed

        def _is_probable_url(value: str) -> bool:
            if not isinstance(value, str):
                return False
            parsed = urlparse(value.strip())
            if not parsed.scheme:
                return False
            if parsed.scheme.lower() == "file":
                return bool(parsed.path)
            return bool(parsed.netloc)

        workbook_dir: Optional[Path] = None
        try:
            workbook_path = getattr(actions.book, "fullname", "") or getattr(actions.book, "full_name", "")
        except Exception:
            workbook_path = ""
        if workbook_path:
            try:
                workbook_dir = Path(workbook_path).resolve().parent
            except Exception:
                workbook_dir = None

        target_sheet, normalized_range = _split_sheet_and_range(cell_range, sheet_name)
        source_rows, source_cols = _parse_range_dimensions(normalized_range)

        raw_original = actions.read_range(normalized_range, target_sheet)
        original_data = _reshape_to_dimensions(raw_original, source_rows, source_cols)
        original_data = _unescape_matrix_values(original_data)

        if source_rows == 0 or source_cols == 0:
            return f"Range '{cell_range}' has no usable cells to translate."

        source_matrix = [row[:] for row in original_data]
        range_adjustment_note: Optional[str] = None
        writing_to_source_directly = translation_output_range is None
        include_context_columns = output_mode != "translation_only"
        should_apply_formatting = output_mode != "translation_only"
        translation_block_width = (
            _MIN_CONTEXT_BLOCK_WIDTH if include_context_columns else 1
        )

        effective_length_ratio_limit: Optional[float] = None
        effective_length_ratio_min: Optional[float] = None
        if length_ratio_limit is not None:
            try:
                candidate_limit = float(length_ratio_limit)
            except (TypeError, ValueError):
                raise ToolExecutionError("length_ratio_limit は数値で指定してください。") from None
            if not math.isfinite(candidate_limit) or candidate_limit <= 0:
                raise ToolExecutionError("length_ratio_limit には 0 より大きい有限の数値を指定してください。")
            effective_length_ratio_limit = candidate_limit
        if length_ratio_min is not None:
            try:
                candidate_min = float(length_ratio_min)
            except (TypeError, ValueError):
                raise ToolExecutionError("length_ratio_min は数値で指定してください。") from None
            if not math.isfinite(candidate_min) or candidate_min < 0:
                raise ToolExecutionError("length_ratio_min には 0 以上の有限の数値を指定してください。")
            effective_length_ratio_min = candidate_min
        if (
            effective_length_ratio_min is not None
            and effective_length_ratio_limit is not None
            and effective_length_ratio_min > effective_length_ratio_limit
        ):
            raise ToolExecutionError("length_ratio_min は length_ratio_limit 以下にしてください。")

        enforce_length_limit = (
            output_mode == "translation_only"
            and (
                effective_length_ratio_limit is not None
                or effective_length_ratio_min is not None
            )
        )

        length_metrics: Dict[Tuple[int, int], Dict[str, Any]] = {}
        length_limit_violations: List[str] = []
        length_violation_positions: Set[Tuple[int, int]] = set()

        def _compute_length_ratio(source_units: int, translated_units: int) -> float:
            if source_units <= 0:
                return 0.0 if translated_units <= 0 else math.inf
            return translated_units / source_units

        def _ratio_violation_kind(ratio_value: float) -> Optional[str]:
            if effective_length_ratio_min is not None and ratio_value < effective_length_ratio_min:
                return "below"
            if effective_length_ratio_limit is not None and ratio_value > effective_length_ratio_limit:
                return "above"
            return None

        def _format_ratio_bounds_for_display() -> str:
            if (
                effective_length_ratio_min is not None
                and effective_length_ratio_limit is not None
            ):
                return f"{effective_length_ratio_min:.2f}〜{effective_length_ratio_limit:.2f}"
            if effective_length_ratio_min is not None:
                return f"{effective_length_ratio_min:.2f} 以上"
            if effective_length_ratio_limit is not None:
                return f"{effective_length_ratio_limit:.2f} 以下"
            return ""

        def _format_ratio(value: float) -> str:
            if not math.isfinite(value):
                return "∞"
            return f"{value:.2f}"

        _METADATA_RATIO_TOLERANCE = 0.01

        def _coerce_int(value: Any) -> Optional[int]:
            if isinstance(value, bool):
                return None
            if isinstance(value, int):
                return value
            if isinstance(value, float):
                if not math.isfinite(value):
                    return None
                rounded = int(round(value))
                if abs(value - rounded) < 1e-6:
                    return rounded
                return None
            if isinstance(value, str):
                cleaned = value.strip()
                if not cleaned:
                    return None
                try:
                    if "." in cleaned or "e" in cleaned.lower():
                        candidate = float(cleaned)
                        if not math.isfinite(candidate):
                            return None
                        rounded = int(round(candidate))
                        if abs(candidate - rounded) < 1e-6:
                            return rounded
                        return None
                    return int(cleaned)
                except ValueError:
                    return None
            return None

        def _coerce_float(value: Any) -> Optional[float]:
            if isinstance(value, bool):
                return None
            if isinstance(value, (int, float)):
                if not math.isfinite(float(value)):
                    return None
                return float(value)
            if isinstance(value, str):
                cleaned = value.strip()
                if not cleaned:
                    return None
                try:
                    candidate = float(cleaned)
                except ValueError:
                    return None
                if not math.isfinite(candidate):
                    return None
                return candidate
            return None

        def _extract_reported_length_metadata(payload_item: Mapping[str, Any]) -> Tuple[Optional[int], Optional[int], Optional[float]]:
            reported_source = _coerce_int(payload_item.get("source_length"))
            reported_translated = _coerce_int(payload_item.get("translated_length"))
            reported_ratio = _coerce_float(payload_item.get("length_ratio"))
            if (
                reported_source is None
                or reported_translated is None
                or reported_ratio is None
            ) and isinstance(payload_item.get("length_verification"), Mapping):
                verification = payload_item.get("length_verification")
                result_mapping = verification.get("result") if isinstance(verification.get("result"), Mapping) else None  # type: ignore[assignment]
                if isinstance(result_mapping, Mapping):
                    if reported_source is None:
                        reported_source = _coerce_int(result_mapping.get("source_length"))
                    if reported_translated is None:
                        reported_translated = _coerce_int(result_mapping.get("translated_length"))
                    if reported_ratio is None:
                        reported_ratio = _coerce_float(result_mapping.get("length_ratio"))
            return reported_source, reported_translated, reported_ratio


        def _repair_unescaped_length_verification_result(response_text: Any) -> Tuple[Any, Optional[str]]:
            if not isinstance(response_text, str):
                return response_text, "non_string"
            if '"length_verification"' not in response_text:
                return response_text, "no_length_verification_section"

            cursor = 0
            text = response_text
            marker_token = '"length_verification"'
            attempted_repair = False
            repair_applied = False
            repair_failed = False

            while cursor < len(text):
                marker_index = text.find(marker_token, cursor)
                if marker_index == -1:
                    break

                result_index = text.find('"result"', marker_index)
                if result_index == -1:
                    cursor = marker_index + len(marker_token)
                    continue

                colon_index = text.find(":", result_index)
                if colon_index == -1:
                    cursor = result_index + 7
                    continue

                value_quote_index = text.find('"', colon_index)
                if value_quote_index == -1:
                    cursor = colon_index + 1
                    continue

                inner_start_index = value_quote_index + 1
                if inner_start_index >= len(text) or text[inner_start_index] != "{":
                    cursor = inner_start_index
                    continue

                brace_depth = 0
                position = inner_start_index
                closing_brace_index: Optional[int] = None
                while position < len(text):
                    char = text[position]
                    if char == "{":
                        brace_depth += 1
                    elif char == "}":
                        brace_depth -= 1
                        if brace_depth == 0:
                            closing_brace_index = position
                            break
                    position += 1

                if closing_brace_index is None:
                    break

                closing_quote_index = closing_brace_index + 1
                if closing_quote_index >= len(text) or text[closing_quote_index] != '"':
                    cursor = closing_brace_index + 1
                    continue

                inner_payload = text[inner_start_index : closing_brace_index + 1]
                attempted_repair = True
                try:
                    parsed_payload = json.loads(inner_payload)
                except json.JSONDecodeError:
                    try:
                        parsed_payload = ast.literal_eval(inner_payload)
                    except (ValueError, SyntaxError):
                        repair_failed = True
                        cursor = closing_brace_index + 1
                        continue
                if not isinstance(parsed_payload, (dict, list)):
                    cursor = closing_brace_index + 1
                    continue

                safe_payload = json.dumps(parsed_payload, ensure_ascii=False)
                text = (
                    text[:value_quote_index]
                    + safe_payload
                    + text[closing_quote_index + 1 :]
                )
                if safe_payload != inner_payload:
                    repair_applied = True
                cursor = value_quote_index + len(safe_payload)

            if repair_applied:
                return text, "repaired"
            if repair_failed:
                return text, "repair_failed"
            if attempted_repair:
                return text, "repair_not_needed"
            return text, "no_length_verification_section"

        def _strip_json_code_fences(text: str) -> str:
            if not isinstance(text, str):
                return text
            stripped = text.strip()
            if not stripped.startswith("```"):
                return stripped
            fence_match = re.match(r"^```(?:json)?\s*", stripped, re.IGNORECASE)
            if not fence_match:
                return stripped
            content = stripped[fence_match.end():]
            if content.endswith("```"):
                content = content[:-3].rstrip()
            return content

        def _extract_first_json_payload(response_text: Any) -> Tuple[Optional[Any], Optional[str]]:
            if not isinstance(response_text, str):
                return None, "non_string"
            cleaned = _strip_json_code_fences(response_text)
            decoder = json.JSONDecoder()
            index = 0
            length = len(cleaned)
            found_candidate = False
            decode_error_occurred = False
            decoded_payloads: List[Tuple[Any, int]] = []
            while index < length:
                char = cleaned[index]
                if char in ("{", "["):
                    found_candidate = True
                    try:
                        payload, end_index = decoder.raw_decode(cleaned, index)
                    except json.JSONDecodeError:
                        decode_error_occurred = True
                        index += 1
                        continue
                    else:
                        decoded_payloads.append((payload, end_index))
                        index = end_index
                        continue
                index += 1
            if decoded_payloads:
                for payload, _ in reversed(decoded_payloads):
                    if isinstance(payload, list):
                        return payload, None
                return decoded_payloads[-1][0], None
            stripped = cleaned.strip()
            if not stripped:
                return None, "empty_response"
            if not found_candidate:
                return None, "no_json_marker"
            if decode_error_occurred:
                return None, "json_decode_failed"
            return None, "no_json_found"


        citation_should_include_explanations = writing_to_source_directly and include_context_columns
        out_rows = source_rows
        out_cols = source_cols if writing_to_source_directly else source_cols * translation_block_width
        if writing_to_source_directly and not overwrite_source:
            raise ToolExecutionError(
                "translation_output_range must be provided when overwrite_source is False."
            )
        if writing_to_source_directly:
            if include_context_columns:
                raise ToolExecutionError(
                    "translation_output_range must be provided when references are enabled so that explanations and reference pairs can be written."
                )
            output_sheet = target_sheet
            output_range = normalized_range
            output_matrix = source_matrix
            out_cols = source_cols
        else:
            output_sheet, output_range = _split_sheet_and_range(translation_output_range, target_sheet)
            out_rows, out_cols = _parse_range_dimensions(output_range)
            min_required_width = _MIN_CONTEXT_BLOCK_WIDTH if include_context_columns else 1
            if out_rows < source_rows:
                raise ToolExecutionError(
                    "translation_output_range must span the same number of rows as the source range."
                )
            per_column_width = math.ceil(out_cols / source_cols)
            if per_column_width < min_required_width:
                per_column_width = min_required_width
            adjusted_total_cols = per_column_width * source_cols
            if out_cols != adjusted_total_cols:
                start_row, start_col, _, _ = _parse_range_bounds(output_range)
                adjusted_end_row = start_row + source_rows - 1
                adjusted_end_col = start_col + adjusted_total_cols - 1
                adjusted_range = _build_range_reference(
                    start_row,
                    adjusted_end_row,
                    start_col,
                    adjusted_end_col,
                )
                original_range_display = translation_output_range
                adjusted_range_display = (
                    f"{output_sheet}!{adjusted_range}" if output_sheet else adjusted_range
                )
                range_adjustment_note = (
                    f"translation_output_range '{original_range_display}' was resized to '{adjusted_range_display}' "
                    "to maintain a constant column block per source column."
                )
                output_range = adjusted_range
                out_rows, out_cols = source_rows, adjusted_total_cols
            translation_block_width = out_cols // source_cols
            if include_context_columns and translation_block_width < min_required_width:
                translation_block_width = min_required_width
            raw_output = actions.read_range(output_range, output_sheet)
            try:
                output_matrix = _reshape_to_dimensions(raw_output, out_rows, out_cols)
            except ToolExecutionError:
                output_matrix = [["" for _ in range(out_cols)] for _ in range(out_rows)]
            output_matrix = _unescape_matrix_values(output_matrix)
        max_reference_pairs_per_item = (
            max(0, translation_block_width - 2) if include_context_columns else 0
        )

        _ensure_not_stopped()

        reference_warning_notes: List[str] = []
        translation_cache: Dict[str, Dict[str, Any]] = {}

        def _dedupe_preserve_order(values: List[str]) -> List[str]:
            seen: Set[str] = set()
            ordered: List[str] = []
            for entry in values:
                if entry in seen:
                    continue
                seen.add(entry)
                ordered.append(entry)
            return ordered

        def _collect_reference_entries(raw_values: Optional[List[Any]], label: str) -> Tuple[List[Dict[str, str]], List[str]]:
            entries: List[Dict[str, str]] = []
            invalid_tokens: List[str] = []
            seen_urls: Set[str] = set()

            if not raw_values:
                return entries, []

            for raw_value in raw_values:
                _ensure_not_stopped()
                if raw_value is None:
                    continue
                if not isinstance(raw_value, str):
                    raise ToolExecutionError(f"Each entry in {label} must be a string.")

                original_value = raw_value
                url = _strip_enclosing_quotes(raw_value)
                if not url:
                    continue
                normalized_url = url.strip()
                if not normalized_url:
                    continue

                if not _is_probable_url(normalized_url):
                    invalid_tokens.append(original_value or "(空文字列)")
                    continue

                try:
                    parsed = urlparse(normalized_url)
                    scheme = (parsed.scheme or "").lower()
                    has_remote_netloc = bool(parsed.netloc)
                except Exception:
                    scheme = ""
                    has_remote_netloc = False

                if scheme not in {"http", "https"} or not has_remote_netloc:
                    invalid_tokens.append(original_value or "(空文字列)")
                    continue

                resolved_url = normalized_url
                if resolved_url not in seen_urls:
                    seen_urls.add(resolved_url)
                    entries.append({
                        "id": f"{label[:1].upper()}{len(entries) + 1}",
                        "url": resolved_url,
                    })

            warnings: List[str] = []
            if invalid_tokens:
                invalid_urls = ", ".join(_dedupe_preserve_order(invalid_tokens))
                warnings.append(
                    f"{label} の値 {invalid_urls} は HTTP(S) URL ではないため除外しました。"
                )
            return entries, warnings

        def _canonicalize_source_text(value: Any) -> str:
            if not isinstance(value, str):
                return ""
            normalized = _normalize_cell_value(value)
            if not isinstance(normalized, str):
                return ""
            normalized = normalized.replace("\r\n", "\n").replace("\r", "\n")
            return normalized


        source_reference_inputs: List[Any] = []
        target_reference_inputs: List[Any] = []

        for candidate in (reference_urls, source_reference_urls):
            if candidate is None:
                continue
            if isinstance(candidate, (list, tuple, set)):
                source_reference_inputs.extend(candidate)
            else:
                source_reference_inputs.append(candidate)

        if target_reference_urls is not None:
            if isinstance(target_reference_urls, (list, tuple, set)):
                target_reference_inputs.extend(target_reference_urls)
            else:
                target_reference_inputs.append(target_reference_urls)

        source_reference_url_entries, source_warnings = _collect_reference_entries(source_reference_inputs, "source_reference_urls")
        target_reference_url_entries, target_warnings = _collect_reference_entries(target_reference_inputs, "target_reference_urls")

        reference_warning_notes.extend(source_warnings)
        reference_warning_notes.extend(target_warnings)

        references_requested = bool(source_reference_inputs or target_reference_inputs)
        use_references = bool(source_reference_url_entries or target_reference_url_entries)

        if references_requested and not use_references:
            reference_warning_notes.append(
                "参照が読み取れなかったため、参照なしの翻訳モードで続行しました。"
            )

        def _sanitize_evidence_value(value: str) -> str:
            cleaned = value.strip()
            if cleaned.lower().startswith("source:"):
                cleaned = cleaned.split(":", 1)[1].strip()
            return cleaned

        def _strip_reference_urls_from_quote(text: str) -> str:
            """Remove embedded URLs while preserving bracket-only citation markers."""

            if not isinstance(text, str):
                return text

            cleaned = _BRACKETED_URL_PATTERN.sub(lambda match: f"[{match.group(1)}]", text)
            cleaned = re.sub(r"\((?:https?|ftp)://[^)]+\)", "", cleaned)
            cleaned = re.sub(r"(?:https?|ftp)://\S+", "", cleaned)
            cleaned = re.sub(r"[ \t]+", " ", cleaned)
            cleaned = re.sub(r" \n", "\n", cleaned)
            cleaned = re.sub(r"\n ", "\n", cleaned)
            return cleaned.strip()

        def _expand_keyword_variants(keywords: List[str], max_variants: int) -> List[str]:
            variants: List[str] = []
            for keyword in keywords:
                for variant in _generate_keyword_variants(keyword):
                    if variant not in variants:
                        variants.append(variant)
                    if len(variants) >= max_variants:
                        return variants
            return variants


        ratio_bounds_display = _format_ratio_bounds_for_display() if enforce_length_limit else ""
        prompt_parts: List[str]
        if include_context_columns:
            prompt_parts = [
                f"あなたは日本語原文と参照対訳ペアを受け取り、{target_language} への翻訳を生成するアシスタントです。\n",
                "各行の日本語原文を入力順のまま漏れなく翻訳し、未訳部分を残さないでください。\n",
                "参照ペアに含まれる訳語・表現を最優先で再利用し、同じ概念であれば語彙・句・スタイルを可能な限り一致させてください。\n",
                "参照ペアに完全一致する表現がない場合も、近い表現を組み合わせるなどして用語と表現の一貫性を維持してください。\n",
                f"自然な {target_language} の文章になるよう文法を整えて構いませんが、原文と参照ペアに含まれない情報や語釈を追加しないでください。\n",
                "出力は JSON 配列のみです。各要素には次のキーを必ず含めてください。\n",
                f"- \"translated_text\": {target_language} での翻訳文。\n",
                "- \"source_length\": 原文の UTF-16 コードユニット数（整数）。\n",
                "- \"translated_length\": 訳文の UTF-16 コードユニット数（整数）。\n",
                "- \"length_ratio\": translated_length / source_length（数値）。\n",
                "- \"process_notes_jp\": 日本語で数行の翻訳メモ。訳語の根拠や参照ペアの使い方を簡潔に記述してください。\n",
                "- \"reference_pairs\": 実際に参照したペアの配列。利用しなかった場合は空配列を返してください。\n",
                "余計なコメントやマークダウンを付けず、純粋な JSON だけを返してください。\n",
                "以下に日本語原文の配列、続いて参照ペアの配列を示します。\n",
            ]
            prompt_preamble = "".join(prompt_parts)
        else:

            # translation_only 用プロンプト（初回で制約を満たす）
            prompt_lines = [
                f"以下の日本語テキストを {target_language} に翻訳してください。\n",
                "出力は JSON 配列のみで、要素順は入力テキスト順と一致させてください。各要素には必ず次のキーを含めてください:\n",
                '- "translated_text": 訳文（空文字列禁止、前回応答の同一訳語再利用禁止）。\n',
                '- "source_length": 原文の UTF-16 コードユニット数（整数）。\n',
                '- "translated_length": len(translated_text.encode(\"utf-16-le\")) // 2 で得た整数値。\n',
                '- "length_ratio": translated_length / source_length を小数第2位まで四捨五入した数値。\n',
                '- "length_verification": { "method": "utf16-le", "translated_length_computed": 同じ整数, "length_ratio_computed": 同じ数値, "status": "ok" }\n',
                "UTF-16 長さは UTF-16LE でエンコードしたバイト長を 2 で割った正確な値のみを使用し、概算や重複計算は禁止です。\n",
                "検算では Python 等で `len(translated_text.encode(\"utf-16-le\")) // 2` を実行し、得た値を translated_length と length_verification.translated_length_computed の両方に転記してください。\n",
                "translated_length と length_ratio は必ず同じ計算結果から導き、出力前に再チェックしてズレがあれば修正してください。\n",
                "length_verification.status はすべての値が一致した場合のみ \"ok\" とし、一致しない場合は JSON を出力せず再計算してください。\n",
                "重要: 後続の長さ調整は行いません。1 回の回答で文字数制約を必ず満たしてください。\n",
            ]
            if enforce_length_limit:
                if ratio_bounds_display:
                    prompt_lines.append(f"文字数倍率の目標レンジ: {ratio_bounds_display}。")
                prompt_lines.append("必須手順:\n")
                ratio_midpoint_display: Optional[str] = None
                if (
                    effective_length_ratio_min is not None
                    and effective_length_ratio_limit is not None
                ):
                    ratio_midpoint_display = f"{(effective_length_ratio_min + effective_length_ratio_limit) / 2:.2f}"
                if (
                    effective_length_ratio_min is not None
                    and effective_length_ratio_limit is not None
                ):
                    prompt_lines.append(
                        f"1. source_length に {effective_length_ratio_min:.2f}〜{effective_length_ratio_limit:.2f} を掛けて許容下限・上限を算出し、四捨五入して整数化してください。\n"
                    )
                elif effective_length_ratio_min is not None:
                    prompt_lines.append(
                        f"1. source_length × {effective_length_ratio_min:.2f} を許容下限として計算し、四捨五入して整数化してください。\n"
                    )
                elif effective_length_ratio_limit is not None:
                    prompt_lines.append(
                        f"1. source_length × {effective_length_ratio_limit:.2f} を許容上限として計算し、四捨五入して整数化してください。\n"
                    )
                prompt_lines.extend([
                    "2. 訳文案は許容範囲の中央付近を狙い、出力前に translated_length がレンジ内か確認してください。\n",
                    "3. 上限超過時は語句を圧縮し、下限未満時は意味を保った自然な補足で密度を調整し、再計算後に length_ratio がレンジ内か確認してください。\n",
                    "4. Python 等で再計算した translated_length と length_ratio を length_verification フィールドにも反映させ、値が完全に一致することを確認してください。\n",
                    "5. translated_length が許容上限を 1 でも超える場合は JSON を出力せず再構成してから検算してください。\n",
                ])
                if ratio_midpoint_display:
                    prompt_lines.append(
                        f"補足: 範囲中央 ({ratio_midpoint_display}) × source_length を目安の訳文長とし、自然で簡潔な文体を優先してください。\n"
                    )
                prompt_lines.append(
                    "直訳で長くなりがちな場合は見出し風の簡潔表現や列挙を検討し、等位接続詞や冠詞、重複語を極力省いてください。\n"
                )
                prompt_lines.append(
                    "translated_length が許容上限を 1 でも超える場合は必ず再構成し、レンジ内に収まるまで JSON を出力しないでください。\n"
                )
                prompt_lines.append(
                    "見出し・項目名では 1 語訳や短い同義語を優先し、不要な接続詞（and 等）や語尾の反復を避けてください。\n"
                )
                prompt_lines.append(
                    "列挙や見出しを訳す場合は 'and' を使わずにコンマやスラッシュで区切り、各項目を 1 語程度に簡潔化してください（例示不要）。\n"
                )
                prompt_lines.append(
                    "以前の応答でレンジ外だった訳文を再利用せず、語数や語彙を調整した新しい案を必ず提示してください。\n"
                )
                prompt_lines.append(
                    "直前にレンジ外と判定された語句（例: \"Tariff Impact\" のような長い表現）は再利用禁止です。語彙を圧縮した別案に置き換えてください。\n"
                )
                prompt_lines.append(
                    "短い同義語が必要な場合は 1 語訳を優先的に検討し、UTF-16 長が収まるまで語を調整してください。\n"
                )
                prompt_lines.append(
                    "文章として訳す場合は必要な接続詞の使用を許容しますが、冗長な節や重複表現を削除し、簡潔な語順に再構成してください。\n"
                )
                prompt_lines.append(
                    "複合語が長くなる場合は意味を保ったまま一般的な単語 1 語に置き換えてください。\n"
                )
                prompt_lines.append(
                    "JSON を返す直前に各要素の translated_length を再計算し、許容上限を超えていないことを確認した上で出力してください。\n"
                )
                prompt_lines.append("全行について length_ratio が許容レンジ内であると確信するまで JSON を出力しないでください。\n")
            prompt_lines.append("禁止事項:\n")
            prompt_lines.append("  - JSON 以外の出力、複数 JSON、前置き、マークダウン、冗長な説明。\n")
            prompt_lines.append("  - 許容レンジ外の値を含んだまま出力すること。\n")
            prompt_lines.append("  - 再利用禁止語句や直前にレンジ外と判断された訳語の使い回し。\n")
            prompt_lines.extend([
                "各要素は必ず 1 本の訳文のみを返してください（見出し・注釈を追加しない）。",
            ])
            prompt_preamble = "".join(prompt_lines)

        if references_requested or use_references:
            rows_per_batch = 1
        batch_size = rows_per_batch if rows_per_batch is not None else 1
        if batch_size < 1:
            raise ToolExecutionError("rows_per_batch must be at least 1.")

        source_start_row, source_start_col, _, _ = _parse_range_bounds(normalized_range)
        output_start_row, output_start_col, _, _ = _parse_range_bounds(output_range)
        output_total_cols = out_cols if not writing_to_source_directly else source_cols

        row_dirty_flags: List[bool] = [False] * source_rows
        source_row_dirty_flags: List[bool] = [False] * source_rows
        pending_columns_by_row: Dict[int, Set[int]] = {}
        for row_idx in range(source_rows):
            pending_cols: Set[int] = set()
            for col_idx in range(source_cols):
                cell_value = original_data[row_idx][col_idx]
                if not isinstance(cell_value, str):
                    continue
                normalized_cell = cell_value.replace("\r\n", "\n").replace("\r", "\n")
                if JAPANESE_CHAR_PATTERN.search(normalized_cell):
                    pending_cols.add(col_idx)
            if pending_cols:
                pending_columns_by_row[row_idx] = pending_cols

        completed_rows: Set[int] = set()
        incremental_row_messages: List[str] = []

        def _cell_reference(base_row: int, base_col: int, local_row: int, local_col: int) -> str:
            return _build_range_reference(
                base_row + local_row,
                base_row + local_row,
                base_col + local_col,
                base_col + local_col,
            )

        def _output_row_reference(row_idx: int) -> str:
            end_col = output_start_col + output_total_cols - 1
            return _build_range_reference(
                output_start_row + row_idx,
                output_start_row + row_idx,
                output_start_col,
                end_col,
            )

        def _source_row_reference(row_idx: int) -> str:
            end_col = source_start_col + source_cols - 1
            return _build_range_reference(
                source_start_row + row_idx,
                source_start_row + row_idx,
                source_start_col,
                end_col,
            )

        def _translation_column_index(col_idx: int) -> int:
            return col_idx if writing_to_source_directly else col_idx * translation_block_width

        def _compose_row_progress_message(row_idx: int) -> str:
            excel_row_number = source_start_row + row_idx + 1
            fragments: List[str] = []
            for col_idx in range(source_cols):
                translation_col = _translation_column_index(col_idx)
                if translation_col >= output_total_cols:
                    continue
                cell_address = _cell_reference(output_start_row, output_start_col, row_idx, translation_col)
                try:
                    translation_cell_value = output_matrix[row_idx][translation_col]
                except IndexError:
                    translation_cell_value = ""
                normalized_value = ""
                if translation_cell_value is not None:
                    normalized_value = _normalize_cell_value(translation_cell_value).strip()
                if len(normalized_value) > 80:
                    normalized_value = normalized_value[:77] + "..."
                fragments.append(f"{cell_address}='{normalized_value}'")
            summary = "; ".join(fragments) if fragments else "no translations"
            return f"Row {excel_row_number} translation completed: {summary}"

        def _apply_cached_translation(local_row: int, col_idx: int, source_key: str, cached_entry: Dict[str, Any]) -> None:
            nonlocal output_dirty, source_dirty, any_translation, reused_translation_detected
            if include_context_columns:
                return
            translation_value_raw = cached_entry.get("translation")
            if not isinstance(translation_value_raw, str):
                return
            translation_value = translation_value_raw
            if not isinstance(translation_value, str):
                translation_value = str(translation_value)
            translation_value = _maybe_unescape_html_entities(translation_value)
            if not isinstance(translation_value, str):
                translation_value = str(translation_value)
            translation_value_stripped = translation_value.strip()
            if not translation_value_stripped:
                return

            applied_changes = False
            translation_col_index_seed = col_idx if writing_to_source_directly else col_idx * translation_block_width
            cell_ref_for_metrics = _cell_reference(
                output_start_row,
                output_start_col,
                local_row,
                translation_col_index_seed,
            )

            existing_output_value = ""
            try:
                existing_output_value = output_matrix[local_row][translation_col_index_seed]
            except (IndexError, TypeError):
                existing_output_value = ""

            if translation_value != existing_output_value:
                if local_row < len(output_matrix):
                    row_values = output_matrix[local_row]
                    if translation_col_index_seed < len(row_values):
                        row_values[translation_col_index_seed] = translation_value
                        row_dirty_flags[local_row] = True
                        output_dirty = True
                        applied_changes = True

            if not writing_to_source_directly and overwrite_source:
                try:
                    existing_source_value = source_matrix[local_row][col_idx]
                except (IndexError, TypeError):
                    existing_source_value = ""
                if translation_value != existing_source_value:
                    if local_row < len(source_matrix):
                        source_matrix[local_row][col_idx] = translation_value
                        source_dirty = True
                        source_row_dirty_flags[local_row] = True
                        applied_changes = True

            reused_translation_detected = True
            if applied_changes:
                any_translation = True

            source_text_canonical = source_key
            source_length_units = _measure_utf16_length(source_text_canonical)
            translated_length_units = _measure_utf16_length(translation_value)
            ratio_value = _compute_length_ratio(source_length_units, translated_length_units)

            violation_kind = _ratio_violation_kind(ratio_value) if enforce_length_limit else None

            metric_entry = {
                "source_length": source_length_units,
                "translated_length": translated_length_units,
                "ratio": ratio_value,
                "cell_ref": cell_ref_for_metrics,
                "limit": effective_length_ratio_limit,
                "min_limit": effective_length_ratio_min,
                "status": violation_kind or "ok",
                "reported_source_length": None,
                "reported_translated_length": None,
                "reported_ratio": None,
            }
            length_metrics[(local_row, col_idx)] = metric_entry

            if enforce_length_limit and violation_kind and (local_row, col_idx) not in length_violation_positions:
                direction_label = "上限" if violation_kind == "above" else "下限"
                ratio_text = _format_ratio(ratio_value)
                length_limit_violations.append(f"{cell_ref_for_metrics}: ×{ratio_text} ({direction_label}逸脱)")
                length_violation_positions.add((local_row, col_idx))
                actions.log_progress(
                    f"既存の翻訳 {cell_ref_for_metrics} が文字数倍率制約を満たしていません: ×{ratio_text} ({direction_label}逸脱)。"
                )

            pending_cols = pending_columns_by_row.get(local_row)
            if pending_cols is not None:
                pending_cols.discard(col_idx)
                if not pending_cols:
                    _finalize_row(local_row)

        pending_output_segment: Optional[Dict[str, Any]] = None
        pending_source_segment: Optional[Dict[str, Any]] = None

        def _flush_output_segment() -> None:
            nonlocal pending_output_segment
            if pending_output_segment is None:
                return
            start_row_idx = pending_output_segment["start_row"]
            end_row_idx = pending_output_segment["end_row"]
            rows_data = pending_output_segment["rows"]
            range_ref = _build_range_reference(
                output_start_row + start_row_idx,
                output_start_row + end_row_idx,
                output_start_col,
                output_start_col + output_total_cols - 1,
            )
            write_message = actions.write_range(range_ref, rows_data, output_sheet, apply_formatting=should_apply_formatting)
            incremental_row_messages.append(write_message)
            pending_output_segment = None

        def _flush_source_segment() -> None:
            nonlocal pending_source_segment
            if pending_source_segment is None:
                return
            start_row_idx = pending_source_segment["start_row"]
            end_row_idx = pending_source_segment["end_row"]
            rows_data = pending_source_segment["rows"]
            range_ref = _build_range_reference(
                source_start_row + start_row_idx,
                source_start_row + end_row_idx,
                source_start_col,
                source_start_col + source_cols - 1,
            )
            write_message = actions.write_range(range_ref, rows_data, target_sheet, apply_formatting=should_apply_formatting)
            incremental_row_messages.append(write_message)
            pending_source_segment = None

        def _flush_pending_segments() -> None:
            _flush_output_segment()
            _flush_source_segment()

        def _queue_output_row(row_idx: int, row_slice: List[Any]) -> None:
            nonlocal pending_output_segment
            if pending_output_segment is not None:
                if row_idx == pending_output_segment["end_row"] + 1:
                    pending_output_segment["rows"].append(row_slice)
                    pending_output_segment["end_row"] = row_idx
                    return
                _flush_output_segment()
            pending_output_segment = {
                "start_row": row_idx,
                "end_row": row_idx,
                "rows": [row_slice],
            }

        def _queue_source_row(row_idx: int, row_slice: List[Any]) -> None:
            nonlocal pending_source_segment
            if pending_source_segment is not None:
                if row_idx == pending_source_segment["end_row"] + 1:
                    pending_source_segment["rows"].append(row_slice)
                    pending_source_segment["end_row"] = row_idx
                    return
                _flush_source_segment()
            pending_source_segment = {
                "start_row": row_idx,
                "end_row": row_idx,
                "rows": [row_slice],
            }

        def _write_row_output(row_idx: int) -> None:
            _ensure_not_stopped()
            wrote_anything = False
            if row_dirty_flags[row_idx]:
                row_slice = output_matrix[row_idx][:output_total_cols]
                _queue_output_row(row_idx, row_slice)
                row_dirty_flags[row_idx] = False
                wrote_anything = True
            if overwrite_source and not writing_to_source_directly and source_row_dirty_flags[row_idx]:
                source_slice = source_matrix[row_idx][:source_cols]
                _queue_source_row(row_idx, source_slice)
                source_row_dirty_flags[row_idx] = False
                wrote_anything = True
            progress_message = _compose_row_progress_message(row_idx)
            if not wrote_anything:
                progress_message += " (no changes needed)"
            actions.log_progress(progress_message)
            completed_rows.add(row_idx)

        def _finalize_row(row_idx: int) -> None:
            if row_idx in completed_rows:
                return
            pending_columns_by_row.pop(row_idx, None)
            _write_row_output(row_idx)

        citation_sheet = None
        citation_range = None
        citation_matrix: Optional[List[List[str]]] = None
        cite_start_row = cite_start_col = cite_rows = cite_cols = 0
        citation_mode: Optional[str] = None
        citation_note: Optional[str] = None
        if use_references:
            if not citation_output_range:
                citation_note = (
                    "citation_output_range was not provided; evidence details were retained within the translation output range."
                )
            else:
                citation_sheet, citation_range = _split_sheet_and_range(citation_output_range, target_sheet)
                cite_rows, cite_cols = _parse_range_dimensions(citation_range)
                if cite_rows != source_rows:
                    raise ToolExecutionError(
                        "citation_output_range must span the same number of rows as the source range."
                    )
                if cite_cols == 1:
                    citation_mode = "single_column"
                elif cite_cols == source_cols:
                    citation_mode = "per_cell"
                elif cite_cols == source_cols * 2:
                    citation_mode = "paired_columns"
                elif cite_cols == source_cols * 3:
                    citation_mode = "translation_triplets"
                else:
                    fallback_note = (
                        "指定された citation_output_range の列数がサポート外のため、参照出力を翻訳結果の列に内包します。"
                    )
                    actions.log_progress(fallback_note)
                    citation_note = fallback_note
                    citation_sheet = None
                    citation_range = None
                    citation_matrix = None
                    citation_mode = None
                    cite_rows = cite_cols = cite_start_row = cite_start_col = 0
                if citation_mode is not None:
                    cite_start_row, cite_start_col, _, _ = _parse_range_bounds(citation_range)
                    existing_citation = actions.read_range(citation_range, citation_sheet)
                    try:
                        citation_matrix = _reshape_to_dimensions(existing_citation, cite_rows, cite_cols)
                    except ToolExecutionError:
                        citation_matrix = [["" for _ in range(cite_cols)] for _ in range(cite_rows)]
                    if citation_matrix is not None:
                        citation_matrix = _unescape_matrix_values(citation_matrix)

        messages: List[str] = []
        explanation_fallback_notes: List[str] = []
        any_translation = False
        reused_translation_detected = False
        output_dirty = False
        source_dirty = False

        limit_to_single = references_requested or use_references
        if limit_to_single:
            items_per_request = 1
        else:
            items_per_request = max(1, rows_per_batch or _ITEMS_PER_TRANSLATION_REQUEST)

        def _normalize_for_comparison(text: str) -> str:
            return re.sub(r"\s+", "", text)

        def _extract_source_sentences_batch(
            current_texts: List[str],
        ) -> List[List[str]]:
            if not (use_references and source_reference_url_entries):
                return [[] for _ in current_texts]
            if not current_texts:
                return []

            items_payload: List[Dict[str, Any]] = []
            for source_text in current_texts:
                normalized_source = source_text if isinstance(source_text, str) else ""
                entry = {
                    "source_text": normalized_source,
                }
                items_payload.append(entry)
            items_json = json.dumps(items_payload, ensure_ascii=False)
            source_reference_urls_payload: List[str] = [
                entry["url"]
                for entry in source_reference_url_entries
                if isinstance(entry.get("url"), str) and entry["url"].strip()
            ]
            source_reference_urls_json = json.dumps(source_reference_urls_payload, ensure_ascii=False)

            source_sentence_prompt_sections: List[str] = [
                "目的: items(JSON) に含まれる各 source_text と関連する日本語の引用文を、提供された参照URLからできるだけ多様に集めてください。",
                "",
                "指示:",
                "1. source_text に含まれる名詞・動詞・形容詞・副詞・接続詞・重要な数値・固有名詞を抽出し、同義語や言い換えも含めたキーワード候補を作成してください。",
                "2. 各キーワード候補を軸に参照URL本文を走査し、語句が完全一致または意味的に関連する文を広く収集してください。日本語の語尾違いや活用違いも許容して構いません。",
                "3. 文は参照URLに実際に記載されている日本語をそのまま引用し、語尾・句読点を含めて改変しないでください。要約・翻訳・新規生成は禁止です。",
                "4. 関連度が高い順に最大10件まで source_sentences に並べてください。10件未満しか見つからない場合は、取得できた文のみを返してください。",
                "5. 重複する文や前のアイテムで既に列挙した文は除外し、同じ文書でも視点が異なる文を優先してください。",
                "6. 脚注番号 ([1] など)・リンク・リスト記号など本文以外の装飾は削除し、純粋な文だけを保持してください。",
                "7. 参照URL以外の情報源や外部検索は利用せず、適切な文が見つからない場合は空配列 [] を返してください。",
                "",
                "出力形式:",
                "JSON のみを返してください。例: [{\"source_sentences\": [\"...\"]}]. items(JSON) と同じ順序で並べてください。",
                "source_sentences には文字列のリストだけを入れ、追加のキーやテキストは不要です。",
                "",
                "items(JSON):",
                items_json,
            ]
            source_sentence_prompt_sections = [
                "目的: items(JSON) に含まれる各 source_text と関連する日本語の引用文を、提供された参照URLからできるだけ多様に集めてください。",
                "",
                "指示:",
                "1. source_text に含まれる名詞・動詞・形容詞・副詞・接続詞・重要な数値・固有名詞を抽出し、同義語や言い換えも含めたキーワード候補を作成してください。",
                "2. 作成したキーワード候補を用いて主要語→補助語→関連語の順で複数ラウンドの検索を行い、語形変化や複合語も組み合わせてください。source_text の異なる側面が拾えるよう、キーワードの組み合わせをこまめに切り替えてください。",
                "3. 同じ参照URL内でも段落・見出し・視点が異なる箇所を優先し、既出の引用と語彙・言い回し・文構造が似通う文は次候補に回してください。必要に応じて別の参照URLにも切り替え、網羅的に探索してください。",
                "4. 文は参照URLに実際に記載されている日本語をそのまま引用し、語尾・句読点を含めて改変しないでください。要約・翻訳・新規生成は禁止です。",
                "5. 関連度が高い順に最大10件まで source_sentences に並べてください。10件未満しか見つからない場合は、取得できた文のみを返してください。",
                "6. 重複する文や前のアイテムで既に列挙した文は除外し、同じ文書でも視点が異なる文を優先してください。source_text と同じ語句ばかりが並ばないように調整してください。",
                "7. 脚注番号 ([1] など)・リンク・リスト記号など本文以外の装飾は削除し、純粋な文だけを保持してください。",
                "8. 参照URL以外の情報源や外部検索は利用せず、適切な文が見つからない場合は空配列 [] を返してください。",
                "",
                "出力形式:",
                "JSON のみを返してください。例: [{\"source_sentences\": [\"...\"]}]. items(JSON) と同じ順序で並べてください。",
                "source_sentences には文字列のリストだけを入れ、追加のキーやテキストは不要です。",
                "",
                "items(JSON):",
                items_json,
            ]

            if source_reference_urls_payload:
                source_sentence_prompt_sections.extend(
                    [
                        "",
                        "source_reference_urls(JSON):",
                        source_reference_urls_json,
                    ]
                )

            source_sentence_prompt = "\n".join(source_sentence_prompt_sections)
            _ensure_not_stopped()
            actions.log_progress("日本語参照文章抽出: Copilotに依頼中...")
            source_sentence_response = browser_manager.ask(source_sentence_prompt, stop_event=stop_event)
            source_sentence_items, parse_error_code = _extract_first_json_payload(source_sentence_response)
            if source_sentence_items is None:
                error_label = parse_error_code or "unknown"
                raise ToolExecutionError(
                    f"Failed to parse source reference response as JSON ({error_label}): {source_sentence_response}"
                )
            if not isinstance(source_sentence_items, list) or len(source_sentence_items) != len(current_texts):
                raise ToolExecutionError(
                    "Source reference response must be a list with one entry per source text."
                )

            cleaned_results: List[List[str]] = [[] for _ in current_texts]
            for item_index, entry in enumerate(source_sentence_items):
                raw_sentences: List[str] = []
                if isinstance(entry, dict):
                    raw_sentences = entry.get("source_sentences") or entry.get("sentences") or []
                elif isinstance(entry, list):
                    raw_sentences = entry
                if not isinstance(raw_sentences, list):
                    raw_sentences = []
                cleaned_sentences: List[str] = []
                original_text = current_texts[item_index] if item_index < len(current_texts) else ""
                original_normalized = _normalize_for_comparison(original_text) if isinstance(original_text, str) else ""

                for sentence in raw_sentences:
                    if not isinstance(sentence, str):
                        continue
                    stripped = sentence.strip()
                    if not stripped or stripped in cleaned_sentences:
                        continue
                    stripped = _strip_reference_urls_from_quote(stripped)
                    stripped = re.sub(r"\[\d+\]", "", stripped)
                    stripped = re.sub(r"\s{2,}", " ", stripped).strip()
                    if not stripped:
                        continue
                    candidate_normalized = _normalize_for_comparison(stripped)
                    if original_normalized and candidate_normalized == original_normalized:
                        continue
                    cleaned_sentences.append(stripped)
                    if len(cleaned_sentences) >= 10:
                        break
                cleaned_results[item_index] = cleaned_sentences
            return cleaned_results
        def _pair_target_sentences_batch(
            source_references_per_item: List[List[str]],
            current_texts: List[str],
        ) -> List[List[Dict[str, str]]]:
            if not (use_references and target_reference_url_entries):
                return [[] for _ in source_references_per_item]
            if not source_references_per_item:
                return [[] for _ in source_references_per_item]

            extraction_payload: List[Dict[str, Any]] = []
            for idx, source_sentences in enumerate(source_references_per_item):
                extraction_payload.append(
                    {
                        "source_sentences": source_sentences,
                        "source_text": current_texts[idx] if idx < len(current_texts) and isinstance(current_texts[idx], str) else "",
                    }
                )
            extraction_items_json = json.dumps(extraction_payload, ensure_ascii=False)
            target_reference_urls_payload: List[str] = [
                entry["url"]
                for entry in target_reference_url_entries
                if isinstance(entry.get("url"), str) and entry["url"].strip()
            ]
            target_reference_urls_json = json.dumps(target_reference_urls_payload, ensure_ascii=False)



            extraction_prompt_sections: List[str] = [
                (
                    f"タスク: items(JSON) の各要素について、`source_sentences` に含まれる日本語引用文と、指定された `target_reference_urls` からそのまま引用した {target_language} の文を必要なだけ対応付けてください。"
                ),
                "",
                "手順:",
                "- items(JSON) の順番に従って処理し、`source_text` は話題の把握だけに利用してください。",
                "- `target_reference_urls` で指定されたページ内のみを探索し、ナビゲーションやヘッダー、目次など本文外の要素は無視してください。",
                "- 文は掲載されているとおりにコピーし、翻訳・要約・言い換え・句読点や大小文字の変更は行わないでください。",
                "- 固有名詞や数値など特徴的な語が一致する文を優先し、曖昧または一般的な一致は採用しないでください。",
                "- 信頼できる一致が見つからない場合は、そのアイテムの `pairs` を空のままにしてください。",
                "",
                "出力形式:",
                "- 応答は `items(JSON)` と同じ長さ・順序の JSON 配列にしてください。",
                '- 各要素は `{"pairs": [{"source_sentence": "...", "target_sentence": "..."}]}` 形式のオブジェクトにしてください。',
                f"- `target_sentence` には参照資料からコピーした {target_language} の文を、`source_sentence` には対応する日本語引用文をそのまま記載してください。",
                "- 適切な一致が無い場合は `pairs` を空配列にしてください。",
                "",
                "items(JSON):",
                extraction_items_json,
            ]
            if target_reference_urls_payload:
                extraction_prompt_sections.extend(
                    [
                        "",
                        "target_reference_urls(JSON):",
                        target_reference_urls_json,
                    ]
                )

            extraction_prompt = "\n".join(extraction_prompt_sections)
            _ensure_not_stopped()
            actions.log_progress("英語参照文ペア抽出: Copilotに依頼中...")

            def _request_extraction(prompt: str) -> Tuple[Optional[List[Any]], str]:
                response = browser_manager.ask(prompt, stop_event=stop_event)
                payload, parse_error_code = _extract_first_json_payload(response)
                if payload is None or not isinstance(payload, list):
                    return None, response
                return payload, response

            extraction_items, raw_extraction_response = _request_extraction(extraction_prompt)
            if extraction_items is None:
                snippet = raw_extraction_response.strip().replace("\n", " ")
                actions.log_progress(
                    f"英語参照文ペア抽出が失敗しました: {snippet[:180]}{'...' if len(snippet) > 180 else ''}"
                )
                _ensure_not_stopped()
                retry_prompt_sections = [
                    "重要: 応答は JSON のみで返してください。",
                    "- 出力は `items(JSON)` と同じ順序・件数の JSON 配列にしてください。",
                    '- 各要素は {\"pairs\": [{\"source_sentence\": \"...\", \"target_sentence\": \"...\"}]} 形式にしてください。',
                    "- 文は元テキストをそのまま用い、一致しない場合は `pairs` を空配列にしてください。",
                    '- 例: [{\"pairs\": [{\"source_sentence\": \"…\", \"target_sentence\": \"…\"}]}, {\"pairs\": []}]',
                    "",
                    f"以下の日本語引用文に対応する {target_language} の参照文を再度抽出してください。",
                    "",
                    "items(JSON):",
                    extraction_items_json,
                ]
                if target_reference_urls_payload:
                    retry_prompt_sections.extend(
                        [
                            '',
                            'target_reference_urls(JSON):',
                            target_reference_urls_json,
                        ]
                    )
                retry_prompt = "\n".join(retry_prompt_sections)
                actions.log_progress(
                    '英語参照文ペア抽出をJSON指定で再試行します。'
                )
                extraction_items, raw_extraction_response = _request_extraction(retry_prompt)

            if extraction_items is None:
                snippet = raw_extraction_response.strip().replace("\n", " ")
                raise ToolExecutionError(
                    "ターゲット参照ペアの応答をJSONとして解析できませんでした: "
                    f"{snippet[:200]}{'…' if len(snippet) > 200 else ''}"
                )
            if not isinstance(extraction_items, list):
                raise ToolExecutionError("Target reference pair response must be a list.")
            if len(extraction_items) < len(source_references_per_item):
                extraction_items.extend({"pairs": []} for _ in range(len(source_references_per_item) - len(extraction_items)))

            cleaned_results: List[List[Dict[str, str]]] = [[] for _ in source_references_per_item]

            cleaned_results: List[List[Dict[str, str]]] = [[] for _ in source_references_per_item]
            for item_index, entry in enumerate(extraction_items):
                raw_pairs: List[Any] = []
                if isinstance(entry, dict):
                    raw_pairs = entry.get("pairs") or entry.get("reference_pairs") or []
                elif isinstance(entry, list):
                    raw_pairs = entry
                if not isinstance(raw_pairs, list):
                    raw_pairs = []
                cleaned_pairs: List[Dict[str, str]] = []
                seen_keys: Set[Tuple[str, str]] = set()
                for pair in raw_pairs:
                    if not isinstance(pair, dict):
                        continue
                    source_sentence = pair.get("source_sentence") or pair.get("jp") or ""
                    target_sentence = pair.get("target_sentence") or pair.get("translated") or pair.get("en") or ""
                    if not isinstance(source_sentence, str) or not isinstance(target_sentence, str):
                        continue
                    source_clean = source_sentence.strip()
                    target_clean = _strip_reference_urls_from_quote(target_sentence.strip())
                    if not source_clean or not target_clean:
                        continue
                    key = (source_clean, target_clean)
                    if key in seen_keys:
                        continue
                    seen_keys.add(key)
                    cleaned_pairs.append(
                        {
                            "source_sentence": source_clean,
                            "target_sentence": target_clean,
                        }
                    )
                cleaned_results[item_index] = cleaned_pairs
            return cleaned_results

        for row_start in range(0, source_rows, batch_size):
            _ensure_not_stopped()
            row_end = min(row_start + batch_size, source_rows)
            chunk_texts: List[str] = []
            chunk_positions: List[Tuple[int, int]] = []

            for local_row in range(row_start, row_end):
                _ensure_not_stopped()
                for col_idx in range(source_cols):
                    _ensure_not_stopped()
                    cell_value = original_data[local_row][col_idx]
                    if not isinstance(cell_value, str):
                        continue

                    canonical_source = _canonicalize_source_text(cell_value)
                    if not canonical_source:
                        pending_cols = pending_columns_by_row.get(local_row)
                        if pending_cols is not None:
                            pending_cols.discard(col_idx)
                            if not pending_cols:
                                _finalize_row(local_row)
                        continue

                    if not JAPANESE_CHAR_PATTERN.search(canonical_source):
                        pending_cols = pending_columns_by_row.get(local_row)
                        if pending_cols is not None:
                            pending_cols.discard(col_idx)
                            if not pending_cols:
                                _finalize_row(local_row)
                        continue

                    cached_entry_for_source = translation_cache.get(canonical_source)
                    if not include_context_columns and cached_entry_for_source:
                        _apply_cached_translation(local_row, col_idx, canonical_source, cached_entry_for_source)
                        continue

                    translation_col_index_seed = _translation_column_index(col_idx)
                    existing_output_value: Any = ""
                    existing_output_raw = ""
                    if not writing_to_source_directly:
                        try:
                            existing_output_value = output_matrix[local_row][translation_col_index_seed]
                        except (IndexError, TypeError):
                            existing_output_value = ""
                        if isinstance(existing_output_value, str):
                            normalized_output = _normalize_cell_value(existing_output_value)
                            if isinstance(normalized_output, str):
                                existing_output_raw = normalized_output
                    existing_output_normalized = existing_output_raw.strip() if isinstance(existing_output_raw, str) else ""

                    if existing_output_normalized and not _contains_japanese(existing_output_normalized):
                        reused_translation_detected = True
                        cell_ref_for_metrics = _cell_reference(
                            output_start_row,
                            output_start_col,
                            local_row,
                            translation_col_index_seed,
                        )

                        if cached_entry_for_source and isinstance(cached_entry_for_source.get("translation"), str):
                            translation_value_existing = (
                                existing_output_value
                                if isinstance(existing_output_value, str)
                                else existing_output_normalized
                            )
                            if not isinstance(translation_value_existing, str):
                                translation_value_existing = str(translation_value_existing)

                            source_length_units = _measure_utf16_length(canonical_source)
                            translated_length_units = _measure_utf16_length(translation_value_existing)
                            ratio_value = _compute_length_ratio(source_length_units, translated_length_units)
                            violation_kind = _ratio_violation_kind(ratio_value) if enforce_length_limit else None

                            metric_entry_existing = {
                                "source_length": source_length_units,
                                "translated_length": translated_length_units,
                                "ratio": ratio_value,
                                "cell_ref": cell_ref_for_metrics,
                                "limit": effective_length_ratio_limit,
                                "min_limit": effective_length_ratio_min,
                                "status": violation_kind or "ok",
                                "reported_source_length": None,
                                "reported_translated_length": None,
                                "reported_ratio": None,
                            }
                            length_metrics[(local_row, col_idx)] = metric_entry_existing

                            if enforce_length_limit and violation_kind and (local_row, col_idx) not in length_violation_positions:
                                direction_label = "上限" if violation_kind == "above" else "下限"
                                ratio_text = _format_ratio(ratio_value)
                                length_limit_violations.append(f"{cell_ref_for_metrics}: ×{ratio_text} ({direction_label}逸脱)")
                                length_violation_positions.add((local_row, col_idx))
                                actions.log_progress(
                                    f"既存の翻訳 {cell_ref_for_metrics} が文字数倍率制約を満たしていません: ×{ratio_text} ({direction_label}逸脱)。"
                                )

                            pending_cols = pending_columns_by_row.get(local_row)
                            if pending_cols is not None:
                                pending_cols.discard(col_idx)
                                if not pending_cols:
                                    _finalize_row(local_row)
                            continue

                        if not include_context_columns:
                            actions.log_progress(
                                f"{cell_ref_for_metrics} に既存の訳文が見つかりましたが、新しい訳文を取得します。"
                            )
                        # fall through to request a fresh translation

                    chunk_texts.append(canonical_source)
                    chunk_positions.append((local_row, col_idx))

            if not chunk_texts:
                continue

            chunk_cell_evidences: Dict[Tuple[int, int], Dict[str, Any]] = {}
            row_evidence_details: Dict[int, List[Dict[str, Any]]] = {}

            if include_context_columns:
                chunk_entries: List[Dict[str, Any]] = [
                    {"text": text, "positions": [position]}
                    for text, position in zip(chunk_texts, chunk_positions)
                ]
            else:
                chunk_entries = []
                dedup_index: Dict[str, int] = {}
                for text, position in zip(chunk_texts, chunk_positions):
                    cached_entry = translation_cache.get(text)
                    if cached_entry:
                        row_idx_cached, col_idx_cached = position
                        _apply_cached_translation(row_idx_cached, col_idx_cached, text, cached_entry)
                        continue
                    entry_idx = dedup_index.get(text)
                    if entry_idx is None:
                        dedup_index[text] = len(chunk_entries)
                        chunk_entries.append({"text": text, "positions": [position]})
                    else:
                        chunk_entries[entry_idx]["positions"].append(position)

            if not chunk_entries:
                continue

            budget_adjustment_notified = False
            entry_index = 0
            total_entries = len(chunk_entries)
            while entry_index < total_entries:
                _ensure_not_stopped()
                available_budget = _TRANSLATION_UTF16_BUDGET
                budget_trim_applied = False
                current_texts: List[str] = []
                current_position_groups: List[List[Tuple[int, int]]] = []

                while entry_index < total_entries and len(current_texts) < items_per_request:
                    entry = chunk_entries[entry_index]
                    text = entry["text"]
                    text_units = _measure_utf16_length(text)
                    if (
                        available_budget is not None
                        and text_units > available_budget
                        and current_texts
                    ):
                        budget_trim_applied = True
                        break

                    current_texts.append(text)
                    current_position_groups.append(entry["positions"])
                    entry_index += 1

                    if available_budget is not None:
                        available_budget = max(0, available_budget - text_units)
                        if available_budget == 0:
                            if entry_index < total_entries:
                                budget_trim_applied = True
                            break

                if not current_texts and entry_index < total_entries:
                    entry = chunk_entries[entry_index]
                    current_texts = [entry["text"]]
                    current_position_groups = [entry["positions"]]
                    entry_index += 1

                if not current_texts:
                    break

                if (
                    budget_trim_applied
                    and not budget_adjustment_notified
                    and _TRANSLATION_UTF16_BUDGET is not None
                ):
                    actions.log_progress(
                        f"翻訳バッチをUTF-16長さ {_TRANSLATION_UTF16_BUDGET} 以内に調整しました (現在 {len(current_texts)} 件)。"
                    )
                    budget_adjustment_notified = True


                source_references_per_item = _extract_source_sentences_batch(current_texts)
                reference_pairs_context = _pair_target_sentences_batch(
                    source_references_per_item,
                    current_texts,
                )

                source_lengths: Optional[List[int]] = None
                source_lengths_json: Optional[str] = None
                if enforce_length_limit:
                    source_lengths = [_measure_utf16_length(text) for text in current_texts]
                    source_lengths_json = json.dumps(source_lengths, ensure_ascii=False)

                texts_json = json.dumps(current_texts, ensure_ascii=False)

                translation_context = []
                for index, _ in enumerate(current_texts):
                    translation_context.append(
                        {
                            "source_sentences": source_references_per_item[index] if index < len(source_references_per_item) else [],
                            "reference_pairs": reference_pairs_context[index] if index < len(reference_pairs_context) else [],
                        }
                    )
                translation_context_json = json.dumps(translation_context, ensure_ascii=False)
                has_supporting_data = (
                    include_context_columns
                    or use_references
                )

                def _run_translation_request(
                    extra_notice: Optional[str],
                    include_supporting: bool,
                ) -> Tuple[Optional[List[Any]], Any, Optional[str], Optional[str]]:
                    prompt_sections: List[str] = [prompt_preamble]
                    if extra_notice:
                        prompt_sections.append(extra_notice)
                    prompt_sections.extend([
                        "Source sentences:",
                        texts_json,
                    ])
                    if source_lengths_json is not None:
                        prompt_sections.extend([
                            "Source UTF-16 lengths (JSON):",
                            source_lengths_json,
                        ])
                    if include_supporting and has_supporting_data:
                        prompt_sections.extend([
                            "Supporting data (JSON):",
                            translation_context_json,
                        ])
                    prompt_text = "\n".join(prompt_sections) + "\n"
                    _ensure_not_stopped()
                    response_local = browser_manager.ask(prompt_text, stop_event=stop_event)
                    repaired_response_local, repair_status_local = _repair_unescaped_length_verification_result(response_local)
                    parsed_payload_local, parse_error_code_local = _extract_first_json_payload(repaired_response_local)
                    if parsed_payload_local is None:
                        fallback_payload_local, fallback_error_code_local = _extract_first_json_payload(response_local)
                        parse_error_code_local = parse_error_code_local or fallback_error_code_local
                        parsed_payload_local = fallback_payload_local
                    return parsed_payload_local, response_local, parse_error_code_local, repair_status_local

                def _obtain_translation_payload(extra_notice: Optional[str]) -> List[Any]:
                    parsed_payload_local, response_local, parse_error_code_local, repair_status_local = _run_translation_request(
                        extra_notice,
                        include_supporting=True,
                    )
                    if parsed_payload_local is None and has_supporting_data:
                        parsed_payload_local, response_local, parse_error_code_local, repair_status_local = _run_translation_request(
                            extra_notice,
                            include_supporting=False,
                        )
                    if parsed_payload_local is None:
                        error_label_local = parse_error_code_local or repair_status_local or "unknown"
                        raise ToolExecutionError(
                            f"Failed to parse translation response as JSON ({error_label_local}): {response_local}"
                        )
                    if not isinstance(parsed_payload_local, list) or len(parsed_payload_local) != len(current_texts):
                        raise ToolExecutionError(
                            "Translation response must be a list with one entry per source text."
                        )
                    return parsed_payload_local

                def _find_empty_translation_indexes(payload: List[Any]) -> List[int]:
                    empty_indexes: List[int] = []
                    for idx, payload_item in enumerate(payload):
                        candidate: Optional[Any]
                        if isinstance(payload_item, dict):
                            candidate = (
                                payload_item.get("translated_text")
                                or payload_item.get("translation")
                                or payload_item.get("output")
                            )
                        elif isinstance(payload_item, str):
                            candidate = payload_item
                        elif isinstance(payload_item, (int, float)):
                            candidate = str(payload_item)
                        else:
                            candidate = None
                        if candidate is None:
                            empty_indexes.append(idx)
                            continue
                        candidate_raw = _maybe_unescape_html_entities(candidate)
                        if not isinstance(candidate_raw, str):
                            candidate_raw = str(candidate_raw)
                        if not candidate_raw.strip():
                            empty_indexes.append(idx)
                    return empty_indexes


                length_retry_count = 0
                extra_ratio_notice: Optional[str] = None
                max_length_retries = 3
                best_violation_updates: Dict[Tuple[int, int], Dict[str, Any]] = {}
                best_violation_distances: Dict[Tuple[int, int], float] = {}

                def _violation_distance(ratio_value: float, violation_kind: Optional[str]) -> float:
                    if violation_kind == "above":
                        bound = effective_length_ratio_limit if effective_length_ratio_limit is not None else ratio_value
                        return max(0.0, ratio_value - bound)
                    if violation_kind == "below":
                        bound = effective_length_ratio_min if effective_length_ratio_min is not None else ratio_value
                        return max(0.0, bound - ratio_value)
                    if effective_length_ratio_limit is not None and ratio_value > effective_length_ratio_limit:
                        return ratio_value - effective_length_ratio_limit
                    if effective_length_ratio_min is not None and ratio_value < effective_length_ratio_min:
                        return effective_length_ratio_min - ratio_value
                    return 0.0

                def _record_violation_candidates(updates: List[Dict[str, Any]]) -> None:
                    for candidate in updates:
                        violation_kind = candidate.get("violation_kind")
                        if not violation_kind:
                            continue
                        key = (candidate.get("local_row"), candidate.get("col_idx"))
                        if key[0] is None or key[1] is None:
                            continue
                        distance = _violation_distance(candidate.get("ratio_value", 0.0), violation_kind)
                        previous = best_violation_distances.get(key)
                        if previous is None or distance < previous:
                            best_violation_distances[key] = distance
                            best_violation_updates[key] = copy.deepcopy(candidate)

                while True:
                    parsed_payload = _obtain_translation_payload(extra_ratio_notice)
                    extra_ratio_notice = None

                    empty_retry_count = 0
                    while True:
                        empty_indexes = _find_empty_translation_indexes(parsed_payload)
                        if not empty_indexes:
                            break
                        if empty_retry_count >= max_empty_retries:
                            raise ToolExecutionError("Translation response returned an empty 'translated_text' value.")
                        empty_retry_count += 1
                        actions.log_progress(
                            f"翻訳応答に空の translated_text が含まれているため再リクエストします（再試行 {empty_retry_count}/{max_empty_retries}）。"
                        )
                        retry_notice = (
                            "前回の応答で translated_text が空の項目がありました。すべての translated_text フィールドに非空の翻訳文を記入してください。"
                        )
                        parsed_payload = _obtain_translation_payload(extra_notice=retry_notice)

                    pending_updates: List[Dict[str, Any]] = []
                    metadata_corrections: List[str] = []
                    ratio_violations_local: List[Dict[str, Any]] = []

                    for item_index, (item, position_group) in enumerate(zip(parsed_payload, current_position_groups)):
                        if not position_group:
                            continue
                        local_row, col_idx = position_group[0]
                        translation_value: Optional[str] = None
                        process_notes_jp = ""
                        reference_pairs_output: List[Dict[str, str]] = []
                        if isinstance(item, dict):
                            translation_value = (
                                item.get("translated_text")
                                or item.get("translation")
                                or item.get("output")
                            )
                            if include_context_columns:
                                evidence_dict = item.get("evidence") if isinstance(item.get("evidence"), dict) else None
                                raw_process_notes = (
                                    item.get("process_notes_jp")
                                    or item.get("process_notes")
                                    or item.get("explanation_jp")
                                    or item.get("explanation")
                                )
                                if raw_process_notes is None and evidence_dict:
                                    raw_process_notes = (
                                        evidence_dict.get("process_notes_jp")
                                        or evidence_dict.get("process_notes")
                                        or evidence_dict.get("explanation_jp")
                                        or evidence_dict.get("explanation")
                                    )
                                if isinstance(raw_process_notes, (str, int, float)):
                                    process_notes_source_value = str(raw_process_notes)
                                    sanitized_process_notes = _sanitize_evidence_value(process_notes_source_value)
                                    if sanitized_process_notes:
                                        process_notes_jp = sanitized_process_notes
                                    else:
                                        process_notes_jp = process_notes_source_value.strip()

                                raw_pairs_candidate = item.get("reference_pairs") or item.get("pairs")
                                if raw_pairs_candidate is None and evidence_dict:
                                    raw_pairs_candidate = (
                                        evidence_dict.get("reference_pairs")
                                        or evidence_dict.get("pairs")
                                    )
                                if isinstance(raw_pairs_candidate, list):
                                    cleaned_pairs: List[Dict[str, str]] = []
                                    for pair in raw_pairs_candidate:
                                        if not isinstance(pair, dict):
                                            continue
                                        source_sentence = pair.get("source_sentence") or pair.get("jp") or ""
                                        target_sentence = pair.get("target_sentence") or pair.get("translated") or pair.get("en") or ""
                                        if not isinstance(source_sentence, str) or not isinstance(target_sentence, str):
                                            continue
                                        source_clean = source_sentence.strip()
                                        target_clean = _strip_reference_urls_from_quote(target_sentence.strip())
                                        if not source_clean or not target_clean:
                                            continue
                                        cleaned_pairs.append({
                                            "source_sentence": source_clean,
                                            "target_sentence": target_clean,
                                        })
                                    reference_pairs_output = cleaned_pairs
                        elif isinstance(item, str):
                            translation_value = item
                        elif isinstance(item, (int, float)):
                            translation_value = str(item)

                        if not isinstance(translation_value, str):
                            raise ToolExecutionError(
                                "Translation response must include a 'translated_text' string for each item."
                            )

                        translation_value_raw = _maybe_unescape_html_entities(translation_value)
                        if not isinstance(translation_value_raw, str):
                            translation_value_raw = str(translation_value_raw)
                        translation_value_stripped = translation_value_raw.strip()
                        if not translation_value_stripped:
                            raise ToolExecutionError("Translation response returned an empty 'translated_text' value.")
                        translation_value = translation_value_raw

                        translation_col_index_seed = col_idx if writing_to_source_directly else col_idx * translation_block_width
                        cell_ref_for_metrics = _cell_reference(
                            output_start_row,
                            output_start_col,
                            local_row,
                            translation_col_index_seed,
                        )
                        source_text_canonical = current_texts[item_index]
                        source_cell_value = source_text_canonical
                        source_length_units = _measure_utf16_length(source_cell_value)
                        translated_length_units = _measure_utf16_length(translation_value)
                        ratio_value = _compute_length_ratio(source_length_units, translated_length_units)

                        source_value_for_comparison = source_cell_value.strip()
                        if translation_value_stripped == source_value_for_comparison and _contains_japanese(translation_value_stripped):
                            raise ToolExecutionError(
                                "翻訳結果が原文テキストと同一で日本語のままです。翻訳が完了していません。"
                            )
                        if target_language and target_language.lower().startswith("english") and _contains_japanese(translation_value_stripped):
                            raise ToolExecutionError(
                                "翻訳結果に日本語が含まれているため、英語への翻訳が完了していません。"
                            )

                        violation_kind: Optional[str] = _ratio_violation_kind(ratio_value) if enforce_length_limit else None
                        reported_source_length, reported_translated_length, reported_ratio = (
                            _extract_reported_length_metadata(item) if isinstance(item, Mapping) else (None, None, None)
                        )
                        metadata_fragments: List[str] = []
                        if reported_source_length is not None and reported_source_length != source_length_units:
                            metadata_fragments.append(f"source_length {reported_source_length} -> {source_length_units}")
                        if reported_translated_length is not None and reported_translated_length != translated_length_units:
                            metadata_fragments.append(f"translated_length {reported_translated_length} -> {translated_length_units}")
                        if (
                            reported_ratio is not None
                            and math.isfinite(reported_ratio)
                            and abs(reported_ratio - ratio_value) > _METADATA_RATIO_TOLERANCE
                        ):
                            metadata_fragments.append(
                                f"length_ratio {_format_ratio(reported_ratio)} -> {_format_ratio(ratio_value)}"
                            )
                        if metadata_fragments:
                            metadata_corrections.append(
                                f"{cell_ref_for_metrics}: length metadata auto-corrected ({', '.join(metadata_fragments)})"
                            )

                        process_notes_text = (
                            _maybe_unescape_html_entities(process_notes_jp).strip()
                            if include_context_columns
                            else ""
                        )
                        context_pairs: List[Dict[str, str]] = []
                        if include_context_columns and reference_pairs_context and item_index < len(reference_pairs_context):
                            context_pairs = [
                                pair for pair in reference_pairs_context[item_index]
                                if isinstance(pair, dict)
                            ]
                        reference_pairs_list: List[Dict[str, str]] = []
                        if include_context_columns:
                            reference_pairs_list = list(context_pairs)
                            if reference_pairs_output:
                                merged: List[Dict[str, str]] = []
                                seen_keys: Set[Tuple[str, str]] = set()
                                for candidate in reference_pairs_output:
                                    if not isinstance(candidate, dict):
                                        continue
                                    src = candidate.get("source_sentence")
                                    tgt = candidate.get("target_sentence") or candidate.get("translated") or candidate.get("en")
                                    if not isinstance(src, str) or not isinstance(tgt, str):
                                        continue
                                    key = (src.strip(), tgt.strip())
                                    if key in seen_keys:
                                        continue
                                    seen_keys.add(key)
                                    merged.append(
                                        {
                                            "source_sentence": src.strip(),
                                            "target_sentence": tgt.strip(),
                                        }
                                    )
                                for ctx_pair in context_pairs:
                                    src = ctx_pair.get("source_sentence") if isinstance(ctx_pair, dict) else None
                                    tgt = ctx_pair.get("target_sentence") if isinstance(ctx_pair, dict) else None
                                    if not isinstance(src, str) or not isinstance(tgt, str):
                                        continue
                                    key = (src.strip(), tgt.strip())
                                    if key in seen_keys:
                                        continue
                                    seen_keys.add(key)
                                    merged.append(
                                        {
                                            "source_sentence": src.strip(),
                                            "target_sentence": tgt.strip(),
                                        }
                                    )
                                reference_pairs_list = merged

                        sanitized_pairs: List[Dict[str, str]] = []
                        if include_context_columns:
                            seen_pair_keys: Set[Tuple[str, str]] = set()
                            for pair in reference_pairs_list or []:
                                if not isinstance(pair, dict):
                                    continue
                                source_sentence = pair.get("source_sentence")
                                target_sentence = pair.get("target_sentence")
                                if not isinstance(source_sentence, str) or not isinstance(target_sentence, str):
                                    continue
                                source_clean = source_sentence.strip()
                                target_clean = target_sentence.strip()
                                if not source_clean or not target_clean:
                                    continue
                                key = (source_clean, target_clean)
                                if key in seen_pair_keys:
                                    continue
                                seen_pair_keys.add(key)
                                sanitized_pairs.append({
                                    "source_sentence": source_clean,
                                    "target_sentence": target_clean,
                                })

                        if enforce_length_limit and violation_kind:
                            ratio_violations_local.append(
                                {
                                    "cell_ref": cell_ref_for_metrics,
                                    "ratio": ratio_value,
                                    "kind": violation_kind,
                                    "source_units": source_length_units,
                                    "translated_units": translated_length_units,
                                }
                            )

                        pending_updates.append(
                            {
                                "local_row": local_row,
                                "col_idx": col_idx,
                                "position_group": position_group,
                                "translation_value": translation_value,
                                "translation_value_stripped": translation_value_stripped,
                                "process_notes_text": process_notes_text,
                                "reference_pairs_list": reference_pairs_list,
                                "sanitized_pairs": sanitized_pairs,
                                "source_text_canonical": source_text_canonical,
                                "source_length_units": source_length_units,
                                "translated_length_units": translated_length_units,
                                "ratio_value": ratio_value,
                                "violation_kind": violation_kind,
                                "cell_ref": cell_ref_for_metrics,
                                "reported_source_length": reported_source_length,
                                "reported_translated_length": reported_translated_length,
                                "reported_ratio": reported_ratio,
                                "item_index": item_index,
                                "total_items": len(current_texts),
                                "expected_pair_count": len(source_references_per_item[item_index]) if include_context_columns and item_index < len(source_references_per_item) else 0,
                            }
                        )

                    if enforce_length_limit and ratio_violations_local:
                        _record_violation_candidates(pending_updates)
                        if length_retry_count >= max_length_retries:
                            if best_violation_updates:
                                fallback_updates: List[Dict[str, Any]] = []
                                fallback_seen_keys: Set[Tuple[int, int]] = set()

                                for update in pending_updates:
                                    key = (update.get("local_row"), update.get("col_idx"))
                                    if update.get("violation_kind") and key in best_violation_updates:
                                        best_update = copy.deepcopy(best_violation_updates[key])
                                        fallback_updates.append(best_update)
                                        fallback_seen_keys.add(key)
                                    else:
                                        fallback_updates.append(update)
                                        if key[0] is not None and key[1] is not None:
                                            fallback_seen_keys.add(key)

                                for key, best_update in best_violation_updates.items():
                                    if key not in fallback_seen_keys:
                                        fallback_updates.append(copy.deepcopy(best_update))
                                        fallback_seen_keys.add(key)

                                pending_updates = fallback_updates
                                ratio_violations_local = []
                                actions.log_progress(
                                    "翻訳応答が長さ制約に収まりませんでしたが、最も近い訳文を採用します。"
                                )
                            else:
                                violation_messages = "; ".join(
                                    f"{entry['cell_ref']}: {_format_ratio(entry['ratio'])}" for entry in ratio_violations_local[:5]
                                )
                                raise ToolExecutionError(
                                    f"Length ratio constraint violation persisted after {max_length_retries} retries: {violation_messages}"
                                )
                        if ratio_violations_local:
                            length_retry_count += 1
                            notice_lines = [
                                "前回の応答で以下の項目が文字数制約を満たしていません。",
                            ]
                            for entry in ratio_violations_local:
                                direction = "上限" if entry["kind"] == "above" else "下限"
                                notice_lines.append(
                                    f"- {entry['cell_ref']}: 実測倍率 {_format_ratio(entry['ratio'])} が{direction}を外れています "
                                    f"(原文 {entry['source_units']}、訳文 {entry['translated_units']})。"
                                )
                            if ratio_bounds_display:
                                notice_lines.append(f"文字数倍率の目標レンジ: {ratio_bounds_display}。")
                            notice_lines.extend([
                                "各訳文を調整し、translated_length と length_ratio を制約内に収めてください。",
                                "調整ヒント:",
                                "- 見出しや列挙では冠詞や 'and' を省き、短い名詞・同義語に置き換えてください。",
                                "- 文として訳す場合は冗長な節や重複語を削除し、簡潔な語順へ再構成してください。",
                                "- 長い複合語は意味を保ったまま 1 語の一般的な語に言い換えるか、不要な修飾語を省いてください。",
                                "- 下限を下回る場合は意味を変えずに必要な情報を自然な文で補完してください。",
                                "- translated_length は UTF-16LE で再計算し、length_ratio は translated_length / source_length を小数第2位まで四捨五入した値に揃えてください。",
                                "- length_verification.translated_length_computed と length_verification.length_ratio_computed も同じ結果を記載し、他の値と一致しない場合は JSON を提出しないでください。",
                                "- length_verification.status は一致時のみ \"ok\" にし、不一致なら \"mismatch\" と記録して再翻訳してください。",
                                "- 1 語訳が長すぎる場合は Hit/Toll/Drag など別の短い語に置き換えて必ずレンジに収めてください。",
                                "再回答前に全項目の translated_length / source_length を再計算し、指定レンジ内であることを確認してから JSON を返してください。",
                            ])
                            extra_ratio_notice = "\n".join(notice_lines)
                            actions.log_progress(
                                f"翻訳応答が長さ制約を満たしていません（再リクエスト {length_retry_count}/{max_length_retries}）。"
                            )
                            continue

                    for message in metadata_corrections:
                        actions.log_progress(message)

                    for update in pending_updates:
                        local_row = update["local_row"]
                        col_idx = update["col_idx"]
                        position_group = update["position_group"]
                        translation_value = update["translation_value"]
                        translation_value_stripped = update["translation_value_stripped"]
                        translation_col_index_seed = col_idx if writing_to_source_directly else col_idx * translation_block_width
                        if writing_to_source_directly:
                            translation_col_index = translation_col_index_seed
                            explanation_col_index = None
                            pair_start_index = None
                            pair_end_index = None
                        else:
                            translation_col_index = translation_col_index_seed
                            if include_context_columns:
                                explanation_col_index = translation_col_index + 1
                                pair_start_index = translation_col_index + 2
                                pair_end_index = translation_col_index + translation_block_width - 1
                            else:
                                explanation_col_index = None
                                pair_start_index = None
                                pair_end_index = None

                        existing_output_value = output_matrix[local_row][translation_col_index]
                        if translation_value != existing_output_value:
                            output_matrix[local_row][translation_col_index] = translation_value
                            output_dirty = True
                            row_dirty_flags[local_row] = True
                        if include_context_columns:
                            preview = translation_value_stripped or translation_value
                            if len(preview) > 120:
                                preview = preview[:117] + "..."
                            actions.log_progress(
                                f"翻訳結果プレビュー ({update['item_index'] + 1}/{update['total_items']}): {preview}"
                            )
                        if not writing_to_source_directly and overwrite_source:
                            existing_source_value = source_matrix[local_row][col_idx]
                            if translation_value != existing_source_value:
                                source_matrix[local_row][col_idx] = translation_value
                                source_dirty = True
                                source_row_dirty_flags[local_row] = True

                        if include_context_columns:
                            sanitized_pairs = update["sanitized_pairs"]
                            expected_pairs = update["expected_pair_count"]
                            actions.log_progress(
                                f"参照ペア整理結果 ({update['item_index'] + 1}/{update['total_items']}): {len(sanitized_pairs)} 件 / source_sentences={expected_pairs}"
                            )
                            if sanitized_pairs:
                                for idx_pair, pair in enumerate(sanitized_pairs, start=1):
                                    actions.log_progress(
                                        f"参照ペア[{update['item_index'] + 1}-{idx_pair}]: {pair['source_sentence']} -> {pair['target_sentence']}"
                                    )
                            else:
                                actions.log_progress(
                                    f"参照ペア整理結果 ({update['item_index'] + 1}/{update['total_items']}): 0 件 (参照資料に一致する文が見つかりませんでした)"
                                )

                        if not include_context_columns:
                            translation_cache[update["source_text_canonical"]] = {
                                "translation": translation_value,
                            }
                            cached_entry = translation_cache[update["source_text_canonical"]]
                            for extra_position in position_group[1:]:
                                extra_row, extra_col = extra_position
                                _apply_cached_translation(extra_row, extra_col, update["source_text_canonical"], cached_entry)

                        metric_entry = {
                            "source_length": update["source_length_units"],
                            "translated_length": update["translated_length_units"],
                            "ratio": update["ratio_value"],
                            "cell_ref": update["cell_ref"],
                            "limit": effective_length_ratio_limit,
                            "min_limit": effective_length_ratio_min,
                            "status": update["violation_kind"] or "ok",
                            "reported_source_length": update["reported_source_length"],
                            "reported_translated_length": update["reported_translated_length"],
                            "reported_ratio": update["reported_ratio"],
                        }
                        length_metrics[(local_row, col_idx)] = metric_entry
                        any_translation = True

                        pending_cols = pending_columns_by_row.get(local_row)
                        if pending_cols is not None:
                            pending_cols.discard(col_idx)
                            if not pending_cols:
                                _finalize_row(local_row)

                    if pending_updates:
                        _flush_pending_segments()
                    break
            if use_references and citation_matrix is not None:
                if citation_mode == "paired_columns":
                    for local_row in range(row_start, row_end):
                        for col_offset in range(cite_cols):
                            citation_matrix[local_row][col_offset] = ""
                    for (local_row, col_idx), data in chunk_cell_evidences.items():
                        base_col = col_idx * 2
                        if base_col + 1 >= cite_cols:
                            continue
                        process_notes = (data.get("process_notes") or "").strip()
                        pair_lines = data.get("reference_pair_lines", [])
                        pairs_text = "\n".join(pair_lines)
                        citation_matrix[local_row][base_col] = (
                            process_notes if citation_should_include_explanations else ""
                        )
                        citation_matrix[local_row][base_col + 1] = pairs_text
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
                elif citation_mode == "translation_triplets":
                    for local_row in range(row_start, row_end):
                        for col_idx in range(source_cols):
                            base_col = col_idx * 3
                            if base_col + 2 >= cite_cols:
                                continue
                            citation_matrix[local_row][base_col + 1] = ""
                            citation_matrix[local_row][base_col + 2] = ""
                    for (local_row, col_idx), data in chunk_cell_evidences.items():
                        base_col = col_idx * 3
                        if base_col + 2 >= cite_cols:
                            continue
                        process_notes = (data.get("process_notes") or "").strip()
                        pair_lines = data.get("reference_pair_lines", [])
                        pairs_text = "\n".join(pair_lines)
                        citation_matrix[local_row][base_col + 1] = pairs_text
                        citation_matrix[local_row][base_col + 2] = (
                            process_notes if citation_should_include_explanations else ""
                        )
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
                elif citation_mode == "per_cell":
                    for local_row in range(row_start, row_end):
                        for col_idx in range(cite_cols):
                            citation_matrix[local_row][col_idx] = ""
                    for (local_row, col_idx), data in chunk_cell_evidences.items():
                        process_notes = (data.get("process_notes") or "").strip()
                        pair_lines = data.get("reference_pair_lines", [])
                        combined_lines: List[str] = []
                        if citation_should_include_explanations and process_notes:
                            combined_lines.append(f"説明: {process_notes}")
                        combined_lines.extend(pair_lines or [])
                        citation_matrix[local_row][col_idx] = "\n".join(combined_lines).strip()
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
                elif citation_mode == "single_column":
                    for local_row in range(row_start, row_end):
                        entries = row_evidence_details.get(local_row, [])
                        blocks: List[str] = []
                        for data in entries:
                            process_notes = (data.get("process_notes") or "").strip()
                            pair_lines = data.get("reference_pair_lines", [])
                            lines: List[str] = []
                            if citation_should_include_explanations and process_notes:
                                lines.append(f"説明: {process_notes}")
                            lines.extend(pair_lines or [])
                            block_text = "\n".join(line for line in lines if line).strip()
                            if block_text:
                                blocks.append(block_text)
                        citation_matrix[local_row][0] = "\n".join(blocks).strip()
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


        for row_idx, is_dirty in enumerate(row_dirty_flags):
            if is_dirty and row_idx not in completed_rows:
                _finalize_row(row_idx)

        for remaining_row in list(pending_columns_by_row.keys()):
            _finalize_row(remaining_row)

        _flush_pending_segments()

        output_dirty = any(row_dirty_flags)
        if overwrite_source and not writing_to_source_directly:
            source_dirty = any(source_row_dirty_flags)
        else:
            source_dirty = False

        if not any_translation and not reused_translation_detected:
            return f"No translatable text was found in range '{cell_range}'."

        if include_context_columns and explanation_fallback_notes:
            messages.insert(0, "process_notes_jp が不足していたセルに既定の説明文を補いました: " + " / ".join(explanation_fallback_notes))

        if enforce_length_limit and length_limit_violations:
            bounds_text = _format_ratio_bounds_for_display() or "設定範囲"
            failure_summary = "; ".join(length_limit_violations[:5])
            if len(length_limit_violations) > 5:
                failure_summary += "; ..."
            raise ToolExecutionError(
                f"Length ratio constraint violation in {len(length_limit_violations)} cell(s) (expected {bounds_text}): {failure_summary}"
            )

        write_messages: List[str] = []

        if range_adjustment_note:
            write_messages.append(range_adjustment_note)
        if citation_note:
            write_messages.append(citation_note)
        write_messages.extend(incremental_row_messages)
        if reference_warning_notes:
            write_messages.extend(reference_warning_notes)

        if include_context_columns and length_metrics:
            total_length_entries = len(length_metrics)
            sample_entries = sorted(length_metrics.values(), key=lambda entry: entry.get("cell_ref", ""))[:5]
            preview = [
                f"{entry['cell_ref']}: ×{_format_ratio(entry['ratio'])}"
                for entry in sample_entries
                if entry.get('cell_ref')
            ]
            if preview:
                write_messages.append("翻訳文字数メトリクス: " + "; ".join(preview))

        _ensure_not_stopped()

        if output_dirty:
            translation_message = actions.write_range(output_range, output_matrix, output_sheet, apply_formatting=should_apply_formatting)
            write_messages.append(translation_message)

        if overwrite_source and not writing_to_source_directly and source_dirty:
            overwrite_message = actions.write_range(normalized_range, source_matrix, target_sheet, apply_formatting=should_apply_formatting)
            write_messages.append(overwrite_message)

        if not write_messages:
            write_messages.append("翻訳結果は既存のセル内容と一致していたため、ブックへの書き込みは不要でした。")

        messages = write_messages + messages

        return "\n".join(messages)

    except UserStopRequested:
        raise
    except ToolExecutionError:
        raise
    except Exception as exc:
        raise ToolExecutionError(f"範囲 '{cell_range}' の更新中にエラーが発生しました: {exc}") from exc

def translate_range_without_references(
    actions: ExcelActions,
    browser_manager: BrowserCopilotManager,
    cell_range: str,
    target_language: str = "English",
    sheet_name: Optional[str] = None,
    translation_output_range: Optional[str] = None,
    overwrite_source: bool = False,
    length_ratio_limit: Optional[float] = None,
    length_ratio_min: Optional[float] = None,
    rows_per_batch: Optional[int] = None,
    stop_event: Optional[Event] = None,
) -> str:
    """Translate ranges without using external reference material.

    Optionally enforces a UTF-16 ベースの文字数倍率制限を適用できます。
    Supports defining both upper (length_ratio_limit) and lower (length_ratio_min) bounds."""
    if rows_per_batch is not None and rows_per_batch < 1:
        raise ToolExecutionError("rows_per_batch must be at least 1 when provided.")

    if rows_per_batch is None:
        rows_per_batch = max(4, _ITEMS_PER_TRANSLATION_REQUEST)

    return translate_range_contents(
        actions=actions,
        browser_manager=browser_manager,
        cell_range=cell_range,
        target_language=target_language,
        sheet_name=sheet_name,
        citation_output_range=None,
        reference_urls=None,
        translation_output_range=translation_output_range,
        overwrite_source=overwrite_source,
        length_ratio_limit=length_ratio_limit,
        length_ratio_min=length_ratio_min,
        rows_per_batch=rows_per_batch,
        stop_event=stop_event,
        output_mode="translation_only",
    )


def translate_range_with_references(
    actions: ExcelActions,
    browser_manager: BrowserCopilotManager,
    cell_range: str,
    target_language: str = "English",
    sheet_name: Optional[str] = None,
    source_reference_urls: Optional[List[str]] = None,
    target_reference_urls: Optional[List[str]] = None,
    reference_urls: Optional[List[str]] = None,
    translation_output_range: Optional[str] = None,
    citation_output_range: Optional[str] = None,
    overwrite_source: bool = False,
    stop_event: Optional[Event] = None,
    length_ratio_limit: Optional[float] = None,
    length_ratio_min: Optional[float] = None,
) -> str:
    """Translate ranges while consulting paired reference materials in both languages.

    Args:
        actions: Excel automation helper injected by the agent runtime.
        browser_manager: Shared browser manager used for translation API calls.
        cell_range: Range containing the source Japanese text.
        target_language: Target language for the translation (default \"English\").
        sheet_name: Optional sheet override; defaults to the active sheet.
        source_reference_urls: URLs for original-language reference materials.
        target_reference_urls: URLs for target-language reference materials to pair with the originals.
        reference_urls: Legacy collection of reference URLs (treated as source-language items).
        translation_output_range: Range where translation, process explanation, and reference pairs are written.
        citation_output_range: Optional range for structured citation output.
        overwrite_source: Whether to overwrite the source cells directly.
        stop_event: Optional cancellation event set when the user interrupts execution.
        length_ratio_limit: Optional upper bound for UTF-16 length ratio enforced in translation-only mode.
        length_ratio_min: Optional lower bound for UTF-16 length ratio enforced in translation-only mode.
    """
    normalized_source_refs = source_reference_urls or reference_urls
    if not normalized_source_refs and not target_reference_urls:
        raise ToolExecutionError(
            "translate_range_with_references requires source_reference_urls or target_reference_urls."
        )

    return translate_range_contents(
        actions=actions,
        browser_manager=browser_manager,
        cell_range=cell_range,
        target_language=target_language,
        sheet_name=sheet_name,
        citation_output_range=citation_output_range,
        reference_urls=reference_urls,
        source_reference_urls=source_reference_urls,
        target_reference_urls=target_reference_urls,
        translation_output_range=translation_output_range,
        overwrite_source=overwrite_source,
        length_ratio_limit=length_ratio_limit,
        length_ratio_min=length_ratio_min,
        rows_per_batch=1,
        stop_event=stop_event,
    )

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
    stop_event: Optional[Event] = None,
) -> str:
    """Review translated text against the source and record quality outcomes.

    Args:
        actions: Excel automation helper injected by the agent runtime.
        browser_manager: Shared browser manager used for LLM review prompts.
        source_range: Range containing the Japanese source text.
        translated_range: Range containing the draft translations under review.
        status_output_range: Range where OK/REVISE status markers are written.
        issue_output_range: Range used to write reviewer issue notes.
        corrected_output_range: Optional range where corrected translations are written.
        highlight_output_range: Optional range that receives diff-highlight output.
        sheet_name: Optional sheet override; defaults to the active sheet.
        stop_event: Optional cancellation event set when the user interrupts the operation.
    """
    try:
        def _log_debug(message: str) -> None:
            _review_debug(f"[check_translation_quality] {message}")

        def _ensure_not_stopped() -> None:
            if stop_event and stop_event.is_set():
                raise UserStopRequested("ユーザーによって処理が中断されました。")

        _ensure_not_stopped()

        src_rows, src_cols = _parse_range_dimensions(source_range)
        trans_rows, trans_cols = _parse_range_dimensions(translated_range)
        status_rows, status_cols = _parse_range_dimensions(status_output_range)
        issue_rows, issue_cols = _parse_range_dimensions(issue_output_range)
        highlight_rows = highlight_cols = None
        if highlight_output_range:
            highlight_rows, highlight_cols = _parse_range_dimensions(highlight_output_range)

        correction_note: Optional[str] = None
        if corrected_output_range:
            correction_note = (
                "corrected_output_range was provided but corrections are no longer written; the review now outputs only status, issues, and highlight columns."
            )

        if (src_rows, src_cols) != (trans_rows, trans_cols):
            raise ToolExecutionError("Source range and translated range sizes do not match.")
        if (src_rows, src_cols) != (status_rows, status_cols) or (src_rows, src_cols) != (issue_rows, issue_cols):
            raise ToolExecutionError("Output ranges must match the source range size.")
        if highlight_output_range and (src_rows, src_cols) != (highlight_rows, highlight_cols):
            raise ToolExecutionError("Highlight output range must match the source range size.")

        status_sheet_name, status_inner_range = _split_sheet_and_range(status_output_range, sheet_name)
        status_start_row, status_start_col, _, _ = _parse_range_bounds(status_inner_range)
        issue_sheet_name, issue_inner_range = _split_sheet_and_range(issue_output_range, sheet_name)
        issue_start_row, issue_start_col, _, _ = _parse_range_bounds(issue_inner_range)
        highlight_sheet_name: Optional[str] = None
        highlight_start_row = highlight_start_col = None
        if highlight_output_range:
            highlight_sheet_name, highlight_inner_range = _split_sheet_and_range(highlight_output_range, sheet_name)
            highlight_start_row, highlight_start_col, _, _ = _parse_range_bounds(highlight_inner_range)

        source_data = _reshape_to_dimensions(actions.read_range(source_range, sheet_name), src_rows, src_cols)
        translated_data = _reshape_to_dimensions(actions.read_range(translated_range, sheet_name), src_rows, src_cols)
        source_data = _unescape_matrix_values(source_data)
        translated_data = _unescape_matrix_values(translated_data)

        _ensure_not_stopped()

        status_matrix = [["" for _ in range(src_cols)] for _ in range(src_rows)]
        issue_matrix = [["" for _ in range(src_cols)] for _ in range(src_rows)]
        highlight_matrix = [] if highlight_output_range else None
        highlight_styles = [] if highlight_output_range else None

        supports_rich_diff_colors = getattr(actions, "supports_diff_highlight_colors", lambda: True)()

        if highlight_matrix is not None:
            for r in range(src_rows):
                _ensure_not_stopped()
                highlight_row = []
                styles_row = [] if highlight_styles is not None else None
                for c in range(src_cols):
                    _ensure_not_stopped()
                    base_value = _normalize_cell_value(translated_data[r][c])
                    highlight_row.append(base_value)
                    if styles_row is not None:
                        styles_row.append([])
                highlight_matrix.append(highlight_row)
                if styles_row is not None:
                    highlight_styles.append(styles_row)

        def _cell_reference(base_row: int, base_col: int, local_row: int, local_col: int) -> str:
            return _build_range_reference(
                base_row + local_row,
                base_row + local_row,
                base_col + local_col,
                base_col + local_col,
            )

        def _row_reference(base_row: int, base_col: int, row_idx: int, width: int) -> str:
            start_col = base_col
            end_col = base_col + width - 1
            return _build_range_reference(
                base_row + row_idx,
                base_row + row_idx,
                start_col,
                end_col,
            )

        incremental_updates = False

        def _write_single_entry(row_idx: int, col_idx: int) -> None:
            nonlocal incremental_updates
            _ensure_not_stopped()
            incremental_updates = True
            row_width = src_cols
            status_row_ref = _row_reference(status_start_row, status_start_col, row_idx, row_width)
            _log_debug(f"write_range status -> {status_row_ref}")
            actions.write_range(status_row_ref, [status_matrix[row_idx]], status_sheet_name)
            issue_row_ref = _row_reference(issue_start_row, issue_start_col, row_idx, row_width)
            _log_debug(f"write_range issues -> {issue_row_ref}")
            actions.write_range(issue_row_ref, [issue_matrix[row_idx]], issue_sheet_name)
            if highlight_matrix is not None and highlight_start_row is not None and highlight_start_col is not None:
                highlight_row_ref = _row_reference(highlight_start_row, highlight_start_col, row_idx, row_width)
                _log_debug(f"write_range highlight -> {highlight_row_ref}")
                actions.write_range(highlight_row_ref, [highlight_matrix[row_idx]], highlight_sheet_name)
                if highlight_styles is not None:
                    try:
                        _log_debug(f"apply_diff_highlight_colors row={row_idx} ref={highlight_row_ref}")
                        actions.apply_diff_highlight_colors(
                            highlight_row_ref,
                            [highlight_styles[row_idx]],
                            highlight_sheet_name,
                            addition_color_hex="#1565C0",
                            deletion_color_hex="#C62828",
                        )
                    except ToolExecutionError as color_err:
                        error_message = (
                            f"Diff coloring failed for row {row_idx + 1} ({highlight_row_ref}): {color_err}"
                        )
                        _log_debug(error_message)
                        actions.log_progress(error_message)
                    except Exception as unexpected_color_err:
                        error_message = (
                            f"Diff coloring raised unexpected error for row {row_idx + 1} ({highlight_row_ref}): {unexpected_color_err}"
                        )
                        _log_debug(error_message)
                        actions.log_progress(error_message)
            if col_idx == src_cols - 1:
                row_number = status_start_row + row_idx + 1
                status_summaries: List[str] = []
                issue_summaries: List[str] = []
                for summary_col in range(src_cols):
                    status_cell = _cell_reference(status_start_row, status_start_col, row_idx, summary_col)
                    issue_cell = _cell_reference(issue_start_row, issue_start_col, row_idx, summary_col)
                    status_value = status_matrix[row_idx][summary_col] or ""
                    issue_value = issue_matrix[row_idx][summary_col] or ""
                    status_summaries.append(f"{status_cell}:{status_value}")
                    if issue_value:
                        issue_summaries.append(f"{issue_cell}:{issue_value}")
                status_summary = ", ".join(status_summaries)
                issue_summary = ", ".join(issue_summaries) if issue_summaries else "(no notes)"
                progress_message = (
                    f"Row {row_number} review complete. Status -> {status_summary}. Issues -> {issue_summary}"
                )
                actions.log_progress(progress_message)

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
            _ensure_not_stopped()
            for c in range(src_cols):
                _ensure_not_stopped()
                original_text = source_data[r][c]
                translated_text = translated_data[r][c]
                normalized_translation = _normalize_cell_value(translated_text)
                normalized_translation = _maybe_fix_mojibake(normalized_translation)
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
                        status_matrix[r][c] = "MISSING"
                        issue_matrix[r][c] = "Source cell contains Japanese text but the translation cell is blank."
                        needs_revision_count += 1
                        if highlight_matrix is not None:
                            highlight_matrix[r][c] = normalized_translation
                        if highlight_styles is not None:
                            highlight_styles[r][c] = []
                        _write_single_entry(r, c)
                else:
                    status_matrix[r][c] = ""
                    issue_matrix[r][c] = ""
                    if highlight_matrix is not None:
                        if highlight_styles is not None:
                            highlight_styles[r][c] = []
                        highlight_matrix[r][c] = normalized_translation
                    _write_single_entry(r, c)

        if not review_entries:
            _ensure_not_stopped()
            actions.write_range(status_output_range, status_matrix, sheet_name)
            actions.write_range(issue_output_range, issue_matrix, sheet_name)
            if highlight_matrix is not None and highlight_output_range:
                actions.write_range(highlight_output_range, highlight_matrix, sheet_name)
                if highlight_styles is not None:
                    actions.apply_diff_highlight_colors(
                        highlight_output_range,
                        highlight_styles,
                        sheet_name,
                        addition_color_hex="#1565C0",
                        deletion_color_hex="#C62828",
                    )
            return "No review entries were generated; nothing to audit."

        ordered_ids = [entry["id"] for entry in review_entries]
        id_to_index: Dict[str, int] = {entry_id: idx for idx, entry_id in enumerate(ordered_ids)}

        for entry in review_entries:
            _ensure_not_stopped()
            batch = [entry]
            payload = json.dumps(batch, ensure_ascii=False)
            preview_sections: List[str] = []
            for preview_index, preview_item in enumerate(batch, start=1):
                original_preview = preview_item.get("original_text") or ""
                translated_preview = preview_item.get("translated_text") or ""
                item_label = f"Item {preview_index}"
                preview_sections.append(f"{item_label} Original (Japanese):")
                preview_sections.append(original_preview)
                preview_sections.append("")
                preview_sections.append(f"{item_label} Translation (English):")
                preview_sections.append(translated_preview)
                preview_sections.append("")
            preview_text = "\n".join(preview_sections).strip()
            preview_block = f"\n\nOriginal / Translation Preview:\n{preview_text}" if preview_text else ""
            _diff_debug(f"check_translation_quality payload={_shorten_debug(payload)}")
            analysis_prompt = (
                "You are reviewing Japanese-to-English translations.\n"
                "Exactly one review item is provided at a time. Focus only on that single item.\n"
                "Do not attempt to operate Excel or any other applications; only analyze the text and respond in JSON.\n"
                "Each review item includes 'id', 'original_text' (Japanese source text), and 'translated_text' (English translation under review).\n"
                "Treat 'original_text' as the authoritative Japanese source and 'translated_text' as the English draft under review.\n"
                "Assess factual accuracy and overall translation quality. Respect intentional localization, tone choices, and stylistic adjustments when they convey the source meaning and align with expected guidelines; only propose changes for clear mistranslations, omissions, or wording that would impede understanding.\n"
                "Respond with a JSON array containing exactly one object. Provide: 'id', 'status', 'notes', 'corrected_text', and 'highlighted_text'.\n"
                "Optionally include 'before_text', 'after_text', or an 'edits' array (each element with fields 'type', 'text', and 'reason').\n"
                "Use status 'OK' when the draft translation already reflects the source intent (including acceptable paraphrasing); otherwise respond with 'REVISE'.\n"
                "Write 'notes' in Japanese using the exact pattern 'Issue: ... / Suggestion: ...'. Keep them concise and actionable.\n"
                "Set 'corrected_text' to the fully corrected English sentence. For status 'OK', repeat the original translation unchanged.\n"
                "Populate 'highlighted_text' to show the difference versus the current translation: wrap deletions as `[DEL]削除テキスト[DEL]` and additions as `[ADD]追加テキスト[ADD]`. Do not use closing tags like [/DEL] or [/ADD]. Leave it empty for status 'OK'.\n"
                "Do not wrap the JSON in code fences or add commentary outside the array.\n"
            )
            def _parse_batch_response(response_text: str) -> Optional[List[Any]]:
                _diff_debug(f"check_translation_quality parse raw={_shorten_debug(response_text)}")
                if not response_text:
                    return None
                stripped = response_text.strip()
                if not stripped:
                    return None

                cleaned = stripped
                if cleaned.startswith("```"):
                    lines = cleaned.splitlines()
                    removed = False
                    while lines and lines[0].strip().startswith("```"):
                        lines.pop(0)
                        removed = True
                    while lines and lines[-1].strip().startswith("```"):
                        lines.pop()
                        removed = True
                    cleaned = "\n".join(lines).strip()
                    if not cleaned:
                        return None
                    if removed:
                        _diff_debug("check_translation_quality stripped markdown code fences before parsing")

                decoder = json.JSONDecoder(strict=False)
                marker = "Review item (JSON):"
                candidate_texts = [cleaned]
                if marker in cleaned:
                    after_marker = cleaned.split(marker, 1)[1].strip()
                    if after_marker:
                        candidate_texts.insert(0, after_marker)
                for candidate in candidate_texts:
                    potential_starts = [idx for idx, ch in enumerate(candidate) if ch in {'[', '{'}]
                    if not potential_starts:
                        continue
                    for start_idx in potential_starts:
                        if candidate[:start_idx].strip():
                            _diff_debug('check_translation_quality leading non-JSON content detected before payload')
                            continue
                        try:
                            parsed, end_idx = decoder.raw_decode(candidate[start_idx:])
                        except json.JSONDecodeError as decode_error:
                            _diff_debug(f"check_translation_quality decode error start={start_idx} err={decode_error}")
                            continue
                        trailing = candidate[start_idx + end_idx:].strip()
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
                (
                    analysis_prompt
                    + preview_block
                    + "\n\nReview item (JSON):\n"
                    + payload
                    + "\n"
                ),
                (
                    analysis_prompt
                    + "\n\nSTRICT OUTPUT REMINDER: Return exactly one JSON array immediately. Do not include Final Answer, Thought, extra commentary, or multiple JSON payloads."
                    + preview_block
                    + "\n\nReview item (JSON):\n"
                    + payload
                    + "\n"
                ),
                (
                    "You are reviewing Japanese-to-English translations. "
                    "Exactly one review item is supplied at a time; focus only on that single item. "
                    "Do not attempt to control Excel or describe UI steps; respond only with JSON. "
                    "Treat 'original_text' as the Japanese source and 'translated_text' as the English translation under review. "
                    "Carefully evaluate both accuracy and expression; flag unnatural tone, awkward phrasing, register mismatches, or inconsistent terminology even when the literal meaning is broadly correct, and propose smoother alternatives. "
                    "Reply with a single JSON array containing exactly one object. Each object must contain 'id', 'status', 'notes', "
                    "'highlighted_text', 'corrected_text', 'before_text', and 'after_text'. "
                    "Use status 'OK' when the translation is acceptable (notes empty or a short remark). Only select 'OK' when you are certain there are no issues. "
                    "Use status 'REVISE' when changes are needed and write notes in Japanese as 'Issue: ... / Suggestion: ...'. If unsure, choose 'REVISE'. "
                    "Set 'corrected_text' to the fully corrected English sentence. Build 'highlighted_text' from corrected_text, "
                    "marking deletions as `[DEL]削除テキスト[DEL]` and additions as `[ADD]追加テキスト[ADD]` (no closing tags such as [/ADD]). "
                    "Return exactly one JSON array and nothing else."
                    + preview_block
                    + f"\n\nReview item (JSON):\n{payload}\n"
                ),
            ]

            response = ""
            batch_results: Optional[List[Any]] = None
            for prompt_variant in prompt_variants:
                _ensure_not_stopped()
                response = browser_manager.ask(prompt_variant, stop_event=stop_event)
                _diff_debug(f"check_translation_quality response={_shorten_debug(response)}")
                if response and any(indicator in response for indicator in REFUSAL_PATTERNS):
                    _diff_debug('check_translation_quality detected refusal response, trying next prompt variant')
                    continue
                batch_results = _parse_batch_response(response)
                if batch_results is not None:
                    break

            if batch_results is None:
                _diff_debug(f"check_translation_quality unable to parse response={_shorten_debug(response)}")
                raise ToolExecutionError(
                    f"Failed to parse translation quality response as JSON: {response}"
                )

            if not isinstance(batch_results, list):
                raise ToolExecutionError(
                    "Translation quality response must be returned as a JSON array."
                )

            if len(batch_results) > 1:
                _diff_debug("check_translation_quality trimming extra responses to the first item")
                batch_results = batch_results[:1]

            ok_statuses = {"OK", "PASS", "GOOD"}
            revise_statuses = {"REVISE", "NG", "FAIL", "ISSUE"}
            assigned_ids: Set[str] = set()
            assigned_indices: Set[int] = set()
            for item in batch_results:
                _ensure_not_stopped()
                if not isinstance(item, dict):
                    raise ToolExecutionError(
                        "Each translation quality entry must be a JSON object."
                    )
                for key, value in list(item.items()):
                    if isinstance(value, str):
                        fixed_value = _maybe_fix_mojibake(value)
                        if fixed_value != value:
                            item[key] = fixed_value


                raw_item_id = item.get("id")
                candidate_id = str(raw_item_id).strip() if raw_item_id is not None else ""
                resolved_id: Optional[str] = None
                candidate_index: Optional[int] = None
                if candidate_id and candidate_id in id_to_position and candidate_id not in assigned_ids:
                    resolved_id = candidate_id
                    candidate_index = id_to_index.get(candidate_id)
                else:
                    if candidate_id and candidate_id.isdigit():
                        numeric_index = int(candidate_id) - 1
                        if 0 <= numeric_index < len(ordered_ids) and numeric_index not in assigned_indices:
                            candidate_index = numeric_index
                    if candidate_index is None:
                        for idx, entry_id in enumerate(ordered_ids):
                            if idx in assigned_indices:
                                continue
                            candidate_index = idx
                            resolved_id = entry_id
                            break
                    elif resolved_id is None and candidate_index is not None:
                        resolved_id = ordered_ids[candidate_index]
                    if resolved_id and candidate_id and candidate_id != resolved_id:
                        _diff_debug(f"check_translation_quality applying fallback id={candidate_id} -> {resolved_id}")

                if resolved_id is None:
                    _diff_debug(f"check_translation_quality skipping entry with unknown id={candidate_id}")
                    continue
                assigned_ids.add(resolved_id)
                if candidate_index is None:
                    candidate_index = id_to_index.get(resolved_id)
                if candidate_index is not None:
                    assigned_indices.add(candidate_index)
                item_id = resolved_id

                status_raw = item.get("status", "")
                if isinstance(status_raw, str):
                    status_raw = _maybe_fix_mojibake(status_raw)
                    item["status"] = status_raw
                status_value = str(status_raw).strip().upper()

                notes_raw = item.get("notes", "")
                if isinstance(notes_raw, str):
                    notes_raw = _maybe_fix_mojibake(notes_raw)
                    item["notes"] = notes_raw
                notes_value = str(notes_raw).strip()

                before_text = item.get("before_text")
                if isinstance(before_text, str):
                    before_text = _maybe_fix_mojibake(before_text)
                    item["before_text"] = before_text

                after_text = item.get("after_text")
                if isinstance(after_text, str):
                    after_text = _maybe_fix_mojibake(after_text)
                    item["after_text"] = after_text


                row_idx, col_idx = id_to_position[item_id]
                base_translation = translated_data[row_idx][col_idx]
                base_text = _normalize_cell_value(base_translation)
                sanitized_base_text = _maybe_fix_mojibake(base_text)
                corrected_text = _infer_corrected_text(base_text, item)
                corrected_text_str = _normalize_cell_value(corrected_text)
                corrected_text_str = _maybe_fix_mojibake(corrected_text_str)
                is_ok_status = status_value in ok_statuses
                if is_ok_status or not corrected_text_str.strip():
                    corrected_text_str = sanitized_base_text

                if highlight_matrix is not None:
                    if is_ok_status:
                        highlight_matrix[row_idx][col_idx] = ""
                        if highlight_styles is not None:
                            highlight_styles[row_idx][col_idx] = []
                    else:
                        ai_highlight_raw = item.get("highlighted_text") or item.get("highlighted_translation")
                        highlight_text: str
                        highlight_spans: List[Dict[str, int]]
                        highlight_text = ""
                        highlight_spans: List[Dict[str, int]] = []
                        if isinstance(ai_highlight_raw, str) and ("[DEL]" in ai_highlight_raw or "[ADD]" in ai_highlight_raw):
                            parsed_text = _maybe_fix_mojibake(ai_highlight_raw)
                            highlight_text, highlight_spans = _parse_highlight_markup(parsed_text)

                        if not highlight_spans:
                            highlight_text, highlight_spans = _build_diff_highlight(
                                sanitized_base_text,
                                corrected_text_str,
                            )
                        if not highlight_text:
                            highlight_text = corrected_text_str
                        highlight_text, highlight_spans = _attach_highlight_labels(highlight_text, highlight_spans)
                        if highlight_spans and not supports_rich_diff_colors:
                            notify_unavailable = getattr(actions, "notify_diff_colors_unavailable", None)
                            if callable(notify_unavailable):
                                notify_unavailable()
                            highlight_text = _render_textual_diff_markup(highlight_text, highlight_spans)
                            highlight_spans = []
                        # Keep highlight_text unchanged after span generation so offsets stay accurate.
                        _log_debug(f"highlight entry r={row_idx} c={col_idx} text={highlight_text} spans={highlight_spans}")
                        highlight_matrix[row_idx][col_idx] = highlight_text
                        if highlight_styles is not None:
                            highlight_styles[row_idx][col_idx] = highlight_spans

                if is_ok_status:
                    status_matrix[row_idx][col_idx] = "OK"
                    issue_matrix[row_idx][col_idx] = ""
                elif status_value in revise_statuses:
                    status_matrix[row_idx][col_idx] = "REVISE"
                    issue_matrix[row_idx][col_idx] = _format_issue_notes(notes_value)
                    needs_revision_count += 1
                else:
                    status_matrix[row_idx][col_idx] = status_value or "UNKNOWN"
                    issue_matrix[row_idx][col_idx] = _format_issue_notes(notes_value)
                    needs_revision_count += 1

                _write_single_entry(row_idx, col_idx)

        if not incremental_updates:
            _ensure_not_stopped()
            actions.write_range(status_output_range, status_matrix, sheet_name)
            actions.write_range(issue_output_range, issue_matrix, sheet_name)

        processed_items = len(review_entries)
        message = (
            f"Reviewed {processed_items} items; flagged {needs_revision_count} for revision. "
            f"Wrote status to '{status_output_range}' and issues to '{issue_output_range}'."
        )
        if highlight_matrix is not None and highlight_output_range:
            if not incremental_updates:
                _ensure_not_stopped()
                _log_debug(f"write_range highlight bulk -> {highlight_output_range}")
                actions.write_range(highlight_output_range, highlight_matrix, sheet_name)
                if highlight_styles is not None:
                    _ensure_not_stopped()
                    try:
                        _log_debug(f"apply_diff_highlight_colors bulk ref={highlight_output_range}")
                        actions.apply_diff_highlight_colors(
                            highlight_output_range,
                            highlight_styles,
                            sheet_name,
                            addition_color_hex="#1565C0",
                            deletion_color_hex="#C62828",
                        )
                    except ToolExecutionError as color_err:
                        error_message = (
                            f"Diff coloring failed for range {highlight_output_range}: {color_err}"
                        )
                        _log_debug(error_message)
                        actions.log_progress(error_message)
                    except Exception as unexpected_color_err:
                        error_message = (
                            f"Diff coloring raised unexpected error for range {highlight_output_range}: {unexpected_color_err}"
                        )
                        _log_debug(error_message)
                        actions.log_progress(error_message)
            message += f" Highlight output written to '{highlight_output_range}'."
        if correction_note:
            message += f" {correction_note}"
        return message

    except UserStopRequested:
        raise
    except ToolExecutionError:
        raise
    except Exception as e:
        raise ToolExecutionError(f"Translation quality review failed: {e}") from e



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
        original_matrix = _unescape_matrix_values(original_matrix)
        revised_matrix = _unescape_matrix_values(revised_matrix)

        highlight_matrix: List[List[str]] = []
        highlight_styles: List[List[List[Dict[str, int]]]] = []
        supports_rich_diff_colors = getattr(actions, "supports_diff_highlight_colors", lambda: True)()

        for r in range(original_rows):
            text_row: List[str] = []
            style_row: List[List[Dict[str, int]]] = []
            for c in range(original_cols):
                before_text = _normalize_cell_value(original_matrix[r][c])
                after_text = _normalize_cell_value(revised_matrix[r][c])
                _diff_debug(f"highlight_text_differences cell=({r},{c}) before={_shorten_debug(before_text)} after={_shorten_debug(after_text)}")
                highlight_text, spans = _build_diff_highlight(before_text, after_text)
                _diff_debug(f"highlight_text_differences spans= {spans}")
                if spans and not supports_rich_diff_colors:
                    notify_unavailable = getattr(actions, "notify_diff_colors_unavailable", None)
                    if callable(notify_unavailable):
                        notify_unavailable()
                    highlight_text = _render_textual_diff_markup(highlight_text, spans)
                    spans = []
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
            f"Diff highlight written to '{output_range}' using addition color {addition_color_hex} "
            f"and deletion color {deletion_color_hex}."
        )
    except ToolExecutionError:
        raise
    except Exception as exc:
        _diff_debug(f"highlight_text_differences exception={exc}")
        raise ToolExecutionError(f"diff highlight generation failed: {exc}") from exc

def insert_shape(actions: ExcelActions,
                 cell_range: str,
                 shape_type: str,
                 sheet_name: Optional[str] = None,
                 fill_color_hex: Optional[str] = None,
                 line_color_hex: Optional[str] = None) -> str:
    """Insert a drawing shape anchored to the specified range.

    

    Args:

        actions: Excel automation helper injected by the agent runtime.

        cell_range: Anchor range whose top-left corner is used for placement.

        shape_type: Excel shape type name (for example Rectangle).

        sheet_name: Optional sheet override; defaults to the active sheet.

        fill_color_hex: Optional fill colour specified as #RRGGBB.

        line_color_hex: Optional outline colour specified as #RRGGBB.

    """
    return actions.insert_shape_in_range(cell_range, shape_type, sheet_name, fill_color_hex, line_color_hex)

def format_shape(actions: ExcelActions, fill_color_hex: Optional[str] = None, line_color_hex: Optional[str] = None, sheet_name: Optional[str] = None) -> str:
    """Format the most recently inserted shape.

    

    Args:

        actions: Excel automation helper injected by the agent runtime.

        fill_color_hex: Optional fill colour specified as #RRGGBB.

        line_color_hex: Optional outline colour specified as #RRGGBB.

        sheet_name: Optional sheet override; defaults to the active sheet.

    """
    return actions.format_last_shape(fill_color_hex, line_color_hex, sheet_name)
