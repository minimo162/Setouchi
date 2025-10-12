import html
import re
import difflib
import logging
import math
import os
import string
from threading import Event
from typing import List, Any, Optional, Dict, Tuple, Set
from pathlib import Path
from urllib.parse import urlparse, unquote

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
DEFAULT_REFERENCE_PAIR_COLUMNS = 6
_MIN_CONTEXT_BLOCK_WIDTH = 2 + DEFAULT_REFERENCE_PAIR_COLUMNS
_REFERENCE_FALLBACK_ROOTS: List[Path] = []

_REFERENCE_FALLBACK_ROOTS.append(Path.cwd())
downloads_dir = Path.home() / "Downloads"
if downloads_dir.is_dir():
    _REFERENCE_FALLBACK_ROOTS.append(downloads_dir)
extra_dirs = os.getenv("COPILOT_REFERENCE_DIRS", "")
if extra_dirs:
    for token in extra_dirs.split(os.pathsep):
        token = token.strip()
        if not token:
            continue
        candidate = Path(token).expanduser()
        if candidate.is_dir():
            _REFERENCE_FALLBACK_ROOTS.append(candidate)

if _DIFF_DEBUG_ENABLED and not logging.getLogger().handlers:
    logging.basicConfig(level=logging.DEBUG)



try:
    _ITEMS_PER_TRANSLATION_REQUEST = max(
        1, int(os.getenv('EXCEL_COPILOT_TRANSLATION_ITEMS_PER_REQUEST', '1'))
    )
except ValueError:
    _ITEMS_PER_TRANSLATION_REQUEST = 1



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
    try:
        candidate = text.encode('cp932', errors='strict').decode('utf-8', errors='strict')
    except UnicodeError:
        return text
    if candidate == text:
        return text
    original_penalty = _mojibake_penalty(text)
    candidate_penalty = _mojibake_penalty(candidate)
    original_japanese = _count_japanese_characters(text)
    candidate_japanese = _count_japanese_characters(candidate)
    if candidate_penalty > original_penalty and candidate_japanese <= original_japanese:
        return text
    return candidate


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



def _extract_primary_quoted_phrase(text: str) -> Optional[str]:
    match = re.search(r"「([^」]+)」", text)
    if not match:
        return None
    candidate = match.group(1).strip()
    return candidate or None


def _enrich_search_keywords(source_text: str, base_keywords: List[str], max_keywords: int = 12) -> List[str]:
    """Return the AI-supplied keywords with deduplication and opener diversity."""
    del source_text  # retained for signature compatibility

    cleaned_keywords: List[str] = []
    seen_lower: Set[str] = set()
    leading_pairs: Set[str] = set()
    candidates: List[Tuple[str, str]] = []
    fallback_candidates: List[Tuple[str, str]] = []

    for keyword in base_keywords:
        cleaned = (keyword or "").strip()
        if not cleaned:
            continue
        lowered = cleaned.lower()
        if lowered in seen_lower:
            continue
        seen_lower.add(lowered)
        tokens = lowered.split()
        leading_pair = " ".join(tokens[:2]) if tokens else ""
        candidates.append((cleaned, leading_pair))

    for cleaned, leading_pair in candidates:
        if leading_pair and leading_pair in leading_pairs:
            fallback_candidates.append((cleaned, leading_pair))
            continue
        leading_pairs.add(leading_pair)
        cleaned_keywords.append(cleaned)
        if len(cleaned_keywords) >= max_keywords:
            return cleaned_keywords

    for cleaned, leading_pair in fallback_candidates:
        if cleaned in cleaned_keywords:
            continue
        cleaned_keywords.append(cleaned)
        if len(cleaned_keywords) >= max_keywords:
            break

    return cleaned_keywords



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

JAPANESE_CHAR_PATTERN = re.compile(r'[ぁ-んァ-ヶ一-龯]')



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

import json

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

        def _resolve_local_reference_copy(url: str) -> Optional[Path]:
            try:
                parsed = urlparse(url)
            except Exception:
                return None
            filename = unquote(Path(parsed.path).name)
            if not filename:
                return None
            candidates: List[Path] = []
            for root in _REFERENCE_FALLBACK_ROOTS:
                candidates.append(root / filename)
            # allow nested directories encoded in URL
            if "/" in filename:
                path_fragment = Path(filename)
                for root in _REFERENCE_FALLBACK_ROOTS:
                    candidates.append(root / path_fragment)
            for candidate in candidates:
                if candidate.exists():
                    return candidate.resolve()
            return None

        def _resolve_file_reference(value: str) -> Optional[Path]:
            candidate = value.strip()
            if not candidate:
                return None
            path_value = Path(candidate).expanduser()
            search_paths: List[Path] = []
            if path_value.is_absolute():
                search_paths.append(path_value)
            else:
                if workbook_dir:
                    search_paths.append(workbook_dir / path_value)
                search_paths.append(Path.cwd() / path_value)
            for path_candidate in search_paths:
                try:
                    if path_candidate.exists():
                        return path_candidate.resolve()
                except Exception:
                    continue
            return None
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
        translation_block_width = (
            _MIN_CONTEXT_BLOCK_WIDTH if include_context_columns else 1
        )
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
            coerced_notes: List[str] = []
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
                if _is_probable_url(normalized_url):
                    local_copy = _resolve_local_reference_copy(normalized_url)
                    if local_copy is not None:
                        resolved_url = local_copy.as_uri()
                        coerced_notes.append(
                            f"{label} の値 '{original_value}' をローカルファイル '{local_copy.name}' に置き換えて利用します。"
                        )
                    else:
                        resolved_url = normalized_url
                    if normalized_url not in seen_urls:
                        seen_urls.add(resolved_url)
                        entries.append({
                            "id": f"{label[:1].upper()}{len(entries) + 1}",
                            "url": resolved_url,
                        })
                    continue
                resolved_path = _resolve_file_reference(normalized_url)
                if resolved_path:
                    coerced_notes.append(
                        f"{label} の値 '{original_value}' をファイル '{resolved_path.name}' として利用しました。"
                    )
                    resolved_url = resolved_path.as_uri()
                    if resolved_url not in seen_urls:
                        seen_urls.add(resolved_url)
                        entries.append({
                            "id": f"{label[:1].upper()}{len(entries) + 1}",
                            "url": resolved_url,
                        })
                    continue
                invalid_tokens.append(original_value or "(空文字列)")

            warnings: List[str] = []
            if invalid_tokens:
                invalid_urls = ", ".join(_dedupe_preserve_order(invalid_tokens))
                warnings.append(
                    f"{label} で指定された値 ({invalid_urls}) はURLや既存ファイルとして解釈できなかったため無視しました。"
                )
            if coerced_notes:
                warnings.extend(_dedupe_preserve_order(coerced_notes))
            return entries, warnings

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








        prompt_parts: List[str]
        if include_context_columns:
            prompt_parts = [
                "You are given Japanese source sentences together with extracted reference sentence pairs.\n",
                "Translate each sentence into natural English while keeping the original order and translating all content; never leave Japanese text untranslated.\n",
                "Use the provided reference_pairs as factual evidence while composing the translations. You may generate fluent English output, but do not introduce new facts or supporting sentences that are not grounded in the source or these pairs.\n",
                "Ignore bibliographies, reference lists, document metadata, and any material that does not convey the core meaning of the source sentences.\n",
                "Return a JSON array. Each element must contain exactly these keys:\n",
                "- \"translated_text\": the English translation.\n",
                "- \"process_notes_jp\": 2-6 Japanese sentences summarizing key translation decisions and terminology choices. Use natural Japanese.\n",
                "- \"reference_pairs\": an array of objects `{ \"source_sentence\": \"...\", \"target_sentence\": \"...\" }` listing the reference sentence pairs you actually relied on (use an empty array when none are available).\n",
                "Do not include markup, comments, or extra keys. The output must be valid JSON only.\n",
                "Source sentences are provided below as a JSON array in the original order.\n",
                "Supporting reference_pairs for each sentence are listed after the source list.\n",
            ]
            prompt_preamble = "".join(prompt_parts)
        else:
            prompt_preamble = (
                "Translate each Japanese entry below into English while preserving order and meaning.\n"
                "Return a JSON array of strings matching the input order. Each element must be the natural English translation only; no explanations, quotes, or additional keys.\n"
                "Output must be pure JSON with no extra text before or after the array.\n"
            )
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
                if "\n" in normalized_cell:
                    segments = normalized_cell.split("\n")
                    if any(JAPANESE_CHAR_PATTERN.search(segment) for segment in segments):
                        pending_cols.add(col_idx)
                elif JAPANESE_CHAR_PATTERN.search(cell_value):
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

        def _write_row_output(row_idx: int) -> None:
            _ensure_not_stopped()
            wrote_anything = False
            row_reference = _output_row_reference(row_idx)
            row_slice = output_matrix[row_idx][:output_total_cols]
            if row_dirty_flags[row_idx]:
                write_message = actions.write_range(row_reference, [row_slice], output_sheet)
                incremental_row_messages.append(write_message)
                row_dirty_flags[row_idx] = False
                wrote_anything = True
            if overwrite_source and not writing_to_source_directly and source_row_dirty_flags[row_idx]:
                source_reference = _source_row_reference(row_idx)
                overwrite_message = actions.write_range(
                    source_reference,
                    [source_matrix[row_idx][:source_cols]],
                    target_sheet,
                )
                incremental_row_messages.append(overwrite_message)
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
        output_dirty = False
        source_dirty = False

        limit_to_single = references_requested or use_references
        if limit_to_single:
            items_per_request = 1
        else:
            items_per_request = max(1, rows_per_batch or _ITEMS_PER_TRANSLATION_REQUEST)

        def _generate_key_phrases_batch(current_texts: List[str]) -> List[List[str]]:
            if not include_context_columns:
                return [[] for _ in current_texts]
            if not current_texts:
                return []

            texts_json = json.dumps(current_texts, ensure_ascii=False)
            keyphrase_prompt_sections: List[str] = [
                "以下の日本語原文それぞれについて、主要な論点・関係者・背景を押さえるキーフレーズを日本語で6個作成してください。",
                "ルール:",
                "- 各キーフレーズは3〜12文字程度で簡潔にまとめる。",
                "- 同じ語や言い回しを繰り返さず、視点や焦点を変えて多面的にカバーする。",
                "- 番号や箇条書き記号は使わず文章で記載する。",
                "- 出力はJSON配列のみ。各要素は {\"key_phrases\": [...]} 形式で、入力順と同じ長さにする。",
                "",
                "日本語原文(JSON):",
                texts_json,
            ]
            keyphrase_prompt = "\n".join(keyphrase_prompt_sections)
            _ensure_not_stopped()
            actions.log_progress("キーフレーズ生成: Copilotに依頼中...")
            keyphrase_response = browser_manager.ask(keyphrase_prompt, stop_event=stop_event)
            try:
                match = re.search(r'{.*}|\[.*\]', keyphrase_response, re.DOTALL)
                keyphrase_payload = match.group(0) if match else keyphrase_response
                keyphrase_items = json.loads(keyphrase_payload)
            except json.JSONDecodeError as exc:
                raise ToolExecutionError(
                    f"Failed to parse key phrase generation response as JSON: {keyphrase_response}"
                ) from exc
            if not isinstance(keyphrase_items, list) or len(keyphrase_items) != len(current_texts):
                raise ToolExecutionError(
                    "Key phrase response must be a list with one entry per source text."
                )

            cleaned_results: List[List[str]] = [[] for _ in current_texts]
            for item_index, item in enumerate(keyphrase_items):
                raw_phrases: List[str] = []
                if isinstance(item, dict):
                    raw_phrases = item.get("key_phrases") or item.get("keywords") or []
                elif isinstance(item, list):
                    raw_phrases = item
                if not isinstance(raw_phrases, list):
                    raw_phrases = []
                cleaned_phrases: List[str] = []
                for phrase in raw_phrases:
                    if not isinstance(phrase, str):
                        continue
                    stripped = phrase.strip()
                    if not stripped or stripped in cleaned_phrases:
                        continue
                    cleaned_phrases.append(stripped)
                    if len(cleaned_phrases) >= 6:
                        break
                if not cleaned_phrases:
                    fallback_phrase = current_texts[item_index][:30].strip()
                    cleaned_phrases = [fallback_phrase] if fallback_phrase else ["原文"]
                cleaned_results[item_index] = cleaned_phrases
            return cleaned_results

        def _extract_source_sentences_batch(
            current_texts: List[str],
            key_phrases_per_item: List[List[str]],
        ) -> List[List[str]]:
            if not (use_references and source_reference_url_entries):
                return [[] for _ in current_texts]
            if not current_texts:
                return []

            items_payload: List[Dict[str, Any]] = [
                {
                    "japanese": source_text,
                    "key_phrases": key_phrases_per_item[idx] if idx < len(key_phrases_per_item) else [],
                }
                for idx, source_text in enumerate(current_texts)
            ]
            items_json = json.dumps(items_payload, ensure_ascii=False)
            source_reference_urls_payload: List[str] = [
                entry["url"]
                for entry in source_reference_url_entries
                if isinstance(entry.get("url"), str) and entry["url"].strip()
            ]
            source_reference_urls_json = json.dumps(source_reference_urls_payload, ensure_ascii=False)

            source_sentence_prompt_sections: List[str] = [
                "以下の入力に基づき、直前に抽出したキーフレーズと最も関連する参照資料の文章を抽出してください。",
                "",
                "手順:",
                "- 各項目の key_phrases を手掛かりに参照URLを開き、キーフレーズの意味・用語・事実を裏付ける本文の文を最大6文まで特定する。",
                "- 必ず参照資料本文に実際に存在する文をそのまま引用し、新しい文章の生成や要約、原文の翻訳で置き換えない。語尾・句読点も原文どおりに保持する。",
                "- 1つの段落に複数の意味が含まれる場合は「。」や改行などで区切り、意味が異なる部分は別々の文として抽出する。",
                "- 文中のURLや脚注記号（例: [1]）は除去し、本文だけを残す。",
                "- 参考文献一覧や書誌情報、ヘッダー/フッター、要約・メタ説明など本文以外は抽出しない。",
                "- items(JSON) に含まれる原文テキストや key_phrases をそのままコピーしない。参照資料内で一致する文のみを引用する。",
                "- 一致する文が見つからない場合は空配列で返し、推測や翻訳文を入れない。",
                "- 出力はJSON配列のみ。各要素は {\"source_sentences\": [...]} 形式で、入力順と一致させる。",
                "- `source_sentences` は参照資料から抜き出した引用文の配列であり、原文を翻訳した文ではないことを明示的に守る。",
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
            try:
                match = re.search(r'{.*}|\[.*\]', source_sentence_response, re.DOTALL)
                source_sentence_payload = match.group(0) if match else source_sentence_response
                source_sentence_items = json.loads(source_sentence_payload)
            except json.JSONDecodeError as exc:
                raise ToolExecutionError(
                    f"Failed to parse source reference response as JSON: {source_sentence_response}"
                ) from exc
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
                    cleaned_sentences.append(stripped)
                    if len(cleaned_sentences) >= 6:
                        break
                cleaned_results[item_index] = cleaned_sentences
            return cleaned_results

        def _pair_target_sentences_batch(
            current_texts: List[str],
            key_phrases_per_item: List[List[str]],
            source_references_per_item: List[List[str]],
        ) -> List[List[Dict[str, str]]]:
            if not (use_references and target_reference_url_entries):
                return [[] for _ in current_texts]
            if not current_texts:
                return []

            pairing_payload: List[Dict[str, Any]] = [
                {
                    "japanese": source_text,
                    "key_phrases": key_phrases_per_item[idx] if idx < len(key_phrases_per_item) else [],
                    "source_sentences": source_references_per_item[idx] if idx < len(source_references_per_item) else [],
                }
                for idx, source_text in enumerate(current_texts)
            ]
            pairing_items_json = json.dumps(pairing_payload, ensure_ascii=False)
            target_reference_urls_payload: List[str] = [
                entry["url"]
                for entry in target_reference_url_entries
                if isinstance(entry.get("url"), str) and entry["url"].strip()
            ]
            target_reference_urls_json = json.dumps(target_reference_urls_payload, ensure_ascii=False)

            pairing_prompt_sections: List[str] = [
                f"以下の情報を使い、原文参照文と{target_language}参照文のペアを作成してください。",
                "",
                "手順:",
                "- source_sentences の各文に対して、参照資料から意味的に最も対応する翻訳先言語の文を1文選ぶ。",
                "- 必ず参照資料に実際に存在する文をそのまま引用し、新たに文章を生成したり書き換えたりしない。語尾や句読点も資料の表記を保つ。",
                "- source_sentences を翻訳した推測文を target_sentence として出力しない。target_reference_urls で示した参照資料に存在する文のみを引用する。",
                "- 対応する文が複数ある場合は、最も直接的に一致するものを選択する。",
                "- 適合する文が見つからないsource_sentenceはペアに含めない。",
                "- 文は資料からそのまま引用し、不要な引用符やURLを含めない。",
                "- source_sentences が複数ある場合は、それぞれに対応する文を探し、可能な限り同数のペアを作成する。",
                "- 参考文献や書誌情報、翻訳以外の付録・注記はペアに含めない。",
                "- 出力は純粋なJSON配列のみとし、前後に説明やコメントを付けない。",
                "",
                "出力形式:",
                "- JSON配列。各要素は {\"pairs\": [{\"source_sentence\": \"...\", \"target_sentence\": \"...\"}, ...]}。",
                "- ペアは source_sentence の順序を保ち、入力長と一致させる。",
                "",
                "items(JSON):",
                pairing_items_json,
            ]
            if target_reference_urls_payload:
                pairing_prompt_sections.extend(
                    [
                        "",
                        "target_reference_urls(JSON):",
                        target_reference_urls_json,
                    ]
                )

            pairing_prompt = "\n".join(pairing_prompt_sections)
            _ensure_not_stopped()
            actions.log_progress("対になる英語参照文抽出: Copilotに依頼中...")

            def _request_pairing(prompt: str) -> Tuple[Optional[List[Any]], str]:
                response = browser_manager.ask(prompt, stop_event=stop_event)
                try:
                    match = re.search(r'{.*}|\[.*\]', response, re.DOTALL)
                    payload_json = match.group(0) if match else response
                    return json.loads(payload_json), response
                except json.JSONDecodeError:
                    return None, response

            pairing_items, raw_pairing_response = _request_pairing(pairing_prompt)
            if pairing_items is None:
                snippet = raw_pairing_response.strip().replace("\n", " ")
                actions.log_progress(
                    f"参照ペア応答解析失敗: {snippet[:180]}{'…' if len(snippet) > 180 else ''}"
                )
                _ensure_not_stopped()
                retry_prompt_sections = pairing_prompt_sections + [
                    "",
                    "IMPORTANT:",
                    "- Respond with a pure JSON array only. Do not include explanations before or after the array.",
                    "- If no pairs are found, reply with [] (an empty JSON array).",
                ]
                retry_prompt = "\n".join(retry_prompt_sections)
                actions.log_progress(
                    "参照ペア応答がJSON形式ではなかったため、JSON限定指示で再試行します。"
                )
                pairing_items, raw_pairing_response = _request_pairing(retry_prompt)

            if pairing_items is None:
                snippet = raw_pairing_response.strip().replace("\n", " ")
                actions.log_progress(
                    f"参照ペア再試行も失敗: {snippet[:180]}{'…' if len(snippet) > 180 else ''}"
                )
                actions.log_progress(
                    "参照ペア応答をJSONとして解釈できなかったため、参照ペアなしとして処理を継続します。"
                )
                pairing_items = [{"pairs": []} for _ in current_texts]
            if not isinstance(pairing_items, list) or len(pairing_items) != len(current_texts):
                actions.log_progress(
                    "参照ペア応答の形式が想定外だったため、参照ペアなしとして処理を継続します。"
                )
                pairing_items = [{"pairs": []} for _ in current_texts]

            cleaned_results: List[List[Dict[str, str]]] = [[] for _ in current_texts]
            for item_index, entry in enumerate(pairing_items):
                raw_pairs: List[Any] = []
                if isinstance(entry, dict):
                    raw_pairs = entry.get("pairs") or entry.get("reference_pairs") or []
                elif isinstance(entry, list):
                    raw_pairs = entry
                if not isinstance(raw_pairs, list):
                    raw_pairs = []
                cleaned_pairs: List[Dict[str, str]] = []
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
                    cleaned_pairs.append(
                        {
                            "source_sentence": source_clean,
                            "target_sentence": target_clean,
                        }
                    )
                    if len(cleaned_pairs) >= 6:
                        break
                cleaned_results[item_index] = cleaned_pairs
            return cleaned_results

        for row_start in range(0, source_rows, batch_size):
            _ensure_not_stopped()
            row_end = min(row_start + batch_size, source_rows)
            chunk_texts: List[str] = []
            chunk_positions: List[Tuple[int, int, Optional[int]]] = []
            multi_line_segments: Dict[Tuple[int, int], Dict[str, Any]] = {}

            for local_row in range(row_start, row_end):
                _ensure_not_stopped()
                for col_idx in range(source_cols):
                    cell_value = original_data[local_row][col_idx]
                    if not isinstance(cell_value, str):
                        continue

                    cell_key = (local_row, col_idx)
                    normalized_cell = cell_value.replace('\r\n', '\n').replace('\r', '\n')
                    if '\n' in normalized_cell:
                        segments = normalized_cell.split('\n')
                        pending_indexes: Set[int] = set()
                        translated_segments: List[Optional[str]] = []
                        for seg_index, segment_text in enumerate(segments):
                            if JAPANESE_CHAR_PATTERN.search(segment_text):
                                chunk_texts.append(segment_text)
                                chunk_positions.append((local_row, col_idx, seg_index))
                                pending_indexes.add(seg_index)
                                translated_segments.append(None)
                            else:
                                translated_segments.append(segment_text)
                        if pending_indexes:
                            multi_line_segments[cell_key] = {
                                "segments": segments,
                                "translated_segments": translated_segments,
                                "pending_indexes": pending_indexes,
                                "process_notes": {},
                                "reference_pairs": {},
                            }
                            continue

                    if JAPANESE_CHAR_PATTERN.search(cell_value):
                        chunk_texts.append(cell_value)
                        chunk_positions.append((local_row, col_idx, None))

            if not chunk_texts:
                continue

            chunk_cell_evidences: Dict[Tuple[int, int], Dict[str, Any]] = {}
            row_evidence_details: Dict[int, List[Dict[str, Any]]] = {}

            chunk_entries = list(zip(chunk_texts, chunk_positions))
            for entry_start in range(0, len(chunk_entries), items_per_request):
                _ensure_not_stopped()
                entry_slice = chunk_entries[entry_start:entry_start + items_per_request]
                current_texts = [text for text, _ in entry_slice]
                current_positions = [pos for _, pos in entry_slice]

                key_phrases_per_item = _generate_key_phrases_batch(current_texts)
                source_references_per_item = _extract_source_sentences_batch(current_texts, key_phrases_per_item)
                reference_pairs_context = _pair_target_sentences_batch(
                    current_texts,
                    key_phrases_per_item,
                    source_references_per_item,
                )

                texts_json = json.dumps(current_texts, ensure_ascii=False)

                translation_context = [
                    {
                        "reference_pairs": reference_pairs_context[index] if index < len(reference_pairs_context) else [],
                    }
                    for index, _ in enumerate(current_texts)
                ]
                translation_context_json = json.dumps(translation_context, ensure_ascii=False)

                final_prompt_parts = [
                    prompt_preamble,
                    "Source sentences:",
                    texts_json,
                    "Supporting data (JSON):",
                    translation_context_json,
                ]
                final_prompt = "\n".join(final_prompt_parts) + "\n"
                _ensure_not_stopped()
                response = browser_manager.ask(final_prompt, stop_event=stop_event)

                try:
                    match = re.search(r'{.*}|\[.*\]', response, re.DOTALL)
                    json_payload = match.group(0) if match else response
                    parsed_payload = json.loads(json_payload)
                except json.JSONDecodeError:
                    final_prompt = f"{prompt_preamble}{texts_json}"
                    _ensure_not_stopped()
                    response = browser_manager.ask(final_prompt, stop_event=stop_event)
                    try:
                        match = re.search(r'{.*}|\[.*\]', response, re.DOTALL)
                        json_payload = match.group(0) if match else response
                        parsed_payload = json.loads(json_payload)
                    except json.JSONDecodeError as exc:
                        raise ToolExecutionError(
                            f"Failed to parse translation response as JSON: {response}"
                        ) from exc

                if not isinstance(parsed_payload, list) or len(parsed_payload) != len(current_texts):
                    raise ToolExecutionError(
                        "Translation response must be a list with one entry per source text."
                    )

                for item_index, (item, position) in enumerate(zip(parsed_payload, current_positions)):
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
                                or
                                item.get("explanation_jp")
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

                    translation_value = translation_value.strip()
                    translation_value = _maybe_unescape_html_entities(translation_value)
                    if not translation_value:
                        raise ToolExecutionError("Translation response returned an empty 'translated_text' value.")

                    source_cell_value = _normalize_cell_value(original_data[position[0]][position[1]]).strip()
                    if translation_value == source_cell_value and _contains_japanese(translation_value):
                        raise ToolExecutionError(
                            "翻訳結果が元のテキストと同一で日本語のままです。翻訳が完了していません。"
                        )
                    if target_language and target_language.lower().startswith("english") and _contains_japanese(translation_value):
                        raise ToolExecutionError(
                            "翻訳結果に日本語が含まれているため、英語への翻訳が完了していません。"
                        )

                    if len(position) == 3:
                        local_row, col_idx, segment_index = position
                    else:
                        local_row, col_idx = position
                        segment_index = None

                    if writing_to_source_directly:
                        translation_col_index = col_idx
                        explanation_col_index = None
                        pair_start_index = None
                        pair_end_index = None
                    else:
                        translation_col_index = col_idx * translation_block_width
                        if include_context_columns:
                            explanation_col_index = translation_col_index + 1
                            pair_start_index = translation_col_index + 2
                            pair_end_index = translation_col_index + translation_block_width - 1
                        else:
                            explanation_col_index = None
                            pair_start_index = None
                            pair_end_index = None

                    process_notes_text = (
                        _maybe_unescape_html_entities(process_notes_jp).strip()
                        if include_context_columns
                        else ""
                    )
                    context_pairs: List[Dict[str, str]] = []
                    if include_context_columns:
                        if reference_pairs_context and item_index < len(reference_pairs_context):
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
                    cell_key = (local_row, col_idx)
                    multi_segment_state = multi_line_segments.get(cell_key)
                    if multi_segment_state and segment_index is not None:
                        multi_segment_state['translated_segments'][segment_index] = translation_value
                        if process_notes_text:
                            multi_segment_state.setdefault('process_notes', {})[segment_index] = process_notes_text
                        if reference_pairs_list:
                            multi_segment_state.setdefault('reference_pairs', {})[segment_index] = list(reference_pairs_list)
                        pending = multi_segment_state.get('pending_indexes')
                        if pending is not None:
                            pending.discard(segment_index)
                            if pending:
                                continue
                        translated_segments = []
                        for idx, segment in enumerate(multi_segment_state.get('translated_segments', [])):
                            if segment is None:
                                translated_segments.append(multi_segment_state['segments'][idx])
                            else:
                                translated_segments.append(segment)
                        translation_value = "\n".join(translated_segments)
                        ordered_notes = [
                            multi_segment_state.get('process_notes', {}).get(idx, '')
                            for idx in range(len(multi_segment_state.get('segments', [])))
                        ]
                        process_notes_text = "\n".join([entry for entry in ordered_notes if entry]).strip()
                        aggregated_pairs: List[Dict[str, str]] = []
                        for idx in range(len(multi_segment_state.get('segments', []))):
                            aggregated_pairs.extend(
                                multi_segment_state.get('reference_pairs', {}).get(idx, [])
                            )
                        reference_pairs_list = aggregated_pairs

                    existing_output_value = output_matrix[local_row][translation_col_index]
                    if translation_value != existing_output_value:
                        output_matrix[local_row][translation_col_index] = translation_value
                        output_dirty = True
                        row_dirty_flags[local_row] = True
                    if include_context_columns:
                        preview = translation_value
                        if len(preview) > 120:
                            preview = preview[:117] + "..."
                        actions.log_progress(f"翻訳結果プレビュー ({item_index + 1}/{len(current_texts)}): {preview}")
                    if not writing_to_source_directly and overwrite_source:
                        existing_source_value = source_matrix[local_row][col_idx]
                        if translation_value != existing_source_value:
                            source_matrix[local_row][col_idx] = translation_value
                            source_dirty = True
                            source_row_dirty_flags[local_row] = True

                    any_translation = True

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
                        if include_context_columns:
                            expected_pairs = len(source_references_per_item[item_index]) if item_index < len(source_references_per_item) else 0
                            actions.log_progress(
                                f"参照ペア整理結果 ({item_index + 1}/{len(current_texts)}): {len(sanitized_pairs)} 件 / source_sentences={expected_pairs}"
                            )
                            if sanitized_pairs:
                                for idx, pair in enumerate(sanitized_pairs, start=1):
                                    actions.log_progress(
                                        f"参照ペア[{item_index + 1}-{idx}]: {pair['source_sentence']} -> {pair['target_sentence']}"
                                    )
                            else:
                                actions.log_progress(
                                    f"参照ペア整理結果 ({item_index + 1}/{len(current_texts)}): 0 件 (参照資料に一致する文が見つかりませんでした)"
                                )

                    formatted_pairs: List[str] = []
                    if include_context_columns:
                        for pair in sanitized_pairs:
                            formatted_pairs.append(f"{pair['source_sentence']}\n---\n{pair['target_sentence']}")
                        if not formatted_pairs:
                            formatted_pairs = [_NO_QUOTES_PLACEHOLDER]

                    fallback_reason: Optional[str] = None
                    if include_context_columns and use_references:
                        default_explanation = "参照資料の内容を踏まえ、原文の意味と語調を保つように訳語を選定しました。"
                        if not process_notes_text:
                            process_notes_text = default_explanation
                            fallback_reason = "process_notes_jp が欠落していたため既定の説明文を補いました。"
                        elif not JAPANESE_CHAR_PATTERN.search(process_notes_text):
                            process_notes_text = default_explanation
                            fallback_reason = "process_notes_jp に日本語が含まれていなかったため既定の説明文を補いました。"
                        elif len(process_notes_text) < 20:
                            process_notes_text = (
                                process_notes_text + "。原文の語調と用語整合性を確認して訳語を決定しました。"
                            ).strip()
                            if len(process_notes_text) < 20 or not JAPANESE_CHAR_PATTERN.search(process_notes_text):
                                process_notes_text = default_explanation
                                fallback_reason = "process_notes_jp が短すぎたため既定の説明文を補いました。"
                            else:
                                fallback_reason = "process_notes_jp が短かったため補足説明を追加しました。"

                        if fallback_reason:
                            absolute_row = source_start_row + local_row
                            absolute_col = source_start_col + col_idx
                            cell_ref = _build_range_reference(
                                absolute_row,
                                absolute_row,
                                absolute_col,
                                absolute_col,
                            )
                            if target_sheet:
                                cell_ref = f"{target_sheet}!{cell_ref}"
                            explanation_fallback_notes.append(f"{cell_ref}: {fallback_reason}")

                    if include_context_columns:
                        if use_references:
                            if not process_notes_text:
                                raise ToolExecutionError("Translation response must include a 'process_notes_jp' string for each item.")
                            if not JAPANESE_CHAR_PATTERN.search(process_notes_text):
                                raise ToolExecutionError("process_notes_jp の記載は必ず日本語で行ってください。")
                            if len(process_notes_text) < 20:
                                raise ToolExecutionError("process_notes_jp には翻訳判断を具体的に説明してください (20文字以上)。")
                        if pair_start_index is not None:
                            allowed_pairs = max_reference_pairs_per_item
                            if allowed_pairs < len(formatted_pairs):
                                raise ToolExecutionError(
                                    f"translation_output_range does not have enough columns to record {len(formatted_pairs)} reference pairs. "
                                    f"Provide at least {len(formatted_pairs)} pair columns (currently {allowed_pairs})."
                                )

                    if explanation_col_index is not None and include_context_columns:
                        if output_matrix[local_row][explanation_col_index] != process_notes_text:
                            output_matrix[local_row][explanation_col_index] = process_notes_text
                            output_dirty = True
                            row_dirty_flags[local_row] = True
                    if include_context_columns and pair_start_index is not None and pair_end_index is not None:
                        total_pair_slots = pair_end_index - pair_start_index + 1
                        for offset in range(total_pair_slots):
                            absolute_col = pair_start_index + offset
                            desired_value = formatted_pairs[offset] if offset < len(formatted_pairs) else ""
                            if output_matrix[local_row][absolute_col] != desired_value:
                                output_matrix[local_row][absolute_col] = desired_value
                                output_dirty = True
                                row_dirty_flags[local_row] = True

                    evidence_record = {
                        "process_notes": process_notes_text,
                        "reference_pairs": sanitized_pairs,
                        "reference_pair_lines": formatted_pairs,
                    }

                    if use_references:
                        if citation_mode in {"paired_columns", "translation_triplets"}:
                            chunk_cell_evidences[(local_row, col_idx)] = evidence_record
                        elif citation_mode == "per_cell":
                            chunk_cell_evidences[(local_row, col_idx)] = evidence_record
                        elif citation_mode == "single_column":
                            row_evidence_details.setdefault(local_row, []).append(evidence_record)

                    pending_cols = pending_columns_by_row.get(local_row)
                    if pending_cols is not None:
                        pending_cols.discard(col_idx)
                        if not pending_cols:
                            _finalize_row(local_row)
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

        output_dirty = any(row_dirty_flags)
        if overwrite_source and not writing_to_source_directly:
            source_dirty = any(source_row_dirty_flags)
        else:
            source_dirty = False

        if not any_translation:
            return f"No translatable text was found in range '{cell_range}'."

        if include_context_columns and explanation_fallback_notes:
            messages.insert(0, "process_notes_jp が不足していたセルに既定の説明文を補いました: " + " / ".join(explanation_fallback_notes))

        write_messages: List[str] = []

        if range_adjustment_note:
            write_messages.append(range_adjustment_note)
        if citation_note:
            write_messages.append(citation_note)
        write_messages.extend(incremental_row_messages)
        if reference_warning_notes:
            write_messages.extend(reference_warning_notes)

        _ensure_not_stopped()

        if output_dirty:
            translation_message = actions.write_range(output_range, output_matrix, output_sheet)
            write_messages.append(translation_message)

        if overwrite_source and not writing_to_source_directly and source_dirty:
            overwrite_message = actions.write_range(normalized_range, source_matrix, target_sheet)
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
    rows_per_batch: Optional[int] = None,
    stop_event: Optional[Event] = None,
) -> str:
    """Translate ranges without using external reference material."""
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
    """Translate text in a range and write the output plus optional context.

    

    Args:

        actions: Excel automation helper injected by the agent runtime.

        browser_manager: Shared browser manager used for translation API calls.

        cell_range: Range containing the source text.

        target_language: Target language name, defaults to English.

        sheet_name: Optional sheet override; defaults to the active sheet.

        translation_output_range: Optional range for translated rows (three columns per source column).

        overwrite_source: Whether to overwrite the source range directly.

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
                "Assess both factual accuracy and translation quality: look for unnatural tone, awkward phrasing, register mismatches, overly literal wording, or inconsistent terminology even when the meaning is roughly correct, and propose smoother alternatives.\n"
                "Respond with a JSON array containing exactly one object. Provide: 'id', 'status', 'notes', 'corrected_text', and 'highlighted_text'.\n"
                "Optionally include 'before_text', 'after_text', or an 'edits' array (each element with fields 'type', 'text', and 'reason').\n"
                "Use status 'OK' only when the draft translation requires no changes; otherwise respond with 'REVISE'.\n"
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
                    status_matrix[row_idx][col_idx] = "要修正"
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
