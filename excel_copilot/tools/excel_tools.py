import re
import difflib
import logging
import os
import string
from threading import Event
from typing import List, Any, Optional, Dict, Tuple, Set

from excel_copilot.core.browser_copilot_manager import BrowserCopilotManager
from excel_copilot.core.exceptions import ToolExecutionError, UserStopRequested

from .actions import ExcelActions


_logger = logging.getLogger(__name__)
_DIFF_DEBUG_ENABLED = os.getenv('EXCEL_COPILOT_DEBUG_DIFF', '').lower() in {'1', 'true', 'yes'}

_NO_QUOTES_PLACEHOLDER = "引用なし"

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
        return cell
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
    trailing_len = len(segment.rstrip()) - len(segment.strip())
    core_start = leading_len
    core_end = len(segment) - trailing_len if trailing_len else len(segment)
    core = segment[core_start:core_end]
    if not core:
        _diff_debug(f"_format_diff_segment no core text label={label}")
        return segment, None, None
    prefix = segment[:leading_len]
    suffix = segment[core_end:]
    marker_prefix = f"[{label}]"
    marker_suffix = ""
    formatted = f'{prefix}{marker_prefix}{core}{marker_suffix}{suffix}'
    highlight_start_offset = len(prefix)
    highlight_length = len(marker_prefix) + len(core) + len(marker_suffix)
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

            if trimmed_removed and not trimmed_added and suffix:
                suffix_leading = 0
                while suffix_leading < len(suffix) and suffix[suffix_leading].isspace():
                    suffix_leading += 1
                suffix_end = suffix_leading
                while suffix_end < len(suffix) and not suffix[suffix_end].isspace():
                    suffix_end += 1
                if suffix_end > 0:
                    trimmed_added = suffix[:suffix_end]
                    suffix = suffix[suffix_end:]

            if trimmed_removed and not trimmed_added:
                trimmed_added = '（追加なし）'

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

        span_start = current_length
        output_segments.append(segment_text)
        span_length = len(segment_text)
        current_length += span_length
        if span_length > 0:
            spans.append({"start": span_start, "length": span_length, "type": open_type.upper()})

        cursor = match.end()

    if cursor < len(raw_text):
        trailing_text = raw_text[cursor:]
        output_segments.append(trailing_text)

    clean_text = "".join(output_segments)
    return clean_text, spans


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
    reference_ranges: Optional[List[str]] = None,
    citation_output_range: Optional[str] = None,
    reference_urls: Optional[List[str]] = None,
    translation_output_range: Optional[str] = None,
    overwrite_source: bool = False,
    rows_per_batch: Optional[int] = None,
    stop_event: Optional[Event] = None,
    output_mode: str = "translation_with_context",
) -> str:
    '''Translate Japanese text for a range with optional references and controlled output.'''

    try:
        def _ensure_not_stopped() -> None:
            if stop_event and stop_event.is_set():
                raise UserStopRequested("ユーザーによって処理が中断されました。")

        _ensure_not_stopped()
        target_sheet, normalized_range = _split_sheet_and_range(cell_range, sheet_name)
        source_rows, source_cols = _parse_range_dimensions(normalized_range)

        raw_original = actions.read_range(normalized_range, target_sheet)
        original_data = _reshape_to_dimensions(raw_original, source_rows, source_cols)

        if source_rows == 0 or source_cols == 0:
            return f"Range '{cell_range}' has no usable cells to translate."

        source_matrix = [row[:] for row in original_data]
        range_adjustment_note: Optional[str] = None
        writing_to_source_directly = translation_output_range is None
        include_context_columns = output_mode != "translation_only"
        translation_block_width = 3 if include_context_columns else 1
        citation_should_include_explanations = writing_to_source_directly and include_context_columns
        if writing_to_source_directly and not overwrite_source:
            raise ToolExecutionError(
                "translation_output_range must be provided when overwrite_source is False."
            )
        if writing_to_source_directly:
            output_sheet = target_sheet
            output_range = normalized_range
            output_matrix = source_matrix
        else:
            output_sheet, output_range = _split_sheet_and_range(translation_output_range, target_sheet)
            out_rows, out_cols = _parse_range_dimensions(output_range)
            required_output_cols = source_cols * translation_block_width
            if out_rows < source_rows or out_cols < required_output_cols:
                if include_context_columns:
                    requirement_text = "three columns (translation, quotes, explanation)"
                else:
                    requirement_text = "one column (translation)"
                raise ToolExecutionError(
                    f"translation_output_range must provide {requirement_text} per source column."
                )
            if out_rows != source_rows or out_cols != required_output_cols:
                start_row, start_col, _, _ = _parse_range_bounds(output_range)
                adjusted_end_row = start_row + source_rows - 1
                adjusted_end_col = start_col + required_output_cols - 1
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
                if include_context_columns:
                    requirement_text = "translation / quotes / explanation layout"
                else:
                    requirement_text = "translation-only layout"
                range_adjustment_note = (
                    f"translation_output_range '{original_range_display}' did not match the required {requirement_text}; "
                    f"using '{adjusted_range_display}' instead."
                )
                output_range = adjusted_range
                out_rows, out_cols = source_rows, required_output_cols
            raw_output = actions.read_range(output_range, output_sheet)
            try:
                output_matrix = _reshape_to_dimensions(raw_output, out_rows, out_cols)
            except ToolExecutionError:
                output_matrix = [["" for _ in range(out_cols)] for _ in range(out_rows)]

        _ensure_not_stopped()

        reference_entries: List[Dict[str, Any]] = []
        reference_text_pool: List[str] = []
        if reference_ranges:
            range_list = [reference_ranges] if isinstance(reference_ranges, str) else list(reference_ranges)
            for raw_range in range_list:
                _ensure_not_stopped()
                ref_sheet, ref_range = _split_sheet_and_range(raw_range, target_sheet)
                try:
                    ref_data = actions.read_range(ref_range, ref_sheet)
                except ToolExecutionError as exc:
                    raise ToolExecutionError(f"指定した範囲 '{raw_range}' の検証中にエラーが発生しました: {exc}") from exc

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
                raise ToolExecutionError(
                    "No usable reference content found in the provided reference_ranges."
                )

        reference_url_entries: List[Dict[str, str]] = []
        if reference_urls:
            url_list = [reference_urls] if isinstance(reference_urls, str) else list(reference_urls)
            for raw_url in url_list:
                _ensure_not_stopped()
                if not isinstance(raw_url, str):
                    raise ToolExecutionError(
                        "Each reference_urls entry must be a string."
                    )
                url = raw_url.strip()
                if not url:
                    continue
                reference_url_entries.append({
                    "id": f"U{len(reference_url_entries) + 1}",
                    "url": url,
                })

        references_requested = bool(reference_ranges) or bool(reference_urls)
        use_references = bool(reference_entries or reference_url_entries)
        reference_text_pool = [text for text in reference_text_pool if text]

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
        if use_references:
            prompt_parts = [
                "Translate each Japanese entry below into English; keep the order and stay faithful to the source.\n",
                "Every translated_text must be written in natural English; do not copy or leave any Japanese text untranslated.\n",
                "Use the references/URLs only to keep terminology consistent and never emit citation markers in the translation output.\n",
                "When copying supporting material, remove any embedded URLs or hyperlink targets; if citations are unavoidable, retain only bracketed numbers like [1].\n",
                "Borrow phrasing and sentence structure from the supporting quotes whenever it improves the English rendering of the Japanese text. Limit this borrowing strictly to wording—do not import additional facts, subjects, or entities from the quotes, and never swap in the quote's subject or perspective. If the Japanese source states a subject, translate that subject explicitly; only omit a subject when the Japanese sentence omits it.\n",
                "Treat the supporting quotes purely as style references for idiomatic English; preserve every entity that appears in the Japanese sentence, and do not introduce new ones from the quotes.\n",
                "Workflow: make English search keywords, scan the references, and reuse wording only when it supports the same fact.\n",
                "Output must be pure JSON with no commentary, preambles, or Markdown—only the requested array.\n",
            ]
            if reference_entries:
                prompt_parts.append(f"Reference passages:\n{json.dumps(reference_entries, ensure_ascii=False)}\n")
            if reference_url_entries:
                prompt_parts.append(f"Reference URLs:\n{json.dumps(reference_url_entries, ensure_ascii=False)}\n")
            prompt_parts.append(
                "Return a JSON array of objects with 'translated_text' and 'explanation_jp' (2-6 Japanese sentences explaining key terminology and tone). No other keys or markdown.\n"
            )
            prompt_preamble = "".join(prompt_parts)
        else:
            if include_context_columns:
                prompt_preamble = (
                    "Translate each Japanese entry below into English while preserving order and meaning.\n"
                    "Every translated_text must be written in natural English; do not copy or leave any Japanese text untranslated.\n"
                    "Reuse supporting expressions when they help, but never add facts or entities not found in the Japanese sentence.\n"
                    "Return a JSON array of the same length, with no commentary or markdown.\n"
                    "Output must be pure JSON with no additional text before or after the array.\n"
                )
            else:
                prompt_preamble = (
                    "Translate each Japanese entry below into English while preserving order and meaning.\n"
                    "Return a JSON array of strings matching the input order. Each element must be the natural English translation only—no explanations, quotes, or additional keys.\n"
                    "Output must be pure JSON with no extra text before or after the array.\n"
                )
        if references_requested or use_references:
            rows_per_batch = 1
        batch_size = rows_per_batch if rows_per_batch is not None else 1
        if batch_size < 1:
            raise ToolExecutionError("rows_per_batch must be at least 1.")

        source_start_row, source_start_col, _, _ = _parse_range_bounds(normalized_range)
        output_start_row, output_start_col, _, _ = _parse_range_bounds(output_range)

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
                    raise ToolExecutionError(
                        "citation_output_range must have either one column, match the source column count, or provide two or three columns per source column (for explanation/quotes or the translation output layout)."
                    )
                cite_start_row, cite_start_col, _, _ = _parse_range_bounds(citation_range)
                existing_citation = actions.read_range(citation_range, citation_sheet)
                try:
                    citation_matrix = _reshape_to_dimensions(existing_citation, cite_rows, cite_cols)
                except ToolExecutionError:
                    citation_matrix = [["" for _ in range(cite_cols)] for _ in range(cite_rows)]

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
                                "explanations": {},
                                "quotes": {},
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

                texts_json = json.dumps(current_texts, ensure_ascii=False)
                normalized_quotes_per_item: List[List[str]] = []

                if include_context_columns:
                    keyword_prompt = (
                        "For each Japanese item in the JSON array below, craft exactly ten distinct English search phrases.\n"
                        "Ensure the set of ten phrases collectively covers the major actors, actions, consequences, motivations, audiences, and surrounding context described in the item, while keeping each phrasing distinct.\n"
                        "Use synonyms and descriptive stand‑ins so that proper nouns appear only when essential, and alternate with role- or category-based wording (e.g., 'the automaker', 'the sports car') so no single substantive word appears in every phrase.\n"
                        "Mix shorter three-to-four word queries with longer six-to-nine word phrasing, and ensure no two phrases share the same first two words.\n"
                        "Do not invent information beyond the item, and return a JSON array matching the input order where each element exposes only a 'keywords' list.\n"
                        f"{texts_json}"
                    )
                    _ensure_not_stopped()
                    keyword_response = browser_manager.ask(keyword_prompt, stop_event=stop_event)
                    try:
                        match = re.search(r'{.*}|\[.*\]', keyword_response, re.DOTALL)
                        keyword_payload = match.group(0) if match else keyword_response
                        keyword_items = json.loads(keyword_payload)
                    except json.JSONDecodeError as exc:
                        raise ToolExecutionError(
                            f"Failed to parse keyword generation response as JSON: {keyword_response}"
                        ) from exc
                    if not isinstance(keyword_items, list) or len(keyword_items) != len(current_texts):
                        raise ToolExecutionError(
                            "Keyword response must be a list with one entry per source text."
                        )

                    normalized_keywords: List[List[str]] = []
                    for item in keyword_items:
                        if isinstance(item, dict):
                            raw_keywords = item.get("keywords")
                        elif isinstance(item, list):
                            raw_keywords = item
                        else:
                            raw_keywords = None
                        if not raw_keywords or not isinstance(raw_keywords, list):
                            raise ToolExecutionError(
                                "Each keyword entry must contain a 'keywords' list."
                            )
                        keyword_list = []
                        for keyword in raw_keywords:
                            if isinstance(keyword, str):
                                cleaned = keyword.strip()
                                if cleaned:
                                    keyword_list.append(cleaned)
                        if not keyword_list:
                            raise ToolExecutionError(
                                "Each keyword entry must include at least one non-empty keyword."
                            )
                        normalized_keywords.append(keyword_list)

                    search_keywords_per_item: List[List[str]] = []
                    for source_text, base_keywords in zip(current_texts, normalized_keywords):
                        enriched_keywords = _enrich_search_keywords(source_text, list(base_keywords))
                        search_keywords_per_item.append(enriched_keywords)

                    keyword_plan_lines: List[str] = []
                    for index, (source_text, keywords) in enumerate(zip(current_texts, search_keywords_per_item), start=1):
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
                        reference_urls_text = (
                            "Refer to the following URLs only if you need the original passages. "
                            "Do not include these URLs in any quote output; when citations are unavoidable, leave only bracketed numbers like [1].\n"
                            + "\n".join(
                                f"  - {entry['url']}" for entry in reference_url_entries if entry.get("url")
                            )
                        )

                    if use_references:
                        evidence_prompt_sections: List[str] = [
                            "Use the keywords to pull English sentences from the provided materials that support each Japanese item.",
                            "When exact matches are scarce, include sentences that share overlapping entities, themes, or context even if the linkage is loose.",
                            "Copy sentences verbatim (punctuation, casing, numerals) and aim for 4-8 varied quotes drawn from different paragraphs or sources whenever possible; only return an empty array if genuinely nothing relevant appears.",
                            "Avoid citing near-duplicate sentences or repeating the same clause unless no other material exists.",
                            "Return JSON only—no introductions, labels, or commentary before or after the array.",
                            "If you lack supporting material, respond with an empty JSON array `[]` and nothing else.",
                            "Return a JSON array matching the order. Each element needs 'quotes' and an 'explanation_jp' string with at least two Japanese sentences explaining the support.",
                            "",
                            "Japanese texts with search keywords:",
                            keyword_plan_text,
                            "",
                        ]
                        if reference_passage_text:
                            evidence_prompt_sections.extend(["Reference passages:", reference_passage_text, ""])
                        if reference_urls_text:
                            evidence_prompt_sections.extend([
                                "Reference URLs (for lookup only; strip the URLs from quotes and keep at most bracket numbers like [1]):",
                                reference_urls_text,
                                "",
                            ])
                    else:
                        evidence_prompt_sections = [
                            "Use the keywords to draft 3-5 concise English candidate sentences per item that could guide the translation.",
                            "Return JSON only—no introductions, labels, or commentary before or after the array.",
                            "Return a JSON array matching the order. Each element must include a 'quotes' array and an 'explanation_jp' string (>=2 Japanese sentences).",
                            "",
                            "Japanese texts with search keywords:",
                            keyword_plan_text,
                            "",
                        ]
                    evidence_prompt = "\n".join(evidence_prompt_sections)
                    _ensure_not_stopped()
                    evidence_response = browser_manager.ask(evidence_prompt, stop_event=stop_event)
                    try:
                        match = re.search(r'{.*}|\[.*\]', evidence_response, re.DOTALL)
                        evidence_payload = match.group(0) if match else evidence_response
                        evidence_items = json.loads(evidence_payload)
                    except json.JSONDecodeError as exc:
                        raise ToolExecutionError(
                            f"Failed to parse evidence response as JSON: {evidence_response}"
                        ) from exc
                    if not isinstance(evidence_items, list) or len(evidence_items) != len(current_texts):
                        raise ToolExecutionError(
                            "Evidence response must be a list with one entry per source text."
                        )

                    for quotes_entry in evidence_items:
                        if isinstance(quotes_entry, dict):
                            raw_quotes = quotes_entry.get("quotes")
                            if raw_quotes is None:
                                translated_candidate = quotes_entry.get("translated_text")
                                if isinstance(translated_candidate, str) and translated_candidate.strip():
                                    raw_quotes = [translated_candidate]
                        elif isinstance(quotes_entry, list):
                            raw_quotes = quotes_entry
                        else:
                            raw_quotes = None
                        quotes_list: List[str] = []
                        if isinstance(raw_quotes, list):
                            for quote in raw_quotes:
                                if isinstance(quote, str):
                                    cleaned_quote = _strip_reference_urls_from_quote(quote)
                                    if cleaned_quote:
                                        quotes_list.append(cleaned_quote)
                        normalized_quotes_per_item.append(quotes_list)

                    translation_context = [
                        {
                            "source_text": text,
                            "keywords": keywords,
                            "quotes": normalized_quotes_per_item[index] if index < len(normalized_quotes_per_item) else []
                        }
                        for index, (text, keywords) in enumerate(zip(current_texts, search_keywords_per_item))
                    ]
                    translation_context_json = json.dumps(translation_context, ensure_ascii=False)

                    final_prompt = (
                        f"{prompt_preamble}{texts_json}"
                        "Write natural English translations that stay faithful to each Japanese sentence.\n"
                        "Use the supporting expressions only when they fit; do not add or omit facts.\n"
                        "Return a JSON array where every element has 'translated_text' and 'explanation_jp' (>=2 Japanese sentences on terminology and tone). No extra keys, quote arrays, or markdown.\n"
                        f"Supporting expressions (JSON): {translation_context_json}\n"
                    )
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
                else:
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
                    explanation_jp = ""

                    if isinstance(item, dict):
                        translation_value = (
                            item.get("translated_text")
                            or item.get("translation")
                            or item.get("output")
                        )
                        if include_context_columns:
                            evidence_value = None
                            raw_explanation = (
                                item.get("explanation_jp")
                                or item.get("explanation")
                            )
                            if raw_explanation is None:
                                evidence_value = item.get("evidence")
                            if isinstance(evidence_value, dict):
                                raw_explanation = (
                                    evidence_value.get("explanation_jp")
                                    or evidence_value.get("explanation")
                                )
                            if isinstance(raw_explanation, (str, int, float)):
                                explanation_jp = _sanitize_evidence_value(str(raw_explanation))
                    elif isinstance(item, str):
                        translation_value = item
                    elif isinstance(item, (int, float)):
                        translation_value = str(item)

                    if not isinstance(translation_value, str):
                        raise ToolExecutionError(
                            "Translation response must include a 'translated_text' string for each item."
                        )

                    translation_value = translation_value.strip()
                    if not translation_value:
                        raise ToolExecutionError("Translation response returned an empty 'translated_text' value.")

                    source_cell_value = _normalize_cell_value(original_data[position[0]][position[1]]).strip()
                    if not include_context_columns:
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
                        quotes_col_index = None
                        explanation_col_index = None
                    else:
                        translation_col_index = col_idx * translation_block_width
                        if include_context_columns:
                            quotes_col_index = translation_col_index + 1
                            explanation_col_index = translation_col_index + 2
                        else:
                            quotes_col_index = None
                            explanation_col_index = None

                    explanation_text = explanation_jp.strip() if include_context_columns else ""
                    quote_candidates = []
                    if include_context_columns and item_index < len(normalized_quotes_per_item):
                        raw_candidates = normalized_quotes_per_item[item_index]
                        if isinstance(raw_candidates, list):
                            quote_candidates = raw_candidates

                    cell_key = (local_row, col_idx)
                    multi_segment_state = multi_line_segments.get(cell_key)
                    if multi_segment_state and segment_index is not None:
                        multi_segment_state['translated_segments'][segment_index] = translation_value
                        if explanation_text:
                            multi_segment_state.setdefault('explanations', {})[segment_index] = explanation_text
                        if quote_candidates:
                            multi_segment_state.setdefault('quotes', {})[segment_index] = quote_candidates
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
                        ordered_explanations = [
                            multi_segment_state.get('explanations', {}).get(idx, '')
                            for idx in range(len(multi_segment_state.get('segments', [])))
                        ]
                        explanation_text = "\n".join([entry for entry in ordered_explanations if entry]).strip()
                        aggregated_quotes = []
                        for idx in range(len(multi_segment_state.get('segments', []))):
                            aggregated_quotes.extend(
                                multi_segment_state.get('quotes', {}).get(idx, [])
                            )
                        quote_candidates = aggregated_quotes

                    existing_output_value = output_matrix[local_row][translation_col_index]
                    if translation_value != existing_output_value:
                        output_matrix[local_row][translation_col_index] = translation_value
                        output_dirty = True
                    if not writing_to_source_directly and overwrite_source:
                        existing_source_value = source_matrix[local_row][col_idx]
                        if translation_value != existing_source_value:
                            source_matrix[local_row][col_idx] = translation_value
                            source_dirty = True

                    any_translation = True

                    final_quotes: List[str] = []
                    seen_quotes: Set[str] = set()
                    if include_context_columns:
                        for candidate in quote_candidates or []:
                            if not isinstance(candidate, str):
                                continue
                            cleaned_candidate = candidate.strip()
                            if not cleaned_candidate or cleaned_candidate in seen_quotes:
                                continue
                            seen_quotes.add(cleaned_candidate)
                            final_quotes.append(cleaned_candidate)

                    formatted_quotes: List[str]
                    if include_context_columns:
                        if final_quotes:
                            formatted_quotes = [
                                f"引用{idx}: {quote}"
                                for idx, quote in enumerate(final_quotes, start=1)
                            ]
                        else:
                            formatted_quotes = [_NO_QUOTES_PLACEHOLDER]
                    else:
                        formatted_quotes = []

                    fallback_reason: Optional[str] = None
                    if include_context_columns and use_references:
                        default_explanation = "参照資料の内容を踏まえ、原文の意味と語調を保つように訳語を選定しました。"
                        if not explanation_text:
                            explanation_text = default_explanation
                            fallback_reason = "explanation_jp が欠落していたため既定の説明文を補いました。"
                        elif not JAPANESE_CHAR_PATTERN.search(explanation_text):
                            explanation_text = default_explanation
                            fallback_reason = "explanation_jp に日本語が含まれていなかったため既定の説明文を補いました。"
                        elif len(explanation_text) < 20:
                            explanation_text = (
                                explanation_text + "。原文の語調と用語整合性を確認して訳語を決定しました。"
                            ).strip()
                            if len(explanation_text) < 20 or not JAPANESE_CHAR_PATTERN.search(explanation_text):
                                explanation_text = default_explanation
                                fallback_reason = "explanation_jp が短すぎたため既定の説明文を補いました。"
                            else:
                                fallback_reason = "explanation_jp が短かったため補足説明を追加しました。"

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

                    if quotes_col_index is not None and explanation_col_index is not None and include_context_columns:
                        quotes_text = "\n".join(formatted_quotes)
                        if output_matrix[local_row][quotes_col_index] != quotes_text:
                            output_matrix[local_row][quotes_col_index] = quotes_text
                            output_dirty = True
                        if output_matrix[local_row][explanation_col_index] != explanation_text:
                            output_matrix[local_row][explanation_col_index] = explanation_text
                            output_dirty = True

                    if use_references:
                        if not explanation_text:
                            raise ToolExecutionError("Translation response must include an 'explanation_jp' string for each item.")
                        if not JAPANESE_CHAR_PATTERN.search(explanation_text):
                            raise ToolExecutionError("explanation_jp の記載は必ず日本語で行ってください。")
                        if len(explanation_text) < 20:
                            raise ToolExecutionError("explanation_jp には翻訳判断を具体的に説明してください (20文字以上)。")

                    quotes_lines: List[str] = list(formatted_quotes)

                    evidence_record = {
                        "explanation": explanation_text,
                        "quotes_lines": quotes_lines,
                    }

                    if use_references:
                        if citation_mode in {"paired_columns", "translation_triplets"}:
                            chunk_cell_evidences[(local_row, col_idx)] = evidence_record
                        elif citation_mode == "per_cell":
                            chunk_cell_evidences[(local_row, col_idx)] = evidence_record
                        elif citation_mode == "single_column":
                            row_evidence_details.setdefault(local_row, []).append(evidence_record)
            if use_references and citation_matrix is not None:
                if citation_mode == "paired_columns":
                    for local_row in range(row_start, row_end):
                        for col_offset in range(cite_cols):
                            citation_matrix[local_row][col_offset] = ""
                    for (local_row, col_idx), data in chunk_cell_evidences.items():
                        base_col = col_idx * 2
                        if base_col + 1 >= cite_cols:
                            continue
                        explanation_text = (data.get("explanation") or "").strip()
                        quotes_text = "\n".join(data.get("quotes_lines", []))
                        citation_matrix[local_row][base_col] = (
                            explanation_text if citation_should_include_explanations else ""
                        )
                        citation_matrix[local_row][base_col + 1] = quotes_text
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
                        explanation_text = (data.get("explanation") or "").strip()
                        quotes_text = "\n".join(data.get("quotes_lines", []))
                        citation_matrix[local_row][base_col + 1] = quotes_text
                        citation_matrix[local_row][base_col + 2] = (
                            explanation_text if citation_should_include_explanations else ""
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
                        explanation_text = (data.get("explanation") or "").strip()
                        quotes_lines = data.get("quotes_lines", [])
                        combined_lines: List[str] = []
                        if citation_should_include_explanations and explanation_text:
                            combined_lines.append(f"説明: {explanation_text}")
                        combined_lines.extend(quotes_lines)
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
                            explanation_text = (data.get("explanation") or "").strip()
                            quotes_lines = data.get("quotes_lines", [])
                            lines: List[str] = []
                            if citation_should_include_explanations and explanation_text:
                                lines.append(f"説明: {explanation_text}")
                            lines.extend(quotes_lines)
                            if lines:
                                blocks.append("\n".join(lines))
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


        if not any_translation:
            return f"No translatable text was found in range '{cell_range}'."

        if include_context_columns and explanation_fallback_notes:
            messages.insert(0, "explanation_jp が不足していたセルに既定の説明文を補いました: " + " / ".join(explanation_fallback_notes))

        write_messages: List[str] = []

        if range_adjustment_note:
            write_messages.append(range_adjustment_note)
        if citation_note:
            write_messages.append(citation_note)

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
        reference_ranges=None,
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
    reference_ranges: Optional[List[str]] = None,
    reference_urls: Optional[List[str]] = None,
    translation_output_range: Optional[str] = None,
    citation_output_range: Optional[str] = None,
    overwrite_source: bool = False,
    stop_event: Optional[Event] = None,
) -> str:
    """Translate ranges while consulting the supplied references per cell."""
    if not reference_ranges and not reference_urls:
        raise ToolExecutionError(
            "Either reference_ranges or reference_urls must be provided when using translate_range_with_references."
        )

    return translate_range_contents(
        actions=actions,
        browser_manager=browser_manager,
        cell_range=cell_range,
        target_language=target_language,
        sheet_name=sheet_name,
        reference_ranges=reference_ranges,
        citation_output_range=citation_output_range,
        reference_urls=reference_urls,
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

        reference_ranges: Optional list of ranges containing reference material.

        citation_output_range: Optional range used to store citation markers.

        reference_urls: Optional list of reference URLs to include in the output.

        translation_output_range: Optional range for translated rows (three columns per source column).

        overwrite_source: Whether to overwrite the source range directly.

    """
    try:
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

        _ensure_not_stopped()

        status_matrix = [["" for _ in range(src_cols)] for _ in range(src_rows)]
        issue_matrix = [["" for _ in range(src_cols)] for _ in range(src_rows)]
        highlight_matrix = [] if highlight_output_range else None
        highlight_styles = [] if highlight_output_range else None

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
            actions.write_range(status_row_ref, [status_matrix[row_idx]], status_sheet_name)
            issue_row_ref = _row_reference(issue_start_row, issue_start_col, row_idx, row_width)
            actions.write_range(issue_row_ref, [issue_matrix[row_idx]], issue_sheet_name)
            if highlight_matrix is not None and highlight_start_row is not None and highlight_start_col is not None:
                highlight_row_ref = _row_reference(highlight_start_row, highlight_start_col, row_idx, row_width)
                actions.write_range(highlight_row_ref, [highlight_matrix[row_idx]], highlight_sheet_name)
                if highlight_styles is not None:
                    actions.apply_diff_highlight_colors(
                        highlight_row_ref,
                        [highlight_styles[row_idx]],
                        highlight_sheet_name,
                        addition_color_hex="#1565C0",
                        deletion_color_hex="#C62828",
                    )
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
                "Respond with a JSON array containing exactly one object. Provide: 'id', 'status', 'notes', 'corrected_text', and 'highlighted_text'.\n"
                "Optionally include 'before_text', 'after_text', or an 'edits' array (each element with fields 'type', 'text', and 'reason').\n"
                "Use status 'OK' only when the draft translation requires no changes; otherwise respond with 'REVISE'.\n"
                "Write 'notes' in Japanese using the exact pattern 'Issue: ... / Suggestion: ...'. Keep them concise and actionable.\n"
                "Set 'corrected_text' to the fully corrected English sentence. For status 'OK', repeat the original translation unchanged.\n"
                "Populate 'highlighted_text' to show the difference versus the current translation: wrap deletions in '[DEL]...'/wrap additions in '[ADD]...'. Leave it empty for status 'OK'.\n"
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

                decoder = json.JSONDecoder()
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
                    "Reply with a single JSON array containing exactly one object. Each object must contain 'id', 'status', 'notes', "
                    "'highlighted_text', 'corrected_text', 'before_text', and 'after_text'. "
                    "Use status 'OK' when the translation is acceptable (notes empty or a short remark). Only select 'OK' when you are certain there are no issues. "
                    "Use status 'REVISE' when changes are needed and write notes in Japanese as 'Issue: ... / Suggestion: ...'. If unsure, choose 'REVISE'. "
                    "Set 'corrected_text' to the fully corrected English sentence. Build 'highlighted_text' from corrected_text, "
                    "marking additions as [ADD ...] and deletions as [DEL ...]. "
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
                        if 0 <= numeric_index < len(ordered_positions) and numeric_index not in assigned_indices:
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
                        if isinstance(ai_highlight_raw, str) and ("[DEL]" in ai_highlight_raw or "[ADD]" in ai_highlight_raw):
                            parsed_text = _maybe_fix_mojibake(ai_highlight_raw)
                            highlight_text, highlight_spans = _parse_highlight_markup(parsed_text)
                            if not highlight_spans:
                                highlight_text, highlight_spans = _build_diff_highlight(sanitized_base_text, corrected_text_str)
                        else:
                            highlight_text, highlight_spans = _build_diff_highlight(sanitized_base_text, corrected_text_str)
                        highlight_text = _maybe_fix_mojibake(highlight_text)
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
                actions.write_range(highlight_output_range, highlight_matrix, sheet_name)
                if highlight_styles is not None:
                    _ensure_not_stopped()
                    actions.apply_diff_highlight_colors(
                        highlight_output_range,
                        highlight_styles,
                        sheet_name,
                        addition_color_hex="#1565C0",
                        deletion_color_hex="#C62828",
                    )
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
