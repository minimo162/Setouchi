import xlwings as xw
import sys
import subprocess
import logging
import os
import math
from typing import Any, List, Optional, Dict, Tuple, Callable
from ..core.exceptions import ToolExecutionError
from ..core.excel_manager import ExcelManager

_ACTIONS_LOGGER = logging.getLogger(__name__)
_DIFF_DEBUG_ENABLED = os.getenv('EXCEL_COPILOT_DEBUG_DIFF', '').lower() in {'1', 'true', 'yes'}

if _DIFF_DEBUG_ENABLED and not logging.getLogger().handlers:
    logging.basicConfig(level=logging.DEBUG)

_EXCEL_MAX_CELL_CHARS = 32766


def _read_int_env(name: str, default: int, minimum: int = 0, maximum: Optional[int] = None) -> int:
    raw_value = os.getenv(name)
    if raw_value is None or str(raw_value).strip() == '':
        value = default
    else:
        try:
            value = int(float(raw_value))
        except Exception:
            value = default
    if value < minimum:
        value = minimum
    if maximum is not None and value > maximum:
        value = maximum
    return value


_MAX_RICH_TEXT_LENGTH = _read_int_env(
    'EXCEL_COPILOT_MAX_RICH_TEXT_LENGTH',
    24000,
    minimum=1,
    maximum=_EXCEL_MAX_CELL_CHARS,
)
_MAX_RICH_TEXT_SPANS = _read_int_env('EXCEL_COPILOT_MAX_RICH_TEXT_SPANS', 192, minimum=1)
_MAX_RICH_TEXT_TOTAL_SPAN_LENGTH = _read_int_env(
    'EXCEL_COPILOT_MAX_RICH_TEXT_TOTAL',
    max(6400, _MAX_RICH_TEXT_LENGTH * 3),
    minimum=_MAX_RICH_TEXT_LENGTH,
    maximum=_EXCEL_MAX_CELL_CHARS * 3,
)
_MAX_RICH_TEXT_LINE_BREAKS = _read_int_env('EXCEL_COPILOT_MAX_RICH_TEXT_LINE_BREAKS', 256, minimum=0, maximum=2048)
_ENABLE_RICH_DIFF_COLORS = os.getenv('EXCEL_COPILOT_ENABLE_RICH_DIFF_COLORS', '1').lower() in {'1', 'true', 'yes', 'on'}
_FORCE_RICH_DIFF_COLORS = os.getenv('EXCEL_COPILOT_FORCE_RICH_DIFF_COLORS', '').lower() in {'1', 'true', 'yes', 'on'}


_MAX_CHARWISE_DIFF_SPAN = 400


_REVIEW_DEBUG_ENABLED = os.getenv('EXCEL_COPILOT_DEBUG_REVIEW', '').lower() in {'1', 'true', 'yes', 'on'}



def _review_debug(message: str) -> None:
    if not _REVIEW_DEBUG_ENABLED:
        return
    try:
        print(f"[review-debug] {message}")
    except Exception:
        pass


def _diff_debug(message: str) -> None:
    if _DIFF_DEBUG_ENABLED:
        _ACTIONS_LOGGER.debug(message)


def _safe_cell_address(cell: xw.Range, row_idx: int, col_idx: int) -> Optional[str]:
    """Return a best-effort A1-style address for diagnostics."""
    try:
        address = cell.get_address(row_absolute=False, column_absolute=False, include_sheet=False)
        if address:
            return address
    except Exception:
        pass
    try:
        row_number = getattr(cell, 'row', None)
        column_number = getattr(cell, 'column', None)
        if isinstance(row_number, int) and isinstance(column_number, int):
            return f"{xw.utils.col_name(column_number)}{row_number}"
    except Exception:
        pass
    try:
        return f"{xw.utils.col_name(col_idx + 1)}{row_idx + 1}"
    except Exception:
        pass
    return None


def _shorten_debug(value: Any, limit: int = 120) -> str:
    if value is None:
        return ''
    text_value = str(value).replace('\r', '\r').replace('\n', '\n')
    return text_value if len(text_value) <= limit else text_value[:limit] + '…'


class ExcelActions:
    """
    具体的なExcel操作を実行するメソッドを集約したクラス。
    """

    def __init__(self, manager: ExcelManager, progress_callback: Optional[Callable[[str], None]] = None):
        if not manager or not manager.get_active_workbook():
            raise ValueError("有効なExcelManagerインスタンスが必要です。")
        self.book = manager.get_active_workbook()
        self._progress_callback = progress_callback
        self._progress_buffer: List[str] = []
        self._column_width_cap = 90  # characters
        self._column_width_floor = 18  # characters
        self._preferred_column_width = 60  # characters
        self._min_row_height = 16
        self._line_height = 14
        self._max_row_height = 360
        self._diff_colors_platform_ready = _FORCE_RICH_DIFF_COLORS or sys.platform.lower() != 'darwin'
        self._diff_colors_supported = _ENABLE_RICH_DIFF_COLORS and self._diff_colors_platform_ready
        self._diff_color_warning_sent = False

    def set_progress_callback(self, callback: Optional[Callable[[str], None]]) -> None:
        self._progress_callback = callback

    def log_progress(self, message: str) -> None:
        if message is None:
            return
        self._progress_buffer.append(message)
        if self._progress_callback:
            try:
                self._progress_callback(message)
            except Exception as callback_error:
                _ACTIONS_LOGGER.debug(f"log_progress callback error: {callback_error}")
        else:
            try:
                print(message)
            except Exception:
                pass

    def consume_progress_messages(self) -> List[str]:
        messages = self._progress_buffer[:]
        self._progress_buffer.clear()
        return messages

    def supports_diff_highlight_colors(self) -> bool:
        """Return True if rich diff coloring is supported on this platform."""
        return self._diff_colors_supported

    def notify_diff_colors_unavailable(self) -> None:
        """Log a warning once when diff coloring is not supported."""
        if self._diff_colors_platform_ready or self._diff_color_warning_sent:
            return
        self.log_progress(
            "Diff coloring skipped: この環境のExcelはセル内の部分的な文字色変更をサポートしていません。テキストマーカーを残します。"
        )
        self._diff_color_warning_sent = True

    def _get_sheet(self, sheet_name: Optional[str] = None) -> xw.Sheet:
        try:
            return self.book.sheets[sheet_name] if sheet_name else self.book.sheets.active
        except Exception as e:
            raise ToolExecutionError(f"シート '{sheet_name or 'アクティブ'}' の取得に失敗: {e}")

    def write_to_cell(self, cell: str, value: Any, sheet_name: Optional[str] = None) -> str:
        try:
            sheet = self._get_sheet(sheet_name)
            target = sheet.range(cell)
            target.value = value
            self._apply_text_wrapping(target)
            return f"セル {cell} に値 '{value}' を正常に書き込みました。"
        except Exception as e:
            raise ToolExecutionError(f"セル '{cell}' への書き込み中にエラーが発生しました: {e}")

    def read_cell_value(self, cell: str, sheet_name: Optional[str] = None) -> Any:
        try:
            sheet = self._get_sheet(sheet_name)
            value = sheet.range(cell).value
            return f"セル '{cell}' の値は '{value}' です。"
        except Exception as e:
            raise ToolExecutionError(f"セル '{cell}' の読み取り中にエラーが発生しました: {e}")

    def get_sheet_names(self) -> List[str]:
        try:
            return [sheet.name for sheet in self.book.sheets]
        except Exception as e:
            raise ToolExecutionError(f"シート名の取得中にエラーが発生しました: {e}")

    def set_formula(self, cell: str, formula: str, sheet_name: Optional[str] = None) -> str:
        try:
            sheet = self._get_sheet(sheet_name)
            sheet.range(cell).formula = formula
            return f"セル {cell} に数式 '{formula}' を正常に設定しました。"
        except Exception as e:
            raise ToolExecutionError(f"数式 '{formula}' の設定中にエラーが発生しました: {e}")

    def copy_range(self, source_range: str, destination_range: str, sheet_name: Optional[str] = None) -> str:
        try:
            sheet = self._get_sheet(sheet_name)
            sheet.range(source_range).copy(sheet.range(destination_range))
            return f"範囲 '{source_range}' を '{destination_range}' に正常にコピーしました。"
        except Exception as e:
            raise ToolExecutionError(f"範囲のコピー中にエラーが発生しました: {e}")

    def read_range(self, cell_range: str, sheet_name: Optional[str] = None) -> List[List[Any]]:
        try:
            sheet = self._get_sheet(sheet_name)
            return sheet.range(cell_range).value
        except Exception as e:
            raise ToolExecutionError(f"範囲 '{cell_range}' の読み取り中にエラーが発生しました: {e}")

    def write_range(
        self,
        cell_range: str,
        data: List[List[Any]],
        sheet_name: Optional[str] = None,
        apply_formatting: bool = True,
    ) -> str:
        """指定された範囲にデータを書き込む前に次元を検証し、必要に応じて整形を適用する。"""
        try:
            sheet = self._get_sheet(sheet_name)
            target_range = sheet.range(cell_range)

            if not isinstance(data, list) or (data and not isinstance(data[0], list)):
                raise ToolExecutionError("書き込むデータは2次元リストである必要があります。")

            data_rows = len(data)
            data_cols = len(data[0]) if data_rows > 0 else 0
            _review_debug(f"actions.write_range start range={cell_range} rows={data_rows} cols={data_cols}")

            range_rows = target_range.rows.count
            range_cols = target_range.columns.count
            _review_debug(f"actions.write_range target dims rows={range_rows} cols={range_cols}")

            if data_rows != range_rows or data_cols != range_cols:
                error_msg = (
                    f"データの次元が一致しません。書き込み先範囲 ({cell_range}) は {range_rows}行 x {range_cols}列ですが、"
                    f"提供されたデータは {data_rows}行 x {data_cols}列です。読み取ったデータと同じ次元のデータを渡してください。"
                )
                _review_debug(f"actions.write_range dimension mismatch: {error_msg}")
                raise ToolExecutionError(error_msg)

            target_range.value = data
            _review_debug(f"actions.write_range wrote values range={cell_range}")
            if apply_formatting:
                self._apply_text_wrapping(target_range)
                _review_debug(f"actions.write_range completed range={cell_range} formatted=True")
            else:
                _review_debug(f"actions.write_range completed range={cell_range} formatted=False")
            return f"範囲 '{cell_range}' にデータを正常に書き込みました。"
        except Exception as e:
            _review_debug(f"actions.write_range error range={cell_range} error={e}")
            if isinstance(e, ToolExecutionError):
                raise e
            raise ToolExecutionError(f"範囲 '{cell_range}' への書き込み中に予期せぬエラーが発生しました: {e}")


    def _apply_text_wrapping(self, target_range: xw.Range) -> None:
        """Turn on wrapping, align to top-left, and auto-fit within sensible bounds."""

        range_address = None
        try:
            range_address = target_range.address
        except Exception:
            pass
        _review_debug(f"_apply_text_wrapping start address={range_address}")

        def _safe_call(callable_obj, *args, **kwargs):
            try:
                callable_obj(*args, **kwargs)
                return True
            except Exception:
                return False

        target_api = getattr(target_range, 'api', None)
        if target_api is not None:
            _safe_call(setattr, target_api, 'HorizontalAlignment', -4131)
            _safe_call(setattr, target_api, 'VerticalAlignment', -4160)
        _safe_call(setattr, target_range, 'horizontal_alignment', 'left')
        _safe_call(setattr, target_range, 'vertical_alignment', 'top')

        wrap_applied = False
        if hasattr(target_range, 'wrap_text'):
            wrap_applied = _safe_call(setattr, target_range, 'wrap_text', True)
        if not wrap_applied and target_api is not None:
            wrap_applied = _safe_call(setattr, target_api, 'WrapText', True)
            if not wrap_applied:
                cells_api = getattr(target_api, 'Cells', None)
                if cells_api is not None:
                    wrap_applied = _safe_call(setattr, cells_api, 'WrapText', True)

        # Restore moderate column width adjustments.
        try:
            for column in target_range.columns:
                try:
                    existing_width = column.column_width
                except Exception:
                    existing_width = None
                desired = self._preferred_column_width
                if existing_width is not None and existing_width > 0:
                    desired = max(existing_width, self._preferred_column_width)
                desired = max(self._column_width_floor, min(desired, self._column_width_cap))
                if not _safe_call(setattr, column, 'column_width', desired):
                    column_api = getattr(column, 'api', None)
                    if column_api is not None:
                        _safe_call(setattr, column_api, 'ColumnWidth', desired)
        except Exception:
            pass

        # Restore moderate row height adjustments.
        try:
            approx_lines = 1
            values = target_range.value
            if isinstance(values, list) and values and isinstance(values[0], list):
                flat_values = [str(cell or '') for row in values for cell in row]
            else:
                flat_values = [str(values or '')]
            width_chars = max(1, self._preferred_column_width - 2)
            max_lines_cap = max(1, self._max_row_height // self._line_height)
            for text in flat_values:
                if not text:
                    continue
                normalized = text.replace('\r\n', '\n').replace('\r', '\n')
                segments = normalized.split('\n')
                total_lines = 0
                for segment in segments:
                    segment_len = len(segment)
                    wrapped_lines = 1 + ((segment_len - 1) // width_chars) if segment_len else 1
                    total_lines += wrapped_lines
                approx_lines = max(approx_lines, min(total_lines, max_lines_cap))
            desired_height = max(self._min_row_height, min(self._max_row_height, approx_lines * self._line_height))
            _safe_call(setattr, target_range, 'row_height', desired_height)
            target_api = getattr(target_range, 'api', None)
            if target_api is not None:
                _safe_call(setattr, target_api, 'RowHeight', desired_height)
        except Exception:
            pass

        _review_debug(f"_apply_text_wrapping end address={range_address} wrap_applied={wrap_applied}")

    def apply_diff_highlight_colors(self,
                                    cell_range: str,
                                    style_matrix: List[List[List[Dict[str, Any]]]],
                                    sheet_name: Optional[str] = None,
                                    addition_color_hex: str = "#1565C0",
                                    deletion_color_hex: str = "#C62828") -> None:
        """Apply highlight colors to rich diff spans within the target range."""
        try:
            if not style_matrix:
                _diff_debug("apply_diff_highlight_colors empty style matrix")
                _review_debug(f"apply_diff_highlight_colors start range={cell_range} skipped (empty style matrix)")
                return
            if not _ENABLE_RICH_DIFF_COLORS:
                _diff_debug('apply_diff_highlight_colors disabled via configuration')
                self.log_progress(
                    "Skipped diff coloring; EXCEL_COPILOT_ENABLE_RICH_DIFF_COLORS is disabled. Text markers remain."
                )
                return
            if not self._diff_colors_platform_ready:
                has_spans = any(
                    isinstance(row_styles, list) and any(isinstance(cell_spans, list) and cell_spans for cell_spans in row_styles)
                    for row_styles in style_matrix
                )
                if has_spans:
                    self.notify_diff_colors_unavailable()
                _diff_debug("apply_diff_highlight_colors skipped: diff colors unsupported on this platform")
                _review_debug(f"apply_diff_highlight_colors skipped range={cell_range} unsupported platform")
                return
            _diff_debug(
                f"apply_diff_highlight_colors start range={cell_range} sheet={sheet_name} rows={len(style_matrix)}"
            )
            _review_debug(f"apply_diff_highlight_colors start range={cell_range} sheet={sheet_name} rows={len(style_matrix)}")

            sheet = self._get_sheet(sheet_name)
            target_range = sheet.range(cell_range)
            rows = target_range.rows.count
            cols = target_range.columns.count

            def _hex_to_color_tuple(hex_code: str) -> Tuple[int, int, int]:
                hex_code = hex_code.lstrip("#")
                if len(hex_code) != 6:
                    raise ValueError(f"Invalid hex color: {hex_code}")
                r = int(hex_code[0:2], 16)
                g = int(hex_code[2:4], 16)
                b = int(hex_code[4:6], 16)
                return (r, g, b)

            def _color_tuple_to_bgr_int(color_tuple: Tuple[int, int, int]) -> int:
                r, g, b = color_tuple
                return (b << 16) | (g << 8) | r

            addition_color_tuple = _hex_to_color_tuple(addition_color_hex)
            deletion_color_tuple = _hex_to_color_tuple(deletion_color_hex)
            addition_color_value = _color_tuple_to_bgr_int(addition_color_tuple)
            deletion_color_value = _color_tuple_to_bgr_int(deletion_color_tuple)
            type_aliases = {
                "addition": "addition",
                "add": "addition",
                "added": "addition",
                "insert": "addition",
                "inserted": "addition",
                "追加": "addition",
                "deletion": "deletion",
                "delete": "deletion",
                "deleted": "deletion",
                "del": "deletion",
                "remove": "deletion",
                "removed": "deletion",
                "削除": "deletion",
            }

            def _safe_span_length(span: Dict[str, Any]) -> int:
                if not isinstance(span, dict):
                    return 0
                raw_length = span.get("length", 0)
                try:
                    return int(raw_length or 0)
                except Exception:
                    try:
                        return int(float(raw_length))
                    except Exception:
                        return 0

            _diff_debug(
                f"apply_diff_highlight_colors target_range rows={rows} cols={cols} addition={addition_color_hex} deletion={deletion_color_hex}"
            )

            skipped_cells: List[str] = []

            max_rows = min(rows, len(style_matrix))
            for r_idx in range(max_rows):
                row_styles = style_matrix[r_idx]
                if not isinstance(row_styles, list):
                    _diff_debug(f"apply_diff_highlight_colors row {r_idx} invalid styles entry")
                    continue
                max_cols = min(cols, len(row_styles))
                for c_idx in range(max_cols):
                    spans = row_styles[c_idx] if isinstance(row_styles[c_idx], list) else []
                    cell = target_range[r_idx, c_idx]
                    _review_debug(f"apply_diff_highlight_colors cell({r_idx},{c_idx}) spans={spans}")
                    try:
                        cell_value = cell.value or ""
                    except Exception as value_error:
                        _diff_debug(f"apply_diff_highlight_colors cell({r_idx},{c_idx}) value error {value_error}")
                        _review_debug(f"apply_diff_highlight_colors cell({r_idx},{c_idx}) value error {value_error}")
                        continue
                    if not isinstance(cell_value, str):
                        _diff_debug(f"apply_diff_highlight_colors cell({r_idx},{c_idx}) skipped non-str value")
                        _review_debug(f"apply_diff_highlight_colors cell({r_idx},{c_idx}) skipped non-str value")
                        continue
                    value_len = len(cell_value)
                    _diff_debug(
                        f"apply_diff_highlight_colors cell({r_idx},{c_idx}) value_len={value_len} spans={spans}"
                    )
                    _review_debug(f"apply_diff_highlight_colors cell({r_idx},{c_idx}) value_len={value_len}")
                    line_breaks = cell_value.count("\n")
                    total_span_length = sum(
                        max(0, _safe_span_length(span)) for span in spans if isinstance(span, dict)
                    )
                    if (
                        value_len > _MAX_RICH_TEXT_LENGTH
                        or len(spans) > _MAX_RICH_TEXT_SPANS
                        or total_span_length > _MAX_RICH_TEXT_TOTAL_SPAN_LENGTH
                        or line_breaks > _MAX_RICH_TEXT_LINE_BREAKS
                    ):
                        _diff_debug(
                            f"apply_diff_highlight_colors cell({r_idx},{c_idx}) skipped rich text due to limits len={value_len} spans={len(spans)} total_span={total_span_length} line_breaks={line_breaks}"
                        )
                        cell_address = _safe_cell_address(cell, r_idx, c_idx)
                        if cell_address:
                            skipped_cells.append(cell_address)
                        continue
                    if value_len <= 0:
                        continue
                    try:
                        cell.api.Font.ColorIndex = 1
                        cell.api.Font.Color = 0
                        cell.api.Font.Strikethrough = False
                    except Exception as reset_error:
                        _diff_debug(f"apply_diff_highlight_colors cell({r_idx},{c_idx}) reset color failed {reset_error}")
                    try:
                        entire_chars = cell.characters[:]
                        entire_chars_api = getattr(entire_chars, 'api', None)
                        if entire_chars_api is not None:
                            entire_chars_api.Font.ColorIndex = 1
                            entire_chars_api.Font.Color = 0
                            entire_chars_api.Font.Strikethrough = False
                        entire_chars.font.color = (0, 0, 0)
                        entire_chars.font.strikethrough = False
                    except Exception as chars_reset_err:
                        _diff_debug(f"apply_diff_highlight_colors cell({r_idx},{c_idx}) characters reset failed {chars_reset_err}")
                    if not spans:
                        continue

                    value_len = len(cell_value)
                    if value_len <= 0:
                        continue

                    color_map: List[Optional[Tuple[str, int, Tuple[int, int, int]]]] = [None] * value_len
                    valid_span_present = False

                    for span in spans:
                        if not isinstance(span, dict):
                            _diff_debug(f"apply_diff_highlight_colors cell({r_idx},{c_idx}) span invalid {span}")
                            continue
                        start = span.get("start")
                        length = span.get("length")
                        span_type = span.get("type")
                        if start is None or length is None:
                            _diff_debug(f"apply_diff_highlight_colors cell({r_idx},{c_idx}) missing bounds {span}")
                            continue
                        start_idx = max(int(start), 0)
                        length_val = int(length)
                        if length_val <= 0:
                            _diff_debug(f"apply_diff_highlight_colors cell({r_idx},{c_idx}) non-positive length {span}")
                            continue
                        if start_idx >= value_len:
                            _diff_debug(f"apply_diff_highlight_colors cell({r_idx},{c_idx}) start out of range {span}")
                            continue
                        max_available = value_len - start_idx
                        if length_val > max_available:
                            _diff_debug(
                                f"apply_diff_highlight_colors cell({r_idx},{c_idx}) length clamped from {length_val} to {max_available}"
                            )
                            length_val = max_available
                        if not isinstance(span_type, str):
                            _diff_debug(f"apply_diff_highlight_colors cell({r_idx},{c_idx}) span type not str {span_type}")
                            continue
                        span_type_key = span_type.strip().lower()
                        normalized_type = type_aliases.get(span_type_key, span_type_key)
                        if normalized_type == "addition":
                            color_tuple = addition_color_tuple
                            color_value = addition_color_value
                            color_kind = "addition"
                        elif normalized_type == "deletion":
                            color_tuple = deletion_color_tuple
                            color_value = deletion_color_value
                            color_kind = "deletion"
                        else:
                            _diff_debug(f"apply_diff_highlight_colors cell({r_idx},{c_idx}) unknown type {span_type}")
                            continue
                        span_end = start_idx + length_val
                        for idx in range(start_idx, span_end):
                            color_map[idx] = (color_kind, color_value, color_tuple)
                        valid_span_present = True

                    if not valid_span_present:
                        continue

                    segments: List[Tuple[int, int, Optional[Tuple[str, int, Tuple[int, int, int]]]]] = []
                    segment_start = 0
                    current_color = color_map[0]
                    for idx in range(1, value_len):
                        if color_map[idx] != current_color:
                            segments.append((segment_start, idx - segment_start, current_color))
                            segment_start = idx
                            current_color = color_map[idx]
                    segments.append((segment_start, value_len - segment_start, current_color))

                    for seg_start, seg_length, color_info in segments:
                        if seg_length <= 0:
                            continue
                        if color_info is None:
                            continue
                        color_kind, color_value, color_tuple = color_info
                        segment_start = seg_start
                        segment_end = seg_start + seg_length
                        applied = False
                        char_slice = None
                        slice_api = None
                        segment_font = None
                        try:
                            char_slice = cell.characters[segment_start:segment_end]
                            slice_api = getattr(char_slice, 'api', None)
                            segment_font = getattr(char_slice, 'font', None)
                        except Exception as slice_error:
                            _diff_debug(
                                f"apply_diff_highlight_colors cell({r_idx},{c_idx}) segment slice error {slice_error}"
                            )
                            continue
                        if slice_api is not None:
                            try:
                                slice_api.Font.Color = color_value
                                applied = True
                            except Exception as primary_error:
                                _diff_debug(
                                    f"apply_diff_highlight_colors cell({r_idx},{c_idx}) api span color error {primary_error}"
                                )
                        if segment_font is not None:
                            try:
                                segment_font.color = color_tuple
                                applied = True
                            except Exception as span_fallback_error:
                                if not applied:
                                    _diff_debug(
                                        f"apply_diff_highlight_colors cell({r_idx},{c_idx}) characters span color error {span_fallback_error}"
                                    )
                                    continue
                        if applied:
                            is_deletion = color_kind == "deletion"
                            if color_kind == "addition":
                                color_hex = addition_color_hex
                                color_index = 5  # blue
                            elif color_kind == "deletion":
                                color_hex = deletion_color_hex
                                color_index = 3  # red
                            else:
                                color_hex = ""
                                color_index = None
                            if color_hex:
                                _diff_debug(
                                    f"apply_diff_highlight_colors applied span type={color_kind} start={seg_start} length={seg_length} color={color_hex}"
                                )
                                _review_debug(
                                    f"apply_diff_highlight_colors span success cell=({r_idx},{c_idx}) kind={color_kind} start={seg_start} length={seg_length}"
                                )
                                block_colored = False
                                if slice_api is not None:
                                    try:
                                        if color_index is not None:
                                            slice_api.Font.ColorIndex = color_index
                                        slice_api.Font.Color = color_value
                                        slice_api.Font.Strikethrough = is_deletion
                                        block_colored = True
                                    except Exception as span_block_error:
                                        _review_debug(f"apply_diff_highlight_colors span block color failed: {span_block_error}")
                                if segment_font is not None:
                                    try:
                                        segment_font.color = color_tuple
                                        segment_font.strikethrough = is_deletion
                                    except Exception as segment_block_error:
                                        _review_debug(f"apply_diff_highlight_colors segment font assign failed: {segment_block_error}")
                                if not block_colored:
                                    charwise_success = True
                                    chunk_size = _MAX_CHARWISE_DIFF_SPAN
                                    for chunk_start in range(0, seg_length, chunk_size):
                                        chunk_length = min(chunk_size, seg_length - chunk_start)
                                        for offset in range(chunk_length):
                                            absolute_idx = segment_start + chunk_start + offset
                                            try:
                                                char_segment = cell.characters[absolute_idx:absolute_idx + 1]
                                                char_segment_api = getattr(char_segment, 'api', None)
                                                if char_segment_api is not None:
                                                    if color_index is not None:
                                                        char_segment_api.Font.ColorIndex = color_index
                                                    char_segment_api.Font.Color = color_value
                                                    char_segment_api.Font.Strikethrough = is_deletion
                                                char_segment_font = getattr(char_segment, 'font', None)
                                                if char_segment_font is not None:
                                                    char_segment_font.color = color_tuple
                                                    char_segment_font.strikethrough = is_deletion
                                            except Exception as char_error:
                                                charwise_success = False
                                                _review_debug(f"apply_diff_highlight_colors charwise error pos={absolute_idx + 1} err={char_error}")
                                                break
                                        if not charwise_success:
                                            break
                                    if seg_length > _MAX_CHARWISE_DIFF_SPAN and not charwise_success:
                                        _review_debug(f"apply_diff_highlight_colors charwise fallback skipped due to span length {seg_length}")

            if skipped_cells:
                unique_cells = list(dict.fromkeys(skipped_cells))
                preview = ", ".join(unique_cells[:5])
                message = f"Skipped diff highlighting for {len(skipped_cells)} cell(s) due to size limits"
                if preview:
                    if len(unique_cells) > 5:
                        message += f": {preview}, ..."
                    else:
                        message += f": {preview}"
                self.log_progress(message)
        except Exception as e:
            _diff_debug(f"apply_diff_highlight_colors exception={e}")
            _review_debug(f"apply_diff_highlight_colors exception range={cell_range} error={e}")
            raise ToolExecutionError(f"差分ハイライトの色適用中にエラーが発生しました: {e}") from e
        else:
            _review_debug(f"apply_diff_highlight_colors completed range={cell_range} skipped_cells={len(skipped_cells)}")
