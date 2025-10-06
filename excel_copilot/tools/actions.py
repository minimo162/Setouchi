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

_MAX_RICH_TEXT_LENGTH = int(os.getenv('EXCEL_COPILOT_MAX_RICH_TEXT_LENGTH', '800'))
_MAX_RICH_TEXT_SPANS = int(os.getenv('EXCEL_COPILOT_MAX_RICH_TEXT_SPANS', '48'))
_MAX_RICH_TEXT_TOTAL_SPAN_LENGTH = int(os.getenv('EXCEL_COPILOT_MAX_RICH_TEXT_TOTAL', '1200'))
_MAX_RICH_TEXT_LINE_BREAKS = int(os.getenv('EXCEL_COPILOT_MAX_RICH_TEXT_LINE_BREAKS', '12'))


def _diff_debug(message: str) -> None:
    if _DIFF_DEBUG_ENABLED:
        _ACTIONS_LOGGER.debug(message)


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

    def consume_progress_messages(self) -> List[str]:
        messages = self._progress_buffer[:]
        self._progress_buffer.clear()
        return messages

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

    def write_range(self, cell_range: str, data: List[List[Any]], sheet_name: Optional[str] = None) -> str:
        """指定された範囲にデータを書き込む前に、データの次元が範囲と一致するかを検証する。"""
        try:
            sheet = self._get_sheet(sheet_name)
            target_range = sheet.range(cell_range)

            if not isinstance(data, list) or (data and not isinstance(data[0], list)):
                 raise ToolExecutionError("書き込むデータは2次元リストである必要があります。")
            
            data_rows = len(data)
            data_cols = len(data[0]) if data_rows > 0 else 0

            range_rows = target_range.rows.count
            range_cols = target_range.columns.count

            if data_rows != range_rows or data_cols != range_cols:
                error_msg = (
                    f"データの次元が一致しません。書き込み先範囲 ({cell_range}) は {range_rows}行 x {range_cols}列ですが、"
                    f"提供されたデータは {data_rows}行 x {data_cols}列です。読み取ったデータと同じ次元のデータを渡してください。"
                )
                raise ToolExecutionError(error_msg)

            target_range.value = data
            self._apply_text_wrapping(target_range)
            return f"範囲 '{cell_range}' にデータを正常に書き込みました。"
        except Exception as e:
            if isinstance(e, ToolExecutionError):
                raise e
            raise ToolExecutionError(f"範囲 '{cell_range}' への書き込み中に予期せぬエラーが発生しました: {e}")

    def _apply_text_wrapping(self, target_range: xw.Range) -> None:
        """Turn on wrapping, align to top-left, and auto-fit within sensible bounds."""

        try:
            target_range.api.HorizontalAlignment = -4131  # xlLeft
            target_range.api.VerticalAlignment = -4160  # xlTop
        except Exception:
            try:
                for cell in target_range.cells:
                    try:
                        cell.api.HorizontalAlignment = -4131
                        cell.api.VerticalAlignment = -4160
                    except Exception:
                        continue
            except Exception:
                pass

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
                try:
                    column.column_width = desired
                except Exception:
                    try:
                        column.api.ColumnWidth = desired
                    except Exception:
                        continue
        except Exception:
            pass

        try:
            target_range.rows.autofit()
        except Exception:
            try:
                target_range.api.EntireRow.AutoFit()
            except Exception:
                pass

        try:
            row_iterable = list(target_range.rows)
        except Exception:
            row_iterable = []

        def _effective_width(cell: xw.Range) -> float:
            try:
                width_value = cell.column_width
            except Exception:
                width_value = None
            if width_value is None or width_value <= 0:
                width_value = self._preferred_column_width
            return max(self._column_width_floor, min(width_value, self._column_width_cap))

        for row in row_iterable:
            try:
                cells_iterable = list(row.cells)
            except Exception:
                cells_iterable = []

            max_lines = 1
            for cell in cells_iterable:
                try:
                    value = cell.value
                except Exception:
                    value = None
                if value is None:
                    continue
                text = str(value)
                if not text:
                    continue
                width_hint = _effective_width(cell)
                approx_lines = max(1, math.ceil(len(text) / max(1, width_hint - 2)))
                max_lines = max(max_lines, min(approx_lines, self._max_row_height // self._line_height))

            desired_height = max(self._min_row_height, min(self._max_row_height, max_lines * self._line_height))
            try:
                current_height = row.row_height
            except Exception:
                current_height = None

            if current_height is None or current_height < desired_height - 1:
                try:
                    row.row_height = desired_height
                except Exception:
                    try:
                        row.api.RowHeight = desired_height
                    except Exception:
                        continue

        try:
            target_range.wrap_text = True
        except Exception:
            pass

        try:
            target_range.api.WrapText = True
        except Exception:
            pass

        try:
            target_range.api.Cells.WrapText = True
        except Exception:
            pass

        try:
            target_range.api.wrap_text.set(True)
        except Exception:
            pass

        try:
            target_range.api.cells.wrap_text.set(True)
        except Exception:
            pass

        try:
            for column in target_range.columns:
                try:
                    column.wrap_text = True
                except Exception:
                    try:
                        column.api.WrapText = True
                    except Exception:
                        try:
                            column.api.wrap_text.set(True)
                        except Exception:
                            continue
        except Exception:
            pass

        try:
            sheet_obj = target_range.sheet
            address = target_range.address
            if sheet_obj is not None and address:
                try:
                    sheet_range = sheet_obj.range(address)
                    sheet_range.wrap_text = True
                except Exception:
                    pass
        except Exception:
            pass

        try:
            for cell in target_range.cells:
                try:
                    cell.wrap_text = True
                except Exception:
                    try:
                        cell.api.WrapText = True
                    except Exception:
                        try:
                            cell.api.wrap_text.set(True)
                        except Exception:
                            continue
        except Exception:
            pass

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
                return
            _diff_debug(
                f"apply_diff_highlight_colors start range={cell_range} sheet={sheet_name} rows={len(style_matrix)}"
            )
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
                    try:
                        cell_value = cell.value or ""
                    except Exception as value_error:
                        _diff_debug(f"apply_diff_highlight_colors cell({r_idx},{c_idx}) value error {value_error}")
                        continue
                    if not isinstance(cell_value, str):
                        _diff_debug(f"apply_diff_highlight_colors cell({r_idx},{c_idx}) skipped non-str value")
                        continue
                    value_len = len(cell_value)
                    _diff_debug(
                        f"apply_diff_highlight_colors cell({r_idx},{c_idx}) value_len={value_len} spans={spans}"
                    )
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
                        try:
                            cell_address = cell.get_address(row_absolute=False, column_absolute=False, include_sheet=False)
                        except Exception:
                            cell_address = None
                        if not cell_address:
                            cell_address = f"{r_idx}:{c_idx}"
                        skipped_cells.append(cell_address)
                        continue
                    if value_len <= 0:
                        continue
                    try:
                        cell.api.Font.ColorIndex = 1
                        cell.api.Font.Color = 0
                    except Exception as reset_error:
                        _diff_debug(f"apply_diff_highlight_colors cell({r_idx},{c_idx}) reset color failed {reset_error}")
                    try:
                        entire_chars = cell.api.Characters()
                        entire_chars.Font.ColorIndex = 1
                        entire_chars.Font.Color = 0
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
                        start_position = seg_start + 1
                        applied = False
                        try:
                            char_range = cell.api.Characters(start_position, seg_length)
                            char_range.Font.Color = color_value
                            applied = True
                        except Exception as primary_error:
                            _diff_debug(
                                f"apply_diff_highlight_colors cell({r_idx},{c_idx}) api span color error {primary_error}"
                            )
                        try:
                            segment_font = cell.characters[start_position - 1, seg_length].font
                            segment_font.color = color_tuple
                            applied = True
                        except Exception as span_fallback_error:
                            if not applied:
                                _diff_debug(
                                    f"apply_diff_highlight_colors cell({r_idx},{c_idx}) characters span color error {span_fallback_error}"
                                )
                                continue
                        if applied:
                            if color_kind == "addition":
                                color_hex = addition_color_hex
                            elif color_kind == "deletion":
                                color_hex = deletion_color_hex
                            else:
                                color_hex = ""
                            if color_hex:
                                _diff_debug(
                                    f"apply_diff_highlight_colors applied span type={color_kind} start={seg_start} length={seg_length} color={color_hex}"
                                )

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
            raise ToolExecutionError(f"差分ハイライトの色適用中にエラーが発生しました: {e}") from e
