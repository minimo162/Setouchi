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
            _diff_debug(
                f"apply_diff_highlight_colors target_range rows={rows} cols={cols} addition={addition_color_hex} deletion={deletion_color_hex}"
            )

            api_char_supported = True
            characters_char_supported = True

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
                    try:
                        cell.api.Font.ColorIndex = 0
                    except Exception as reset_error:
                        _diff_debug(f"apply_diff_highlight_colors cell({r_idx},{c_idx}) reset color failed {reset_error}")
                    if not spans:
                        continue
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

                        start_position = start_idx + 1
                        length_val = min(length_val, value_len - start_idx)
                        if length_val <= 0:
                            _diff_debug(
                                f"apply_diff_highlight_colors cell({r_idx},{c_idx}) computed non-positive length after clamp start={start_idx} value_len={value_len}"
                            )
                            continue

                        span_applied = False
                        try:
                            cell.api.Characters(start_position, length_val).Font.Color = color_value
                            span_applied = True
                        except Exception as primary_error:
                            _diff_debug(
                                f"apply_diff_highlight_colors cell({r_idx},{c_idx}) api span color error {primary_error}"
                            )

                        if not span_applied:
                            try:
                                cell.characters[start_position - 1, length_val].font.color = color_tuple
                                span_applied = True
                            except Exception as span_fallback_error:
                                _diff_debug(
                                    f"apply_diff_highlight_colors cell({r_idx},{c_idx}) characters span color error {span_fallback_error}"
                                )

                        applied = True
                        for char_offset in range(length_val):
                            char_position = start_position + char_offset
                            colored = False
                            if api_char_supported:
                                try:
                                    cell.api.Characters(char_position, 1).Font.Color = color_value
                                    colored = True
                                except Exception as per_char_error:
                                    api_char_supported = False
                                    _diff_debug(
                                        f"apply_diff_highlight_colors cell({r_idx},{c_idx}) api single char error {per_char_error}"
                                    )
                            if not colored and characters_char_supported:
                                try:
                                    cell.characters[char_position - 1].font.color = color_tuple
                                    colored = True
                                except Exception as per_char_fallback_error:
                                    characters_char_supported = False
                                    _diff_debug(
                                        f"apply_diff_highlight_colors cell({r_idx},{c_idx}) characters single char error {per_char_fallback_error}"
                                    )
                            if not colored:
                                applied = False
                                _diff_debug(
                                    f"apply_diff_highlight_colors cell({r_idx},{c_idx}) unable to color char at {char_position}"
                                )
                                break

                        if not span_applied and not applied:
                            continue

                        if not applied:
                            try:
                                cell.characters[start_position - 1, length_val].font.color = color_tuple
                                applied = True
                            except Exception as span_fallback_error:
                                _diff_debug(
                                    f"apply_diff_highlight_colors cell({r_idx},{c_idx}) characters span color error {span_fallback_error}"
                                )

                        if not applied:
                            applied = True
                            for char_offset in range(length_val):
                                char_position = start_position + char_offset
                                colored = False
                                if api_char_supported:
                                    try:
                                        cell.api.Characters(char_position, 1).Font.Color = color_value
                                        colored = True
                                    except Exception as per_char_error:
                                        api_char_supported = False
                                        _diff_debug(
                                            f"apply_diff_highlight_colors cell({r_idx},{c_idx}) api single char error {per_char_error}"
                                        )
                                if not colored and characters_char_supported:
                                    try:
                                        cell.characters[char_position - 1].font.color = color_tuple
                                        colored = True
                                    except Exception as per_char_fallback_error:
                                        characters_char_supported = False
                                        _diff_debug(
                                            f"apply_diff_highlight_colors cell({r_idx},{c_idx}) characters single char error {per_char_fallback_error}"
                                        )
                                if not colored:
                                    applied = False
                                    _diff_debug(
                                        f"apply_diff_highlight_colors cell({r_idx},{c_idx}) unable to color char at {char_position}"
                                    )
                                    break

                        if not applied:
                            continue

                        color_hex = addition_color_hex if color_kind == "addition" else deletion_color_hex
                        _diff_debug(
                            f"apply_diff_highlight_colors applied type={span_type} kind={color_kind} start={start_idx} length={length_val} color={color_hex}"
                        )

        except Exception as e:
            _diff_debug(f"apply_diff_highlight_colors exception={e}")
            raise ToolExecutionError(f"差分ハイライトの色適用中にエラーが発生しました: {e}") from e
