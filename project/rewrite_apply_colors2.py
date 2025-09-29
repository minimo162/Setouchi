from pathlib import Path
import re

path = Path('excel_copilot/tools/actions.py')
text = path.read_text(encoding='utf-8')
pattern = re.compile(r"    def apply_diff_highlight_colors\(self,\s*\n(?:        .+\n)+?        except Exception as e:\n            raise ToolExecutionError\(f\\"差分ハイライトの色適用中にエラーが発生しました: \\{e\\}\\"\) from e\n", re.DOTALL)
match = pattern.search(text)
if not match:
    raise SystemExit('apply_diff_highlight_colors block not found')
new_block = """    def apply_diff_highlight_colors(self,
                                    cell_range: str,
                                    style_matrix: List[List[List[Dict[str, Any]]]],
                                    sheet_name: Optional[str] = None,
                                    addition_color_hex: str = \"#1565C0\",
                                    deletion_color_hex: str = \"#C62828\") -> None:
        \"\"\"Apply highlight colors to rich diff spans within the target range.\"\"\"
        try:
            if not style_matrix:
                _diff_debug('apply_diff_highlight_colors empty style matrix')
                return
            _diff_debug(
                f\"apply_diff_highlight_colors start range={cell_range} sheet={sheet_name} rows={len(style_matrix)}\"
            )
            sheet = self._get_sheet(sheet_name)
            target_range = sheet.range(cell_range)
            rows = target_range.rows.count
            cols = target_range.columns.count

            def _hex_to_color_value(hex_code: str) -> int:
                hex_code = hex_code.lstrip(\"#\")
                if len(hex_code) != 6:
                    raise ValueError(f\"Invalid hex color: {hex_code}\")
                r = int(hex_code[0:2], 16)
                g = int(hex_code[2:4], 16)
                b = int(hex_code[4:6], 16)
                return r | (g << 8) | (b << 16)

            addition_color = _hex_to_color_value(addition_color_hex)
            deletion_color = _hex_to_color_value(deletion_color_hex)
            _diff_debug(
                f\"apply_diff_highlight_colors target_range rows={rows} cols={cols} addition={addition_color_hex} deletion={deletion_color_hex}\"
            )

            max_rows = min(rows, len(style_matrix))
            for r_idx in range(max_rows):
                row_styles = style_matrix[r_idx]
                if not isinstance(row_styles, list):
                    _diff_debug(f\"apply_diff_highlight_colors row {r_idx} invalid styles entry\")
                    continue
                max_cols = min(cols, len(row_styles))
                for c_idx in range(max_cols):
                    spans = row_styles[c_idx] if isinstance(row_styles[c_idx], list) else []
                    if not spans:
                        _diff_debug(f\"apply_diff_highlight_colors cell({r_idx},{c_idx}) no spans\")
                        continue
                    cell = target_range[r_idx, c_idx]
                    try:
                        cell_value = cell.value
                    except Exception as value_error:
                        _diff_debug(f\"apply_diff_highlight_colors cell({r_idx},{c_idx}) value error {value_error}\")
                        continue
                    if not isinstance(cell_value, str):
                        _diff_debug(f\"apply_diff_highlight_colors cell({r_idx},{c_idx}) skipped non-str value\")
                        continue
                    _diff_debug(
                        f\"apply_diff_highlight_colors cell({r_idx},{c_idx}) value_len={len(cell_value)} spans={spans}\"
                    )
                    for span in spans:
                        if not isinstance(span, dict):
                            _diff_debug(f\"apply_diff_highlight_colors cell({r_idx},{c_idx}) span invalid {span}\")
                            continue
                        start = span.get('start')
                        length = span.get('length')
                        span_type = span.get('type')
                        if start is None or length is None or length <= 0:
                            _diff_debug(f\"apply_diff_highlight_colors cell({r_idx},{c_idx}) invalid bounds {span}\")
                            continue
                        if span_type == '追加':
                            color_value = addition_color
                        elif span_type == '削除':
                            color_value = deletion_color
                        else:
                            color_value = None
                        if color_value is None:
                            _diff_debug(f\"apply_diff_highlight_colors cell({r_idx},{c_idx}) unknown type {span_type}\")
                            continue
                        try:
                            characters = cell.api.Characters(int(start) + 1, int(length))
                            characters.Font.Color = color_value
                            _diff_debug(
                                f\"apply_diff_highlight_colors applied type={span_type} start={start} length={length} color=#{color_value:06X}\"
                            )
                        except Exception as color_error:
                            _diff_debug(
                                f\"apply_diff_highlight_colors failed type={span_type} start={start} length={length} error={color_error}\"
                            )
                            continue
        except Exception as e:
            _diff_debug(f\"apply_diff_highlight_colors exception={e}\")
            raise ToolExecutionError(f\"差分ハイライトの色適用中にエラーが発生しました: {e}\") from e
"""
text = text[:match.start()] + new_block + text[match.end():]
path.write_text(text, encoding='utf-8')
