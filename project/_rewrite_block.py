from pathlib import Path

path = Path('excel_copilot/tools/actions.py')
lines = path.read_text(encoding='utf-8').splitlines()
start = 233
end = 256
replacement = [
    '                        color_value = _color_tuple_to_bgr_int(color_tuple)',
    '                        success = True',
    '                        for char_offset in range(length_val):',
    '                            try:',
    '                                char_obj = cell.characters[start_idx + char_offset + 1]',
    '                                char_obj.api.Font.Color = color_value',
    '                            except Exception as char_error:',
    '                                _diff_debug(',
    '                                    f"apply_diff_highlight_colors char error at offset {char_offset}: {char_error}"',
    '                                )',
    '                                success = False',
    '                                break',
    '                        if not success:',
    '                            continue',
    '                        log_color = f"#{color_value:06X}"',
    '                        _diff_debug(',
    '                            f"apply_diff_highlight_colors applied type={span_type} kind={color_kind} start={start_idx} length={length_val} color={log_color}"',
    '                        )',
]
lines[start:end] = replacement
path.write_text('\n'.join(lines) + '\n', encoding='utf-8')
