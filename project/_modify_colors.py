from pathlib import Path

path = Path('excel_copilot/tools/actions.py')
lines = path.read_text(encoding='utf-8').splitlines()

start = None
for idx, line in enumerate(lines):
    if 'characters = cell.characters[start_idx + 1, length_val]' in line:
        start = idx
        break

if start is None:
    raise SystemExit('target line not found')

# determine end of block (until after log line)
end = start
while end < len(lines) and 'continue' not in lines[end]:
    end += 1
end += 1  # include continue line

new_block = [
"                            color_value = _color_tuple_to_bgr_int(color_tuple)",
"                            success = True",
"                            for char_offset in range(length_val):",
"                                try:",
"                                    char_obj = cell.characters[start_idx + char_offset + 1]",
"                                    char_obj.api.Font.Color = color_value",
"                                except Exception as char_error:",
"                                    _diff_debug(",
"                                        f\"apply_diff_highlight_colors char error at offset {char_offset}: {char_error}\"",
"                                    )",
"                                    success = False",
"                                    break",
"                            if success:",
"                                log_color = f\"#{color_value:06X}\"",
"                                _diff_debug(",
"                                    f\"apply_diff_highlight_colors applied type={span_type} kind={color_kind} start={start_idx} length={length_val} color={log_color}\"",
"                                )",
"                            else:",
"                                continue",
]

lines[start:end] = new_block
path.write_text('\n'.join(lines) + '\n', encoding='utf-8')
