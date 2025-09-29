from pathlib import Path

path = Path('excel_copilot/tools/actions.py')
lines = path.read_text(encoding='utf-8').splitlines()

start = None
for idx, line in enumerate(lines):
    if line.strip().startswith('try:') and 'color_value = _color_tuple_to_bgr_int' in lines[idx+1]:
        start = idx
        break

if start is None:
    raise SystemExit('combined try block not found')

# find end (line with log _diff_debug) etc
end = start
while end < len(lines) and '_diff_debug(' not in lines[end].strip():
    end += 1
while end < len(lines) and not lines[end].strip().startswith(')'):
    end += 1
end += 1  # include closing line

# we will rebuild manually for clarity
replacement = [
    '                        try:',
    '                            color_value = _color_tuple_to_bgr_int(color_tuple)',
    '                            success = True',
    '                            for char_offset in range(length_val):',
    '                                try:',
    '                                    char_obj = cell.characters[start_idx + char_offset + 1]',
    '                                    char_obj.api.Font.Color = color_value',
    '                                except Exception as char_error:',
    '                                    _diff_debug(',
    '                                        f"apply_diff_highlight_colors char error at offset {char_offset}: {char_error}"',
    '                                    )',
    '                                    success = False',
    '                                    break',
    '                        except Exception as color_error:',
    '                            _diff_debug(',
    '                                f"apply_diff_highlight_colors failed type={span_type} start={start_idx} length={length_val} error={color_error}"',
    '                            )',
    '                            continue',
    '                        if not success:',
    '                            continue',
    '                        log_color = f"#{color_value:06X}"',
    '                        _diff_debug(',
    '                            f"apply_diff_highlight_colors applied type={span_type} kind={color_kind} start={start_idx} length={length_val} color={log_color}"',
    '                        )',
]

# Need to replace from start to line before 'if not success' maybe? We'll find old block boundaries explicitly
end = start
while end < len(lines) and not lines[end].strip().startswith('if not success:'):
    end += 1
# include rest until after log block (closing ) ).
while end < len(lines) and ') )' not in lines[end].replace(' ', ''):
    end += 1
end += 1

lines[start:end] = replacement
path.write_text('\n'.join(lines) + '\n', encoding='utf-8')
