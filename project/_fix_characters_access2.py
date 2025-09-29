from pathlib import Path

path = Path('excel_copilot/tools/actions.py')
lines = path.read_text(encoding='utf-8').splitlines()

old_line = "                            characters = cell.api.Characters(start_idx + 1, length_val)"
if old_line not in lines:
    raise SystemExit('old characters access line not found')
idx = lines.index(old_line)
lines[idx] = "                            characters = cell.characters[start_idx + 1, length_val]"
lines[idx + 1] = "                            # xlwings wrapper for character slice; still direct COM for color"

old_color_line = "                            characters.Font.Color = color_value"
if old_color_line not in lines:
    raise SystemExit('old color assignment line not found')
idx = lines.index(old_color_line)
lines[idx] = "                            characters.api.Font.Color = color_value"
lines[idx + 1] = "                            # use COM font to bypass tuple/int comparison"

path.write_text('\n'.join(lines) + '\n', encoding='utf-8')
