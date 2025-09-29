from pathlib import Path

path = Path('excel_copilot/tools/actions.py')
text = path.read_text(encoding='utf-8')
lines = text.splitlines()

target = "                            characters = cell.characters[start_idx + 1, length_val]"
if target not in lines:
    raise SystemExit('target characters access not found')
idx = lines.index(target)
lines[idx] = "                            characters = cell.api.Characters(start_idx + 1, length_val)"
lines.insert(idx + 1, "                            # use COM API directly to avoid tuple comparison in xlwings indexing")

target = "                            characters.api.Font.Color = color_value"
if target not in lines:
    raise SystemExit('target color assignment line not found')
idx = lines.index(target)
lines[idx] = "                            characters.Font.Color = color_value"
lines.insert(idx + 1, "                            # COM Characters already exposes Font property")

path.write_text('\n'.join(lines) + ('\n' if text.endswith('\n') else ''), encoding='utf-8')
