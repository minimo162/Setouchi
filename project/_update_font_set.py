from pathlib import Path

path = Path('excel_copilot/tools/actions.py')
text = path.read_text(encoding='utf-8')
lines = text.splitlines()

target = "                            characters.font.color = color_value"
if target not in lines:
    raise SystemExit('target line not found')
idx = lines.index(target)
lines[idx] = "                            characters.api.Font.Color = color_value"
lines.insert(idx + 1, "                            # api call avoids tuple/int comparison in xlwings setter")
path.write_text('\n'.join(lines) + ('\n' if text.endswith('\n') else ''), encoding='utf-8')
