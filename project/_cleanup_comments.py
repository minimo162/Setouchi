from pathlib import Path

path = Path('excel_copilot/tools/actions.py')
lines = path.read_text(encoding='utf-8').splitlines()
cleanup_strings = [
    "                            # api call avoids tuple/int comparison in xlwings setter"
]
lines = [line for line in lines if line not in cleanup_strings]
path.write_text('\n'.join(lines) + '\n', encoding='utf-8')
