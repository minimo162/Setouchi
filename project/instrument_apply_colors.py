from pathlib import Path

path = Path('excel_copilot/tools/excel_tools.py')
text = path.read_text(encoding='utf-8')
before = "        try:\n            if not style_matrix:\n                return\n            sheet = self._get_sheet(sheet_name)\n"
after = "        try:\n            if not style_matrix:\n                _diff_debug('apply_diff_highlight_colors empty style matrix')\n                return\n            _diff_debug(f\"apply_diff_highlight_colors range={cell_range} rows={len(style_matrix)} cols={(len(style_matrix[0]) if style_matrix else 0)} addition={addition_color_hex} deletion={deletion_color_hex}\")\n            sheet = self._get_sheet(sheet_name)\n"
if before not in text:
    raise SystemExit('target start not found')
text = text.replace(before, after, 1)
path.write_text(text, encoding='utf-8')
