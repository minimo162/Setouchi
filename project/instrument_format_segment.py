from pathlib import Path
import re

path = Path('excel_copilot/tools/excel_tools.py')
text = path.read_text(encoding='utf-8')
pattern = re.compile(r"def _format_diff_segment\(tokens: List\[str\], label: str\) -> Tuple\[str, Optional\[int\], Optional\[int\]\]:\n    if not tokens:\n        return '', None, None\n    segment = ''.join(tokens)\n    if not segment.strip():\n        return segment, None, None\n    leading_len = len(segment) - len(segment.lstrip\(\))\n    trailing_len = len(segment.rstrip\(\)) - len(segment.strip\(\))\n    core_start = leading_len\n    core_end = len(segment) - trailing_len if trailing_len else len(segment)\n    core = segment\[core_start:core_end\]\n    if not core:\n        return segment, None, None\n    prefix = segment\[:leading_len\]\n    suffix = segment\[core_end:\]\n    marker_prefix = f'【{label}：'\n    marker_suffix = '】'\n    formatted = f'{prefix}{marker_prefix}{core}{marker_suffix}{suffix}'\n    highlight_start_offset = len(prefix) + len(marker_prefix)\n    highlight_length = len(core)\n    return formatted, highlight_start_offset, highlight_length\n")
match = pattern.search(text)
if not match:
    raise SystemExit('target block not found')
block = match.group(0)
lines = block.split('\n')
lines.insert(2, "    _diff_debug(f'_format_diff_segment start label={label} tokens={_shorten_debug(tokens)}')")
insert_index = len(lines) - 1
lines.insert(insert_index, "    _diff_debug(f'_format_diff_segment result label={label} formatted={_shorten_debug(formatted)} offset={highlight_start_offset} length={highlight_length}')")
new_block = '\n'.join(lines)
text = text.replace(block, new_block)
path.write_text(text, encoding='utf-8')
