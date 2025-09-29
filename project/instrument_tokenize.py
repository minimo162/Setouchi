from pathlib import Path
import re

path = Path('excel_copilot/tools/excel_tools.py')
text = path.read_text(encoding='utf-8')
pattern = re.compile(r"def _tokenize_for_diff\(text: str\) -> List\[str\]:\n(?:    .+\n){0,80}?    return segments or \[text\]\n", re.MULTILINE)
match = pattern.search(text)
if not match:
    raise SystemExit('tokenize block not found')
block = match.group(0)
block = block.replace('    if not text:\n        return []\n', "    if not text:\n        _diff_debug('_tokenize_for_diff empty input')\n        return []\n")
block = block.replace('    raw_tokens = _BASE_DIFF_TOKEN_PATTERN.findall(text)\n', "    raw_tokens = _BASE_DIFF_TOKEN_PATTERN.findall(text)\n    _diff_debug(f\"_tokenize_for_diff raw_tokens={_shorten_debug(raw_tokens)}\")\n")
block = block.replace("            segments.append(''.join(current_tokens))\n", "            segment = ''.join(current_tokens)\n            segments.append(segment)\n            _diff_debug(f\"_tokenize_for_diff flush segment={_shorten_debug(segment)}\")\n")
block = block.replace('        current_tokens.append(token)\n', "        current_tokens.append(token)\n        _diff_debug(f\"_tokenize_for_diff token={_shorten_debug(token)}\")\n")
block = block.replace("            if '\\r\\n' in token or '\\n' in token:\n                flush()\n            continue\n", "            if '\\r\\n' in token or '\\n' in token:\n                _diff_debug('_tokenize_for_diff flush due to newline token')\n                flush()\n            continue\n")
block = block.replace('    flush()\n    return segments or [text]\n', "    flush()\n    result = segments or [text]\n    _diff_debug(f\"_tokenize_for_diff result={_shorten_debug(result)}\")\n    return result\n")
text = text[:match.start()] + block + text[match.end():]
path.write_text(text, encoding='utf-8')
