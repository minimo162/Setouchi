import pathlib
import importlib.util
import types
import sys
import difflib

sys.stdout.reconfigure(encoding="utf-8")
path = pathlib.Path(r"excel_copilot/tools/excel_tools.py")
lines = path.read_text(encoding="utf-8").splitlines()
package_name = "excel_copilot.tools"
if package_name not in sys.modules:
    pkg = types.ModuleType(package_name)
    pkg.__path__ = [str(pathlib.Path(r"excel_copilot/tools").resolve())]
    sys.modules[package_name] = pkg
spec = importlib.util.spec_from_file_location(f"{package_name}._compiled", r"excel_copilot/tools/__pycache__/excel_tools.cpython-311.pyc")
module = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = module
spec.loader.exec_module(module)
raw_strings = set()
for value in module.__dict__.values():
    if isinstance(value, str):
        raw_strings.add(value)
    elif hasattr(value, "__code__"):
        for const in value.__code__.co_consts:
            if isinstance(const, str):
                raw_strings.add(const)
strings = [s.replace("\n", " ") for s in raw_strings]

def normalize(s: str) -> str:
    replacements = [
        ("EEE", ""),
        ("�", ""),
        ("\u2001", ""),
        ("めE", "め"),
        ("、E", "、"),
        ("\r", ""),
        ("\n", " "),
    ]
    for old, new in replacements:
        s = s.replace(old, new)
    return s
norm_strings = [normalize(s) for s in strings]
for i, line in enumerate(lines, 1):
    if "EEE" in line:
        normalized_line = normalize(line)
        matches = difflib.get_close_matches(normalized_line, norm_strings, n=3, cutoff=0.1)
        print(i, line)
        print("  ->", matches)
