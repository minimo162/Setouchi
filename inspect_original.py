import importlib.util
import sys
import types
import pathlib
sys.stdout.reconfigure(encoding="utf-8")
package_name = "excel_copilot.tools._original"
pyc_path = pathlib.Path(r"excel_copilot/tools/__pycache__/excel_tools_original.cpython-311.pyc")
if not pyc_path.exists():
    raise SystemExit("original pyc not found")
spec = importlib.util.spec_from_file_location(package_name, pyc_path)
module = importlib.util.module_from_spec(spec)
sys.modules[package_name] = module
spec.loader.exec_module(module)
print("LEGACY", module.LEGACY_DIFF_MARKER_PATTERN.pattern)
print("MODERN", module.MODERN_DIFF_MARKER_PATTERN.pattern)
print("BASE", module._BASE_DIFF_TOKEN_PATTERN.pattern)
print("BOUNDARY", ''.join(sorted(module._SENTENCE_BOUNDARY_CHARS)))
print("CLOSING", module._CLOSING_PUNCTUATION)
