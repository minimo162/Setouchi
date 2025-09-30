import importlib
mod = importlib.import_module("excel_copilot.tools.excel_tools")
for idx, c in enumerate(mod.translate_range_contents.__code__.co_consts):
    if isinstance(c, str) and "根拠の説明を日本語で1" in c:
        print(idx, c)
