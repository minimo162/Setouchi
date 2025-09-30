import importlib, dis
mod = importlib.import_module("excel_copilot.tools.excel_tools")
for instr in dis.get_instructions(mod.translate_range_contents):
    if instr.argval == "根拠の説明を日本語で1〜2文記述してください。":
        start = max(0, instr.offset - 200)
        end = instr.offset + 200
        break
else:
    raise SystemExit('not found')
for instr in dis.get_instructions(mod.translate_range_contents):
    if start <= instr.offset <= end:
        print(instr.offset, instr.opname, instr.argval)
