import importlib, dis, sys
sys.stdout.reconfigure(encoding="utf-8")
mod = importlib.import_module("excel_copilot.tools.excel_tools")
for instr in dis.get_instructions(mod.translate_range_contents):
    if 1820 <= instr.offset <= 1895:
        print(instr.offset, instr.opname, instr.argval)
