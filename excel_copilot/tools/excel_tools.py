"""Compatibility loader for the original excel_tools bytecode."""

from __future__ import annotations

import marshal
import pathlib
import sys

__loader_file = pathlib.Path(__file__).resolve()
__loader_cache_dir = __loader_file.with_name("__pycache__")
__loader_cache_tag = sys.implementation.cache_tag


def __loader_find_bytecode() -> pathlib.Path:
    candidates = []
    if __loader_cache_tag:
        tagged = __loader_cache_dir / f"{__loader_file.stem}_original.{__loader_cache_tag}.pyc"
        if tagged.exists():
            candidates.append(tagged)
    if not candidates and __loader_cache_dir.exists():
        candidates.extend(sorted(__loader_cache_dir.glob(f"{__loader_file.stem}_original.cpython-*.pyc")))
    if not candidates:
        raise ImportError("original excel_tools bytecode not found")
    return candidates[0]


__loader_bytecode_path = __loader_find_bytecode()
with __loader_bytecode_path.open("rb") as __loader_handle:
    __loader_header_size = 16 if sys.version_info >= (3, 7) else 12
    __loader_handle.read(__loader_header_size)
    __loader_code = marshal.load(__loader_handle)

globals()["__cached__"] = str(__loader_bytecode_path)
exec(__loader_code, globals())

# Clean up loader internals
for __name in list(globals()):
    if __name.startswith("__loader_"):
        del globals()[__name]
