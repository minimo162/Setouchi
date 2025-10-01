"""High level Excel helper functions exposed to the co-pilot agent."""

from __future__ import annotations

import json
import difflib
import math
import re
from typing import Any, Dict, Iterator, List, Optional, Sequence, Tuple

from excel_copilot.core.browser_copilot_manager import BrowserCopilotManager
from excel_copilot.core.exceptions import ToolExecutionError
from .actions import ExcelActions

__all__ = [
    "write_cell_value",
    "read_cell_value",
    "list_sheet_names",
    "copy_range_values",
    "set_cell_formula",
    "read_range_values",
    "write_range_values",
    "get_active_workbook_context",
    "translate_range_contents",
    "review_translation_quality",
    "highlight_text_differences",
]


# ---------------------------------------------------------------------------
# Public tool functions
# ---------------------------------------------------------------------------


def write_cell_value(
    actions: ExcelActions,
    cell: str,
    value: Any,
    sheet_name: Optional[str] = None,
) -> str:
    """Write a value to a single cell.

    Args:
        actions: Excel automation helper.
        cell: Target cell address such as ``"A1"``.
        value: Value to write. Strings, numbers, booleans and dates are allowed.
        sheet_name: Optional sheet name. Defaults to the active sheet.

    Returns:
        Confirmation text describing the action.
    """

    actions.write_to_cell(cell, value, sheet_name)
    return f"Stored value in cell {cell}."


def read_cell_value(
    actions: ExcelActions,
    cell: str,
    sheet_name: Optional[str] = None,
) -> str:
    """Read the value of a single cell and present it as text."""

    raw = actions.read_range(cell, sheet_name)
    matrix = _ensure_2d(raw)
    value = matrix[0][0] if matrix else None
    return f"Cell {cell} currently contains: '{_coerce_to_string(value)}'."


def list_sheet_names(actions: ExcelActions) -> str:
    """List all sheets that belong to the active workbook."""

    sheet_names = actions.get_sheet_names()
    if not sheet_names:
        return "No sheets were found in the current workbook."
    return "Sheets: " + ", ".join(sheet_names)


def copy_range_values(
    actions: ExcelActions,
    source_range: str,
    destination_range: str,
    sheet_name: Optional[str] = None,
) -> str:
    """Copy values from one range into another range."""

    actions.copy_range(source_range, destination_range, sheet_name)
    return f"Copied values from {source_range} to {destination_range}."


def set_cell_formula(
    actions: ExcelActions,
    cell: str,
    formula: str,
    sheet_name: Optional[str] = None,
) -> str:
    """Assign an Excel formula to a cell."""

    actions.set_formula(cell, formula, sheet_name)
    return f"Assigned formula {formula} to cell {cell}."


def read_range_values(
    actions: ExcelActions,
    cell_range: str,
    sheet_name: Optional[str] = None,
) -> str:
    """Read a range and format the values as plain text."""

    matrix = _ensure_2d(actions.read_range(cell_range, sheet_name))
    if not matrix:
        return "The specified range did not contain any readable values."
    lines: List[str] = []
    for row_index, row in enumerate(matrix, start=1):
        values = ", ".join(_coerce_to_string(cell) for cell in row)
        lines.append(f"Row {row_index}: {values}")
    return "\n".join(lines)


def write_range_values(
    actions: ExcelActions,
    cell_range: str,
    data: List[List[Any]],
    sheet_name: Optional[str] = None,
) -> str:
    """Write a two-dimensional list of data to an Excel range."""

    if not isinstance(data, list) or not all(isinstance(row, list) for row in data):
        raise ToolExecutionError("The data argument must be a two-dimensional list.")
    actions.write_range(cell_range, data, sheet_name)
    return f"Wrote {len(data)} row(s) to range {cell_range}."


def get_active_workbook_context(actions: ExcelActions) -> str:
    """Return the name of the active workbook and sheet."""

    workbook = getattr(actions.book, "name", "Unknown workbook")
    try:
        sheet = actions.book.sheets.active.name
    except Exception:
        sheet = "Unknown sheet"
    return f"Workbook: {workbook}\nSheet: {sheet}"


def translate_range_contents(
    actions: ExcelActions,
    browser_manager: BrowserCopilotManager,
    source_range: str,
    target_range: str,
    target_language: str = "English",
    sheet_name: Optional[str] = None,
    rows_per_batch: int = 5,
) -> str:
    """Translate every cell in a range using the Copilot LLM."""

    if rows_per_batch <= 0:
        raise ToolExecutionError("rows_per_batch must be at least 1.")

    source_matrix = _ensure_2d(actions.read_range(source_range, sheet_name))
    rows = len(source_matrix)
    cols = len(source_matrix[0]) if rows else 0

    flat_values = [_coerce_to_string(value) for row in source_matrix for value in row]
    translations: List[str] = []

    for batch in _chunked(flat_values, rows_per_batch):
        prompt = _build_translation_prompt(batch, target_language)
        response = browser_manager.ask(prompt)
        translations.extend(_parse_json_array(response, expected_length=len(batch)))

    translated_matrix = _reshape_flat_list(translations, rows, cols)
    actions.write_range(target_range, translated_matrix, sheet_name)

    return (
        f"Translated {len(translations)} text(s) into {target_language} "
        f"and wrote the results to range {target_range}."
    )


def review_translation_quality(
    actions: ExcelActions,
    source_range: str,
    translated_range: str,
    status_output_range: str,
    notes_output_range: str,
    sheet_name: Optional[str] = None,
) -> str:
    """Perform a lightweight consistency check across translations."""

    source_matrix = _ensure_2d(actions.read_range(source_range, sheet_name))
    translated_matrix = _ensure_2d(actions.read_range(translated_range, sheet_name))

    if len(source_matrix) != len(translated_matrix) or any(
        len(src_row) != len(dst_row)
        for src_row, dst_row in zip(source_matrix, translated_matrix)
    ):
        raise ToolExecutionError("Source and translated ranges have different shapes.")

    status_rows: List[List[str]] = []
    note_rows: List[List[str]] = []
    total = 0
    revise = 0

    for src_row, dst_row in zip(source_matrix, translated_matrix):
        status_row: List[str] = []
        note_row: List[str] = []
        for src_text, dst_text in zip(src_row, dst_row):
            total += 1
            src_numbers = _extract_numbers(_coerce_to_string(src_text))
            dst_numbers = _extract_numbers(_coerce_to_string(dst_text))
            if src_numbers != dst_numbers:
                status_row.append("REVISE")
                note_row.append("Numbers differ between the source and translation.")
                revise += 1
            else:
                status_row.append("OK")
                note_row.append("")
        status_rows.append(status_row)
        note_rows.append(note_row)

    actions.write_range(status_output_range, status_rows, sheet_name)
    actions.write_range(notes_output_range, note_rows, sheet_name)

    return f"Checked {total} cell(s). REVISE flagged for {revise} item(s)."


def highlight_text_differences(
    actions: ExcelActions,
    original_range: str,
    revised_range: str,
    output_range: str,
    sheet_name: Optional[str] = None,
) -> str:
    """Highlight textual changes between two equally sized ranges."""

    original_matrix = _ensure_2d(actions.read_range(original_range, sheet_name))
    revised_matrix = _ensure_2d(actions.read_range(revised_range, sheet_name))

    if len(original_matrix) != len(revised_matrix) or any(
        len(src_row) != len(dst_row)
        for src_row, dst_row in zip(original_matrix, revised_matrix)
    ):
        raise ToolExecutionError("Original and revised ranges must have the same shape.")

    diff_matrix: List[List[str]] = []
    for original_row, revised_row in zip(original_matrix, revised_matrix):
        diff_row: List[str] = []
        for before, after in zip(original_row, revised_row):
            diff_row.append(_build_highlight_text(_coerce_to_string(before), _coerce_to_string(after)))
        diff_matrix.append(diff_row)

    actions.write_range(output_range, diff_matrix, sheet_name)
    total_cells = sum(len(row) for row in diff_matrix)
    return f"Generated highlighted text for {total_cells} cell(s)."


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _ensure_2d(value: Any) -> List[List[Any]]:
    if isinstance(value, list):
        if not value:
            return [[]]
        if all(isinstance(row, list) for row in value):
            return [list(row) for row in value]
        return [[item] for item in value]
    return [[value]]


def _coerce_to_string(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        if math.isfinite(value) and value.is_integer():
            return str(int(value))
        if math.isfinite(value):
            return str(value)
        return ""
    return str(value)


def _chunked(values: Sequence[str], size: int) -> Iterator[List[str]]:
    for index in range(0, len(values), size):
        yield list(values[index:index + size])


def _build_translation_prompt(texts: Sequence[str], target_language: str) -> str:
    payload = json.dumps(list(texts), ensure_ascii=False)
    return (
        f"Translate {len(texts)} Japanese sentence(s) into {target_language}.\n"
        "Return a JSON array of strings with the same length and order.\n"
        "Do not add extra commentary.\n"
        f"Sentences: {payload}"
    )


def _parse_json_array(payload: str, expected_length: int) -> List[str]:
    try:
        data = json.loads(payload)
    except json.JSONDecodeError:
        match = re.search(r"\[[\s\S]*\]", payload)
        if not match:
            raise ToolExecutionError("Failed to parse a JSON array from the LLM response.")
        data = json.loads(match.group(0))

    if not isinstance(data, list):
        raise ToolExecutionError("The LLM response must be a JSON array.")
    if expected_length and len(data) != expected_length:
        raise ToolExecutionError(
            f"Expected {expected_length} translations but received {len(data)}."
        )
    if any(not isinstance(item, str) for item in data):
        raise ToolExecutionError("Each translation must be returned as a string.")
    return data


def _reshape_flat_list(values: List[str], rows: int, cols: int) -> List[List[str]]:
    if rows * cols != len(values):
        raise ToolExecutionError("Result count does not match the target range dimensions.")
    iterator = iter(values)
    return [[next(iterator) for _ in range(cols)] for _ in range(rows)]


def _extract_numbers(text: str) -> List[str]:
    return sorted(re.findall(r"[-+]?\d+(?:\.\d+)?", text))


def _build_highlight_text(before: str, after: str) -> str:
    if before == after:
        return after

    parts: List[str] = []
    matcher = difflib.SequenceMatcher(None, before, after)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        before_slice = before[i1:i2]
        after_slice = after[j1:j2]
        if tag == "equal":
            parts.append(after_slice)
        elif tag == "insert":
            parts.append(f"[ADD]{after_slice}[/ADD]")
        elif tag == "delete":
            parts.append(f"[DEL]{before_slice}[/DEL]")
        elif tag == "replace":
            if before_slice:
                parts.append(f"[DEL]{before_slice}[/DEL]")
            if after_slice:
                parts.append(f"[ADD]{after_slice}[/ADD]")
    return "".join(parts)
