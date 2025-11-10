#!/usr/bin/env python3
"""Bridge script exposing ExcelManager context as JSON."""

import argparse
import json
import sys
from typing import Any, Dict

from excel_copilot.core.excel_manager import ExcelManager
from excel_copilot.core.exceptions import ExcelConnectionError


def _emit(data: Dict[str, Any]) -> None:
    json.dump(data, sys.stdout, ensure_ascii=False)
    sys.stdout.write("\n")
    sys.stdout.flush()


def refresh_context(workbook: str | None) -> int:
    try:
        with ExcelManager(workbook) as manager:
            active = manager.get_active_workbook_and_sheet()
            workbook_names = manager.list_workbook_names()
            current_book = manager.get_active_workbook()
            current_name = current_book.name
            snapshot: Dict[str, Any] = {
                "active_workbook": active.get("workbook_name"),
                "active_sheet": active.get("sheet_name"),
                "workbooks": [],
            }
            for name in workbook_names:
                try:
                    manager.activate_workbook(name)
                    sheets = manager.list_sheet_names()
                except Exception as exc:  # pragma: no cover - defensive fallback
                    sheets = []
                    print(f"[excel-bridge] failed to inspect workbook {name}: {exc}", file=sys.stderr)
                snapshot["workbooks"].append({"name": name, "sheets": sheets})
            if current_name and current_name != snapshot["active_workbook"]:
                try:
                    manager.activate_workbook(current_name)
                except Exception:
                    pass
    except ExcelConnectionError as exc:
        _emit({"error": str(exc), "kind": "connection"})
        return 2
    except Exception as exc:  # pragma: no cover - defensive
        _emit({"error": str(exc), "kind": "unknown"})
        return 1

    _emit({"context": snapshot})
    return 0


def focus_workbook(workbook: str) -> int:
    try:
        with ExcelManager(workbook) as manager:
            manager.activate_workbook(workbook)
            manager.focus_application_window()
            _emit({"status": "focused", "workbook": workbook})
            return 0
    except Exception as exc:
        _emit({"error": str(exc), "kind": "focus"})
        return 1


def main() -> int:
    parser = argparse.ArgumentParser(description="Excel bridge CLI")
    sub = parser.add_subparsers(dest="command", required=True)

    refresh = sub.add_parser("refresh-context", help="Dump workbook context as JSON")
    refresh.add_argument("--workbook", help="Preferred workbook name", default=None)

    focus = sub.add_parser("focus-workbook", help="Activate and focus a workbook")
    focus.add_argument("name", help="Workbook name to activate")

    args = parser.parse_args()
    if args.command == "refresh-context":
        return refresh_context(args.workbook)
    if args.command == "focus-workbook":
        return focus_workbook(args.name)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
