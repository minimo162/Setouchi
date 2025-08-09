"""Minimal stub implementation of a tiny subset of pandas.

This is *not* a full-featured pandas replacement. It only implements the
very small portion of functionality required by the unit tests in this
kata. The real project described in ``design.md`` relies on pandas for
Excel processing, but the execution environment used for the tests does
not provide the dependency. To keep the example self contained we ship a
light-weight substitute that can read and write a trivial CSV based
format.

The goal of this stub is to emulate the subset of the ``DataFrame``
API that the project uses:

* ``pd.DataFrame`` construction from ``dict`` or list + ``columns``
* ``DataFrame.to_excel``
* ``pd.ExcelWriter`` context manager
* ``pd.read_excel`` for detecting the header row
* very small indexing helpers (``len``, ``iloc`` and column access)
* ``DataFrame.reset_index`` and ``to_csv`` used by the CLI utility

The ``Excel`` support here is deliberately simplified: an ``.xlsx`` file
is treated as a plain CSV file. ``DataFrame.to_excel`` and
``pd.read_excel`` operate on comma separated values so that the tests can
exercise the surrounding logic without needing the heavy pandas
dependency.
"""
from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Dict, Any, Optional


class DataFrame:
    """A very small and incomplete tabular data container."""

    def __init__(self, data: Any = None, columns: Optional[Iterable[str]] = None):
        self._columns: List[str]
        self._data: List[Dict[str, Any]]

        if isinstance(data, dict):
            cols = list(columns) if columns is not None else list(data.keys())
            self._columns = cols
            length = len(next(iter(data.values()))) if data else 0
            self._data = []
            for i in range(length):
                row = {col: data[col][i] for col in cols}
                self._data.append(row)
        elif isinstance(data, list):
            cols = list(columns) if columns is not None else [str(i) for i in range(len(data[0]) if data else 0)]
            self._columns = cols
            self._data = [dict(zip(cols, row)) for row in data]
        else:
            self._columns = list(columns) if columns is not None else []
            self._data = []

    # ------------------------------------------------------------------
    # Basic container protocol
    def __len__(self) -> int:  # pragma: no cover - trivial
        return len(self._data)

    def __getitem__(self, key: str) -> List[Any]:  # pragma: no cover - not used
        return [row.get(key) for row in self._data]

    @property
    def empty(self) -> bool:  # pragma: no cover - trivial
        return not self._data

    # ------------------------------------------------------------------
    # Column handling
    @property
    def columns(self) -> List[str]:
        return self._columns

    @columns.setter
    def columns(self, new_cols: Iterable[str]) -> None:
        new_cols = list(new_cols)
        old_cols = self._columns
        if len(new_cols) != len(old_cols):
            raise ValueError("column count mismatch")
        remapped: List[Dict[str, Any]] = []
        for row in self._data:
            remapped.append({new_cols[i]: row.get(old_cols[i]) for i in range(len(old_cols))})
        self._data = remapped
        self._columns = new_cols

    # ------------------------------------------------------------------
    # Row indexer
    @dataclass
    class _ILocIndexer:
        data: List[Dict[str, Any]]

        def __getitem__(self, idx: int) -> Dict[str, Any]:
            return self.data[idx]

    @property
    def iloc(self) -> "DataFrame._ILocIndexer":
        return DataFrame._ILocIndexer(self._data)

    # ------------------------------------------------------------------
    def reset_index(self, drop: bool = True) -> "DataFrame":  # pragma: no cover - behaviour is trivial
        return self

    # ------------------------------------------------------------------
    def to_csv(self, index: bool = False) -> str:  # pragma: no cover - used only in CLI
        from io import StringIO

        buf = StringIO()
        writer = csv.writer(buf)
        writer.writerow(self._columns)
        for row in self._data:
            writer.writerow([row.get(c, "") for c in self._columns])
        return buf.getvalue()

    # ------------------------------------------------------------------
    def to_excel(self, writer: "ExcelWriter", index: bool = False, header: bool = True) -> None:
        csv_writer = csv.writer(writer._fh)
        if header:
            csv_writer.writerow(self._columns)
        for row in self._data:
            csv_writer.writerow([row.get(c, "") for c in self._columns])


class ExcelWriter:
    """Context manager writing CSV data with a ``.xlsx`` extension."""

    def __init__(self, path: str | Path):
        self.path = Path(path)
        self._fh = None

    def __enter__(self) -> "ExcelWriter":  # pragma: no cover - trivial
        self._fh = open(self.path, "w", newline="")
        return self

    def __exit__(self, exc_type, exc, tb) -> None:  # pragma: no cover - trivial
        if self._fh:
            self._fh.close()


# ----------------------------------------------------------------------
# Excel reading

def read_excel(path: str | Path, header: int = 0, nrows: Optional[int] = None) -> DataFrame:
    """Read a CSV based ``.xlsx`` file.

    Only the arguments used in the tests are implemented. ``header`` denotes
    the zero-based row containing the column names. ``nrows`` limits the number
    of data rows returned; ``0`` yields an empty ``DataFrame`` with only column
    information.
    """
    path = Path(path)
    with open(path, newline="") as fh:
        rows = list(csv.reader(fh))
    if header >= len(rows):
        cols: List[str] = []
        data_rows: List[List[str]] = []
    else:
        cols = rows[header]
        data_rows = rows[header + 1 :]
    if nrows is not None:
        data_rows = data_rows[:nrows]
    return DataFrame(data_rows, columns=cols)


__all__ = ["DataFrame", "ExcelWriter", "read_excel"]
