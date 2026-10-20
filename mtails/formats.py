"""Delimited-text import/export used by the CSV and TSV paths.

The spreadsheet model is the same JSON structure the JavaScript layer owns:
one workbook with ``sheets`` and an address-keyed ``cells`` map per sheet.
"""

from __future__ import annotations

import csv
import io
import re
from typing import Any, List, Optional

MAX_ROWS = 100000
MAX_COLS = 1000

_NUMBER = re.compile(r"^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$")


def _is_number(text: str) -> bool:
    if not text or text.strip() != text:
        return False
    if len(text) > 1 and text[0] == "0" and not text.startswith("0."):
        # Preserve identifiers such as 007 or 00-1 rather than turning them
        # into numbers; Excel shows a smart tag here, we keep the text.
        return False
    return bool(_NUMBER.match(text))


def _convert(text: str) -> Any:
    if text == "":
        return None
    if _is_number(text):
        try:
            value = float(text)
            return int(value) if value.is_integer() and "e" not in text.lower() else value
        except ValueError:
            return text
    if text.upper() == "TRUE":
        return True
    if text.upper() == "FALSE":
        return False
    return text


def detect_delimiter(text: str) -> str:
    """Pick the delimiter that splits the first non-empty line most often."""
    sample = ""
    for line in text.splitlines():
        if line.strip():
            sample = line
            break
    if not sample:
        return ","
    counts = {",": sample.count(","), ";": sample.count(";"), "\t": sample.count("\t"), "|": sample.count("|")}
    best = max(counts, key=counts.get)
    return best if counts[best] else ","


def worksheet_name(name: str) -> str:
    cleaned = re.sub(r"[\\/*?:\[\]]", " ", name or "Sheet1").strip() or "Sheet1"
    return cleaned[:31]


def text_to_sheet(text: str, name: str = "Sheet1", delimiter: Optional[str] = None) -> dict:
    """Parse delimited text into a one-sheet workbook."""
    separator = delimiter or detect_delimiter(text)
    if separator == "\\t":
        separator = "\t"
    reader = csv.reader(io.StringIO(text), delimiter=separator)
    cells: dict = {}
    for row_index, row in enumerate(reader, start=1):
        if row_index > MAX_ROWS:
            break
        for col_index, raw in enumerate(row[:MAX_COLS], start=1):
            value = _convert(raw)
            if value is None:
                continue
            address = _address(row_index, col_index)
            cells[address] = {"v": value}
    return {
        "name": worksheet_name(name),
        "cells": cells,
        "cols": {},
        "rows": {},
        "merges": [],
        "freeze": {"rows": 0, "cols": 0},
        "charts": [],
        "images": [],
        "comments": [],
        "filters": None,
        "validations": [],
        "conditional": [],
        "tables": [],
        "page": default_page(),
        "protection": None,
    }


def _col_name(index: int) -> str:
    name = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def _address(row: int, col: int) -> str:
    return "%s%d" % (_col_name(col), row)


def _bounds(cells: dict) -> tuple[int, int]:
    max_row = 0
    max_col = 0
    for address in cells:
        match = re.match(r"^([A-Z]+)(\d+)$", address)
        if not match:
            continue
        col = 0
        for char in match.group(1):
            col = col * 26 + (ord(char) - 64)
        max_col = max(max_col, col)
        max_row = max(max_row, int(match.group(2)))
    return max_row, max_col


def _display(cell: dict) -> str:
    value = cell.get("v")
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float):
        if value.is_integer() and abs(value) < 1e15:
            return str(int(value))
        return repr(value)
    return str(value)


def sheet_to_delimited(model: dict, sheet_index: int = 0, delimiter: str = ",") -> str:
    """Serialise the used range of a sheet, values only (like Excel CSV)."""
    sheets = model.get("sheets") or []
    if not sheets:
        return ""
    index = min(max(0, int(sheet_index)), len(sheets) - 1)
    cells = sheets[index].get("cells") or {}
    max_row, max_col = _bounds(cells)
    if not max_row or not max_col:
        return ""
    buffer = io.StringIO()
    writer = csv.writer(buffer, delimiter=delimiter, lineterminator="\r\n", quoting=csv.QUOTE_MINIMAL)
    for row in range(1, max_row + 1):
        line: List[str] = []
        for col in range(1, max_col + 1):
            cell = cells.get(_address(row, col))
            line.append(_display(cell) if cell else "")
        writer.writerow(line)
    return buffer.getvalue()


def default_page() -> dict:
    """Page setup defaults shared with the JavaScript model."""
    return {
        "size": "A4",
        "orientation": "portrait",
        "margins": {"t": 72, "r": 72, "b": 72, "l": 72, "header": 30, "footer": 30},
        "scale": 100,
        "gridlines": False,
        "headings": False,
        "header": "",
        "footer": "&P",
        "area": "",
    }
