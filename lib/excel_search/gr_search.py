"""
Excel content searching for GR numbers.

Uses openpyxl (for .xlsx/.xlsm) and xlrd (for .xls) to read Excel files
without requiring a running Excel application. Two search strategies:

- Pattern-based: finds cells matching a regex like ``GR[\\s\\-_]?\\d+``
- Numeric format: finds 8-digit GR numbers (starting with 2), also
  recognizing 'GR Number' column headers and reading values below them.
"""

import os
import re
from typing import List

import openpyxl
import xlrd


def _is_xls(file_path: str) -> bool:
    """Check if file is legacy .xls format."""
    return os.path.splitext(file_path)[1].lower() == ".xls"


def _iter_cells_openpyxl(file_path: str):
    """
    Yield (row_idx, col_idx, cell_text) for every cell in a .xlsx/.xlsm file.

    row_idx and col_idx are 1-based to match Excel conventions.
    """
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            for row_idx, row in enumerate(ws.iter_rows(), start=1):
                for col_idx, cell in enumerate(row, start=1):
                    value = cell.value
                    if value is not None:
                        yield row_idx, col_idx, str(value)
    finally:
        wb.close()


def _iter_cells_xlrd(file_path: str):
    """
    Yield (row_idx, col_idx, cell_text) for every cell in a .xls file.

    row_idx and col_idx are 1-based to match Excel conventions.
    """
    wb = xlrd.open_workbook(file_path)
    for sheet in wb.sheets():
        for row_idx in range(sheet.nrows):
            for col_idx in range(sheet.ncols):
                value = sheet.cell_value(row_idx, col_idx)
                if value not in (None, ""):
                    yield row_idx + 1, col_idx + 1, str(value)


def _iter_cells(file_path: str):
    """Dispatch to the appropriate reader based on file extension."""
    if _is_xls(file_path):
        return _iter_cells_xlrd(file_path)
    return _iter_cells_openpyxl(file_path)


def search_excel_for_gr_pattern(file_path: str, pattern: str = r"GR[\s\-_]?\d+") -> List[str]:
    """
    Search all cells in an Excel file for values matching a regex pattern.

    Reads the file using openpyxl or xlrd (no COM/Excel app required).

    Args:
        file_path: Absolute path to the Excel file to search.
        pattern: Regex pattern to match against cell text content.
                 Defaults to ``GR[\\s\\-_]?\\d+`` which catches formats
                 like "GR 12345", "GR-12345", "GR_12345", "GR12345".

    Returns:
        Deduplicated list of matched strings found across all worksheets.
    """
    compiled = re.compile(pattern)
    gr_numbers_found = []

    try:
        for _row, _col, cell_text in _iter_cells(file_path):
            matches = compiled.findall(cell_text)
            gr_numbers_found.extend(matches)
    except Exception as e:
        print(f"  Error reading file: {e}")

    return list(set(gr_numbers_found))


def search_excel_for_gr_numbers(file_path: str) -> List[str]:
    """
    Search an Excel file for 8-digit GR numbers (format: 2XXXXXXX).

    This function implements a more targeted search strategy:
    1. Looks for cells containing 8-digit numbers or 'GR Number' headers.
    2. Extracts 8-digit numbers that start with '2' (e.g. 26000117).
    3. When a 'GR Number' header is found, also checks the cell directly
       below it for an 8-digit number.

    Args:
        file_path: Absolute path to the Excel file to search.

    Returns:
        Deduplicated list of 8-digit GR number strings found in the file.
    """
    gr_numbers_found = []
    gr_header_positions = []  # (row, col) of "GR Number" headers

    try:
        # First pass: collect all matches and header positions
        cells_cache = {}
        for row, col, cell_text in _iter_cells(file_path):
            cells_cache[(row, col)] = cell_text

            if re.search(r"GR\s*Number", cell_text) or re.search(r"\b\d{8}\b", cell_text):
                # Extract 8-digit numbers starting with 2
                match = re.search(r"\b(2\d{7})\b", cell_text)
                if match:
                    gr_numbers_found.append(match.group(1))

                # Track header positions for next-row lookup
                if re.search(r"GR\s*Number", cell_text):
                    gr_header_positions.append((row, col))

        # Second pass: check cells below GR Number headers
        for row, col in gr_header_positions:
            next_cell = cells_cache.get((row + 1, col), "")
            if next_cell:
                next_match = re.search(r"\b(\d{8})\b", next_cell)
                if next_match:
                    gr_numbers_found.append(next_match.group(1))

    except Exception as e:
        print(f"  Error reading file: {e}")

    return list(set(gr_numbers_found))
