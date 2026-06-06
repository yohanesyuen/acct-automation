"""
Analyze Excel files for column headers matching a substring.

Scans all Excel files in a directory, reads the first row (headers) of each
sheet, and reports any columns whose header contains the search substring.

Usage:
    python scripts/analyze_headers.py
    python scripts/analyze_headers.py --search-term "GR"
    python scripts/analyze_headers.py --search-term "Amount,Date" --source-subdir Excel_Files
"""

import os
import sys
from pathlib import Path

# Ensure project root is on sys.path so 'lib' package resolves correctly.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import openpyxl
import xlrd

from lib.reporting import write_csv_report
from lib.task_config import parse_task_args, unpack_config, get_report_path
from lib.utils import is_excel_file


def _get_headers_openpyxl(file_path: str) -> list:
    """Extract headers (first row) from each sheet in a .xlsx/.xlsm file."""
    results = []
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            for row in ws.iter_rows(min_row=1, max_row=1):
                for col_idx, cell in enumerate(row, start=1):
                    if cell.value is not None:
                        results.append({
                            "sheet": ws.title,
                            "column": col_idx,
                            "header": str(cell.value),
                        })
    finally:
        wb.close()
    return results


def _get_headers_xlrd(file_path: str) -> list:
    """Extract headers (first row) from each sheet in a .xls file."""
    results = []
    wb = xlrd.open_workbook(file_path)
    for sheet in wb.sheets():
        if sheet.nrows == 0:
            continue
        for col_idx in range(sheet.ncols):
            value = sheet.cell_value(0, col_idx)
            if value not in (None, ""):
                results.append({
                    "sheet": sheet.name,
                    "column": col_idx + 1,
                    "header": str(value),
                })
    return results


def get_headers(file_path: str) -> list:
    """Get headers from an Excel file (dispatches by extension)."""
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".xls":
        return _get_headers_xlrd(file_path)
    return _get_headers_openpyxl(file_path)


def analyze_headers(config):
    search_terms, exclude_terms, source_subdir = unpack_config(
        config, "search_term", "exclude_term", "source_subdir=Excel_Files"
    )
    report_file = config.get("report_file", "Header_Analysis_Report.csv")

    source_dir = os.path.join(config["root"], source_subdir)
    report_path = get_report_path(config, report_file)

    # Normalize search_terms to a list
    if search_terms is None or search_terms == []:
        print("Error: search_term is required. Specify in YAML or via --search-term.")
        sys.exit(1)
    elif isinstance(search_terms, str):
        search_list = [search_terms]
    else:
        search_list = list(search_terms)

    # Normalize exclude_terms to a list
    if exclude_terms is None or exclude_terms == []:
        exclude_list = []
    elif isinstance(exclude_terms, str):
        exclude_list = [exclude_terms]
    else:
        exclude_list = list(exclude_terms)

    if not os.path.exists(source_dir):
        print(f"Source directory does not exist: {source_dir}")
        sys.exit(1)

    print(f"Analyzing Excel headers in: {source_dir}")
    print(f"  Search terms: {search_list}")
    if exclude_list:
        print(f"  Exclude terms: {exclude_list}")
    print()

    report = []
    files_scanned = 0
    matches_found = 0

    for root_dir, _dirs, files in os.walk(source_dir):
        for filename in files:
            if not is_excel_file(filename):
                continue

            file_path = os.path.join(root_dir, filename)
            files_scanned += 1

            try:
                headers = get_headers(file_path)
            except Exception as e:
                print(f"  Error reading {filename}: {e}")
                continue

            for header_info in headers:
                header_lower = header_info["header"].lower()
                matched_terms = [t for t in search_list if t.lower() in header_lower]

                if not matched_terms:
                    continue

                # Check exclusions — skip if any exclude term is in the header
                excluded_by = [t for t in exclude_list if t.lower() in header_lower]
                if excluded_by:
                    continue

                matches_found += 1
                print(f"  MATCH: {filename} | Sheet: {header_info['sheet']} | "
                      f"Col {header_info['column']}: \"{header_info['header']}\" "
                      f"(matched: {', '.join(matched_terms)})")

                report.append({
                    "FileName": filename,
                    "FilePath": file_path,
                    "Sheet": header_info["sheet"],
                    "Column": header_info["column"],
                        "Header": header_info["header"],
                        "MatchedTerms": ", ".join(matched_terms),
                    })

    if write_csv_report(report_path, report):
        print(f"\nReport saved to: {report_path}")

    print(f"\nDone. Scanned {files_scanned} file(s), found {matches_found} matching header(s).")


if __name__ == "__main__":
    config = parse_task_args(
        description="Analyze Excel column headers for substring matches.",
        default_task="analyze_headers",
    )
    analyze_headers(config)
