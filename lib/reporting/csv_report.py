"""
CSV report generation utilities.

Provides a simple interface for writing lists of dictionaries to CSV files,
commonly used for extraction result summaries.
"""

import csv
from typing import Dict, List


def write_csv_report(report_path: str, rows: List[Dict[str, str]]) -> bool:
    """
    Write a list of dictionaries to a CSV file.

    Uses the keys of the first row as the CSV header. The file is written
    with UTF-8 encoding and no extra blank lines between rows.

    Args:
        report_path: Absolute path where the CSV file will be written.
                     Parent directories must already exist.
        rows: List of dictionaries, each representing one row. All dicts
              should have the same keys. If the list is empty, no file
              is written and False is returned.

    Returns:
        True if the report was written successfully, False if rows was empty.
    """
    if not rows:
        return False

    with open(report_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=rows[0].keys())
        writer.writeheader()
        writer.writerows(rows)

    return True
