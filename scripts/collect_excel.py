"""
Collect Excel files from the Attachments directory into a single Excel folder.

Scans Root/Attachments (and subfolders) for Excel files and copies them
into Root/Excel_Files with a flat structure (date-prefixed filenames).

Usage:
    python scripts/collect_excel.py
    python scripts/collect_excel.py --task extract_attachments
"""

import os
import sys
import shutil
import glob
from pathlib import Path

# Ensure project root is on sys.path so 'lib' package resolves correctly.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from lib.reporting import write_csv_report
from lib.task_config import parse_task_args, unpack_config, get_output_dir, get_report_path
from lib.utils import is_excel_file


def collect_excel(config):
    attachments_subdir, excel_subdir = unpack_config(
        config, "attachments_subdir=Attachments", "excel_subdir=Excel_Files"
    )
    report_file = config.get("report_file", "Excel_Collect_Report.csv")

    source_dir = os.path.join(config["root"], attachments_subdir)
    dest_dir = get_output_dir(config, excel_subdir)
    report_path = get_report_path(config, report_file)

    if not os.path.exists(source_dir):
        print(f"Source directory does not exist: {source_dir}")
        sys.exit(1)

    print(f"Collecting Excel files from: {source_dir}")
    print(f"  Destination: {dest_dir}\n")

    report = []
    copied = 0
    skipped = 0

    for root_dir, _dirs, files in os.walk(source_dir):
        for filename in files:
            if not is_excel_file(filename):
                continue

            src_path = os.path.join(root_dir, filename)
            # Preserve subfolder name in the destination filename to avoid collisions
            rel_folder = os.path.relpath(root_dir, source_dir)
            if rel_folder == ".":
                dest_filename = filename
            else:
                # Prefix with the subfolder name (sanitized)
                folder_prefix = rel_folder.replace(os.sep, "_")
                dest_filename = f"{folder_prefix}__{filename}"

            dest_path = os.path.join(dest_dir, dest_filename)

            # Handle duplicates
            if os.path.exists(dest_path):
                skipped += 1
                continue

            try:
                shutil.copy2(src_path, dest_path)
                copied += 1
                print(f"  Copied: {dest_filename}")

                report.append({
                    "FileName": dest_filename,
                    "OriginalPath": src_path,
                    "SourceFolder": rel_folder,
                    "DestPath": dest_path,
                })
            except Exception as e:
                print(f"  Failed: {filename} - {e}")

    if write_csv_report(report_path, report):
        print(f"\nReport saved to: {report_path}")

    print(f"\nDone. {copied} Excel file(s) copied, {skipped} skipped (already exist).")


if __name__ == "__main__":
    config = parse_task_args(
        description="Collect Excel files from Attachments into a single Excel folder.",
        default_task="collect_excel",
    )
    collect_excel(config)
