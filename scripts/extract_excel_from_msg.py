"""
Extract Excel attachments from all .msg files in a directory.

Usage:
    python scripts/extract_excel_from_msg.py
"""

import sys
from pathlib import Path

# Ensure project root is on sys.path so 'lib' package resolves correctly.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from lib.email_extraction.extractors import get_msg_files, extract_attachments_from_multiple_msg
from lib.email_extraction.filters import create_filename_extension_filter
from lib.task_config import parse_task_args, unpack_config, get_output_dir


def extract_excel_attachments_from_msg_directory(config):
    attachments_subdir, excel_subdir = unpack_config(
        config, "attachments_subdir=Attachments", "excel_subdir=Excel_Files"
    )
    source_dir = get_output_dir(config, attachments_subdir)
    dest_dir = get_output_dir(config, excel_subdir)

    msg_files = get_msg_files(source_dir, recursive=True)
    if not msg_files:
        print(f"No .msg files found under {source_dir}")
        return {}

    extension_filter = create_filename_extension_filter([".xls", ".xlsx", ".xlsm"])
    results = extract_attachments_from_multiple_msg(
        msg_files,
        dest_dir,
        filename_filter=extension_filter,
        verbose=True,
    )

    total = sum(len(items) for items in results.values())
    print(f"\nProcessed {len(msg_files)} .msg files. Extracted {total} Excel attachment(s) to {dest_dir}")
    return results


if __name__ == "__main__":
    config = parse_task_args(
        description="Extract Excel attachments from all .msg files in a directory.",
        default_task="main",
    )
    extract_excel_attachments_from_msg_directory(config)
