"""
Extract Excel attachments that contain GR numbers in their cell content.

Scans the actual spreadsheet data (not just filename/email body) to decide
whether to keep each file. Uses a regex-based GR pattern search.

Usage:
    python scripts/extract_grn_search_content.py
"""

import os
import sys
import shutil
from pathlib import Path

# Ensure project root is on sys.path so 'lib' package resolves correctly.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from lib.outlook import (
    get_outlook_inbox,
    filter_emails_by_sender_and_keyword,
    iter_attachments,
    save_attachment_to_temp,
    cleanup_temp_file,
)
from lib.excel_search import search_excel_for_gr_pattern
from lib.reporting import write_csv_report
from lib.task_config import parse_task_args, get_output_dir, get_report_path
from lib.utils import is_excel_file, make_date_prefixed_filename


def extract_grn_search_content(config):
    sender_email = config.get("sender_email", "sender@example.com")
    keyword = config.get("keyword", "GRN")
    output_folder = get_output_dir(config, config.get("excel_subdir", "Excel_Files"))
    report_path = get_report_path(config, config.get("report_file", "GRN_Report.csv"))

    inbox = get_outlook_inbox()
    emails = filter_emails_by_sender_and_keyword(inbox, sender_email, keyword)

    print("Searching for Excel files with GR numbers in content...")

    report = []

    for info in iter_attachments(emails, extension_filter=is_excel_file):
        temp_path = save_attachment_to_temp(info.attachment_ref, info.filename)
        print(f"  Checking: {info.filename}...")

        try:
            gr_numbers = search_excel_for_gr_pattern(temp_path)

            if gr_numbers:
                new_filename = make_date_prefixed_filename(info.received_time, info.filename)
                dest_path = os.path.join(output_folder, new_filename)
                shutil.copy2(temp_path, dest_path)

                print(f"  SAVED: {info.filename} (GR: {', '.join(gr_numbers)})")

                report.append({
                    "FileName": new_filename,
                    "OriginalFileName": info.filename,
                    "EmailSubject": info.email_subject,
                    "ReceivedDate": str(info.received_time),
                    "GRNumbers": ", ".join(gr_numbers),
                    "FilePath": dest_path,
                })
            else:
                print(f"  Skipped: no GR numbers in {info.filename}")
        finally:
            cleanup_temp_file(temp_path)

    if write_csv_report(report_path, report):
        print(f"Report saved to: {report_path}")

    print(f"\nDone. {len(report)} file(s) saved to {output_folder}")


if __name__ == "__main__":
    config = parse_task_args(
        description="Extract Excel attachments with GR numbers in cell content.",
        default_task="extract_grn_search_content",
    )
    extract_grn_search_content(config)
