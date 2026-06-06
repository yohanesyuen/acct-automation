"""
Find GR numbers (8-digit format like 26000117) inside Excel attachments.

Similar to extract_grn_search_content but uses the stricter 8-digit GR
number search (numbers starting with '2'). No keyword filter on emails —
just filters by sender.

Usage:
    python scripts/find_gr_numbers.py
"""

import os
import sys
import shutil
import subprocess
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
from lib.excel_search import search_excel_for_gr_numbers
from lib.reporting import write_csv_report
from lib.task_config import parse_task_args, get_output_dir, get_report_path
from lib.utils import is_excel_file, make_date_prefixed_filename


def find_gr_numbers(config):
    sender_email = config.get("sender_email", "sender@example.com")
    output_folder = get_output_dir(config, config.get("excel_subdir", "Excel_Files"))
    report_path = get_report_path(config, config.get("report_file", "GRN_Report.csv"))

    inbox = get_outlook_inbox()
    emails = filter_emails_by_sender_and_keyword(inbox, sender_email, keyword=None)

    print(f"Searching emails from {sender_email} for Excel files with 8-digit GR numbers...")

    report = []

    for info in iter_attachments(emails, extension_filter=is_excel_file):
        temp_path = save_attachment_to_temp(info.attachment_ref, info.filename)
        print(f"  Checking: {info.filename}...")

        try:
            gr_numbers = search_excel_for_gr_numbers(temp_path)

            if gr_numbers:
                new_filename = make_date_prefixed_filename(info.received_time, info.filename)
                dest_path = os.path.join(output_folder, new_filename)
                shutil.copy2(temp_path, dest_path)

                print(f"  SUCCESS: {info.filename} (GR: {', '.join(gr_numbers)})")

                report.append({
                    "FileName": new_filename,
                    "OriginalFileName": info.filename,
                    "EmailSubject": info.email_subject,
                    "SenderEmail": info.email_sender,
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

    if config.get("open_folder_on_complete", False) and report:
        subprocess.Popen(["explorer.exe", output_folder])


if __name__ == "__main__":
    config = parse_task_args(
        description="Find GR numbers (8-digit format) inside Excel attachments.",
        default_task="find_gr_numbers",
    )
    find_gr_numbers(config)
