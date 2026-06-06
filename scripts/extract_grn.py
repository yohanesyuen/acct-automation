"""
Extract GRN Excel attachments from Outlook inbox.

Saves Excel attachments whose filename or parent email body references
a GR number pattern. Generates a CSV report of extracted files.

Usage:
    python scripts/extract_grn.py
    python scripts/extract_grn.py --task extract_grn_excel
"""

import os
import sys
from pathlib import Path

# Ensure project root is on sys.path so 'lib' package resolves correctly.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from lib.outlook import (
    get_outlook_inbox,
    filter_emails_by_sender_and_keyword,
    iter_attachments,
    save_attachment,
)
from lib.reporting import write_csv_report
from lib.task_config import parse_task_args, unpack_config, get_output_dir, get_report_path
from lib.utils import (
    is_excel_file,
    extract_gr_numbers_from_text,
    make_date_prefixed_filename,
    GR_PATTERN,
)


def extract_grn(config):
    sender_email, keyword = unpack_config(config, "sender_email", "keyword")
    output_folder = get_output_dir(config, config.get("excel_subdir", "Excel_Files"))
    report_path = get_report_path(config, config.get("report_file", "GRN_Report.csv"))

    inbox = get_outlook_inbox()
    emails = filter_emails_by_sender_and_keyword(inbox, sender_email, keyword)

    print(f"Searching for Excel files from {sender_email} (keyword: {keyword})...")

    report = []

    for info in iter_attachments(emails, extension_filter=is_excel_file):
        body = info.attachment_ref.Parent.Body or ""

        # Only save if GR pattern found in filename or email body
        if not (GR_PATTERN.search(info.filename) or GR_PATTERN.search(body)):
            continue

        new_filename = make_date_prefixed_filename(info.received_time, info.filename)
        dest_path = os.path.join(output_folder, new_filename)

        try:
            save_attachment(info.attachment_ref, dest_path)
            print(f"  Saved: {info.filename}")

            report.append({
                "FileName": new_filename,
                "OriginalFileName": info.filename,
                "EmailSubject": info.email_subject,
                "ReceivedDate": str(info.received_time),
                "GRNumbers": ", ".join(extract_gr_numbers_from_text(body)),
                "FilePath": dest_path,
            })
        except Exception as e:
            print(f"  Failed: {info.filename} - {e}")

    if write_csv_report(report_path, report):
        print(f"Report saved to: {report_path}")

    print(f"\nDone. {len(report)} Excel file(s) saved to {output_folder}")


if __name__ == "__main__":
    config = parse_task_args(
        description="Extract GRN Excel attachments from Outlook inbox.",
        default_task="extract_grn",
    )
    extract_grn(config)
