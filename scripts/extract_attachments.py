"""
Extract attachments from emails matching sender addresses, filtered by filename keyword.

Filters emails by sender, then saves only attachments whose filename
contains any of the specified keywords (case-insensitive).
If no keywords are specified, all attachments are saved.

Usage:
    python scripts/extract_attachments.py
    python scripts/extract_attachments.py --sender-email "alice@example.com,bob@example.com"
    python scripts/extract_attachments.py --keyword "GRN,Invoice"
"""

import os
import sys
from pathlib import Path

# Ensure project root is on sys.path so 'lib' package resolves correctly.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from lib.outlook import (
    get_outlook_inbox,
    filter_emails,
    iter_attachments,
    save_attachment,
)
from lib.reporting import write_csv_report
from lib.task_config import parse_task_args, get_output_dir, get_report_path
from lib.utils import make_date_prefixed_filename, sanitize_filename


def extract_attachments(config):
    sender_email = config.get("sender_email")
    keyword = config.get("keyword")
    output_folder = get_output_dir(config, config.get("attachments_subdir", "Attachments"))
    report_path = get_report_path(config, config.get("report_file", "Attachments_Report.csv"))

    # Normalize keyword to a list
    if keyword is None or keyword == []:
        keyword_list = None
    elif isinstance(keyword, str):
        keyword_list = [keyword]
    else:
        keyword_list = list(keyword)

    inbox = get_outlook_inbox()
    # Filter by sender only — keyword matching is done on attachment filenames
    emails = filter_emails(inbox, sender_email=sender_email, keyword=None)

    senders_display = sender_email if isinstance(sender_email, list) else [sender_email or "all"]
    print(f"Extracting attachments from: {', '.join(senders_display)}")
    if keyword_list:
        print(f"  Filename keyword filter: {keyword_list}")
    print(f"  Output: {output_folder}\n")

    report = []

    for info in iter_attachments(emails):
        # If keywords specified, only save attachments whose filename matches
        if keyword_list:
            filename_lower = info.filename.lower()
            if not any(k.lower() in filename_lower for k in keyword_list):
                continue

        new_filename = make_date_prefixed_filename(info.received_time, info.filename)
        dest_path = os.path.join(output_folder, new_filename)

        try:
            save_attachment(info.attachment_ref, dest_path)
            print(f"  Saved: {info.filename} (from: {info.email_sender})")

            report.append({
                "FileName": new_filename,
                "OriginalFileName": info.filename,
                "EmailSubject": info.email_subject,
                "SenderEmail": info.email_sender,
                "ReceivedDate": str(info.received_time),
                "FilePath": dest_path,
            })
        except Exception as e:
            print(f"  Failed: {info.filename} - {e}")

    if write_csv_report(report_path, report):
        print(f"\nReport saved to: {report_path}")

    print(f"\nDone. {len(report)} attachment(s) saved to {output_folder}")


if __name__ == "__main__":
    config = parse_task_args(
        description="Extract all attachments from emails matching sender address(es).",
        default_task="extract_attachments",
    )
    extract_attachments(config)
