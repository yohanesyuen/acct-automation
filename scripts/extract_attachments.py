"""
Extract attachments from emails matching sender addresses and keyword criteria.

An email passes the filter if:
  - It is from any of the specified senders, AND
  - (the body/subject contains any keyword OR any attachment filename contains any keyword)

If no keywords are specified, all attachments from matching senders are saved.
All attachments from a matching email are saved.

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
    # Filter by sender only — keyword check includes body + attachment filenames
    emails = filter_emails(inbox, sender_email=sender_email, keyword=None)

    senders_display = sender_email if isinstance(sender_email, list) else [sender_email or "all"]
    print(f"Extracting attachments from: {', '.join(senders_display)}")
    if keyword_list:
        print(f"  Keyword filter (body or filename): {keyword_list}")
    print(f"  Output: {output_folder}\n")

    report = []

    for info in iter_attachments(emails):
        # If keywords specified, the email must match:
        #   body/subject contains any keyword OR any attachment filename contains any keyword
        if keyword_list:
            try:
                parent_email = info.attachment_ref.Parent
                subject_lower = (parent_email.Subject or "").lower()
                body_lower = (parent_email.Body or "").lower()

                # Check body/subject
                body_match = any(k.lower() in subject_lower or k.lower() in body_lower
                                 for k in keyword_list)

                # Check attachment filenames
                filename_match = False
                if not body_match:
                    for i in range(1, parent_email.Attachments.Count + 1):
                        att_name = parent_email.Attachments.Item(i).FileName.lower()
                        if any(k.lower() in att_name for k in keyword_list):
                            filename_match = True
                            break

                if not body_match and not filename_match:
                    continue
            except Exception:
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
