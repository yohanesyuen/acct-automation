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
    get_outlook_folder,
    filter_emails,
    iter_attachments,
    save_attachment,
)
from lib.reporting import write_csv_report
from lib.task_config import parse_task_args, get_output_dir, get_report_path
from lib.utils import sanitize_filename


def extract_attachments(config):
    sender_email = config.get("sender_email")
    keyword = config.get("keyword")
    file_types = config.get("file_types")
    output_folder = get_output_dir(config, config.get("attachments_subdir", "Attachments"))
    report_path = get_report_path(config, config.get("report_file", "Attachments_Report.csv"))

    # Normalize keyword to a list
    if keyword is None or keyword == []:
        keyword_list = None
    elif isinstance(keyword, str):
        keyword_list = [keyword]
    else:
        keyword_list = list(keyword)

    # Normalize file_types to a list of lowercase extensions with dots
    if file_types is None or file_types == []:
        ext_filter = None
    elif isinstance(file_types, str):
        ext_filter = [file_types if file_types.startswith(".") else f".{file_types}"]
    else:
        ext_filter = [ft if ft.startswith(".") else f".{ft}" for ft in file_types]
    if ext_filter:
        ext_filter = [e.lower() for e in ext_filter]

    inbox = get_outlook_folder(config.get("folder", "inbox"))
    # Filter by sender only — keyword check includes body + attachment filenames
    emails = filter_emails(inbox, sender_email=sender_email, keyword=None, verbose=True)

    senders_display = sender_email if isinstance(sender_email, list) else [sender_email or "all"]
    print(f"Extracting attachments from: {', '.join(senders_display)}")
    if keyword_list:
        print(f"  Keyword filter (body or filename): {keyword_list}")
    if ext_filter:
        print(f"  File type filter: {ext_filter}")
    print(f"  Output: {output_folder}\n")

    report = []
    emails_checked = 0
    keyword_body_matches = 0
    keyword_filename_matches = 0

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

                if body_match:
                    keyword_body_matches += 1
                elif filename_match:
                    keyword_filename_matches += 1
                else:
                    emails_checked += 1
                    if emails_checked <= 5:
                        print(f"  [keyword] SKIP: \"{(parent_email.Subject or '')[:50]}\" "
                              f"(attachments: {', '.join(a.FileName for a in parent_email.Attachments)})")
                    continue
            except Exception as e:
                print(f"  [keyword] ERROR checking email: {e}")
                continue

        new_filename = info.filename

        # File type filter — skip attachments that don't match
        if ext_filter:
            file_ext = os.path.splitext(new_filename)[1].lower()
            if file_ext not in ext_filter:
                continue

        # Create per-email subfolder: <sanitized_subject>_<YYYYMMDD_HHMMSS>/
        safe_subject = sanitize_filename(info.email_subject)
        date_str = info.received_time.strftime("%Y%m%d_%H%M%S")
        subfolder = os.path.join(output_folder, f"{safe_subject}_{date_str}")
        os.makedirs(subfolder, exist_ok=True)
        dest_path = os.path.join(subfolder, new_filename)

        try:
            save_attachment(info.attachment_ref, dest_path)
            print(f"  Saved: {info.filename} (from: {info.email_sender})")

            report.append({
                "FileName": new_filename,
                "OriginalFileName": info.filename,
                "EmailSubject": info.email_subject,
                "SenderEmail": info.email_sender,
                "ReceivedDate": str(info.received_time),
                "SubFolder": f"{safe_subject}_{date_str}",
                "FilePath": dest_path,
            })
        except Exception as e:
            print(f"  Failed: {info.filename} - {e}")

    if write_csv_report(report_path, report):
        print(f"\nReport saved to: {report_path}")

    if keyword_list:
        print(f"\n  [keyword] Summary: body/subject matches={keyword_body_matches}, "
              f"filename matches={keyword_filename_matches}, "
              f"no match skipped={emails_checked}")

    print(f"\nDone. {len(report)} attachment(s) saved to {output_folder}")


if __name__ == "__main__":
    config = parse_task_args(
        description="Extract all attachments from emails matching sender address(es).",
        default_task="extract_attachments",
    )
    extract_attachments(config)
