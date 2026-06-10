"""
Extract attachments from Outlook emails with flexible filtering.

Combines the functionality of extract_grn, extract_grn_emails, and the
original extract_attachments into a single configurable script.

Filtering criteria:
  - Sender(s): email must be from any specified sender
  - Keywords: body/subject OR any attachment filename must match
  - File types: only save attachments matching these extensions
  - Folder: search inbox or sent items

Optionally saves the email itself as .msg alongside attachments.

Usage:
    python scripts/extract_attachments.py
    python scripts/extract_attachments.py --sender-email "alice@example.com,bob@example.com"
    python scripts/extract_attachments.py --keyword "GRN,Invoice"
    python scripts/extract_attachments.py --save-msg true
"""

import os
import sys
from pathlib import Path

# Ensure project root is on sys.path so 'lib' package resolves correctly.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from lib.outlook import (
    get_outlook_folder,
    filter_emails,
    iter_attachments,
    save_attachment,
)
from lib.reporting import write_csv_report
from lib.task_config import parse_task_args, unpack_config, get_output_dir, get_report_path
from lib.utils import sanitize_filename


def extract_attachments(config):
    sender_email, keyword, file_types = unpack_config(
        config, "sender_email", "keyword", "file_types"
    )
    output_folder = get_output_dir(config, config.get("attachments_subdir", "Attachments"))
    report_path = get_report_path(config, config.get("report_file", "Attachments_Report.csv"))
    save_msg = str(config.get("save_msg", "false")).lower() in ("true", "1", "yes")

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

    folder = get_outlook_folder(config.get("folder", "inbox"))
    emails = filter_emails(folder, sender_email=sender_email, keyword=None, verbose=True)

    senders_display = sender_email if isinstance(sender_email, list) else [sender_email or "all"]
    print(f"Extracting attachments from: {', '.join(senders_display)}")
    if keyword_list:
        print(f"  Keyword filter (body or filename): {keyword_list}")
    if ext_filter:
        print(f"  File type filter: {ext_filter}")
    if save_msg:
        print(f"  Also saving .msg files")
    print(f"  Output: {output_folder}\n")

    report = []
    emails_processed = set()  # Track which emails we've saved .msg for
    keyword_body_matches = 0
    keyword_filename_matches = 0
    keyword_skipped = 0

    for info in iter_attachments(emails):
        # Keyword filter: body/subject contains keyword OR any attachment filename matches
        if keyword_list:
            try:
                parent_email = info.attachment_ref.Parent
                subject_lower = (parent_email.Subject or "").lower()
                body_lower = (parent_email.Body or "").lower()

                body_match = any(k.lower() in subject_lower or k.lower() in body_lower
                                 for k in keyword_list)

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
                    keyword_skipped += 1
                    if keyword_skipped <= 5:
                        print(f"  [skip] \"{(parent_email.Subject or '')[:50]}\"")
                    continue
            except Exception as e:
                print(f"  [error] {e}")
                continue

        # File type filter
        if ext_filter:
            file_ext = os.path.splitext(info.filename)[1].lower()
            if file_ext not in ext_filter:
                continue

        # Create per-email subfolder: <sanitized_subject>_<YYYYMMDD_HHMMSS>/
        safe_subject = sanitize_filename(info.email_subject)
        date_str = info.received_time.strftime("%Y%m%d_%H%M%S")
        subfolder_name = f"{safe_subject}_{date_str}"
        subfolder = os.path.join(output_folder, subfolder_name)
        os.makedirs(subfolder, exist_ok=True)

        # Save .msg if requested (once per email)
        if save_msg and subfolder_name not in emails_processed:
            try:
                msg_path = os.path.join(subfolder, f"{safe_subject}_{date_str}.msg")
                info.attachment_ref.Parent.SaveAs(msg_path, 3)
                emails_processed.add(subfolder_name)
            except Exception:
                pass

        dest_path = os.path.join(subfolder, info.filename)

        try:
            save_attachment(info.attachment_ref, dest_path)
            print(f"  Saved: {info.filename} (from: {info.email_sender})")

            report.append({
                "FileName": info.filename,
                "EmailSubject": info.email_subject,
                "SenderEmail": info.email_sender,
                "ReceivedDate": str(info.received_time),
                "SubFolder": subfolder_name,
                "FilePath": dest_path,
            })
        except Exception as e:
            print(f"  Failed: {info.filename} - {e}")

    if write_csv_report(report_path, report):
        print(f"\nReport saved to: {report_path}")

    if keyword_list:
        print(f"\n  [filter] body/subject={keyword_body_matches}, "
              f"filename={keyword_filename_matches}, skipped={keyword_skipped}")

    print(f"\nDone. {len(report)} attachment(s) saved to {output_folder}")


if __name__ == "__main__":
    config = parse_task_args(
        description="Extract attachments from Outlook emails.",
        default_task="extract_attachments",
    )
    extract_attachments(config)
