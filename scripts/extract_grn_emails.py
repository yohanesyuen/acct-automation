"""
Extract GRN-related emails and their attachments from Outlook Inbox.

Saves matching emails as .msg files and extracts all their attachments
into per-email subfolders.

Usage:
    python scripts/extract_grn_emails.py
"""

import os
import sys
from pathlib import Path

# Ensure project root is on sys.path so 'lib' package resolves correctly.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from lib.outlook import get_outlook_inbox, filter_emails_by_keyword
from lib.task_config import parse_task_args, unpack_config, get_output_dir
from lib.utils import sanitize_filename


def extract_grn_emails(config):
    (keyword,) = unpack_config(config, "keyword")
    email_folder = get_output_dir(config, config.get("emails_subdir", "Emails"))
    attachment_folder = get_output_dir(config, config.get("attachments_subdir", "Attachments"))

    inbox = get_outlook_inbox()

    print(f"Searching for emails containing '{keyword}' with attachments...")

    count = 0
    total_attachments = 0

    for email in filter_emails_by_keyword(inbox, keyword):
        subject = email.Subject or ""
        safe_subject = sanitize_filename(subject)
        date_str = email.ReceivedTime.strftime("%Y%m%d_%H%M%S")

        # Save all attachments into a per-email subfolder
        subfolder = os.path.join(attachment_folder, f"{safe_subject}_{date_str}")
        os.makedirs(subfolder, exist_ok=True)

        for i in range(1, email.Attachments.Count + 1):
            attachment = email.Attachments.Item(i)
            try:
                attachment.SaveAsFile(os.path.join(subfolder, attachment.FileName))
                total_attachments += 1
                print(f"  Saved attachment: {attachment.FileName}")
            except Exception as e:
                print(f"  Failed: {attachment.FileName} - {e}")

        # Save email as .msg
        try:
            msg_path = os.path.join(email_folder, f"{safe_subject}_{date_str}.msg")
            email.SaveAs(msg_path, 3)  # 3 = olMSG format
            count += 1
            print(f"Saved email: {subject}")
        except Exception as e:
            print(f"Failed to save email: {subject} - {e}")

    print(f"\nDone. {count} email(s), {total_attachments} attachment(s) saved to {config['root']}")


if __name__ == "__main__":
    config = parse_task_args(
        description="Extract GRN-related emails and their attachments from Outlook Inbox.",
        default_task="extract_grn_emails",
    )
    extract_grn_emails(config)
