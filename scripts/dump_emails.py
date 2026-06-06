"""
Download all emails after a specified date from Inbox and Sent Items.

Saves each email as a .msg file in Raw_Emails/<folder>/ subdirectories.

Usage:
    python scripts/dump_emails.py
    python scripts/dump_emails.py --after-date 2025-01-01
    python scripts/dump_emails.py --after-date 2025-06-01 --folders inbox
"""

import os
import sys
from datetime import datetime
from pathlib import Path

# Ensure project root is on sys.path so 'lib' package resolves correctly.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from lib.outlook import get_outlook_folder
from lib.task_config import parse_task_args, unpack_config, get_output_dir
from lib.utils import sanitize_filename


def dump_emails(config):
    after_date_str, folders_cfg, raw_emails_subdir = unpack_config(
        config, "after_date", "folders", "raw_emails_subdir=Raw_Emails"
    )

    # Parse the date
    if not after_date_str:
        print("Error: after_date is required (format: YYYY-MM-DD).")
        sys.exit(1)

    try:
        after_date = datetime.strptime(after_date_str, "%Y-%m-%d")
    except (ValueError, TypeError):
        print(f"Error: invalid date format '{after_date_str}'. Use YYYY-MM-DD.")
        sys.exit(1)

    # Normalize folders to a list
    if folders_cfg is None or folders_cfg == []:
        folder_names = ["inbox", "sent"]
    elif isinstance(folders_cfg, str):
        folder_names = [folders_cfg]
    else:
        folder_names = list(folders_cfg)

    output_base = get_output_dir(config, raw_emails_subdir)

    print(f"Dumping emails received after: {after_date.strftime('%Y-%m-%d')}")
    print(f"  Folders: {', '.join(folder_names)}")
    print(f"  Output: {output_base}\n")

    total_saved = 0

    for folder_name in folder_names:
        try:
            folder = get_outlook_folder(folder_name)
        except ValueError as e:
            print(f"  Skipping unknown folder '{folder_name}': {e}")
            continue

        folder_output = os.path.join(output_base, folder_name)
        os.makedirs(folder_output, exist_ok=True)

        folder_count = 0
        print(f"  Scanning {folder_name}...")

        for item in folder.Items:
            try:
                received_time = item.ReceivedTime
                # COM datetime to Python datetime comparison
                item_date = datetime(
                    received_time.year, received_time.month, received_time.day,
                    received_time.hour, received_time.minute, received_time.second,
                )
            except Exception:
                continue

            if item_date < after_date:
                continue

            try:
                subject = item.Subject or "No Subject"
                safe_subject = sanitize_filename(subject)
                date_str = item_date.strftime("%Y%m%d_%H%M%S")
                msg_filename = f"{date_str}_{safe_subject}.msg"
                msg_path = os.path.join(folder_output, msg_filename)

                # Skip if already exists
                if os.path.exists(msg_path):
                    continue

                item.SaveAs(msg_path, 3)  # 3 = olMSG format
                folder_count += 1

                if folder_count <= 10:
                    print(f"    Saved: {msg_filename}")
                elif folder_count == 11:
                    print(f"    ... (suppressing further output)")
            except Exception as e:
                print(f"    Failed to save: {subject[:40]} - {e}")

        total_saved += folder_count
        print(f"  {folder_name}: {folder_count} email(s) saved.\n")

    print(f"Done. {total_saved} total email(s) saved to {output_base}")


if __name__ == "__main__":
    config = parse_task_args(
        description="Download all emails after a date from Inbox and Sent Items as .msg files.",
        default_task="dump_emails",
    )
    dump_emails(config)
