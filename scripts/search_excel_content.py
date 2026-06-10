"""
Search Excel attachments from Outlook emails for GR numbers in cell content.

Supports two search strategies (configurable via YAML or --search-strategy):
  - "pattern": regex match for GR references (e.g. GR 12345, GR-12345)
  - "gr_numbers": strict 8-digit numbers starting with 2 (e.g. 26000117)

Usage:
    python scripts/search_excel_content.py
    python scripts/search_excel_content.py --search-strategy gr_numbers
    python scripts/search_excel_content.py --task find_gr_numbers
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
from lib.excel_search import search_excel_for_gr_pattern, search_excel_for_gr_numbers
from lib.reporting import write_csv_report
from lib.task_config import parse_task_args, unpack_config, get_output_dir, get_report_path
from lib.utils import is_excel_file, make_date_prefixed_filename


SEARCH_STRATEGIES = {
    "pattern": search_excel_for_gr_pattern,
    "gr_numbers": search_excel_for_gr_numbers,
}


def search_excel_content(config):
    sender_email, keyword, strategy_name = unpack_config(
        config, "sender_email", "keyword", "search_strategy=pattern"
    )
    output_folder = get_output_dir(config, config.get("excel_subdir", "Excel_Files"))
    report_path = get_report_path(config, config.get("report_file", "GRN_Report.csv"))

    search_fn = SEARCH_STRATEGIES.get(strategy_name)
    if search_fn is None:
        print(f"Error: unknown search_strategy '{strategy_name}'")
        print(f"Available: {', '.join(SEARCH_STRATEGIES.keys())}")
        sys.exit(1)

    inbox = get_outlook_inbox()
    emails = filter_emails_by_sender_and_keyword(inbox, sender_email, keyword=keyword)

    print(f"Searching Excel attachments (strategy: {strategy_name}, sender: {sender_email}"
          f"{f', keyword: {keyword}' if keyword else ''})...")

    report = []

    for info in iter_attachments(emails, extension_filter=is_excel_file):
        temp_path = save_attachment_to_temp(info.attachment_ref, info.filename)
        print(f"  Checking: {info.filename}...")

        try:
            gr_numbers = search_fn(temp_path)

            if gr_numbers:
                new_filename = make_date_prefixed_filename(info.received_time, info.filename)
                dest_path = os.path.join(output_folder, new_filename)
                shutil.copy2(temp_path, dest_path)

                print(f"  SAVED: {info.filename} (GR: {', '.join(gr_numbers)})")

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
        os.startfile(output_folder)


if __name__ == "__main__":
    config = parse_task_args(
        description="Search Excel attachments for GR numbers in cell content.",
        default_task="search_excel_content",
    )
    search_excel_content(config)
