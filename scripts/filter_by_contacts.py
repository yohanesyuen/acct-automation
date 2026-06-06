"""
Filter processed emails by sender or recipient contacts.

Reads unpacked email folders (each containing metadata.yml + attachments)
from a source directory and copies only those where any specified contact
appears as sender, To, or CC. Generates a report.xlsx.

Usage:
    python scripts/filter_by_contacts.py
    python scripts/filter_by_contacts.py --contacts "alice@example.com,bob@example.com"
"""

import os
import sys
import shutil
from pathlib import Path

# Ensure project root is on sys.path so 'lib' package resolves correctly.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import openpyxl
import yaml

from lib.task_config import parse_task_args, unpack_config, get_output_dir, get_report_path
from lib.utils import sanitize_filename


def _write_excel_report(report_path: str, rows: list) -> bool:
    """Write report rows to an Excel file."""
    if not rows:
        return False

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Filtered Emails Report"

    headers = list(rows[0].keys())
    ws.append(headers)

    for row in rows:
        ws.append([row.get(h, "") for h in headers])

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 50)

    wb.save(report_path)
    return True


def _load_metadata(folder_path: str) -> dict:
    """Load metadata.yml from an email subfolder."""
    metadata_path = os.path.join(folder_path, "metadata.yml")
    if not os.path.exists(metadata_path):
        return None
    with open(metadata_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def _contact_matches(metadata: dict, contacts_lower: list) -> bool:
    """Check if any contact appears in from, to, or cc fields."""
    from_field = (metadata.get("from") or "").lower()
    to_field = (metadata.get("to") or "").lower()
    cc_field = (metadata.get("cc") or "").lower()

    all_fields = f"{from_field} {to_field} {cc_field}"

    return any(contact in all_fields for contact in contacts_lower)


def filter_by_contacts(config):
    source_subdir, dest_subdir, contacts = unpack_config(
        config, "source_subdir=GR_Candidates", "dest_subdir=Filtered_Emails", "contacts"
    )
    report_file = config.get("report_file", "report.xlsx")

    source_dir = os.path.join(config["root"], source_subdir)
    dest_dir = get_output_dir(config, dest_subdir)
    report_path = get_report_path(config, report_file)

    # Normalize contacts to a list
    if contacts is None or contacts == []:
        print("Error: contacts is required. Specify in YAML or via --contacts.")
        sys.exit(1)
    elif isinstance(contacts, str):
        contacts_list = [contacts]
    else:
        contacts_list = list(contacts)

    contacts_lower = [c.lower().strip() for c in contacts_list]

    if not os.path.exists(source_dir):
        print(f"Source directory does not exist: {source_dir}")
        sys.exit(1)

    # Find all subfolders that contain a metadata.yml
    email_folders = [
        d for d in sorted(os.listdir(source_dir))
        if os.path.isdir(os.path.join(source_dir, d))
        and os.path.exists(os.path.join(source_dir, d, "metadata.yml"))
    ]

    if not email_folders:
        print(f"No email folders with metadata.yml found in: {source_dir}")
        return

    print(f"Filtering emails by contacts from: {source_dir}")
    print(f"  Contacts: {contacts_list}")
    print(f"  Destination: {dest_dir}")
    print(f"  Found {len(email_folders)} email folder(s)\n")

    report = []
    matched = 0
    skipped = 0

    for folder_name in email_folders:
        folder_path = os.path.join(source_dir, folder_name)
        metadata = _load_metadata(folder_path)

        if metadata is None:
            skipped += 1
            continue

        if not _contact_matches(metadata, contacts_lower):
            skipped += 1
            continue

        matched += 1

        # Copy entire folder to destination
        dest_folder = os.path.join(dest_dir, folder_name)
        if not os.path.exists(dest_folder):
            shutil.copytree(folder_path, dest_folder)
        else:
            # Folder already copied — skip
            pass

        # List attachments (everything except metadata.yml)
        attachments = [
            f for f in os.listdir(folder_path)
            if f != "metadata.yml" and os.path.isfile(os.path.join(folder_path, f))
        ]

        print(f"  MATCH: {folder_name} (from: {metadata.get('from', 'unknown')})")

        report.append({
            "Folder": folder_name,
            "Subject": metadata.get("subject", ""),
            "From": metadata.get("from", ""),
            "To": metadata.get("to", ""),
            "CC": metadata.get("cc", ""),
            "Date": metadata.get("date", ""),
            "Attachments": ", ".join(attachments),
            "DestPath": dest_folder,
        })

    if _write_excel_report(report_path, report):
        print(f"\nReport saved to: {report_path}")

    print(f"\nDone. {matched} matched, {skipped} skipped out of {len(email_folders)} email(s).")


if __name__ == "__main__":
    config = parse_task_args(
        description="Filter processed emails by sender/recipient contacts.",
        default_task="filter_by_contacts",
    )
    filter_by_contacts(config)
