"""
Filter processed emails by sender or recipient contacts.

Reads .msg files from a source directory (e.g. GR_Candidates) and copies
only those where any of the specified contacts appear as sender, To, or CC
into a destination folder. Generates a report.xlsx and metadata.yml per email.

Usage:
    python scripts/filter_by_contacts.py
    python scripts/filter_by_contacts.py --contacts "alice@example.com,bob@example.com"
"""

import os
import sys
import glob
import shutil
from pathlib import Path

# Ensure project root is on sys.path so 'lib' package resolves correctly.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import extract_msg
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


def _write_metadata(folder_path: str, msg) -> None:
    """Write a metadata.yml file with email details."""
    metadata = {
        "from": msg.sender or "",
        "to": msg.to or "",
        "cc": msg.cc or "",
        "subject": msg.subject or "",
        "date": str(msg.date or ""),
        "body": (msg.body or "")[:2000],
    }

    metadata_path = os.path.join(folder_path, "metadata.yml")
    with open(metadata_path, "w", encoding="utf-8") as f:
        yaml.dump(metadata, f, default_flow_style=False, allow_unicode=True, sort_keys=False)


def _contact_matches(msg, contacts_lower: list) -> bool:
    """Check if any contact appears in sender, To, or CC fields."""
    sender = (msg.sender or "").lower()
    to_field = (msg.to or "").lower()
    cc_field = (msg.cc or "").lower()

    all_fields = f"{sender} {to_field} {cc_field}"

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

    # Find all .msg files recursively
    msg_files = glob.glob(os.path.join(source_dir, "**", "*.msg"), recursive=True)

    if not msg_files:
        print(f"No .msg files found in: {source_dir}")
        return

    print(f"Filtering emails by contacts from: {source_dir}")
    print(f"  Contacts: {contacts_list}")
    print(f"  Destination: {dest_dir}")
    print(f"  Found {len(msg_files)} .msg file(s)\n")

    report = []
    matched = 0
    skipped = 0

    for msg_path in msg_files:
        try:
            msg = extract_msg.Message(msg_path)
        except Exception as e:
            print(f"  Error reading {os.path.basename(msg_path)}: {e}")
            continue

        if not _contact_matches(msg, contacts_lower):
            skipped += 1
            continue

        matched += 1

        # Create per-email subfolder: <sanitized_subject>_<YYYYMMDD_HHMMSS>/
        subject = msg.subject or "No Subject"
        safe_subject = sanitize_filename(subject)
        try:
            msg_date = msg.date
            if msg_date:
                date_str = msg_date.strftime("%Y%m%d_%H%M%S")
            else:
                date_str = "00000000_000000"
        except Exception:
            date_str = "00000000_000000"

        subfolder_name = f"{safe_subject}_{date_str}"
        subfolder = os.path.join(dest_dir, subfolder_name)
        os.makedirs(subfolder, exist_ok=True)

        # Copy the .msg file itself
        dest_msg = os.path.join(subfolder, os.path.basename(msg_path))
        if not os.path.exists(dest_msg):
            shutil.copy2(msg_path, dest_msg)

        # Extract all attachments
        for att in msg.attachments:
            filename = att.longFilename or att.shortFilename
            if not filename:
                continue
            output_path = os.path.join(subfolder, filename)
            if not os.path.exists(output_path):
                try:
                    with open(output_path, "wb") as f:
                        f.write(att.data)
                except Exception:
                    pass

        # Write metadata
        try:
            _write_metadata(subfolder, msg)
        except Exception:
            pass

        print(f"  MATCH: {subfolder_name} (from: {msg.sender or 'unknown'})")

        report.append({
            "SubFolder": subfolder_name,
            "Subject": subject,
            "From": msg.sender or "",
            "To": msg.to or "",
            "CC": msg.cc or "",
            "Date": str(msg.date or ""),
            "Attachments": ", ".join(
                (a.longFilename or a.shortFilename or "") for a in msg.attachments
            ),
            "DestPath": subfolder,
        })

    if _write_excel_report(report_path, report):
        print(f"\nReport saved to: {report_path}")

    print(f"\nDone. {matched} matched, {skipped} skipped out of {len(msg_files)} email(s).")


if __name__ == "__main__":
    config = parse_task_args(
        description="Filter processed emails by sender/recipient contacts.",
        default_task="filter_by_contacts",
    )
    filter_by_contacts(config)
