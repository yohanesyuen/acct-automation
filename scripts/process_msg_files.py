"""
Process .msg files: extract attachments with metadata.

Scans a source directory for .msg files, extracts attachments (filtered by
type), creates per-email subfolders with a metadata.yml, and generates a
report.xlsx.

Default: only extracts emails containing Excel attachments.
Set file_types to empty to extract all.

Usage:
    python scripts/process_msg_files.py
    python scripts/process_msg_files.py --file-types ".xlsx,.pdf"
    python scripts/process_msg_files.py --file-types "" (all types)
    python scripts/process_msg_files.py --source-subdir Raw_Emails/inbox
"""

import os
import sys
import glob
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
    ws.title = "Extracted Attachments"

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


def process_msg_files(config):
    source_subdir, dest_subdir, file_types = unpack_config(
        config, "source_subdir=Raw_Emails", "dest_subdir=Emails_With_Attachments", "file_types"
    )
    report_file = config.get("report_file", "report.xlsx")

    source_dir = os.path.join(config["root"], source_subdir)
    dest_dir = get_output_dir(config, dest_subdir)
    report_path = get_report_path(config, report_file)

    # Ensure source directory exists
    os.makedirs(source_dir, exist_ok=True)

    # Normalize file_types filter — default to Excel extensions
    if file_types is None:
        ext_filter = [".xlsx", ".xls", ".xlsm", ".xlsb"]
    elif file_types == [] or file_types == "":
        ext_filter = None  # No filter — extract all
    elif isinstance(file_types, str):
        ext_filter = [file_types if file_types.startswith(".") else f".{file_types}"]
    else:
        ext_filter = [ft if ft.startswith(".") else f".{ft}" for ft in file_types]
    if ext_filter:
        ext_filter = [e.lower() for e in ext_filter]

    msg_files = glob.glob(os.path.join(source_dir, "**", "*.msg"), recursive=True)

    if not msg_files:
        print(f"No .msg files found in: {source_dir}")
        return

    print(f"Processing .msg files from: {source_dir}")
    if ext_filter:
        print(f"  File type filter: {ext_filter}")
    else:
        print(f"  File type filter: all")
    print(f"  Destination: {dest_dir}")
    print(f"  Found {len(msg_files)} .msg file(s)\n")

    report = []
    total_attachments = 0
    msgs_with_attachments = 0

    for msg_path in msg_files:
        try:
            msg = extract_msg.Message(msg_path)
        except Exception as e:
            print(f"  Error reading {os.path.basename(msg_path)}: {e}")
            continue

        attachments = msg.attachments
        if not attachments:
            continue

        # Filter attachments by extension
        valid_attachments = []
        for att in attachments:
            filename = att.longFilename or att.shortFilename
            if not filename:
                continue
            if ext_filter:
                file_ext = os.path.splitext(filename)[1].lower()
                if file_ext not in ext_filter:
                    continue
            valid_attachments.append((filename, att))

        if not valid_attachments:
            continue

        msgs_with_attachments += 1

        # Create per-email subfolder: <sanitized_subject>_<YYYYMMDD_HHMMSS>/
        subject = msg.subject or "No Subject"
        safe_subject = sanitize_filename(subject)
        try:
            msg_date = msg.date
            date_str = msg_date.strftime("%Y%m%d_%H%M%S") if msg_date else "00000000_000000"
        except Exception:
            date_str = "00000000_000000"

        subfolder_name = f"{safe_subject}_{date_str}"
        subfolder = os.path.join(dest_dir, subfolder_name)
        os.makedirs(subfolder, exist_ok=True)

        # Write metadata.yml
        try:
            _write_metadata(subfolder, msg)
        except Exception as e:
            print(f"  Warning: metadata failed for {subfolder_name}: {e}")

        for filename, att in valid_attachments:
            output_path = os.path.join(subfolder, filename)

            if os.path.exists(output_path):
                base, ext = os.path.splitext(filename)
                counter = 1
                while os.path.exists(output_path):
                    output_path = os.path.join(subfolder, f"{base}_{counter}{ext}")
                    counter += 1

            try:
                with open(output_path, "wb") as f:
                    f.write(att.data)
                total_attachments += 1
                print(f"  Saved: {subfolder_name}/{filename}")

                report.append({
                    "SubFolder": subfolder_name,
                    "FileName": filename,
                    "EmailSubject": subject,
                    "From": msg.sender or "",
                    "To": msg.to or "",
                    "Date": str(msg.date or ""),
                    "FilePath": output_path,
                })
            except Exception as e:
                print(f"  Failed: {filename} - {e}")

    if _write_excel_report(report_path, report):
        print(f"\nReport saved to: {report_path}")

    print(f"\nDone. {msgs_with_attachments} email(s), {total_attachments} attachment(s) extracted.")


if __name__ == "__main__":
    config = parse_task_args(
        description="Process .msg files and extract attachments with metadata.",
        default_task="process_msg_files",
    )
    process_msg_files(config)
