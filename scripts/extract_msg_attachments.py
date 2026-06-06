"""
Extract attachments from raw .msg files into an Emails_With_Attachments folder.

Scans Root/Raw_Emails for .msg files that contain attachments and extracts
them into per-email subfolders under Root/Emails_With_Attachments.

Usage:
    python scripts/extract_msg_attachments.py
    python scripts/extract_msg_attachments.py --source-subdir Raw_Emails
    python scripts/extract_msg_attachments.py --file-types ".xlsx,.pdf"
"""

import os
import sys
import glob
from pathlib import Path

# Ensure project root is on sys.path so 'lib' package resolves correctly.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import extract_msg

from lib.reporting import write_csv_report
from lib.task_config import parse_task_args, unpack_config, get_output_dir, get_report_path
from lib.utils import sanitize_filename


def extract_msg_attachments(config):
    source_subdir, dest_subdir, file_types = unpack_config(
        config, "source_subdir=Raw_Emails", "dest_subdir=Emails_With_Attachments", "file_types"
    )
    report_file = config.get("report_file", "Msg_Attachments_Report.csv")

    source_dir = os.path.join(config["root"], source_subdir)
    dest_dir = get_output_dir(config, dest_subdir)
    report_path = get_report_path(config, report_file)

    # Normalize file_types filter
    if file_types is None or file_types == []:
        ext_filter = None
    elif isinstance(file_types, str):
        ext_filter = [file_types if file_types.startswith(".") else f".{file_types}"]
    else:
        ext_filter = [ft if ft.startswith(".") else f".{ft}" for ft in file_types]
    if ext_filter:
        ext_filter = [e.lower() for e in ext_filter]

    if not os.path.exists(source_dir):
        print(f"Source directory does not exist: {source_dir}")
        sys.exit(1)

    # Find all .msg files recursively
    msg_files = glob.glob(os.path.join(source_dir, "**", "*.msg"), recursive=True)

    if not msg_files:
        print(f"No .msg files found in: {source_dir}")
        return

    print(f"Extracting attachments from .msg files in: {source_dir}")
    if ext_filter:
        print(f"  File type filter: {ext_filter}")
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

        # Filter attachments by extension if specified
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

        # Create per-email subfolder from .msg filename
        msg_basename = os.path.splitext(os.path.basename(msg_path))[0]
        safe_name = sanitize_filename(msg_basename, max_length=80)
        subfolder = os.path.join(dest_dir, safe_name)
        os.makedirs(subfolder, exist_ok=True)

        for filename, att in valid_attachments:
            output_path = os.path.join(subfolder, filename)

            # Handle duplicates
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

                print(f"  Saved: {safe_name}/{filename}")

                report.append({
                    "MsgFile": os.path.basename(msg_path),
                    "AttachmentName": filename,
                    "SubFolder": safe_name,
                    "DestPath": output_path,
                })
            except Exception as e:
                print(f"  Failed: {filename} - {e}")

    if write_csv_report(report_path, report):
        print(f"\nReport saved to: {report_path}")

    print(f"\nDone. {msgs_with_attachments} email(s) with attachments, "
          f"{total_attachments} file(s) extracted to {dest_dir}")


if __name__ == "__main__":
    config = parse_task_args(
        description="Extract attachments from .msg files into Emails_With_Attachments.",
        default_task="extract_msg_attachments",
    )
    extract_msg_attachments(config)
