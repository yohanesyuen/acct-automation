"""
Download invoice-related email attachments from Outlook.

Searches the inbox for emails where the subject, body, or any attachment
filename contains "inv" or "invoice". Saves attachments and a metadata.yml
per email subfolder.

This is a pre-configured variant of download_attachments_by_sender.

Usage:
    python scripts/download_invoice_attachments.py
    python scripts/download_invoice_attachments.py --sender-email "vendor@example.com"
    python scripts/download_invoice_attachments.py --no-gui
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from download_attachments_by_sender import download_attachments_by_sender
from lib.task_config import parse_task_args

if __name__ == "__main__":
    config = parse_task_args(
        description="Download invoice-related email attachments.",
        default_task="download_invoice_attachments",
    )
    download_attachments_by_sender(config)
