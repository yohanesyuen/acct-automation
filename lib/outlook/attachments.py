"""
Attachment iteration and saving utilities.

Provides generators and helpers that flatten the nested email → attachment
loop into a single iterable stream, reducing nesting in calling code.
"""

import os
import shutil
import tempfile
from typing import Generator, List, Optional, Callable, Tuple
from dataclasses import dataclass


@dataclass
class AttachmentInfo:
    """Metadata about an email attachment, extracted for downstream processing."""
    filename: str
    email_subject: str
    email_sender: str
    received_time: object  # datetime-like COM object
    attachment_ref: object  # Outlook Attachment COM object


def iter_attachments(
    emails,
    extension_filter: Optional[Callable[[str], bool]] = None,
) -> Generator[AttachmentInfo, None, None]:
    """
    Flatten email iteration into a stream of attachment metadata.

    Iterates over emails and their attachments, optionally filtering by
    file extension. This eliminates the common two-level for-loop nesting.

    Args:
        emails: Iterable of Outlook MailItem COM objects (e.g. from a filter).
        extension_filter: Optional callable that takes a filename and returns
                         True if the attachment should be yielded.

    Yields:
        AttachmentInfo for each attachment that passes the filter.
    """
    for email in emails:
        try:
            subject = email.Subject or ""
            sender = email.SenderEmailAddress or ""
            received_time = email.ReceivedTime
        except Exception:
            continue

        for i in range(1, email.Attachments.Count + 1):
            attachment = email.Attachments.Item(i)
            filename = attachment.FileName

            if extension_filter and not extension_filter(filename):
                continue

            yield AttachmentInfo(
                filename=filename,
                email_subject=subject,
                email_sender=sender,
                received_time=received_time,
                attachment_ref=attachment,
            )


def save_attachment(attachment_ref, dest_path: str) -> str:
    """
    Save an Outlook attachment to disk.

    Args:
        attachment_ref: Outlook Attachment COM object.
        dest_path: Full destination file path.

    Returns:
        The path the file was saved to.
    """
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    attachment_ref.SaveAsFile(dest_path)
    return dest_path


def save_attachment_to_temp(attachment_ref, filename: str) -> str:
    """
    Save an attachment to a temporary directory for inspection.

    Args:
        attachment_ref: Outlook Attachment COM object.
        filename: Original filename to use in temp dir.

    Returns:
        Full path to the temp file.
    """
    temp_path = os.path.join(tempfile.gettempdir(), filename)
    attachment_ref.SaveAsFile(temp_path)
    return temp_path


def cleanup_temp_file(path: str) -> None:
    """Silently remove a temporary file if it exists."""
    try:
        if os.path.exists(path):
            os.remove(path)
    except OSError:
        pass
