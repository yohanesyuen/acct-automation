from .connection import get_outlook_inbox
from .email_filter import filter_emails, filter_emails_by_sender_and_keyword, filter_emails_by_keyword
from .attachments import (
    AttachmentInfo,
    iter_attachments,
    save_attachment,
    save_attachment_to_temp,
    cleanup_temp_file,
)

__all__ = [
    "get_outlook_inbox",
    "filter_emails",
    "filter_emails_by_sender_and_keyword",
    "filter_emails_by_keyword",
    "AttachmentInfo",
    "iter_attachments",
    "save_attachment",
    "save_attachment_to_temp",
    "cleanup_temp_file",
]
