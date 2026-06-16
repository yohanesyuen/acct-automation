from .connection import get_outlook_inbox, get_outlook_sent_items, get_outlook_folder, iter_subfolders
from .email_filter import filter_emails, filter_emails_by_sender_and_keyword, filter_emails_by_keyword
from .attachments import (
    AttachmentInfo,
    iter_attachments,
    save_attachment,
    save_attachment_to_temp,
    cleanup_temp_file,
)
from .folder_picker import gui_select_outlook_folder

__all__ = [
    "get_outlook_inbox",
    "get_outlook_sent_items",
    "get_outlook_folder",
    "iter_subfolders",
    "filter_emails",
    "filter_emails_by_sender_and_keyword",
    "filter_emails_by_keyword",
    "AttachmentInfo",
    "iter_attachments",
    "save_attachment",
    "save_attachment_to_temp",
    "cleanup_temp_file",
    "gui_select_outlook_folder",
]
