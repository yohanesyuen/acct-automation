"""
Outlook COM connection utilities.

Provides a thin wrapper around the Outlook COM interface for connecting
to the MAPI namespace and retrieving mail folders.
"""

import win32com.client


# Outlook folder IDs
_FOLDER_INBOX = 6
_FOLDER_SENT_ITEMS = 5


def iter_subfolders(folder):
    """Recursively yield a folder and all of its subfolders (depth-first)."""
    yield folder
    try:
        for i in range(1, folder.Folders.Count + 1):
            yield from iter_subfolders(folder.Folders.Item(i))
    except Exception:
        pass


def get_outlook_inbox():
    """
    Connect to the running Outlook instance and return the Inbox folder.

    Uses the Outlook COM automation interface to access the default
    MAPI Inbox folder (folder ID 6).

    Returns:
        A COM object representing the Outlook Inbox folder, supporting
        iteration over .Items and access to mail item properties.

    Raises:
        Exception: If Outlook is not running or COM connection fails.
    """
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(_FOLDER_INBOX)
    return inbox


def get_outlook_sent_items():
    """
    Connect to the running Outlook instance and return the Sent Items folder.

    Uses the Outlook COM automation interface to access the default
    MAPI Sent Items folder (folder ID 5).

    Returns:
        A COM object representing the Outlook Sent Items folder, supporting
        iteration over .Items and access to mail item properties.

    Raises:
        Exception: If Outlook is not running or COM connection fails.
    """
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    sent = namespace.GetDefaultFolder(_FOLDER_SENT_ITEMS)
    return sent


def get_outlook_folder(folder_name: str = "inbox"):
    """
    Connect to the running Outlook instance and return the requested folder.

    Supported folder names: "inbox", "sent", "sent_items".

    Args:
        folder_name: Name of the folder to retrieve (case-insensitive).

    Returns:
        A COM object representing the requested Outlook folder.

    Raises:
        ValueError: If the folder name is not recognized.
        Exception: If Outlook is not running or COM connection fails.
    """
    folder_map = {
        "inbox": _FOLDER_INBOX,
        "sent": _FOLDER_SENT_ITEMS,
        "sent_items": _FOLDER_SENT_ITEMS,
    }

    key = folder_name.lower().strip()
    if key not in folder_map:
        raise ValueError(
            f"Unknown folder '{folder_name}'. "
            f"Available: {', '.join(folder_map.keys())}"
        )

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    return namespace.GetDefaultFolder(folder_map[key])
