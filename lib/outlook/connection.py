"""
Outlook COM connection utilities.

Provides a thin wrapper around the Outlook COM interface for connecting
to the MAPI namespace and retrieving mail folders.
"""

import win32com.client


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
    inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
    return inbox
