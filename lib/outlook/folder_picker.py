"""
Outlook folder selection dialog.

Shows a tkinter Treeview populated from the live Outlook folder tree so the
user can pick which folder (and its subfolders) to search.
"""

import tkinter as tk
from tkinter import ttk, messagebox


def gui_select_outlook_folder(title: str = "Select Outlook Folder") -> object:
    """
    Open a tree dialog showing all Outlook folders and return the selected one.

    Connects to Outlook via COM, walks the mailbox folder tree, and presents it
    in a tkinter Treeview.  The user selects a single folder; the script will
    then search that folder AND all of its subfolders.

    Args:
        title: Window title string.

    Returns:
        The selected Outlook COM folder object, or None if cancelled / error.
    """
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)   # olFolderInbox = 6
        store_root = inbox.Parent               # mailbox root (contains Inbox, Sent, etc.)
    except Exception as e:
        messagebox.showerror("Outlook Error", f"Could not connect to Outlook:\n{e}")
        return None

    # Map treeview node id → COM folder object
    folder_map: dict = {}
    result = {"folder": None}

    window = tk.Tk()
    window.title(title)
    window.resizable(True, True)

    label = ttk.Label(
        window,
        text="Select the folder to search (subfolders are included automatically):",
        font=("Segoe UI", 10),
    )
    label.pack(padx=15, pady=(12, 4), anchor="w")

    frame = ttk.Frame(window)
    frame.pack(padx=15, pady=5, fill=tk.BOTH, expand=True)

    scrollbar_y = ttk.Scrollbar(frame, orient=tk.VERTICAL)
    scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
    scrollbar_x = ttk.Scrollbar(frame, orient=tk.HORIZONTAL)
    scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

    tree = ttk.Treeview(
        frame,
        show="tree",
        selectmode="browse",
        yscrollcommand=scrollbar_y.set,
        xscrollcommand=scrollbar_x.set,
    )
    tree.pack(fill=tk.BOTH, expand=True)
    scrollbar_y.config(command=tree.yview)
    scrollbar_x.config(command=tree.xview)

    def _populate(parent_node: str, folder) -> str:
        try:
            name = folder.Name
        except Exception:
            return ""
        node = tree.insert(parent_node, "end", text=name, open=False)
        folder_map[node] = folder
        try:
            for i in range(1, folder.Folders.Count + 1):
                _populate(node, folder.Folders.Item(i))
        except Exception:
            pass
        return node

    root_node = _populate("", store_root)

    # Expand the top-level mailbox and pre-select Inbox
    if root_node:
        tree.item(root_node, open=True)
        inbox_node = None
        for child in tree.get_children(root_node):
            if tree.item(child, "text").lower() == "inbox":
                inbox_node = child
                break
        if inbox_node:
            tree.selection_set(inbox_node)
            tree.see(inbox_node)
        else:
            tree.selection_set(root_node)

    btn_frame = ttk.Frame(window)
    btn_frame.pack(pady=12)

    def on_ok():
        sel = tree.selection()
        if sel:
            result["folder"] = folder_map.get(sel[0])
        window.destroy()

    def on_cancel():
        window.destroy()

    ok_btn = ttk.Button(btn_frame, text="Select", command=on_ok, width=12)
    ok_btn.pack(side=tk.LEFT, padx=8)
    cancel_btn = ttk.Button(btn_frame, text="Cancel", command=on_cancel, width=12)
    cancel_btn.pack(side=tk.LEFT, padx=8)

    window.bind("<Return>", lambda _e: on_ok())
    window.bind("<Escape>", lambda _e: on_cancel())
    tree.bind("<Double-1>", lambda _e: on_ok())

    window.update_idletasks()
    w, h = 440, 520
    x = (window.winfo_screenwidth() // 2) - (w // 2)
    y = (window.winfo_screenheight() // 2) - (h // 2)
    window.geometry(f"{w}x{h}+{x}+{y}")

    window.mainloop()

    return result["folder"]
