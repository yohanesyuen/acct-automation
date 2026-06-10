"""
Tkinter GUI form for task configuration.

Provides:
  - gui_select_task: dropdown to choose which task to run
  - gui_task_args: form to edit config values before running
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog
from typing import Dict, Any, Optional, List
from pathlib import Path


# User-friendly display names for scripts
TASK_DISPLAY_NAMES = {
    "analyze_headers": "Analyze Excel Headers",
    "collect_excel": "Collect Excel from Attachments",
    "dump_emails": "Download Emails (by Date)",
    "extract_attachments": "Extract Attachments (Outlook)",
    "filter_by_contacts": "Filter Emails by Contacts",
    "process_msg_files": "Process .msg Files (Extract Attachments)",
    "search_excel_content": "Search Excel Content for GR Numbers",
}


def _get_display_name(script_name: str) -> str:
    """Get a user-friendly name for a script."""
    return TASK_DISPLAY_NAMES.get(script_name, script_name.replace("_", " ").title())


def _browse_folder(entry: tk.Entry) -> None:
    """Open a folder picker and set the entry value."""
    path = filedialog.askdirectory()
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)


def gui_select_task() -> Optional[str]:
    """
    Show a dropdown to select which task script to run.

    Returns:
        The script name (e.g. "extract_attachments"), or None if cancelled.
    """
    project_root = Path(__file__).resolve().parent.parent
    scripts_dir = project_root / "scripts"

    # Discover available scripts
    scripts = sorted(
        p.stem for p in scripts_dir.glob("*.py")
        if not p.name.startswith("_")
    )

    if not scripts:
        return None

    result = {"selected": None}

    root = tk.Tk()
    root.title("acct-automation — Select Task")
    root.resizable(False, False)

    # Header
    header = ttk.Label(root, text="Select a task to run:", font=("Segoe UI", 11, "bold"))
    header.pack(padx=20, pady=(15, 10))

    # Build display list
    display_names = [_get_display_name(s) for s in scripts]
    name_to_script = dict(zip(display_names, scripts))

    # Listbox with scrollbar
    frame = ttk.Frame(root)
    frame.pack(padx=20, pady=5, fill=tk.BOTH, expand=True)

    scrollbar = ttk.Scrollbar(frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    listbox = tk.Listbox(
        frame, width=50, height=min(len(display_names), 12),
        font=("Segoe UI", 10), selectmode=tk.SINGLE,
        yscrollcommand=scrollbar.set,
    )
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.config(command=listbox.yview)

    for name in display_names:
        listbox.insert(tk.END, name)

    # Select first item by default
    listbox.selection_set(0)

    # Buttons
    btn_frame = ttk.Frame(root)
    btn_frame.pack(pady=15)

    def on_select():
        selection = listbox.curselection()
        if selection:
            display_name = listbox.get(selection[0])
            result["selected"] = name_to_script[display_name]
        root.destroy()

    def on_cancel():
        result["selected"] = None
        root.destroy()

    select_btn = ttk.Button(btn_frame, text="Next →", command=on_select, width=12)
    select_btn.pack(side=tk.LEFT, padx=10)

    cancel_btn = ttk.Button(btn_frame, text="Cancel", command=on_cancel, width=12)
    cancel_btn.pack(side=tk.LEFT, padx=10)

    # Double-click to select
    listbox.bind("<Double-1>", lambda e: on_select())
    root.bind("<Return>", lambda e: on_select())
    root.bind("<Escape>", lambda e: on_cancel())

    # Center window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")

    root.mainloop()

    return result["selected"]


def gui_task_args(
    description: str,
    default_task: str,
    config_keys: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """
    Show a tkinter form for editing task config and return the merged config.

    Loads the YAML defaults, presents them in a form, and returns the
    user-edited values as a config dictionary. If the user cancels, exits.

    Args:
        description: Script description shown as the window title.
        default_task: Default task config name to load.
        config_keys: Optional list of keys to show. If None, shows all.

    Returns:
        Config dictionary with user-edited values.
    """
    from lib.task_config import load_task_config

    config = load_task_config(default_task)

    # Determine which keys to show
    if config_keys is None:
        all_keys = list(config.keys())
    else:
        all_keys = ["root"] + [k for k in config_keys if k != "root"]

    # Result holder
    result = {"cancelled": True}

    # Build the form
    root = tk.Tk()
    root.title(f"Task: {description}")
    root.resizable(True, False)

    # Style
    style = ttk.Style()
    style.configure("TLabel", padding=5)
    style.configure("TEntry", padding=3)
    style.configure("TButton", padding=5)

    # Header
    header = ttk.Label(root, text=description, font=("Segoe UI", 11, "bold"))
    header.grid(row=0, column=0, columnspan=3, padx=10, pady=(10, 5), sticky="w")

    task_label = ttk.Label(root, text=f"Task config: {default_task}.yml", foreground="gray")
    task_label.grid(row=1, column=0, columnspan=3, padx=10, pady=(0, 10), sticky="w")

    entries = {}
    row_idx = 2

    for key in all_keys:
        value = config.get(key, "")

        # Display label (replace underscores with spaces, title case)
        display_name = key.replace("_", " ").title()
        label = ttk.Label(root, text=display_name + ":")
        label.grid(row=row_idx, column=0, padx=(10, 5), pady=3, sticky="e")

        # Convert list values to comma-separated display
        if isinstance(value, list):
            display_value = ", ".join(str(v) for v in value)
        elif isinstance(value, bool):
            display_value = str(value).lower()
        else:
            display_value = str(value) if value is not None else ""

        entry = ttk.Entry(root, width=60)
        entry.insert(0, display_value)
        entry.grid(row=row_idx, column=1, padx=5, pady=3, sticky="ew")
        entries[key] = entry

        # Add browse button for path-like keys
        if key in ("root",) or key.endswith("_subdir") or key.endswith("_dir"):
            browse_btn = ttk.Button(root, text="Browse", width=8,
                                    command=lambda e=entry: _browse_folder(e))
            browse_btn.grid(row=row_idx, column=2, padx=(0, 10), pady=3)

        row_idx += 1

    # Configure column weights for resizing
    root.columnconfigure(1, weight=1)

    # Buttons
    btn_frame = ttk.Frame(root)
    btn_frame.grid(row=row_idx, column=0, columnspan=3, pady=15)

    def on_run():
        result["cancelled"] = False
        root.destroy()

    def on_cancel():
        result["cancelled"] = True
        root.destroy()

    run_btn = ttk.Button(btn_frame, text="▶ Run", command=on_run, width=12)
    run_btn.pack(side=tk.LEFT, padx=10)

    cancel_btn = ttk.Button(btn_frame, text="Cancel", command=on_cancel, width=12)
    cancel_btn.pack(side=tk.LEFT, padx=10)

    # Bind Enter key to Run
    root.bind("<Return>", lambda e: on_run())
    root.bind("<Escape>", lambda e: on_cancel())

    # Center window
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f"+{x}+{y}")

    root.mainloop()

    if result["cancelled"]:
        print("Cancelled by user.")
        sys.exit(0)

    # Merge edited values back into config
    for key in all_keys:
        entry_value = entries[key].get().strip()
        original_value = config.get(key)

        if isinstance(original_value, list):
            # Split comma-separated back into list
            if entry_value:
                config[key] = [v.strip() for v in entry_value.split(",")]
            else:
                config[key] = []
        elif isinstance(original_value, bool):
            config[key] = entry_value.lower() in ("true", "1", "yes")
        else:
            config[key] = entry_value

    return config
