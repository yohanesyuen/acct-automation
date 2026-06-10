"""
Tkinter GUI form for task configuration.

Provides:
  - gui_select_task: dropdown to choose which task to run
  - gui_task_args: form to edit config values before running

Field types are determined by FIELD_TYPES in task_config.py:
  - "directory": folder browser dialog
  - "file_save:<ext>": file save dialog with extension filter
  - "text": plain text entry (default)
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
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


def _get_field_type(key: str) -> str:
    """Get the field type for a config key."""
    from lib.task_config import FIELD_TYPES
    return FIELD_TYPES.get(key, "text")


def _browse_folder(entry: tk.Entry) -> None:
    """Open a folder picker and set the entry value."""
    path = filedialog.askdirectory()
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)


def _browse_file_save(entry: tk.Entry, extensions: str) -> None:
    """Open a file save dialog and set the entry value."""
    ext_list = [e.strip() for e in extensions.split(",")]
    filetypes = [(f"{e.upper()} files", f"*.{e}") for e in ext_list]
    filetypes.append(("All files", "*.*"))

    path = filedialog.asksaveasfilename(
        defaultextension=f".{ext_list[0]}",
        filetypes=filetypes,
    )
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)


def _is_placeholder(value: str) -> bool:
    """Check if a value is a placeholder that needs user input."""
    # Uppercase-only values without path separators are placeholders
    stripped = value.strip()
    if not stripped:
        return True
    # Match patterns like "OUTPUT_DIR", "RAW_EMAILS" (no path separators, all uppercase/underscore)
    import re
    return bool(re.match(r'^[A-Z][A-Z0-9_]*$', stripped))


def _validate_config(config: Dict[str, Any], all_keys: list) -> Optional[str]:
    """Validate config values. Returns error message or None if valid."""
    field_type = None
    for key in all_keys:
        ft = _get_field_type(key)
        value = config.get(key, "")

        if ft == "directory":
            if isinstance(value, str) and _is_placeholder(value):
                return f"'{key.replace('_', ' ').title()}' must be set to a valid directory path.\n\nPlease use the Browse button to select a folder."

    return None


def gui_select_task() -> Optional[str]:
    """
    Show a dropdown to select which task script to run.

    Runs git fetch on startup to check for updates. Shows a "Pull (N behind)"
    button if the local branch is behind the remote.

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

    # Git fetch and check how far behind we are
    commits_behind = 0
    try:
        from lib.git_ops import get_repo, fetch, get_commits_behind, pull
        repo = get_repo()
        fetch(repo)
        commits_behind = get_commits_behind(repo)
    except Exception:
        pass  # Git not configured or no remote — skip silently

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

    def on_pull():
        """Pull latest changes from remote."""
        try:
            from lib.git_ops import get_repo, pull as git_pull
            repo = get_repo()
            updated = git_pull(repo)
            if updated:
                messagebox.showinfo("Pull Complete", "Updated to latest version.\nThe app will now restart.")
                root.destroy()
                # Re-launch by returning a special sentinel
                result["selected"] = "__restart__"
            else:
                messagebox.showinfo("Pull", "Already up to date.")
        except Exception as e:
            messagebox.showerror("Pull Failed", str(e))

    # Pull button (only show if behind)
    if commits_behind > 0:
        pull_btn = ttk.Button(
            btn_frame, text=f"⬇ Pull ({commits_behind} behind)",
            command=on_pull, width=20,
        )
        pull_btn.pack(side=tk.LEFT, padx=10)

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
    prefilled_config: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """
    Show a tkinter form for editing task config and return the final config.

    Uses FIELD_TYPES to determine which picker to show for each field.
    Validates that directory fields are set to real paths (not placeholders).

    Args:
        description: Script description shown as the window title.
        default_task: Default task config name to load.
        config_keys: Optional list of keys to show. If None, shows all.
        prefilled_config: Optional pre-merged config dict to populate the form.

    Returns:
        Config dictionary with user-edited values.
    """
    from lib.task_config import load_task_config

    if prefilled_config is not None:
        config = prefilled_config
    else:
        config = load_task_config(default_task)

    # Determine which keys to show
    if config_keys is None:
        all_keys = list(config.keys())
    else:
        all_keys = ["root"] + [k for k in config_keys if k != "root"]

    # Result holder
    result = {"cancelled": True}

    # Build the form
    window = tk.Tk()
    window.title(f"Task: {description}")
    window.resizable(True, False)

    # Style
    style = ttk.Style()
    style.configure("TLabel", padding=5)
    style.configure("TEntry", padding=3)
    style.configure("TButton", padding=5)
    style.configure("Required.TEntry", fieldbackground="#fff3cd")

    # Header
    header = ttk.Label(window, text=description, font=("Segoe UI", 11, "bold"))
    header.grid(row=0, column=0, columnspan=3, padx=10, pady=(10, 5), sticky="w")

    task_label = ttk.Label(window, text=f"Config: tasks/{default_task}.yml", foreground="gray")
    task_label.grid(row=1, column=0, columnspan=3, padx=10, pady=(0, 10), sticky="w")

    entries = {}
    row_idx = 2

    for key in all_keys:
        value = config.get(key, "")
        field_type = _get_field_type(key)

        # Display label
        display_name = key.replace("_", " ").title()
        label = ttk.Label(window, text=display_name + ":")
        label.grid(row=row_idx, column=0, padx=(10, 5), pady=3, sticky="e")

        # Convert list values to comma-separated display
        if isinstance(value, list):
            display_value = ", ".join(str(v) for v in value)
        elif isinstance(value, bool):
            display_value = str(value).lower()
        else:
            display_value = str(value) if value is not None else ""

        entry = ttk.Entry(window, width=60)
        entry.insert(0, display_value)
        entry.grid(row=row_idx, column=1, padx=5, pady=3, sticky="ew")
        entries[key] = entry

        # Highlight placeholder values
        if field_type == "directory" and _is_placeholder(display_value):
            entry.configure(style="Required.TEntry")

        # Add appropriate browse button based on field type
        if field_type == "directory":
            browse_btn = ttk.Button(window, text="Browse…", width=8,
                                    command=lambda e=entry: _browse_folder(e))
            browse_btn.grid(row=row_idx, column=2, padx=(0, 10), pady=3)
        elif field_type.startswith("file_save:"):
            extensions = field_type.split(":", 1)[1]
            browse_btn = ttk.Button(window, text="Browse…", width=8,
                                    command=lambda e=entry, ext=extensions: _browse_file_save(e, ext))
            browse_btn.grid(row=row_idx, column=2, padx=(0, 10), pady=3)

        row_idx += 1

    # Configure column weights for resizing
    window.columnconfigure(1, weight=1)

    # Buttons
    btn_frame = ttk.Frame(window)
    btn_frame.grid(row=row_idx, column=0, columnspan=3, pady=15)

    def on_run():
        # Merge values first for validation
        for key in all_keys:
            entry_value = entries[key].get().strip()
            original_value = config.get(key)
            if isinstance(original_value, list):
                config[key] = [v.strip() for v in entry_value.split(",")] if entry_value else []
            elif isinstance(original_value, bool):
                config[key] = entry_value.lower() in ("true", "1", "yes")
            else:
                config[key] = entry_value

        # Validate
        error = _validate_config(config, all_keys)
        if error:
            messagebox.showwarning("Configuration Required", error)
            return  # Don't close — let user fix it

        result["cancelled"] = False
        window.destroy()

    def on_cancel():
        result["cancelled"] = True
        window.destroy()

    run_btn = ttk.Button(btn_frame, text="▶ Run", command=on_run, width=12)
    run_btn.pack(side=tk.LEFT, padx=10)

    cancel_btn = ttk.Button(btn_frame, text="Cancel", command=on_cancel, width=12)
    cancel_btn.pack(side=tk.LEFT, padx=10)

    # Bind keys
    window.bind("<Return>", lambda e: on_run())
    window.bind("<Escape>", lambda e: on_cancel())

    # Center window
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f"+{x}+{y}")

    window.mainloop()

    if result["cancelled"]:
        print("Cancelled by user.")
        sys.exit(0)

    return config
