"""
Tkinter GUI form for task configuration.

Displays a form with all config keys pre-filled from the YAML defaults.
The user can edit values and click Run to proceed, or Cancel to abort.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Dict, Any, Optional, List


def _browse_folder(entry: tk.Entry) -> None:
    """Open a folder picker and set the entry value."""
    path = filedialog.askdirectory()
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)


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

    run_btn = ttk.Button(btn_frame, text="Run", command=on_run, width=12)
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
