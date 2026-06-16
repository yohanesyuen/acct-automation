"""
Task configuration loader with hardcoded defaults and auto-save.

Each task has a hardcoded default config. On first run, a YAML file is
generated in tasks/. CLI args and GUI edits automatically update the YAML.
"""

import argparse
import os
from pathlib import Path
from typing import Dict, Any, List, Optional

import yaml


PROJECT_ROOT = Path(__file__).resolve().parent.parent
TASKS_DIR = PROJECT_ROOT / "tasks"


# ---------------------------------------------------------------------------
# Hardcoded default configs for each task
# ---------------------------------------------------------------------------

TASK_DEFAULTS: Dict[str, Dict[str, Any]] = {
    "download_attachments_by_sender": {
        "root": "OUTPUT_DIR",
        "folder": "inbox",
        "sender_email": [],
        "keyword": [],
        "file_types": [],
        "save_msg": False,
        "attachments_subdir": "ATTACHMENTS",
        "report_file": "Attachments_Report.csv",
    },
    "download_invoice_attachments": {
        "root": "OUTPUT_DIR",
        "folder": "inbox",
        "sender_email": [],
        "keyword": ["inv", "invoice"],
        "file_types": [],
        "save_msg": False,
        "attachments_subdir": "INVOICE_ATTACHMENTS",
        "report_file": "Invoice_Attachments_Report.csv",
    },
    "dump_emails": {
        "root": "OUTPUT_DIR",
        "after_date": "2025-01-01",
        "folders": ["inbox", "sent"],
        "raw_emails_subdir": "RAW_EMAILS",
    },
    "process_msg_files": {
        "root": "OUTPUT_DIR",
        "source_subdir": "RAW_EMAILS",
        "dest_subdir": "EMAILS_WITH_ATTACHMENTS",
        "file_types": [".xlsx", ".xls", ".xlsm", ".xlsb"],
        "report_file": "report.xlsx",
    },
    "filter_by_contacts": {
        "root": "OUTPUT_DIR",
        "source_subdir": "GR_CANDIDATES",
        "dest_subdir": "FILTERED_EMAILS",
        "contacts": [],
        "report_file": "report.xlsx",
    },
    "search_excel_content": {
        "root": "OUTPUT_DIR",
        "sender_email": [],
        "keyword": "GRN",
        "search_strategy": "pattern",
        "excel_subdir": "EXCEL_FILES",
        "report_file": "GRN_Report.csv",
        "open_folder_on_complete": False,
    },
    "collect_excel": {
        "root": "OUTPUT_DIR",
        "attachments_subdir": "ATTACHMENTS",
        "excel_subdir": "EXCEL_FILES",
        "report_file": "Excel_Collect_Report.csv",
    },
    "analyze_headers": {
        "root": "OUTPUT_DIR",
        "search_term": ["GR"],
        "exclude_term": [],
        "source_subdir": "EXCEL_FILES",
        "report_file": "Header_Analysis_Report.csv",
    },
    "extract_invoices": {
        "root": "OUTPUT_DIR",
        "pdf_dir": "",
        "auto_detect_vendor": True,
        "vendor": "UIE Industrial",
    },
}


# Field type hints for the GUI — determines which widget to show
# "directory": folder browser
# "file_save:<ext>": file save dialog with extension filter
# "checkbox": boolean checkbutton
# "dropdown:<opt1,opt2,...>": read-only combobox with the given options
# "text": plain text entry (default)
FIELD_TYPES: Dict[str, str] = {
    "root": "directory",
    "pdf_dir": "directory",
    "auto_detect_vendor": "checkbox",
    "vendor": "dropdown:UIE Industrial,LP Construction,MJM Services",
    "attachments_subdir": "text",
    "excel_subdir": "text",
    "raw_emails_subdir": "text",
    "source_subdir": "text",
    "dest_subdir": "text",
    "report_file": "file_save:csv,xlsx",
}


# Fields that are only shown in the GUI when a controlling field holds a
# specific value. Maps a field key -> (controller_key, value_that_shows_it).
# When the controller is a checkbox, toggling it dynamically shows/hides the
# dependent field without rebuilding the form.
CONDITIONAL_FIELDS: Dict[str, tuple] = {
    "vendor": ("auto_detect_vendor", False),
}


# ---------------------------------------------------------------------------
# Config loading / saving
# ---------------------------------------------------------------------------

def _get_default_config(task_name: str) -> Dict[str, Any]:
    """Get the hardcoded default config for a task."""
    if task_name in TASK_DEFAULTS:
        return dict(TASK_DEFAULTS[task_name])  # shallow copy
    # Fallback minimal config
    return {"root": "OUTPUT_DIR"}


def _is_valid_config(config: Any) -> bool:
    """Check if a loaded config is valid (dict with a 'root' key)."""
    return isinstance(config, dict) and "root" in config


def save_task_config(task_name: str, config: Dict[str, Any], config_dir: str = None) -> str:
    """
    Save a task config dictionary to its YAML file.

    Args:
        task_name: Task name (becomes the filename).
        config: Config dictionary to write.
        config_dir: Optional directory path. Defaults to tasks/.

    Returns:
        Path to the written file.
    """
    if config_dir is None:
        config_dir = str(TASKS_DIR)

    os.makedirs(config_dir, exist_ok=True)
    config_path = os.path.join(config_dir, f"{task_name}.yml")

    with open(config_path, "w", encoding="utf-8") as f:
        yaml.dump(config, f, default_flow_style=False, allow_unicode=True, sort_keys=False)

    return config_path


def load_task_config(task_name: str, config_dir: str = None) -> Dict[str, Any]:
    """
    Load a task configuration YAML file, creating from defaults if needed.

    If the file doesn't exist or is invalid, generates a fresh one from
    hardcoded defaults and saves it.

    Args:
        task_name: Name of the task (without .yml extension).
        config_dir: Optional directory path. Defaults to tasks/.

    Returns:
        Dictionary with all config keys.
    """
    if config_dir is None:
        config_dir = str(TASKS_DIR)

    config_path = os.path.join(config_dir, f"{task_name}.yml")

    # Try to load existing config
    if os.path.exists(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                config = yaml.safe_load(f)
            if _is_valid_config(config):
                defaults = _get_default_config(task_name)
                changed = False

                # Ensure all default keys exist (for forward compatibility)
                for key, value in defaults.items():
                    if key not in config:
                        config[key] = value
                        changed = True

                # Prune keys removed from the schema (e.g. retired fields) so
                # stale entries don't linger in the YAML or show up in the GUI.
                # Only prune for known tasks where defaults define the schema.
                if task_name in TASK_DEFAULTS:
                    for key in [k for k in config if k not in defaults]:
                        del config[key]
                        changed = True

                # Persist the cleaned-up config so the file self-heals.
                if changed:
                    save_task_config(task_name, config, config_dir)
                return config
        except Exception:
            pass
        # Invalid config — overwrite with defaults
        print(f"  [config] Invalid YAML in tasks/{task_name}.yml — regenerating from defaults.")

    # Generate from hardcoded defaults
    config = _get_default_config(task_name)
    os.makedirs(config_dir, exist_ok=True)
    save_task_config(task_name, config, config_dir)
    print(f"  [config] Created tasks/{task_name}.yml — edit for your environment.")

    return config


# ---------------------------------------------------------------------------
# CLI + GUI integration
# ---------------------------------------------------------------------------

def parse_task_args(
    description: str,
    default_task: str,
    config_keys: Optional[List[str]] = None,
    argv: Optional[List[str]] = None,
    no_gui: bool = False,
) -> Dict[str, Any]:
    """
    Parse CLI arguments, merge with YAML, show GUI, and auto-save changes.

    Flow:
    1. Load YAML (or generate from defaults if missing/invalid)
    2. Parse CLI args and merge overrides
    3. Show GUI form pre-populated with merged values (skipped if no_gui=True)
    4. Save final config back to YAML

    Args:
        description: Script description for --help and GUI title.
        default_task: Default task config name.
        config_keys: Optional list of keys to expose. If None, uses all.
        argv: Optional argument list (defaults to sys.argv[1:]).
        no_gui: If True, skip the GUI dialog and use YAML + CLI values directly.

    Returns:
        Final config dictionary (after GUI confirmation).
    """
    import sys as _sys
    check_argv = argv if argv is not None else _sys.argv[1:]

    # First pass: determine which task to load
    pre_parser = argparse.ArgumentParser(add_help=False)
    pre_parser.add_argument("--task", default=default_task)
    pre_args, _remaining = pre_parser.parse_known_args(check_argv)

    task_name = pre_args.task
    config = load_task_config(task_name)

    # Determine which keys to expose
    if config_keys is None:
        config_keys = [k for k in config.keys() if k != "root"]

    all_keys = ["root"] + [k for k in config_keys if k != "root"]

    # Build argparse for --help and CLI parsing
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument(
        "--task",
        default=default_task,
        help=f"Task config name (default: {default_task}).",
    )

    for key in all_keys:
        flag = f"--{key.replace('_', '-')}"
        default_val = config.get(key)
        if isinstance(default_val, bool):
            parser.add_argument(flag, default=None, action="store_true",
                                help=f"'{key}' (current: {default_val}).")
        elif isinstance(default_val, list):
            display_val = ", ".join(str(v) for v in default_val)
            parser.add_argument(flag, default=None, type=str,
                                help=f"'{key}' comma-separated (current: [{display_val}]).")
        else:
            parser.add_argument(flag, default=None, type=str,
                                help=f"'{key}' (current: {default_val}).")

    args = parser.parse_args(check_argv)

    # Merge CLI overrides
    for key in all_keys:
        cli_value = getattr(args, key, None)
        if cli_value is not None:
            if isinstance(cli_value, str) and "," in cli_value:
                config[key] = [v.strip() for v in cli_value.split(",")]
            else:
                yaml_val = config.get(key)
                if isinstance(yaml_val, list) and isinstance(cli_value, str):
                    config[key] = [cli_value.strip()]
                else:
                    config[key] = cli_value

    if no_gui:
        final_config = config
    else:
        from lib.gui_config import gui_task_args
        final_config = gui_task_args(description, task_name, config_keys, prefilled_config=config)

    # Auto-save the final config back to YAML
    save_task_config(task_name, final_config)

    # Expose no_gui flag to the script without persisting it
    final_config["_no_gui"] = no_gui

    return final_config


# ---------------------------------------------------------------------------
# Utility functions
# ---------------------------------------------------------------------------

def unpack_config(config: Dict[str, Any], *keys: str) -> tuple:
    """
    Unpack multiple config values in a single assignment via tuple unpacking.

    Args:
        config: Task configuration dictionary.
        *keys: Key names to extract. Use "key=default" syntax for defaults.

    Returns:
        Tuple of values in the same order as the requested keys.
    """
    values = []
    for key in keys:
        if "=" in key:
            key_name, default = key.split("=", 1)
            values.append(config.get(key_name.strip(), default.strip()))
        else:
            values.append(config.get(key.strip()))
    return tuple(values)


def get_output_dir(config: Dict[str, Any], *subdirs: str) -> str:
    """
    Construct an output directory path from the task root and subdirectories.

    Creates the directory on disk if it does not already exist.
    """
    path = os.path.join(config["root"], *subdirs)
    os.makedirs(path, exist_ok=True)
    return path


def get_report_path(config: Dict[str, Any], filename: str) -> str:
    """
    Construct a report file path under the task root.

    Parent directory is ensured to exist.
    """
    path = os.path.join(config["root"], filename)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    return path
