"""
Task configuration loader.

Reads a YAML file with a ``root`` key defining the base output directory.
All other paths in a task are constructed relative to this root.

Also provides CLI argument parsing that allows config values to be
overridden from the command line.
"""

import argparse
import os
from pathlib import Path
from typing import Dict, Any, List, Optional

import yaml


def load_task_config(task_name: str, config_dir: str = None) -> Dict[str, Any]:
    """
    Load a task configuration YAML file.

    Looks for ``<task_name>.yml`` in the config directory (defaults to a
    ``tasks/`` folder next to this package). The YAML must contain at
    minimum a ``root`` key specifying the base output directory.

    Args:
        task_name: Name of the task (without .yml extension).
        config_dir: Optional path to the directory containing task YAML
                    files. Defaults to ``<project_root>/tasks/``.

    Returns:
        Dictionary with all keys from the YAML file.

    Raises:
        FileNotFoundError: If the YAML file does not exist.
        ValueError: If the YAML file is missing the required ``root`` key.
    """
    if config_dir is None:
        # Default to <project_root>/tasks/
        project_root = Path(__file__).resolve().parent.parent
        config_dir = str(project_root / "tasks")

    config_path = os.path.join(config_dir, f"{task_name}.yml")

    if not os.path.exists(config_path):
        # Try to auto-copy from tasks_defaults/
        import shutil
        project_root = Path(__file__).resolve().parent.parent
        default_path = project_root / "tasks_defaults" / f"{task_name}.yml"
        if default_path.exists():
            os.makedirs(config_dir, exist_ok=True)
            shutil.copy2(str(default_path), config_path)
            print(f"Created tasks/{task_name}.yml from defaults — edit it for your environment.")
        else:
            raise FileNotFoundError(
                f"Task config not found: {config_path}\n"
                f"Run 'python main.py init' to create default configs."
            )

    with open(config_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)

    if not config or "root" not in config:
        raise ValueError(
            f"Task config '{config_path}' must contain a 'root' key."
        )

    return config


def parse_task_args(
    description: str,
    default_task: str,
    config_keys: Optional[List[str]] = None,
    argv: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """
    Parse CLI arguments, merge with YAML defaults, and show a GUI form.

    The GUI form is always shown, pre-populated with values from:
    1. YAML task config (base defaults)
    2. CLI arguments (override YAML values)

    The user can then review/edit all values in the form before running.

    Pass --help to see available CLI flags without launching GUI.

    Args:
        description: Script description shown in ``--help`` and GUI title.
        default_task: Default task config name if ``--task`` is not provided.
        config_keys: List of config keys to expose as CLI flags. If None,
                     the keys are inferred from the YAML file after loading.
        argv: Optional argument list (defaults to sys.argv[1:]).

    Returns:
        Merged config dictionary after GUI confirmation.
    """
    import sys as _sys
    check_argv = argv if argv is not None else _sys.argv[1:]

    # First pass: parse --task to load the right YAML
    pre_parser = argparse.ArgumentParser(add_help=False)
    pre_parser.add_argument("--task", default=default_task)
    pre_args, _remaining = pre_parser.parse_known_args(check_argv)

    config = load_task_config(pre_args.task)

    # Determine which keys to expose
    if config_keys is None:
        config_keys = [k for k in config.keys() if k != "root"]

    all_keys = ["root"] + [k for k in config_keys if k != "root"]

    # Build full parser for --help and CLI override parsing
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument(
        "--task",
        default=default_task,
        help=f"Task config name to load from tasks/ (default: {default_task}).",
    )

    for key in all_keys:
        flag = f"--{key.replace('_', '-')}"
        default_val = config.get(key)
        if isinstance(default_val, bool):
            parser.add_argument(flag, default=None, action="store_true",
                                help=f"Override '{key}' (default from YAML: {default_val}).")
        elif isinstance(default_val, list):
            display_val = ", ".join(str(v) for v in default_val)
            parser.add_argument(flag, default=None, type=str,
                                help=f"Override '{key}' as comma-separated values "
                                     f"(default from YAML: [{display_val}]).")
        else:
            parser.add_argument(flag, default=None, type=str,
                                help=f"Override '{key}' (default from YAML: {default_val}).")

    args = parser.parse_args(check_argv)

    # Merge CLI overrides into config
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

    # Show GUI form pre-populated with merged config
    from lib.gui_config import gui_task_args
    return gui_task_args(description, pre_args.task, config_keys, prefilled_config=config)


def unpack_config(config: Dict[str, Any], *keys: str) -> tuple:
    """
    Unpack multiple config values in a single assignment via tuple unpacking.

    Args:
        config: Task configuration dictionary.
        *keys: Key names to extract. Use "key=default" syntax to specify
               a default value if the key is missing.

    Returns:
        Tuple of values in the same order as the requested keys.

    Example:
        sender_email, keyword, excel_subdir = unpack_config(
            config, "sender_email", "keyword", "excel_subdir=Excel_Files"
        )
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

    Args:
        config: Task configuration dictionary (must have a ``root`` key).
        *subdirs: Zero or more subdirectory names to append to root.

    Returns:
        Absolute path to the (now-existing) output directory.
    """
    path = os.path.join(config["root"], *subdirs)
    os.makedirs(path, exist_ok=True)
    return path


def get_report_path(config: Dict[str, Any], filename: str) -> str:
    """
    Construct a report file path under the task root.

    Args:
        config: Task configuration dictionary (must have a ``root`` key).
        filename: Report filename (e.g. ``GRN_Report.csv``).

    Returns:
        Absolute path for the report file (parent directory is ensured to exist).
    """
    path = os.path.join(config["root"], filename)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    return path
