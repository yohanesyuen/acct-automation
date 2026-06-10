# acct-automation — Project Context

You are assisting with a general-purpose accounting automation tool. The
project automates repetitive tasks around email processing, attachment
extraction, Excel analysis, and data filtering — primarily interfacing
with Outlook and Excel files on Windows.

Below is a complete reference of the project's library modules and
available task scripts. Use this context to understand what already exists
before proposing changes or new code.

**Your response rules:**
- Do NOT immediately write code or provide a solution.
- Instead, reply with a focused follow-up question asking what specific
  task or problem we want to accomplish.
- Only after receiving a clear answer should you proceed with implementation.

---

## Architecture

- **GUI-first**: Running any script opens a tkinter form pre-populated with saved config. CLI args pre-fill the form but don't bypass it.
- **Auto-save**: After the user clicks "Run", config is saved back to `tasks/<task>.yml`. Next run remembers the last-used values.
- **Self-bootstrapping**: No setup required. If `tasks/<task>.yml` doesn't exist or is invalid, it's auto-generated from hardcoded defaults in `lib/task_config.py`.
- **Dynamic loading**: Scripts are loaded via `importlib` — no subprocess spawning.
- **`--help`** still works without GUI for documentation.

## CLI Interface

```
python main.py                       # GUI task selector → config form → run
python main.py generate              # Regenerate this pre_prompt.md
python main.py list                  # List available task scripts
python main.py run <script>          # Open GUI for a specific script
python main.py run <script> --help   # Show CLI flags (no GUI)
python main.py commit -m "msg"       # Git commit with auto file list
python main.py push                  # Git push using GITHUB_PAT from .env
```

### CLI args pre-fill the GUI

```
python main.py run extract_attachments --sender-email "a@x.com,b@y.com" --keyword "GRN"
```

This opens the GUI with sender_email and keyword already filled in. The user
can review/edit and click Run. Final values are saved to YAML.

## Config System

Precedence (highest wins):
1. GUI edits
2. CLI arguments
3. Saved YAML (`tasks/<task>.yml`)
4. Hardcoded defaults (`lib/task_config.py` → `TASK_DEFAULTS`)

### Conventions

- List values: YAML arrays or comma-separated CLI (`--keyword "GRN,Invoice"`)
- Empty list `[]` = disabled/no filter
- Path defaults: uppercase underscore-delimited placeholders (`OUTPUT_DIR`) — forces user to select via Browse
- Booleans: `true`/`false` in YAML
- Field types enforced via `FIELD_TYPES` in `lib/task_config.py` (directory picker, file save dialog, text)

---

## Adding a New Script

1. Add defaults to `TASK_DEFAULTS` in `lib/task_config.py`
2. Add field types to `FIELD_TYPES` if the task has directory/file path fields
3. Add display name to `TASK_DISPLAY_NAMES` in `lib/gui_config.py`
4. Create `scripts/<name>.py`:

```python
"""
Short description of what this script does.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from lib.task_config import parse_task_args, unpack_config, get_output_dir, get_report_path


def my_task(config):
    value_a, value_b = unpack_config(config, "key_a", "key_b=default")
    output_folder = get_output_dir(config, config.get("output_subdir", "OUTPUT"))
    # ... task logic ...


if __name__ == "__main__":
    config = parse_task_args(
        description="Short description of what this script does.",
        default_task="my_task",
    )
    my_task(config)
```

The YAML auto-generates on first run. GUI picks it up automatically.

---

{{SCRIPTS_LISTING}}

---

{{LIB_LISTING}}
