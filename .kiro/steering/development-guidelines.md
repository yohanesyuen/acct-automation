# Development Guidelines

## Outlook / COM Automation

- Do NOT run or execute anything that invokes Outlook (or any Office COM automation) on the development machine.
- Treat Outlook and the Exchange/MAPI layer as a black box. We write code against its interface but never call it locally.
- Testing and validation should use mocks, stubs, or sample data rather than live COM connections.
- When verifying scripts, limit checks to syntax/compilation — do not attempt runtime execution of Outlook-dependent code.

## Git Workflow

- Create a git commit after every set of changes resulting from a single user prompt.
- Use a concise, descriptive commit message summarizing what was done.
- Stage only the files that were modified as part of that prompt's work.

## Architecture

- **GUI-first**: All scripts open a tkinter form pre-populated with config values. CLI args pre-fill the form but don't bypass it.
- **Config auto-save**: After GUI confirmation, the final config is automatically saved back to `tasks/<task>.yml`.
- **Hardcoded defaults**: Each task has defaults in `lib/task_config.py` → `TASK_DEFAULTS` dict. No `tasks_defaults/` directory exists.
- **Self-bootstrapping**: If a task YAML doesn't exist or is invalid, it's auto-generated from `TASK_DEFAULTS`.
- **Path defaults**: Use uppercase underscore-delimited names for directory values (e.g. `RAW_EMAILS`, `EXCEL_FILES`, `ATTACHMENTS`).

## Adding a New Script

1. Add hardcoded defaults to `TASK_DEFAULTS` in `lib/task_config.py`.
2. Add a display name to `TASK_DISPLAY_NAMES` in `lib/gui_config.py`.
3. Create `scripts/<name>.py` using `parse_task_args()` which handles CLI parsing + GUI form + auto-save.
4. Use `unpack_config(config, "key1", "key2=default")` for clean config extraction.
5. The `tasks/<name>.yml` will auto-generate on first run.

## Script Conventions

- All scripts start with `sys.path.insert(0, ...)` to resolve `lib`.
- Main function accepts a `config` dict (not raw args).
- Entry point uses `parse_task_args(description, default_task)`.
- `sender_email`, `keyword`, `file_types`, `contacts` support lists (YAML arrays or comma-separated CLI).
- Empty list `[]` means "no filter" / "disabled".
- Boolean config values: use `true`/`false` in YAML, `store_true` in argparse.

## Config Precedence

1. Hardcoded defaults (`TASK_DEFAULTS`)
2. Saved YAML (`tasks/<task>.yml`)
3. CLI arguments (pre-fill the GUI)
4. GUI edits (final values, auto-saved back to YAML)

## Dependencies

- `openpyxl` + `xlrd`: Excel reading (no COM)
- `extract_msg`: .msg file parsing
- `pygit2`: git operations
- `pyyaml`: config files
- `win32com.client`: Outlook COM (never run locally in dev)
- `tkinter`: GUI forms (stdlib)
