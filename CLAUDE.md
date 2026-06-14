# acct-automation

General-purpose accounting automation tool. Automates repetitive tasks: email processing, attachment extraction, Excel analysis, data filtering, and report generation. New tasks are standalone scripts following the established patterns.

## Architecture

- **GUI-first**: Scripts open a tkinter form pre-populated with config values. CLI args pre-fill the form but don't bypass it — unless `--no-gui` is passed, which skips the dialog entirely.
- **No-arg scripts**: Entry functions that take no parameters run directly without any config GUI.
- **Config auto-save**: After GUI confirmation, the final config is saved back to `tasks/<task>.yml`.
- **Hardcoded defaults**: Each task has defaults in `lib/task_config.py` → `TASK_DEFAULTS`. No `tasks_defaults/` directory.
- **Self-bootstrapping**: Missing or invalid task YAMLs are auto-generated from `TASK_DEFAULTS`.
- **Path defaults**: Use uppercase underscore-delimited placeholders (e.g. `OUTPUT_DIR`, `RAW_EMAILS`) to force the user to browse for a real path.
- **Field types**: `FIELD_TYPES` in `task_config.py` controls GUI behavior (directory picker, file save dialog, or text entry).
- **Dynamic loading**: Scripts are loaded via `importlib` — no subprocess.
- **CLI module**: Subcommand handlers and parser live in `lib/cli.py`. `main.py` is the entry point and introspection only.

## Config Precedence

1. Hardcoded defaults (`TASK_DEFAULTS`)
2. Saved YAML (`tasks/<task>.yml`)
3. CLI arguments (pre-fill the GUI)
4. GUI edits (final values, auto-saved back to YAML)

## Adding a New Script

1. Add hardcoded defaults to `TASK_DEFAULTS` in `lib/task_config.py`.
2. Add field type entries to `FIELD_TYPES` for any directory/file fields.
3. Add a display name to `TASK_DISPLAY_NAMES` in `lib/gui_config.py`.
4. Create `scripts/<name>.py` using `parse_task_args()`.
5. Use `unpack_config(config, "key1", "key2=default")` for clean config extraction.
6. The `tasks/<name>.yml` auto-generates on first run.

## Script Conventions

- All scripts start with `sys.path.insert(0, ...)` to resolve `lib`.
- Main function accepts a `config` dict, not raw args.
- Entry point uses `parse_task_args(description, default_task)`.
- `sender_email`, `keyword`, `file_types`, `contacts` support lists (YAML arrays or comma-separated CLI).
- Empty list `[]` means "no filter" / "disabled".
- Boolean config values: `true`/`false` in YAML, `store_true` in argparse.

## Outlook / COM Automation

- Do NOT run or execute anything that invokes Outlook or any Office COM automation.
- Treat Outlook and the Exchange/MAPI layer as a black box — write code against its interface but never call it locally.
- Test with mocks, stubs, or sample data. Never use live COM connections.
- When verifying scripts, limit checks to syntax/import — do not attempt runtime execution of Outlook-dependent code.

## Python Notebooks

- When editing `.ipynb` files, create new cells or update the last empty cell — do not overwrite cells with existing outputs.
- Notebooks in `study/` are for ad-hoc exploration and are gitignored.
- Keep cell outputs (they document results). Use `display()` for DataFrame rendering.

## Git Workflow

- Commit after every set of changes from a single user prompt.
- Use a concise, descriptive commit message summarizing what was done.
- Stage only the files modified as part of that prompt's work.

## Dependencies

- `openpyxl` + `xlrd`: Excel reading (no COM)
- `extract_msg`: .msg file parsing
- `pygit2`: git operations
- `pyyaml`: config files
- `win32com.client`: Outlook COM (never run locally in dev)
- `tkinter`: GUI forms (stdlib)
