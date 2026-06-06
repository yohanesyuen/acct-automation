# acct-automation — Project Context

You are assisting with an accounting automation project that extracts
GRN (Goods Received Note) data from Outlook emails and Excel attachments.

Below is a complete reference of the project's library modules and
available task scripts. Use this context to understand what already exists
before proposing changes or new code.

**Your response rules:**
- Do NOT immediately write code or provide a solution.
- Instead, reply with a focused follow-up question asking what specific
  task or problem we want to accomplish.
- Only after receiving a clear answer should you proceed with implementation.

---

## CLI Interface

The project is driven by `main.py` with the following subcommands:

```
python main.py generate          # Regenerate this pre_prompt.md
python main.py list              # List available task scripts
python main.py run               # List scripts (same as list)
python main.py run <script>      # Run a task script
python main.py run <script> --help   # Show all available args for a script
```

### Script argument convention

Each script loads its defaults from a YAML file in `tasks/` and exposes
every config key as a CLI override flag:

```
python main.py run extract_grn --sender-email someone@example.com --keyword Invoice
python main.py run find_gr_numbers --root "D:\Output" --report-file custom.csv
```

- `--task <name>` selects which YAML config to load (default per script).
- All other YAML keys become `--<key-with-hyphens>` flags.
- CLI values override YAML defaults; YAML values override hardcoded defaults.

### Adding a new task script

Create a YAML config in `tasks/` and a script in `scripts/`:

**tasks/my_new_task.yml**
```yaml
root: "C:\\Users\\...\\Output"

sender_email: "someone@example.com"
keyword: "Invoice"
excel_subdir: "Excel_Files"
report_file: "My_Report.csv"
```

**scripts/my_new_task.py**
```python
"""
Short description of what this script does.

Usage:
    python scripts/my_new_task.py
    python scripts/my_new_task.py --sender-email other@example.com
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from lib.outlook import get_outlook_inbox, filter_emails, iter_attachments
from lib.excel_search import search_excel_for_gr_pattern
from lib.reporting import write_csv_report
from lib.task_config import parse_task_args, get_output_dir, get_report_path
from lib.utils import is_excel_file


def my_new_task(config):
    sender_email = config.get("sender_email")
    keyword = config.get("keyword")
    output_folder = get_output_dir(config, config.get("excel_subdir", "Excel_Files"))
    report_path = get_report_path(config, config.get("report_file", "My_Report.csv"))

    inbox = get_outlook_inbox()
    emails = filter_emails(inbox, sender_email=sender_email, keyword=keyword)

    # ... process emails, build report rows ...

    write_csv_report(report_path, report)


if __name__ == "__main__":
    config = parse_task_args(
        description="Short description of what this script does.",
        default_task="my_new_task",
    )
    my_new_task(config)
```

All YAML keys automatically become `--flag` overrides at the CLI.

---

{{SCRIPTS_LISTING}}

---

{{LIB_LISTING}}
