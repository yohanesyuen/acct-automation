# acct-automation

Accounting automation toolkit for extracting GR (Goods Receipt) numbers from Outlook `.msg` emails and Excel attachments, and generating CSV reports.

## What it does

- Extracts attachments (Excel files) from `.msg` email files
- Searches Excel files for GR numbers using pattern matching and header-aware strategies
- Connects to live Outlook via COM for email filtering and dumping
- Generates CSV summary reports of extraction results
- Provides a GUI task selector and CLI for running scripts

## Tech stack

- Python 3.13
- `extract_msg` — parse Outlook `.msg` files
- `openpyxl` / `xlrd` — read `.xlsx`/`.xlsm`/`.xls` without Excel
- `win32com` (pywin32) — live Outlook COM connection
- Standard library: `csv`, `re`, `tkinter` (GUI)

## Setup

1. Install dependencies (no `requirements.txt` is present — install manually):
   ```
   pip install extract-msg openpyxl xlrd pywin32
   ```

2. Configure the working directory in `tasks/main.yml`:
   ```yaml
   root: "C:\path\to\GRN_Extract"
   attachments_subdir: "Attachments"
   excel_subdir: "Excel_Files"
   ```

## Usage

```bash
# Open GUI task selector
python main.py

# List available scripts
python main.py list

# Run a specific script
python main.py run <script_name>

# Show help for a script
python main.py run <script_name> --help
```

Available scripts in `scripts/`:

| Script | Purpose |
|--------|---------|
| `extract_attachments` | Extract Excel attachments from `.msg` files |
| `search_excel_content` | Search Excel files for GR numbers |
| `process_msg_files` | Process `.msg` files end-to-end |
| `collect_excel` | Collect Excel files from a directory |
| `dump_emails` | Dump email metadata to CSV |
| `filter_by_contacts` | Filter emails by sender contacts |
| `analyze_headers` | Analyse Excel header structures |
