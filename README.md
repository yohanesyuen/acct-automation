# acct-automation

Accounting automation toolkit for extracting GR (Goods Receipt) numbers from Outlook `.msg` emails and Excel attachments, generating CSV reports, and parsing PDF invoices.

## What it does

- Extracts attachments (Excel files) from `.msg` email files
- Searches Excel files for GR numbers using pattern matching and header-aware strategies
- Connects to live Outlook via COM for email filtering and dumping
- Generates CSV summary reports of extraction results
- Provides a GUI task selector and CLI for running scripts
- Parses PDF invoices: text extraction, table detection, field pattern matching

## Tech stack

- Python 3.13+
- `extract-msg` — parse Outlook `.msg` files
- `openpyxl` / `xlrd` — read `.xlsx`/`.xlsm`/`.xls` without Excel
- `pywin32` — live Outlook COM connection
- `PyYAML` — task configuration files
- `pygit2` — git operations
- `pdfplumber` — text and table extraction from PDFs
- `PyMuPDF` (fitz) — PDF rendering and rasterisation for OCR
- `winrt` — Windows Runtime OCR engine (`Windows.Media.Ocr`)
- `pandas` — tabular data manipulation
- Standard library: `csv`, `re`, `tkinter` (GUI)

## Setup

### 1. Create and activate the project virtual environment

```powershell
# Create (first time only)
python -m virtualenv .venv

# Activate — PowerShell
.venv\Scripts\Activate.ps1

# Activate — bash/sh
source .venv/Scripts/activate
```

> **Note:** `main.py` will print a warning if it detects you are running in `conda base`
> or the system Python instead of a dedicated venv or named conda environment.

### 2. Install dependencies

```powershell
pip install -r requirements.txt
```

### 3. Configure task paths

Task configs live in `tasks/<script_name>.yml` and are auto-generated on first run from
the hardcoded defaults in `lib/task_config.py`. Edit the generated YAML or use the GUI
to update paths.

## Usage

```powershell
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

## PDF / OCR approach (`study/`)

The `study/` scripts explore invoice PDF parsing. The current approach layers two techniques:

| Technique | Library | When to use |
|-----------|---------|-------------|
| Text extraction | `pdfplumber` | Native/digital PDFs — fast, no GPU needed |
| Rasterise + OCR | `PyMuPDF` + `winrt` | Scanned or image-embedded PDFs |

### Alternative OCR libraries

| Library | Pip package | Notes |
|---------|-------------|-------|
| **Tesseract** | `pytesseract` | Needs [Tesseract binary](https://github.com/tesseract-ocr/tesseract) installed; free, widely supported |
| **EasyOCR** | `easyocr` | PyTorch-based, 80+ languages, GPU-optional; heavier than Tesseract |
| **PaddleOCR** | `paddleocr` | PaddlePaddle-based; excellent on tables and structured docs; heavy install |
| **Surya** | `surya-ocr` | Transformer-based, offline; strong on multi-column documents; GPU-optional |
| **docTR** | `python-doctr` | Mindee's document OCR (PyTorch or TF); layout-aware, table support |
| **Camelot** | `camelot-py[cv]` | Table extraction only (not OCR); works on native PDFs; `lattice` mode for bordered tables |
| **Tabula** | `tabula-py` | JVM-based table extractor; alternative to camelot for native PDFs |
| **Azure Doc Intelligence** | `azure-ai-documentintelligence` | Cloud API; pre-built invoice/receipt models; best accuracy on structured forms |
| **Google Cloud Vision** | `google-cloud-vision` | Cloud API; strong general OCR; pay-per-call |

**Recommendation for invoice parsing:**
- Native digital PDFs → `pdfplumber` (already in use) or `camelot` for complex tables
- Scanned invoices (offline) → `surya-ocr` or `easyocr` as drop-in replacements for `winrt`
- High-accuracy structured forms → Azure Document Intelligence (pre-built invoice model)
