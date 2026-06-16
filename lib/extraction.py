import re
import abc
from pathlib import Path


def _read_pdf_text(pdf_path: str) -> str:
    import pdfplumber
    page_texts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_texts.append(page.extract_text() or '')
    return '\n'.join(page_texts)


def find_pdfs(folder: str):
    """Recursively yield Path objects for every PDF file under folder."""
    yield from sorted(Path(folder).rglob("*.pdf"))


def _extractor_by_folder(pdf_path: str):
    """
    Return an extractor class if any ancestor folder name matches a FOLDER_HINTS
    entry, otherwise None.  Checks all path components so a hint in any parent
    folder (not just the immediate parent) is detected.
    """
    parts = [p.lower() for p in Path(pdf_path).parts]
    candidates = [UIEIndustrialExtractor, LPConstructionExtractor, MJMServicesExtractor]
    for cls in candidates:
        if any(hint in part for hint in cls.FOLDER_HINTS for part in parts):
            return cls
    return None


def guess_extractor(pdf_path: str):
    """
    Return the best-matching InvoiceExtractor subclass for a PDF, or None.

    Resolution order:
    1. Folder hint: if any ancestor directory name contains a company shorthand
       (e.g. "mjm", "uie", "lp"), that extractor is used immediately.
    2. Fingerprint scoring: pdfplumber text is checked against each extractor's
       FINGERPRINTS regex list; highest score wins.
    3. Sparse-text fallback: if pdfplumber returns < 100 chars, the PDF is
       likely scanned — MJMServicesExtractor (OCR) is used.
    """
    _SPARSE_THRESHOLD = 100

    hint_cls = _extractor_by_folder(pdf_path)
    if hint_cls is not None:
        return hint_cls

    try:
        text = _read_pdf_text(pdf_path)
    except Exception:
        return None

    if len(text.strip()) < _SPARSE_THRESHOLD:
        return MJMServicesExtractor

    candidates = [UIEIndustrialExtractor, LPConstructionExtractor, MJMServicesExtractor]
    scores = {
        cls: sum(1 for fp in cls.FINGERPRINTS if re.search(fp, text, re.IGNORECASE))
        for cls in candidates
    }
    best = max(scores, key=scores.get)
    return best if scores[best] > 0 else None


class InvoiceExtractor(abc.ABC):
    """Base class for regex-based PDF invoice extractors.

    extract() returns a single dict with invoice fields at the top level and
    a 'line_items' key containing a list-of-lists (headers row first), or None.

    Example return value::

        {
            "invoice_number": "INV-001",
            "invoice_date":   "01/01/2024",
            ...
            "line_items": [
                ["Item", "Description", "Qty", "UOM", "Unit Price", "Disc", "Amount"],
                ["1",    "Steel pipe",  "10",  "PC",  "100.00",     "0.00", "1000.00"],
            ]
        }
    """

    INVOICE_PATTERNS: dict = {}
    LINE_ITEM_COLS: list = []
    FINGERPRINTS: list = []   # Regex patterns that identify this company's PDFs
    FOLDER_HINTS: list = []   # Lowercase substrings matched against folder names
    FLAGS: int = re.IGNORECASE

    def extract_fields(self, text: str) -> dict:
        results = {}
        for field, pattern in self.INVOICE_PATTERNS.items():
            m = re.search(pattern, text, self.FLAGS)
            results[field] = m.group(1).strip() if m else 'NOT FOUND'
        return results

    @abc.abstractmethod
    def extract_line_items(self, *args, **kwargs):
        """Return list-of-lists (headers + rows), or None if none found."""

    @abc.abstractmethod
    def extract(self, pdf_path: str) -> dict:
        """Run the full pipeline; return merged fields dict with 'line_items' key."""


class UIEIndustrialExtractor(InvoiceExtractor):
    FOLDER_HINTS = ['uie']
    FINGERPRINTS = [
        r'UIE',
        r'NET\s*TOTAL\s*\[SGD\]',
        r'UEN:\s*\w',       # "UEN: T12..." (not "Paynow UEN")
        r'QT[\d/]+',        # quotation reference number
    ]

    INVOICE_PATTERNS = {
        'invoice_number': r'NO\.\s*([\w/]+)',
        'invoice_date':   r'DATE\s+(\d{2}/\d{2}/\d{4})',
        'due_date':       r'DUE\s*DATE\s+(\d{2}/\d{2}/\d{4})',
        'po_number':      r'(\d{7}\s*-\s*OP\s*-\s*\d+)',
        'sub_total':      r'SUB\s*TOTAL\s+([\d,]+\.\d{2})',
        'gst':            r'GST\s*\(\d+%\)\s+([\d,]+\.\d{2})',
        'net_total':      r'NET\s*TOTAL\s*\[SGD\]\s+([\d,]+\.\d{2})',
        'uen':            r'UEN:\s*([\w]+)',
        'gst_reg':        r'GST\s*REG\s*NO:\s*([\w\-]+)',
        'terms':          r'(\d+)\s*DAYS',
        'attn':           r'Attn\s*:\s*(.+)',
        'account_number': r'([\w\-]+\(HQ\))',
        'reference':      r'(QT[\d/]+)',
    }

    LINE_ITEM_PATTERN = r'(\d+)\s+[A-Z]\s+(.+?)\s+(\d+)\s+(PC)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)'
    LINE_ITEM_COLS = ['Item', 'Description', 'Qty', 'UOM', 'Unit Price', 'Disc', 'Amount']

    def extract_line_items(self, text: str):
        rows = re.findall(self.LINE_ITEM_PATTERN, text)
        if not rows:
            return None
        return [self.LINE_ITEM_COLS] + [list(r) for r in rows]

    def extract(self, pdf_path: str) -> dict:
        full_text = _read_pdf_text(pdf_path)
        result = self.extract_fields(full_text)
        result['line_items'] = self.extract_line_items(full_text)
        return result


class LPConstructionExtractor(InvoiceExtractor):
    # MULTILINE needed: LINE_ITEM_PATTERN uses ^ and $; 'company' pattern spans \n
    FLAGS = re.IGNORECASE | re.MULTILINE

    FOLDER_HINTS = ['lp']
    FINGERPRINTS = [
        r'L\.?\s*P\.?\s*CONSTRUCTION',
        r'Paynow\s*UEN',
        r'DBS\s*Account',
        r'Sub\s*Total\s*\(Excluding\s*Tax\)',
    ]

    INVOICE_PATTERNS = {
        'invoice_number': r'Ref No\.\s*:\s*([\w\-]+)',
        'invoice_date':   r'Date\s*:\s*(\d{2}/\d{2}/\d{4})',
        'page':           r'Page\s*:\s*(\d+ of \d+)',
        'po_number':      r'Purchase Order:\s*([\d\-A-Z]+)',
        'sub_total':      r'Sub Total \(Excluding Tax\)\s+([\d,]+\.\d{2})',
        'gst':            r'GST payable @ \d+%\s+([\d,]+\.\d{2})',
        'net_total':      r'Total \(Inclusive of Tax\)\s+([\d,]+\.\d{2})',
        'uen':            r'Paynow UEN\s*:\s*([\w]+)',
        'bank_account':   r'DBS Account No:\s*([\d\-]+)',
        'company':        r'payable to\s*\n([\w &]+)',
    }

    LINE_ITEM_PATTERN = r'^:(.+?)\s+([\d,]+\.\d{2})\s*$'
    LINE_ITEM_COLS = ['Description', 'Amount']

    def extract_line_items(self, text: str):
        rows = re.findall(self.LINE_ITEM_PATTERN, text, self.FLAGS)
        if not rows:
            return None
        return [self.LINE_ITEM_COLS] + [list(r) for r in rows]

    def extract(self, pdf_path: str) -> dict:
        full_text = _read_pdf_text(pdf_path)
        result = self.extract_fields(full_text)
        result['line_items'] = self.extract_line_items(full_text)
        return result


class MJMServicesExtractor(InvoiceExtractor):
    # Primary detection is via the sparse-text fallback in guess_extractor;
    # these fingerprints handle edge cases where pdfplumber does extract text.
    FOLDER_HINTS = ['mjm']
    FINGERPRINTS = [
        r'MJM',
        r'Inv\s*No',
        r'Net\s*Total',
    ]

    TESS_CONFIG = '--psm 6'
    TESSERACT_CMD = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

    INVOICE_PATTERNS = {
        'invoice_number': r'Inv\s*No[.:]\s*([\w/\-]+)',
        'invoice_date':   r'Date\s*[.:]\s*(\d{2}/\d{2}/\d{4})',
        'gst_reg':        r'GST\s*No[.:]\s*([\w]+)',
        'attn':           r'Attn\s*[.:]\s*([^,]+)',
        # Anchored to label so inline line-item amounts don't pollute the result
        'sub_total':      r'Sub\s*Total\s+\$?\s*([\d,]+\.\d{2})',
        'gst':            r'GST\s*(?:\(\d+%\))?\s+\$?\s*([\d,]+\.\d{2})',
        'net_total':      r'Net\s*Total\s+\$?\s*([\d,]+\.\d{2})',
    }

    LINE_ITEM_COLS = ['No', 'Description', 'Amount']
    _LINE_ITEM_RE = re.compile(r'^(.+?)\s+\$([\d,]+\.\d{2})\s*$')

    def extract_line_items(self, lines: list):
        sub_total_idx = next(
            (i for i, line in enumerate(lines) if re.search(r'sub\s*total', line, re.IGNORECASE)),
            len(lines),
        )
        rows = []
        for line in lines[:sub_total_idx]:
            m = self._LINE_ITEM_RE.match(line.strip())
            if m:
                rows.append([str(len(rows) + 1), m.group(1).strip(), m.group(2)])
        if not rows:
            return None
        return [self.LINE_ITEM_COLS] + rows

    @staticmethod
    def _render_page(pdf_path: str, page_index: int = 0, dpi: int = 300):
        import fitz
        import numpy as np
        doc = fitz.open(pdf_path)
        mat = fitz.Matrix(dpi / 72, dpi / 72)
        pix = doc[page_index].get_pixmap(matrix=mat)
        buf = np.frombuffer(pix.samples, dtype=np.uint8)
        img = buf.reshape(pix.height, pix.width, pix.n)
        return img[:, :, :3]

    @staticmethod
    def _preprocess(img, *, denoise: bool = True):
        import cv2
        gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
        if denoise:
            gray = cv2.fastNlMeansDenoising(gray, h=10)
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        return binary

    @staticmethod
    def _ocr(img, tess_config: str, lang: str = 'eng'):
        import pytesseract
        from pytesseract import Output
        text = pytesseract.image_to_string(img, lang=lang, config=tess_config)
        data = pytesseract.image_to_data(img, lang=lang, config=tess_config,
                                         output_type=Output.DATAFRAME)
        words = data[(data.conf > 0) & (data.text.str.strip() != '')].copy()
        return text, words.sort_values(['top', 'left']).reset_index(drop=True)

    @staticmethod
    def _reconstruct_lines(words, y_tol: int = 12) -> list:
        if words.empty:
            return []
        lines = []
        group = [words.iloc[0]]
        for i in range(1, len(words)):
            row = words.iloc[i]
            if abs(row['top'] - group[0]['top']) <= y_tol:
                group.append(row)
            else:
                lines.append(' '.join(r['text'] for r in sorted(group, key=lambda r: r['left'])))
                group = [row]
        if group:
            lines.append(' '.join(r['text'] for r in sorted(group, key=lambda r: r['left'])))
        return lines

    def extract(self, pdf_path: str) -> dict:
        import pytesseract
        pytesseract.pytesseract.tesseract_cmd = self.TESSERACT_CMD
        img = self._render_page(pdf_path)
        processed = self._preprocess(img)
        full_text, words = self._ocr(processed, self.TESS_CONFIG)
        lines = self._reconstruct_lines(words)
        result = self.extract_fields(full_text)
        result['line_items'] = self.extract_line_items(lines)
        return result
