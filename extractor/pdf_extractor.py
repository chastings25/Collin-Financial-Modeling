"""Extract text and tables from a 10-K PDF using pdfplumber."""
from __future__ import annotations
import re
from pathlib import Path


# Anchor phrases that mark the start of financial statement sections
_FINANCIAL_ANCHORS = [
    "CONSOLIDATED STATEMENTS OF OPERATIONS",
    "CONSOLIDATED STATEMENTS OF INCOME",
    "CONSOLIDATED STATEMENTS OF EARNINGS",
    "CONSOLIDATED BALANCE SHEETS",
    "CONSOLIDATED STATEMENTS OF FINANCIAL POSITION",
    "CONSOLIDATED STATEMENTS OF CASH FLOWS",
    "STATEMENTS OF OPERATIONS",
    "STATEMENTS OF INCOME",
    "BALANCE SHEETS",
    "STATEMENTS OF CASH FLOWS",
]


def _find_financial_page_range(pages_text: list[str]) -> tuple[int, int]:
    """Return (start_page_index, end_page_index) for the financial statements section."""
    first_anchor_page = len(pages_text)
    for i, text in enumerate(pages_text):
        upper = text.upper()
        for anchor in _FINANCIAL_ANCHORS:
            if anchor in upper:
                first_anchor_page = i
                break
        if first_anchor_page < len(pages_text):
            break
    end_page = min(first_anchor_page + 60, len(pages_text))
    return first_anchor_page, end_page


def get_financial_section_text(pdf_path: str) -> str:
    """
    Extract text from the financial statements section of a 10-K PDF.
    Anchors on the first financial statement heading and takes ~60 pages.
    Returns a plain-text string suitable for sending to Claude.
    """
    import pdfplumber

    path = Path(pdf_path)
    if not path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    with pdfplumber.open(str(path)) as pdf:
        pages_text = []
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=2, y_tolerance=2) or ""
            pages_text.append(text)

    start, end = _find_financial_page_range(pages_text)
    financial_text = "\n\n--- PAGE BREAK ---\n\n".join(pages_text[start:end])

    if not financial_text.strip():
        # Fallback: return full document text (PDF may have no detectable anchors)
        financial_text = "\n\n--- PAGE BREAK ---\n\n".join(pages_text)

    return financial_text


def extract_tables_from_pdf(pdf_path: str) -> list[dict]:
    """
    Extract tables from the financial section pages.
    Returns a list of {page_num, table_text} dicts.
    """
    import pdfplumber

    path = Path(pdf_path)
    if not path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    with pdfplumber.open(str(path)) as pdf:
        pages_text = [
            (page.extract_text(x_tolerance=2, y_tolerance=2) or "")
            for page in pdf.pages
        ]
        start, end = _find_financial_page_range(pages_text)

        results = []
        for i in range(start, end):
            page = pdf.pages[i]
            tables = page.extract_tables()
            for table in tables:
                if not table or len(table) < 4:
                    continue
                rows_text = []
                for row in table:
                    clean_row = [str(cell).strip() if cell else "" for cell in row]
                    if any(re.search(r"[\d,\.\(\)]+", c) for c in clean_row):
                        rows_text.append("  |  ".join(clean_row))
                if rows_text:
                    results.append({
                        "page_num": i + 1,
                        "table_text": "\n".join(rows_text),
                    })

    return results


def extract_full_text(pdf_path: str) -> str:
    """Return all text from the PDF (used as fallback or for metadata detection)."""
    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        return "\n\n".join(
            (page.extract_text(x_tolerance=2, y_tolerance=2) or "")
            for page in pdf.pages
        )
