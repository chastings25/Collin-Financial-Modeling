"""Parse SEC EDGAR HTML filings to extract financial statement text."""
from __future__ import annotations
import re
from typing import Optional


# Section anchors used to locate each financial statement
_SECTION_PATTERNS = {
    "income_statement": [
        r"CONSOLIDATED\s+STATEMENTS?\s+OF\s+(OPERATIONS|INCOME|EARNINGS|COMPREHENSIVE\s+INCOME)",
        r"STATEMENTS?\s+OF\s+(OPERATIONS|INCOME|EARNINGS)",
    ],
    "balance_sheet": [
        r"CONSOLIDATED\s+(BALANCE\s+SHEETS?|STATEMENTS?\s+OF\s+FINANCIAL\s+POSITION)",
        r"BALANCE\s+SHEETS?",
    ],
    "cash_flow": [
        r"CONSOLIDATED\s+STATEMENTS?\s+OF\s+CASH\s+FLOWS?",
        r"STATEMENTS?\s+OF\s+CASH\s+FLOWS?",
    ],
}


def strip_xbrl_inline_tags(html: str) -> str:
    """Remove iXBRL inline namespace tags while preserving their text content."""
    # Remove ix: namespace tags: <ix:nonfraction ...>VALUE</ix:nonfraction>
    html = re.sub(r"<ix:[^>]+>", "", html, flags=re.IGNORECASE)
    html = re.sub(r"</ix:[^>]+>", "", html, flags=re.IGNORECASE)
    # Remove other common XBRL wrappers
    html = re.sub(r"</?xbrli:[^>]*>", "", html, flags=re.IGNORECASE)
    return html


def _table_to_text(table) -> str:
    """Convert a BeautifulSoup table element to aligned plain text."""
    rows = []
    for tr in table.find_all("tr"):
        cells = []
        for td in tr.find_all(["td", "th"]):
            text = td.get_text(separator=" ", strip=True)
            # Clean up whitespace artifacts and footnote markers
            text = re.sub(r"\s+", " ", text)
            text = re.sub(r"[\(\d+\)]$", "", text).strip()
            cells.append(text)
        if any(cells):
            rows.append("  |  ".join(cells))
    return "\n".join(rows)


def _is_financial_table(table) -> bool:
    """Heuristic: does this table look like a financial statement table?"""
    text = table.get_text()
    # Must have some numeric content (dollar amounts)
    has_numbers = bool(re.search(r"\d{3,}", text))
    # Must have enough rows
    rows = table.find_all("tr")
    has_rows = len(rows) >= 5
    return has_numbers and has_rows


def _find_section_start(soup, patterns: list[str]) -> Optional[object]:
    """Find the first element matching any of the section patterns."""
    for pattern in patterns:
        for el in soup.find_all(string=re.compile(pattern, re.IGNORECASE)):
            return el
    return None


def extract_financial_tables_from_html(html: str) -> str:
    """
    Parse an SEC EDGAR HTML filing and extract the three financial statement tables
    as plain text. Returns a single string with all three sections labelled.
    """
    from bs4 import BeautifulSoup

    html = strip_xbrl_inline_tags(html)
    soup = BeautifulSoup(html, "lxml")

    # Remove script/style noise
    for tag in soup(["script", "style", "meta", "link"]):
        tag.decompose()

    sections = {}
    for section_name, patterns in _SECTION_PATTERNS.items():
        anchor = _find_section_start(soup, patterns)
        if anchor is None:
            continue

        # Walk up to find the containing block, then look for the next table
        el = anchor
        found_table = None
        # Search forward from anchor for nearest financial table
        for _ in range(20):
            el = el.parent if el else None
            if el is None:
                break
            # Look for sibling or descendant tables
            tables = el.find_all("table")
            for tbl in tables:
                if _is_financial_table(tbl):
                    found_table = tbl
                    break
            if found_table:
                break

        if found_table:
            sections[section_name] = _table_to_text(found_table)

    # Build combined output
    parts = []

    section_labels = {
        "income_statement": "=== CONSOLIDATED STATEMENTS OF OPERATIONS ===",
        "balance_sheet": "=== CONSOLIDATED BALANCE SHEETS ===",
        "cash_flow": "=== CONSOLIDATED STATEMENTS OF CASH FLOWS ===",
    }

    for key in ["income_statement", "balance_sheet", "cash_flow"]:
        if key in sections:
            parts.append(section_labels[key])
            parts.append(sections[key])
            parts.append("")

    if not parts:
        # Fallback: extract all financial tables from the document
        parts.append("=== FINANCIAL TABLES (section headers not detected) ===")
        for tbl in soup.find_all("table"):
            if _is_financial_table(tbl):
                parts.append(_table_to_text(tbl))
                parts.append("")

    return "\n".join(parts)
