"""Fetch 10-K filings from SEC EDGAR using the public REST API."""
from __future__ import annotations
import json
import re
import time
from typing import Optional

import requests

import config as cfg


_SESSION = requests.Session()
_SESSION.headers.update({"User-Agent": cfg.EDGAR_USER_AGENT})


def _get(url: str, **kwargs) -> requests.Response:
    time.sleep(cfg.EDGAR_RATE_LIMIT_SLEEP)
    resp = _SESSION.get(url, timeout=30, **kwargs)
    resp.raise_for_status()
    return resp


# ── Ticker / company name → CIK ──────────────────────────────────────────────

def _load_ticker_map() -> dict[str, dict]:
    """Download and return the SEC company_tickers.json mapping."""
    resp = _get(cfg.EDGAR_TICKERS_URL)
    raw = resp.json()
    # Keys are string integers; values are {cik_str, ticker, title}
    return {v["ticker"].upper(): v for v in raw.values()}


def get_cik_from_ticker(ticker: str) -> str:
    """
    Resolve a stock ticker to a zero-padded 10-digit SEC CIK string.
    Raises ValueError if the ticker is not found.
    """
    mapping = _load_ticker_map()
    entry = mapping.get(ticker.upper())
    if not entry:
        raise ValueError(
            f"Ticker '{ticker}' not found in SEC company_tickers.json. "
            "Try --company with the full company name instead."
        )
    return str(entry["cik_str"]).zfill(10)


def search_cik_by_company_name(company_name: str) -> tuple[str, str]:
    """
    Search EDGAR for a company by name.
    Returns (cik_padded, confirmed_company_name).
    Raises ValueError if no match found.
    """
    url = (
        "https://efts.sec.gov/LATEST/search-index?q="
        + requests.utils.quote(f'"{company_name}"')
        + "&forms=10-K&dateRange=custom&startdt=2010-01-01"
    )
    resp = _get(url)
    data = resp.json()
    hits = data.get("hits", {}).get("hits", [])
    if not hits:
        # Retry without quotes for a broader match
        url2 = (
            "https://efts.sec.gov/LATEST/search-index?q="
            + requests.utils.quote(company_name)
            + "&forms=10-K"
        )
        resp2 = _get(url2)
        hits = resp2.json().get("hits", {}).get("hits", [])
    if not hits:
        raise ValueError(
            f"No 10-K filing found for company '{company_name}' on EDGAR."
        )
    source = hits[0].get("_source", {})
    cik = str(source.get("entity_id", source.get("ciks", [""])[0])).zfill(10)
    name = source.get("display_names", [company_name])[0]
    if not cik.strip("0"):
        raise ValueError(f"Could not resolve CIK for '{company_name}'.")
    return cik, name


# ── CIK → latest 10-K filing ─────────────────────────────────────────────────

def get_latest_10k_filing(cik: str) -> dict:
    """
    Fetch the latest 10-K filing metadata for a company.
    Returns dict with keys: accession_number, filing_date, primary_document, cik
    """
    url = f"{cfg.EDGAR_BASE_URL}/submissions/CIK{cik}.json"
    data = _get(url).json()

    filings = data.get("filings", {}).get("recent", {})
    forms = filings.get("form", [])
    accessions = filings.get("accessionNumber", [])
    dates = filings.get("filingDate", [])
    primary_docs = filings.get("primaryDocument", [])

    for i, form in enumerate(forms):
        if form in ("10-K", "10-K405", "10-KSB"):
            return {
                "accession_number": accessions[i],
                "filing_date": dates[i],
                "primary_document": primary_docs[i],
                "cik": cik,
                "company_name": data.get("name", ""),
            }

    # Check older filings if not in recent
    raise ValueError(
        f"No 10-K filing found in recent filings for CIK {cik}. "
        "The company may have filed under a different form type."
    )


# ── Filing index → primary HTML document ─────────────────────────────────────

def _accession_no_hyphens(accession_number: str) -> str:
    return accession_number.replace("-", "")


def get_primary_html_url(cik: str, accession_number: str, primary_document: str) -> str:
    """
    Build the direct URL to the primary HTML document of a filing.
    Falls back to scraping the filing index if the primary_document path doesn't work.
    """
    acc_clean = _accession_no_hyphens(accession_number)
    direct_url = f"{cfg.EDGAR_ARCHIVES_URL}/{cik.lstrip('0')}/{acc_clean}/{primary_document}"

    # Verify the direct URL is reachable
    try:
        resp = _SESSION.head(direct_url, timeout=15)
        if resp.status_code == 200:
            return direct_url
    except Exception:
        pass

    # Scrape index page to find primary document
    index_url = f"{cfg.EDGAR_ARCHIVES_URL}/{cik.lstrip('0')}/{acc_clean}/{acc_clean}-index.htm"
    try:
        resp = _get(index_url)
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(resp.text, "lxml")
        for row in soup.find_all("tr"):
            cells = row.find_all("td")
            if len(cells) >= 4:
                doc_type = cells[3].get_text(strip=True)
                if doc_type in ("10-K", "10-K405", "10-KSB"):
                    link = cells[2].find("a")
                    if link and link.get("href"):
                        href = link["href"]
                        if not href.startswith("http"):
                            href = "https://www.sec.gov" + href
                        return href
    except Exception:
        pass

    return direct_url


# ── Download HTML filing ──────────────────────────────────────────────────────

def download_filing_html(url: str) -> str:
    """Download and return the HTML content of a filing."""
    resp = _get(url)
    # Try to detect encoding from headers or content
    encoding = resp.encoding or "utf-8"
    try:
        return resp.content.decode(encoding)
    except (UnicodeDecodeError, LookupError):
        return resp.content.decode("latin-1")


# ── High-level convenience function ──────────────────────────────────────────

def fetch_10k_html(ticker: str | None = None, company_name: str | None = None,
                   cik: str | None = None) -> tuple[str, str, str]:
    """
    Resolve company to CIK, find the latest 10-K, and return (html_text, company_name, filing_url).
    Provide exactly one of ticker, company_name, or cik.
    """
    if cik:
        cik = cik.zfill(10)
        company_display = f"CIK {cik}"
    elif ticker:
        cik = get_cik_from_ticker(ticker)
        company_display = ticker.upper()
    elif company_name:
        cik, company_display = search_cik_by_company_name(company_name)
    else:
        raise ValueError("Provide ticker, company_name, or cik.")

    filing = get_latest_10k_filing(cik)
    company_display = filing.get("company_name", company_display)
    url = get_primary_html_url(filing["cik"], filing["accession_number"], filing["primary_document"])
    html = download_filing_html(url)
    return html, company_display, url
