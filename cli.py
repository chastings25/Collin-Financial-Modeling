"""Entry point for the 10-K → 3-Statement Excel Model Builder."""
from __future__ import annotations
import argparse
import sys
import os
from pathlib import Path


def create_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="10k_model",
        description="Build a full analyst 3-statement Excel model from a 10-K SEC filing.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  Phase 1 — local PDF:
    python -m 10k_model --pdf path/to/10k.pdf --output model.xlsx

  Phase 2 — SEC EDGAR by ticker:
    python -m 10k_model --ticker AAPL --output aapl_model.xlsx

  Phase 2 — SEC EDGAR by company name:
    python -m 10k_model --company "Apple Inc" --output apple_model.xlsx

  Phase 2 — SEC EDGAR by CIK:
    python -m 10k_model --cik 0000320193 --output model.xlsx
        """,
    )

    source = p.add_mutually_exclusive_group(required=True)
    source.add_argument("--pdf", metavar="PATH",
                        help="Path to a 10-K PDF file (Phase 1)")
    source.add_argument("--ticker", metavar="SYMBOL",
                        help="Stock ticker symbol — fetches from SEC EDGAR (Phase 2)")
    source.add_argument("--company", metavar="NAME",
                        help="Company name — fetches from SEC EDGAR (Phase 2)")
    source.add_argument("--cik", metavar="CIK",
                        help="SEC CIK number — fetches from SEC EDGAR (Phase 2)")

    p.add_argument("--output", "-o", metavar="PATH",
                   help="Output Excel file path (default: <ticker/company>_10k_model.xlsx)")
    p.add_argument("--api-key", metavar="KEY",
                   help="Anthropic API key (overrides ANTHROPIC_API_KEY env var)")
    p.add_argument("--no-cache", action="store_true",
                   help="Disable Claude prompt caching")
    p.add_argument("--verbose", "-v", action="store_true",
                   help="Print detailed extraction progress")
    return p


def _default_output(args: argparse.Namespace) -> str:
    if args.ticker:
        stem = args.ticker.upper()
    elif args.company:
        stem = args.company.replace(" ", "_")[:30]
    elif args.cik:
        stem = f"CIK{args.cik}"
    else:
        stem = Path(args.pdf).stem
    return f"{stem}_10k_model.xlsx"


def main() -> int:
    parser = create_parser()
    args = parser.parse_args()

    # Apply API key override
    if args.api_key:
        os.environ["ANTHROPIC_API_KEY"] = args.api_key

    # Validate env
    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("ERROR: ANTHROPIC_API_KEY is not set. "
              "Set it in your .env file or pass --api-key.", file=sys.stderr)
        return 1

    output_path = args.output or _default_output(args)
    use_cache = not args.no_cache

    # ── PHASE 1: PDF ──────────────────────────────────────────────────────────
    if args.pdf:
        pdf_path = args.pdf
        if not Path(pdf_path).exists():
            print(f"ERROR: PDF file not found: {pdf_path}", file=sys.stderr)
            return 1

        print(f"[1/3] Extracting financial text from PDF: {pdf_path}")
        from extractor.pdf_extractor import get_financial_section_text, extract_tables_from_pdf
        financial_text = get_financial_section_text(pdf_path)
        if args.verbose:
            tables = extract_tables_from_pdf(pdf_path)
            print(f"      Found {len(tables)} financial tables")
            print(f"      Extracted {len(financial_text):,} characters of financial text")

        print("[2/3] Parsing financial statements with Claude...")
        from extractor.claude_parser import extract_financials_with_claude
        data = extract_financials_with_claude(
            financial_text,
            source_hint=f"PDF: {Path(pdf_path).name}",
            use_cache=use_cache,
        )

    # ── PHASE 2: EDGAR ────────────────────────────────────────────────────────
    else:
        identifier = args.ticker or args.company or args.cik
        print(f"[1/3] Fetching 10-K from SEC EDGAR for: {identifier}")
        from extractor.edgar_fetcher import fetch_10k_html
        html, company_name, filing_url = fetch_10k_html(
            ticker=args.ticker,
            company_name=args.company,
            cik=args.cik,
        )
        if args.verbose:
            print(f"      Company: {company_name}")
            print(f"      Filing URL: {filing_url}")
            print(f"      HTML size: {len(html):,} characters")

        print("[2/3] Parsing financial statements with Claude...")
        from extractor.html_parser import extract_financial_tables_from_html
        financial_text = extract_financial_tables_from_html(html)
        if args.verbose:
            print(f"      Extracted {len(financial_text):,} characters of table text")

        from extractor.claude_parser import extract_financials_with_claude
        data = extract_financials_with_claude(
            financial_text,
            source_hint=f"SEC EDGAR: {company_name} — {filing_url}",
            use_cache=use_cache,
        )
        data.metadata.source_url = filing_url

    # ── Print summary ──────────────────────────────────────────────────────────
    years = data.sorted_years()
    print(f"\n      Company    : {data.metadata.company_name}")
    print(f"      Fiscal years: {', '.join('FY' + y for y in years)}")
    print(f"      Currency    : {data.metadata.currency} ({data.metadata.units})")
    if not years:
        print("ERROR: No financial statements were extracted. "
              "Check the filing format or try --verbose for details.", file=sys.stderr)
        return 1

    print(f"\n[3/3] Building Excel workbook → {output_path}")
    from excel.workbook_builder import build_workbook
    saved_path = build_workbook(data, output_path)

    print(f"\n✓ Done. Saved to: {saved_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
