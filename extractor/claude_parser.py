"""Use Claude API to parse financial statement text into structured FinancialData."""
from __future__ import annotations
import json
import re
from typing import Optional

import anthropic

import config as cfg
from models.financial_data import (
    FinancialData, CompanyMetadata,
    IncomeStatement, BalanceSheet, CashFlowStatement,
)


_CLIENT: Optional[anthropic.Anthropic] = None


def _client() -> anthropic.Anthropic:
    global _CLIENT
    if _CLIENT is None:
        _CLIENT = anthropic.Anthropic(api_key=cfg.ANTHROPIC_API_KEY or None)
    return _CLIENT


# ── Prompt construction ───────────────────────────────────────────────────────

_SYSTEM_PROMPT = """You are a senior financial analyst specializing in extracting structured data from SEC 10-K annual filings.

Your task: parse financial statement text and return a single JSON object that exactly matches the schema below.

CRITICAL RULES:
1. Extract ONLY values explicitly stated in the filing. Never estimate, interpolate, or invent values.
2. All monetary values must use the SAME unit scale declared in the filing header (e.g., "in millions, except per share data"). Do NOT convert units — report as-is from the filing.
3. Use JSON null for any line item not present in the filing. Never substitute 0 for a missing item.
4. Negative values: preserve the sign convention from the filing. If the filing shows interest expense as a positive number (common), keep it positive; if negative, keep negative.
5. If the filing shows multiple fiscal years in a single comparative table, return one object per fiscal year in each array.
6. For fiscal_year fields, use the fiscal year end date in YYYY format (e.g., "2024").
7. For currency use ISO 4217 (e.g., "USD"). For units use "millions", "thousands", or "billions".
8. Return ONLY the JSON code block — no explanation, no preamble, no trailing text.

JSON SCHEMA:
{
  "metadata": {
    "company_name": "string",
    "ticker": "string or null",
    "fiscal_year_end": "string (e.g. December 31) or null",
    "currency": "string (e.g. USD)",
    "units": "string (millions/thousands/billions)",
    "filing_date": "string (YYYY-MM-DD) or null",
    "period_of_report": "string (YYYY-MM-DD) or null",
    "auditor": "string or null"
  },
  "income_statements": [
    {
      "fiscal_year": "YYYY",
      "revenue": number,
      "cost_of_revenue": number,
      "gross_profit": number,
      "research_and_development": number or null,
      "selling_general_admin": number or null,
      "operating_expenses_other": number or null,
      "depreciation_amortization": number or null,
      "ebitda": number or null,
      "ebit": number or null,
      "interest_expense": number or null,
      "interest_income": number or null,
      "other_income_expense": number or null,
      "pretax_income": number or null,
      "income_tax": number or null,
      "net_income": number,
      "shares_basic": number or null,
      "shares_diluted": number or null,
      "eps_basic": number or null,
      "eps_diluted": number or null
    }
  ],
  "balance_sheets": [
    {
      "fiscal_year": "YYYY",
      "cash_and_equivalents": number,
      "short_term_investments": number or null,
      "accounts_receivable": number or null,
      "inventory": number or null,
      "prepaid_other_current": number or null,
      "total_current_assets": number or null,
      "ppe_gross": number or null,
      "accumulated_depreciation": number or null,
      "ppe_net": number or null,
      "goodwill": number or null,
      "intangibles": number or null,
      "other_noncurrent_assets": number or null,
      "total_noncurrent_assets": number or null,
      "total_assets": number,
      "accounts_payable": number or null,
      "short_term_debt": number or null,
      "accrued_liabilities": number or null,
      "deferred_revenue_current": number or null,
      "other_current_liabilities": number or null,
      "total_current_liabilities": number or null,
      "long_term_debt": number or null,
      "deferred_tax_liabilities": number or null,
      "other_noncurrent_liabilities": number or null,
      "total_noncurrent_liabilities": number or null,
      "total_liabilities": number or null,
      "common_stock": number or null,
      "additional_paid_in_capital": number or null,
      "retained_earnings": number or null,
      "accumulated_other_comprehensive": number or null,
      "treasury_stock": number or null,
      "total_equity": number,
      "total_liabilities_and_equity": number or null
    }
  ],
  "cash_flow_statements": [
    {
      "fiscal_year": "YYYY",
      "net_income": number,
      "depreciation_amortization": number or null,
      "stock_based_compensation": number or null,
      "changes_in_working_capital": number or null,
      "change_accounts_receivable": number or null,
      "change_inventory": number or null,
      "change_accounts_payable": number or null,
      "change_other_working_capital": number or null,
      "other_operating_activities": number or null,
      "net_cash_operating": number,
      "capital_expenditures": number or null,
      "acquisitions": number or null,
      "purchases_investments": number or null,
      "sales_investments": number or null,
      "other_investing": number or null,
      "net_cash_investing": number or null,
      "debt_issuance": number or null,
      "debt_repayment": number or null,
      "dividends_paid": number or null,
      "share_repurchases": number or null,
      "share_issuance": number or null,
      "other_financing": number or null,
      "net_cash_financing": number or null,
      "net_change_in_cash": number or null,
      "beginning_cash": number or null,
      "ending_cash": number or null
    }
  ]
}"""


def _build_user_prompt(financial_text: str, source_hint: str = "") -> str:
    hint_line = f"Filing source: {source_hint}\n\n" if source_hint else ""
    return (
        f"{hint_line}"
        "Extract all three financial statements from the text below. "
        "First determine the unit scale from the table headers (e.g., 'in millions'). "
        "The filing may show multiple fiscal years in comparative tables — return one entry per year.\n\n"
        "FILING TEXT:\n"
        f"{financial_text}"
    )


# ── Main extraction function ──────────────────────────────────────────────────

def extract_financials_with_claude(
    financial_text: str,
    source_hint: str = "",
    use_cache: bool = True,
) -> FinancialData:
    """
    Send financial statement text to Claude and parse the response into FinancialData.
    """
    system_content: list = [
        {
            "type": "text",
            "text": _SYSTEM_PROMPT,
            **({"cache_control": {"type": "ephemeral"}} if use_cache else {}),
        }
    ]

    extra_headers = {}
    if use_cache:
        extra_headers["anthropic-beta"] = "prompt-caching-2024-07-31"

    response = _client().messages.create(
        model=cfg.CLAUDE_MODEL,
        max_tokens=8192,
        system=system_content,
        messages=[
            {"role": "user", "content": _build_user_prompt(financial_text, source_hint)}
        ],
        extra_headers=extra_headers if extra_headers else None,
    )

    raw_text = response.content[0].text
    return _parse_response(raw_text)


def _extract_json_block(text: str) -> str:
    """Pull the JSON object out of a response that may contain surrounding text."""
    # Try fenced code block first
    m = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, re.DOTALL)
    if m:
        return m.group(1)
    # Try bare JSON object
    m = re.search(r"\{.*\}", text, re.DOTALL)
    if m:
        return m.group(0)
    raise ValueError("No JSON object found in Claude response.")


def _parse_response(response_text: str) -> FinancialData:
    json_str = _extract_json_block(response_text)
    try:
        data = json.loads(json_str)
    except json.JSONDecodeError as e:
        raise ValueError(f"Claude returned invalid JSON: {e}\n\nRaw:\n{json_str[:500]}")

    meta_raw = data.get("metadata", {})
    metadata = CompanyMetadata(
        company_name=meta_raw.get("company_name", "Unknown"),
        ticker=meta_raw.get("ticker"),
        fiscal_year_end=meta_raw.get("fiscal_year_end"),
        currency=meta_raw.get("currency", "USD"),
        units=meta_raw.get("units", "millions"),
        filing_date=meta_raw.get("filing_date"),
        period_of_report=meta_raw.get("period_of_report"),
        auditor=meta_raw.get("auditor"),
    )

    income_statements = [
        IncomeStatement(**{k: v for k, v in raw.items() if k in IncomeStatement.__dataclass_fields__})
        for raw in data.get("income_statements", [])
    ]

    balance_sheets = [
        BalanceSheet(**{k: v for k, v in raw.items() if k in BalanceSheet.__dataclass_fields__})
        for raw in data.get("balance_sheets", [])
    ]

    cash_flows = [
        CashFlowStatement(**{k: v for k, v in raw.items() if k in CashFlowStatement.__dataclass_fields__})
        for raw in data.get("cash_flow_statements", [])
    ]

    result = FinancialData(
        metadata=metadata,
        income_statements=income_statements,
        balance_sheets=balance_sheets,
        cash_flow_statements=cash_flows,
    )

    return validate_and_repair(result)


# ── Post-parse validation ─────────────────────────────────────────────────────

def validate_and_repair(data: FinancialData) -> FinancialData:
    """
    Run integrity checks and derive missing values where possible.
    Issues are printed as warnings; the data is never discarded.
    """
    for is_ in data.income_statements:
        # Derive gross profit if missing
        if is_.gross_profit is None and is_.revenue is not None and is_.cost_of_revenue is not None:
            is_.gross_profit = is_.revenue - is_.cost_of_revenue

        # Derive EBITDA if not reported
        if is_.ebitda is None and is_.ebit is not None and is_.depreciation_amortization is not None:
            is_.ebitda = is_.ebit + is_.depreciation_amortization

        # Warn if gross profit check fails
        if (is_.revenue and is_.cost_of_revenue and is_.gross_profit is not None):
            derived = is_.revenue - is_.cost_of_revenue
            if abs(derived - is_.gross_profit) / max(abs(is_.revenue), 1) > 0.02:
                print(f"  [WARN] {is_.fiscal_year}: Gross profit mismatch "
                      f"(stated={is_.gross_profit:,.1f}, derived={derived:,.1f})")

    for bs in data.balance_sheets:
        if bs.total_assets and bs.total_equity and bs.total_liabilities:
            implied_le = bs.total_liabilities + bs.total_equity
            if abs(implied_le - bs.total_assets) / max(abs(bs.total_assets), 1) > 0.02:
                print(f"  [WARN] {bs.fiscal_year}: Balance sheet does not balance "
                      f"(assets={bs.total_assets:,.1f}, L+E={implied_le:,.1f})")

    for cf in data.cash_flow_statements:
        if cf.ending_cash is not None and cf.beginning_cash is not None and cf.net_change_in_cash is not None:
            implied = cf.beginning_cash + cf.net_change_in_cash
            if abs(implied - cf.ending_cash) / max(abs(cf.ending_cash), 1) > 0.02:
                print(f"  [WARN] {cf.fiscal_year}: Cash flow reconciliation mismatch "
                      f"(ending={cf.ending_cash:,.1f}, implied={implied:,.1f})")

    return data
