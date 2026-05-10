"""
Fetch structured financial data from SEC EDGAR XBRL API.
No API key required — all public data from data.sec.gov.
"""
from __future__ import annotations
import time
from typing import Optional

import requests

import config as cfg
from models.financial_data import (
    FinancialData, CompanyMetadata,
    IncomeStatement, BalanceSheet, CashFlowStatement,
)

_SESSION = requests.Session()
_SESSION.headers.update({"User-Agent": cfg.EDGAR_USER_AGENT})


def _get(url: str) -> dict:
    time.sleep(cfg.EDGAR_RATE_LIMIT_SLEEP)
    r = _SESSION.get(url, timeout=30)
    r.raise_for_status()
    return r.json()


# ── Concept name fallback lists (tried in order) ──────────────────────────
# Each entry is tried until a value is found for the given fiscal year.

_CONCEPTS = {
    # Income Statement
    "revenue": [
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        "Revenues", "SalesRevenueNet", "RevenueFromContractWithCustomerIncludingAssessedTax",
        "SalesRevenueGoodsNet", "RevenueFromRelatedParties",
    ],
    "cost_of_revenue": [
        "CostOfRevenue", "CostOfGoodsSold", "CostOfGoodsSoldAndServicesCost",
        "CostOfGoodsAndServicesSold",
    ],
    "gross_profit": ["GrossProfit"],
    "research_and_development": ["ResearchAndDevelopmentExpense"],
    "selling_general_admin": [
        "SellingGeneralAndAdministrativeExpense",
        "GeneralAndAdministrativeExpense",
    ],
    "depreciation_amortization": [
        "DepreciationDepletionAndAmortization", "DepreciationAndAmortization",
        "Depreciation",
    ],
    "ebit": ["OperatingIncomeLoss"],
    "interest_expense": ["InterestExpense", "InterestExpenseDebt"],
    "interest_income": ["InvestmentIncomeInterest", "InterestIncomeOperating"],
    "pretax_income": [
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments",
    ],
    "income_tax": ["IncomeTaxExpenseBenefit"],
    "net_income": [
        "NetIncomeLoss",
        "NetIncomeLossAvailableToCommonStockholdersBasic",
    ],
    "eps_basic": ["EarningsPerShareBasic"],
    "eps_diluted": ["EarningsPerShareDiluted"],
    "shares_basic": ["WeightedAverageNumberOfSharesOutstandingBasic", "CommonStockSharesOutstanding"],
    "shares_diluted": ["WeightedAverageNumberOfDilutedSharesOutstanding"],
    # Balance Sheet — Assets
    "cash_and_equivalents": ["CashAndCashEquivalentsAtCarryingValue", "Cash"],
    "short_term_investments": ["ShortTermInvestments", "AvailableForSaleSecuritiesCurrent"],
    "accounts_receivable": ["AccountsReceivableNetCurrent", "ReceivablesNetCurrent"],
    "inventory": ["InventoryNet", "InventoryGross"],
    "prepaid_other_current": ["PrepaidExpenseAndOtherAssetsCurrent", "OtherAssetsCurrent"],
    "total_current_assets": ["AssetsCurrent"],
    "ppe_net": ["PropertyPlantAndEquipmentNet"],
    "goodwill": ["Goodwill"],
    "intangibles": ["IntangibleAssetsNetExcludingGoodwill", "FiniteLivedIntangibleAssetsNet"],
    "other_noncurrent_assets": ["OtherAssetsNoncurrent"],
    "total_noncurrent_assets": ["AssetsNoncurrent"],
    "total_assets": ["Assets"],
    # Balance Sheet — Liabilities
    "accounts_payable": ["AccountsPayableCurrent"],
    "short_term_debt": ["ShortTermBorrowings", "DebtCurrent", "NotesPayableCurrent"],
    "accrued_liabilities": ["AccruedLiabilitiesCurrent", "EmployeeRelatedLiabilitiesCurrent"],
    "deferred_revenue_current": ["DeferredRevenueCurrent", "ContractWithCustomerLiabilityCurrent"],
    "other_current_liabilities": ["OtherLiabilitiesCurrent"],
    "total_current_liabilities": ["LiabilitiesCurrent"],
    "long_term_debt": ["LongTermDebtNoncurrent", "LongTermDebt", "LongTermNotesPayable"],
    "deferred_tax_liabilities": ["DeferredIncomeTaxLiabilitiesNet", "DeferredTaxLiabilitiesNoncurrent"],
    "other_noncurrent_liabilities": ["OtherLiabilitiesNoncurrent"],
    "total_noncurrent_liabilities": ["LiabilitiesNoncurrent"],
    "total_liabilities": ["Liabilities"],
    # Balance Sheet — Equity
    "common_stock": ["CommonStockValue"],
    "additional_paid_in_capital": ["AdditionalPaidInCapital", "AdditionalPaidInCapitalCommonStock"],
    "retained_earnings": ["RetainedEarningsAccumulatedDeficit"],
    "accumulated_other_comprehensive": ["AccumulatedOtherComprehensiveIncomeLossNetOfTax"],
    "treasury_stock": ["TreasuryStockValue"],
    "total_equity": [
        "StockholdersEquity",
        "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest",
    ],
    # Cash Flow
    "net_cash_operating": ["NetCashProvidedByUsedInOperatingActivities"],
    "stock_based_compensation": ["ShareBasedCompensation", "AllocatedShareBasedCompensationExpense"],
    "capital_expenditures": ["PaymentsToAcquirePropertyPlantAndEquipment"],
    "acquisitions": ["PaymentsToAcquireBusinessesNetOfCashAcquired"],
    "purchases_investments": ["PaymentsToAcquireInvestments", "PaymentsToAcquireAvailableForSaleSecurities"],
    "sales_investments": ["ProceedsFromSaleOfAvailableForSaleSecurities", "ProceedsFromMaturitiesPrepaymentsAndCallsOfAvailableForSaleSecurities"],
    "net_cash_investing": ["NetCashProvidedByUsedInInvestingActivities"],
    "debt_repayment": ["RepaymentsOfLongTermDebt", "RepaymentsOfDebt"],
    "debt_issuance": ["ProceedsFromIssuanceOfLongTermDebt", "ProceedsFromDebtNetOfIssuanceCosts"],
    "dividends_paid": ["PaymentsOfDividends", "PaymentsOfDividendsCommonStock"],
    "share_repurchases": ["PaymentsForRepurchaseOfCommonStock"],
    "share_issuance": ["ProceedsFromIssuanceOfCommonStock"],
    "net_cash_financing": ["NetCashProvidedByUsedInFinancingActivities"],
    "net_change_in_cash": [
        "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect",
        "NetCashProvidedByUsedInContinuingOperations",
    ],
    "ending_cash": [
        "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
        "CashAndCashEquivalentsAtCarryingValue",
    ],
}


# ── Core XBRL extraction ──────────────────────────────────────────────────

def _extract_annual_values(facts: dict, concepts: list[str],
                            target_years: set[int]) -> dict[int, float]:
    """
    Search through a list of concept names (in priority order) and return
    a {fiscal_year: value} dict for each target year found.
    Values are in the filing's original unit (USD for most, shares for counts).
    """
    results: dict[int, float] = {}
    us_gaap = facts.get("us-gaap", {})

    for concept in concepts:
        if concept not in us_gaap:
            continue
        entries = us_gaap[concept].get("units", {})
        # USD or shares
        unit_data = entries.get("USD") or entries.get("shares") or []

        for entry in unit_data:
            if entry.get("form") not in ("10-K", "10-K405") or entry.get("fp") != "FY":
                continue
            fy = entry.get("fy")
            if fy not in target_years:
                continue
            if fy not in results:
                results[fy] = entry["val"]

        # If we found all target years with this concept, stop trying alternatives
        if all(yr in results for yr in target_years):
            break

    return results


def _pick(values: dict[int, float], year: int,
          divisor: float = 1.0) -> Optional[float]:
    v = values.get(year)
    return v / divisor if v is not None else None


# ── Main public function ──────────────────────────────────────────────────

def fetch_financials_from_xbrl(
    ticker: str | None = None,
    company_name: str | None = None,
    cik: str | None = None,
    num_years: int = 3,
) -> FinancialData:
    """
    Fetch the latest 10-K financial data from SEC EDGAR XBRL.
    Returns a FinancialData object — same shape as the Claude parser output.
    No API key required.
    """
    # Resolve to CIK
    from extractor.edgar_fetcher import get_cik_from_ticker, search_cik_by_company_name

    if cik:
        cik_padded = cik.zfill(10)
        company_display = f"CIK {cik_padded}"
    elif ticker:
        cik_padded = get_cik_from_ticker(ticker)
        company_display = ticker.upper()
    elif company_name:
        cik_padded, company_display = search_cik_by_company_name(company_name)
    else:
        raise ValueError("Provide ticker, company_name, or cik.")

    # Fetch company facts (all XBRL data)
    facts_url = f"{cfg.EDGAR_BASE_URL}/api/xbrl/companyfacts/CIK{cik_padded}.json"
    print(f"  Fetching XBRL facts for {company_display}...")
    raw = _get(facts_url)

    company_name_official = raw.get("entityName", company_display)
    facts = raw.get("facts", {})

    # Determine which fiscal years are available (look at NetIncomeLoss as anchor)
    anchor_concepts = ["NetIncomeLoss", "Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax"]
    available_years: set[int] = set()
    us_gaap = facts.get("us-gaap", {})
    for concept in anchor_concepts:
        if concept in us_gaap:
            for entry in us_gaap[concept].get("units", {}).get("USD", []):
                if entry.get("form") in ("10-K", "10-K405") and entry.get("fp") == "FY":
                    available_years.add(entry["fy"])
            if available_years:
                break

    if not available_years:
        raise ValueError(f"No annual 10-K XBRL data found for {company_display}.")

    target_years = sorted(available_years, reverse=True)[:num_years]
    target_set = set(target_years)

    # Determine units (most companies report in USD; divide to get millions)
    # We'll use millions as the standard unit
    divisor = 1_000_000.0
    units_label = "millions"

    # EPS and shares are not in millions — handle separately
    eps_divisor = 1.0
    shares_divisor = 1_000_000.0  # convert shares to millions

    # ── Extract all concepts ──────────────────────────────────────────────
    def get(field: str) -> dict[int, float]:
        return _extract_annual_values(facts, _CONCEPTS[field], target_set)

    # Pre-fetch all fields
    rev_vals = get("revenue")
    cogs_vals = get("cost_of_revenue")
    gp_vals = get("gross_profit")
    rd_vals = get("research_and_development")
    sga_vals = get("selling_general_admin")
    da_vals = get("depreciation_amortization")
    ebit_vals = get("ebit")
    int_exp_vals = get("interest_expense")
    int_inc_vals = get("interest_income")
    ebt_vals = get("pretax_income")
    tax_vals = get("income_tax")
    ni_vals = get("net_income")
    eps_basic_vals = get("eps_basic")
    eps_dil_vals = get("eps_diluted")
    shares_basic_vals = get("shares_basic")
    shares_dil_vals = get("shares_diluted")

    cash_vals = get("cash_and_equivalents")
    sti_vals = get("short_term_investments")
    ar_vals = get("accounts_receivable")
    inv_vals = get("inventory")
    prepaid_vals = get("prepaid_other_current")
    tca_vals = get("total_current_assets")
    ppe_vals = get("ppe_net")
    gw_vals = get("goodwill")
    intang_vals = get("intangibles")
    other_nca_vals = get("other_noncurrent_assets")
    tnca_vals = get("total_noncurrent_assets")
    ta_vals = get("total_assets")

    ap_vals = get("accounts_payable")
    std_vals = get("short_term_debt")
    accr_vals = get("accrued_liabilities")
    def_rev_vals = get("deferred_revenue_current")
    other_cl_vals = get("other_current_liabilities")
    tcl_vals = get("total_current_liabilities")
    ltd_vals = get("long_term_debt")
    def_tax_vals = get("deferred_tax_liabilities")
    other_ncl_vals = get("other_noncurrent_liabilities")
    tncl_vals = get("total_noncurrent_liabilities")
    tl_vals = get("total_liabilities")
    cs_vals = get("common_stock")
    apic_vals = get("additional_paid_in_capital")
    re_vals = get("retained_earnings")
    aoci_vals = get("accumulated_other_comprehensive")
    ts_vals = get("treasury_stock")
    te_vals = get("total_equity")

    cfo_vals = get("net_cash_operating")
    sbc_vals = get("stock_based_compensation")
    capex_vals = get("capital_expenditures")
    acq_vals = get("acquisitions")
    buy_inv_vals = get("purchases_investments")
    sell_inv_vals = get("sales_investments")
    cfi_vals = get("net_cash_investing")
    debt_rep_vals = get("debt_repayment")
    debt_iss_vals = get("debt_issuance")
    divs_vals = get("dividends_paid")
    buyback_vals = get("share_repurchases")
    share_iss_vals = get("share_issuance")
    cff_vals = get("net_cash_financing")
    net_chg_vals = get("net_change_in_cash")
    end_cash_vals = get("ending_cash")

    # ── Determine fiscal year end ─────────────────────────────────────────
    # Look for the period end date in the submission metadata
    fy_end = None
    try:
        sub_url = f"{cfg.EDGAR_BASE_URL}/submissions/CIK{cik_padded}.json"
        sub = _get(sub_url)
        fy_end = sub.get("fiscalYearEnd", "")  # e.g. "1231" for Dec 31
        if fy_end and len(fy_end) == 4:
            months = {"01":"January","02":"February","03":"March","04":"April",
                      "05":"May","06":"June","07":"July","08":"August",
                      "09":"September","10":"October","11":"November","12":"December"}
            fy_end = f"{months.get(fy_end[:2], fy_end[:2])} {fy_end[2:]}"
    except Exception:
        pass

    # ── Build FinancialData ───────────────────────────────────────────────
    d = divisor

    income_statements = []
    for yr in sorted(target_years):
        # Derive EBITDA if not tagged
        ebit = _pick(ebit_vals, yr, d)
        da = _pick(da_vals, yr, d)
        ebitda = (ebit + da) if (ebit is not None and da is not None) else None

        # Gross profit: derive if not tagged
        rev = _pick(rev_vals, yr, d)
        cogs = _pick(cogs_vals, yr, d)
        gp = _pick(gp_vals, yr, d) or ((rev - cogs) if rev and cogs else None)

        income_statements.append(IncomeStatement(
            fiscal_year=str(yr),
            revenue=rev or 0.0,
            cost_of_revenue=cogs or 0.0,
            gross_profit=gp or 0.0,
            research_and_development=_pick(rd_vals, yr, d),
            selling_general_admin=_pick(sga_vals, yr, d),
            depreciation_amortization=da,
            ebitda=ebitda,
            ebit=ebit,
            interest_expense=_pick(int_exp_vals, yr, d),
            interest_income=_pick(int_inc_vals, yr, d),
            pretax_income=_pick(ebt_vals, yr, d),
            income_tax=_pick(tax_vals, yr, d),
            net_income=_pick(ni_vals, yr, d) or 0.0,
            eps_basic=_pick(eps_basic_vals, yr, eps_divisor),
            eps_diluted=_pick(eps_dil_vals, yr, eps_divisor),
            shares_basic=_pick(shares_basic_vals, yr, shares_divisor),
            shares_diluted=_pick(shares_dil_vals, yr, shares_divisor),
        ))

    balance_sheets = []
    for yr in sorted(target_years):
        balance_sheets.append(BalanceSheet(
            fiscal_year=str(yr),
            cash_and_equivalents=_pick(cash_vals, yr, d) or 0.0,
            short_term_investments=_pick(sti_vals, yr, d),
            accounts_receivable=_pick(ar_vals, yr, d),
            inventory=_pick(inv_vals, yr, d),
            prepaid_other_current=_pick(prepaid_vals, yr, d),
            total_current_assets=_pick(tca_vals, yr, d),
            ppe_net=_pick(ppe_vals, yr, d),
            goodwill=_pick(gw_vals, yr, d),
            intangibles=_pick(intang_vals, yr, d),
            other_noncurrent_assets=_pick(other_nca_vals, yr, d),
            total_noncurrent_assets=_pick(tnca_vals, yr, d),
            total_assets=_pick(ta_vals, yr, d) or 0.0,
            accounts_payable=_pick(ap_vals, yr, d),
            short_term_debt=_pick(std_vals, yr, d),
            accrued_liabilities=_pick(accr_vals, yr, d),
            deferred_revenue_current=_pick(def_rev_vals, yr, d),
            other_current_liabilities=_pick(other_cl_vals, yr, d),
            total_current_liabilities=_pick(tcl_vals, yr, d),
            long_term_debt=_pick(ltd_vals, yr, d),
            deferred_tax_liabilities=_pick(def_tax_vals, yr, d),
            other_noncurrent_liabilities=_pick(other_ncl_vals, yr, d),
            total_noncurrent_liabilities=_pick(tncl_vals, yr, d),
            total_liabilities=_pick(tl_vals, yr, d),
            common_stock=_pick(cs_vals, yr, d),
            additional_paid_in_capital=_pick(apic_vals, yr, d),
            retained_earnings=_pick(re_vals, yr, d),
            accumulated_other_comprehensive=_pick(aoci_vals, yr, d),
            treasury_stock=_pick(ts_vals, yr, d),
            total_equity=_pick(te_vals, yr, d) or 0.0,
        ))

    cash_flows = []
    for yr in sorted(target_years):
        # Derive beginning cash from prior year ending cash
        end_cash = _pick(end_cash_vals, yr, d)
        prior_yr = yr - 1
        beg_cash = _pick(end_cash_vals, prior_yr, d)

        cash_flows.append(CashFlowStatement(
            fiscal_year=str(yr),
            net_income=_pick(ni_vals, yr, d) or 0.0,
            depreciation_amortization=_pick(da_vals, yr, d),
            stock_based_compensation=_pick(sbc_vals, yr, d),
            capital_expenditures=_pick(capex_vals, yr, d),
            acquisitions=_pick(acq_vals, yr, d),
            purchases_investments=_pick(buy_inv_vals, yr, d),
            sales_investments=_pick(sell_inv_vals, yr, d),
            net_cash_operating=_pick(cfo_vals, yr, d) or 0.0,
            net_cash_investing=_pick(cfi_vals, yr, d),
            debt_repayment=_pick(debt_rep_vals, yr, d),
            debt_issuance=_pick(debt_iss_vals, yr, d),
            dividends_paid=_pick(divs_vals, yr, d),
            share_repurchases=_pick(buyback_vals, yr, d),
            share_issuance=_pick(share_iss_vals, yr, d),
            net_cash_financing=_pick(cff_vals, yr, d),
            net_change_in_cash=_pick(net_chg_vals, yr, d),
            beginning_cash=beg_cash,
            ending_cash=end_cash,
        ))

    metadata = CompanyMetadata(
        company_name=company_name_official,
        ticker=ticker.upper() if ticker else None,
        fiscal_year_end=fy_end,
        currency="USD",
        units=units_label,
        source_url=facts_url,
    )

    return FinancialData(
        metadata=metadata,
        income_statements=income_statements,
        balance_sheets=balance_sheets,
        cash_flow_statements=cash_flows,
    )
