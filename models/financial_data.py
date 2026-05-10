from __future__ import annotations
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class IncomeStatement:
    fiscal_year: str
    revenue: float
    cost_of_revenue: float
    gross_profit: float
    research_and_development: Optional[float] = None
    selling_general_admin: Optional[float] = None
    operating_expenses_other: Optional[float] = None
    depreciation_amortization: Optional[float] = None
    ebitda: Optional[float] = None
    ebit: Optional[float] = None
    interest_expense: Optional[float] = None
    interest_income: Optional[float] = None
    other_income_expense: Optional[float] = None
    pretax_income: Optional[float] = None
    income_tax: Optional[float] = None
    net_income: float = 0.0
    shares_basic: Optional[float] = None
    shares_diluted: Optional[float] = None
    eps_basic: Optional[float] = None
    eps_diluted: Optional[float] = None


@dataclass
class BalanceSheet:
    fiscal_year: str
    # Current Assets
    cash_and_equivalents: float = 0.0
    short_term_investments: Optional[float] = None
    accounts_receivable: Optional[float] = None
    inventory: Optional[float] = None
    prepaid_other_current: Optional[float] = None
    total_current_assets: Optional[float] = None
    # Non-current Assets
    ppe_gross: Optional[float] = None
    accumulated_depreciation: Optional[float] = None
    ppe_net: Optional[float] = None
    goodwill: Optional[float] = None
    intangibles: Optional[float] = None
    other_noncurrent_assets: Optional[float] = None
    total_noncurrent_assets: Optional[float] = None
    total_assets: float = 0.0
    # Current Liabilities
    accounts_payable: Optional[float] = None
    short_term_debt: Optional[float] = None
    accrued_liabilities: Optional[float] = None
    deferred_revenue_current: Optional[float] = None
    other_current_liabilities: Optional[float] = None
    total_current_liabilities: Optional[float] = None
    # Long-term Liabilities
    long_term_debt: Optional[float] = None
    deferred_tax_liabilities: Optional[float] = None
    other_noncurrent_liabilities: Optional[float] = None
    total_noncurrent_liabilities: Optional[float] = None
    total_liabilities: Optional[float] = None
    # Equity
    common_stock: Optional[float] = None
    additional_paid_in_capital: Optional[float] = None
    retained_earnings: Optional[float] = None
    accumulated_other_comprehensive: Optional[float] = None
    treasury_stock: Optional[float] = None
    total_equity: float = 0.0
    total_liabilities_and_equity: Optional[float] = None


@dataclass
class CashFlowStatement:
    fiscal_year: str
    # Operating Activities
    net_income: float = 0.0
    depreciation_amortization: Optional[float] = None
    stock_based_compensation: Optional[float] = None
    changes_in_working_capital: Optional[float] = None
    change_accounts_receivable: Optional[float] = None
    change_inventory: Optional[float] = None
    change_accounts_payable: Optional[float] = None
    change_other_working_capital: Optional[float] = None
    other_operating_activities: Optional[float] = None
    net_cash_operating: float = 0.0
    # Investing Activities
    capital_expenditures: Optional[float] = None
    acquisitions: Optional[float] = None
    purchases_investments: Optional[float] = None
    sales_investments: Optional[float] = None
    other_investing: Optional[float] = None
    net_cash_investing: Optional[float] = None
    # Financing Activities
    debt_issuance: Optional[float] = None
    debt_repayment: Optional[float] = None
    dividends_paid: Optional[float] = None
    share_repurchases: Optional[float] = None
    share_issuance: Optional[float] = None
    other_financing: Optional[float] = None
    net_cash_financing: Optional[float] = None
    # Reconciliation
    net_change_in_cash: Optional[float] = None
    beginning_cash: Optional[float] = None
    ending_cash: Optional[float] = None


@dataclass
class CompanyMetadata:
    company_name: str
    ticker: Optional[str] = None
    fiscal_year_end: Optional[str] = None
    currency: str = "USD"
    units: str = "millions"
    filing_date: Optional[str] = None
    period_of_report: Optional[str] = None
    auditor: Optional[str] = None
    source_url: Optional[str] = None


@dataclass
class FinancialData:
    metadata: CompanyMetadata
    income_statements: list[IncomeStatement] = field(default_factory=list)
    balance_sheets: list[BalanceSheet] = field(default_factory=list)
    cash_flow_statements: list[CashFlowStatement] = field(default_factory=list)

    def sorted_years(self) -> list[str]:
        return sorted({is_.fiscal_year for is_ in self.income_statements})

    def income_statement_for_year(self, year: str) -> Optional[IncomeStatement]:
        for is_ in self.income_statements:
            if is_.fiscal_year == year:
                return is_
        return None

    def balance_sheet_for_year(self, year: str) -> Optional[BalanceSheet]:
        for bs in self.balance_sheets:
            if bs.fiscal_year == year:
                return bs
        return None

    def cash_flow_for_year(self, year: str) -> Optional[CashFlowStatement]:
        for cf in self.cash_flow_statements:
            if cf.fiscal_year == year:
                return cf
        return None
