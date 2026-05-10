"""Orchestrate creation of the 5-sheet 3-statement model workbook."""
from __future__ import annotations
from pathlib import Path

from openpyxl import Workbook

import config as cfg
from models.financial_data import FinancialData
from excel.sheet_is import build_income_statement_sheet
from excel.sheet_bs import build_balance_sheet_sheet
from excel.sheet_cf import build_cash_flow_sheet
from excel.sheet_ratios import build_ratios_sheet
from excel.sheet_bridge import build_ebitda_bridge_sheet


def build_workbook(data: FinancialData, output_path: str) -> str:
    """
    Create the full 5-sheet analyst model workbook and save to output_path.
    Returns the resolved output path string.
    """
    wb = Workbook()

    # Remove default sheet and create the five named sheets in order
    default = wb.active
    wb.remove(default)

    for name in [cfg.SHEET_IS, cfg.SHEET_BS, cfg.SHEET_CF, cfg.SHEET_RATIOS, cfg.SHEET_BRIDGE]:
        wb.create_sheet(name)

    # Build sheets in dependency order: IS → BS → CF → Ratios → Bridge
    # (CF and Ratios reference named ranges defined by IS and BS)
    build_income_statement_sheet(wb, data)
    build_balance_sheet_sheet(wb, data)
    build_cash_flow_sheet(wb, data)
    build_ratios_sheet(wb, data)
    build_ebitda_bridge_sheet(wb, data)

    # Set active sheet to Income Statement
    wb.active = wb[cfg.SHEET_IS]

    # Workbook-level metadata
    wb.properties.title = f"{data.metadata.company_name} — 3-Statement Model"
    wb.properties.creator = "10k-model-builder"
    wb.properties.description = (
        f"Auto-generated from SEC 10-K filing. "
        f"Currency: {data.metadata.currency}. Units: {data.metadata.units}."
    )

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    return str(out)
