"""Phase 2 — Flask web server for the 10-K 3-statement model dashboard."""
from __future__ import annotations
import io
import json
import os
import sys
import tempfile
import dataclasses
from pathlib import Path

from flask import Flask, render_template, request, jsonify, send_file
from dotenv import load_dotenv

load_dotenv()
sys.path.insert(0, str(Path(__file__).parent))

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB max upload


def _to_dict(obj) -> dict | list | float | str | None:
    """Recursively convert dataclasses to JSON-serializable dicts."""
    if dataclasses.is_dataclass(obj) and not isinstance(obj, type):
        return {k: _to_dict(v) for k, v in dataclasses.asdict(obj).items()}
    if isinstance(obj, list):
        return [_to_dict(i) for i in obj]
    return obj


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/analyze", methods=["POST"])
def analyze():
    if "pdf" not in request.files:
        return jsonify({"error": "No PDF file uploaded."}), 400

    pdf_file = request.files["pdf"]
    if not pdf_file.filename.lower().endswith(".pdf"):
        return jsonify({"error": "File must be a PDF."}), 400

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        return jsonify({"error": "ANTHROPIC_API_KEY is not configured on the server."}), 500

    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        pdf_file.save(tmp.name)
        tmp_path = tmp.name

    try:
        from extractor.pdf_extractor import get_financial_section_text
        from extractor.claude_parser import extract_financials_with_claude

        financial_text = get_financial_section_text(tmp_path)
        data = extract_financials_with_claude(
            financial_text,
            source_hint=f"PDF: {pdf_file.filename}",
            use_cache=True,
        )

        return jsonify({"ok": True, "data": _to_dict(data)})

    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        Path(tmp_path).unlink(missing_ok=True)


@app.route("/analyze-xbrl", methods=["POST"])
def analyze_xbrl():
    """Fetch financial data from SEC EDGAR XBRL — no API key required."""
    body = request.get_json(force=True) or {}
    ticker = (body.get("ticker") or "").strip().upper() or None
    company = (body.get("company") or "").strip() or None

    if not ticker and not company:
        return jsonify({"error": "Provide a ticker symbol or company name."}), 400

    try:
        from extractor.xbrl_fetcher import fetch_financials_from_xbrl
        data = fetch_financials_from_xbrl(ticker=ticker, company_name=company)
        return jsonify({"ok": True, "data": _to_dict(data)})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/download-excel", methods=["POST"])
def download_excel():
    """Accept JSON financial data, build the Excel workbook, return it as a download."""
    payload = request.get_json(force=True)
    if not payload:
        return jsonify({"error": "No data provided."}), 400

    try:
        from models.financial_data import (
            FinancialData, CompanyMetadata,
            IncomeStatement, BalanceSheet, CashFlowStatement,
        )

        meta_raw = payload.get("metadata", {})
        metadata = CompanyMetadata(**{k: v for k, v in meta_raw.items()
                                      if k in CompanyMetadata.__dataclass_fields__})

        def build_list(raw_list, cls):
            return [cls(**{k: v for k, v in item.items()
                           if k in cls.__dataclass_fields__})
                    for item in (raw_list or [])]

        data = FinancialData(
            metadata=metadata,
            income_statements=build_list(payload.get("income_statements"), IncomeStatement),
            balance_sheets=build_list(payload.get("balance_sheets"), BalanceSheet),
            cash_flow_statements=build_list(payload.get("cash_flow_statements"), CashFlowStatement),
        )

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp_path = tmp.name

        from excel.workbook_builder import build_workbook
        build_workbook(data, tmp_path)

        company = data.metadata.company_name.replace(" ", "_")[:30]
        filename = f"{company}_3statement_model.xlsx"

        return send_file(
            tmp_path,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True, port=5000)
