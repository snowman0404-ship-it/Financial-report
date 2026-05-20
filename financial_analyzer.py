"""
SEC EDGAR Financial Analyzer
Fetches XBRL financial data for any US-listed company and exports
a formatted Excel workbook with PL (YoY) and BS (QoQ) analysis,
risk alerts, and color-coded highlights.

Usage:
    python financial_analyzer.py [TICKER]
    python financial_analyzer.py PARR
    python financial_analyzer.py XOM
    python financial_analyzer.py --demo          # offline test with sample data

Dependencies: pip install pandas requests openpyxl
"""

import sys
import json
import time
import requests
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

USER_AGENT = "FinancialAnalyzer/1.0 (financial-analyzer@example.com)"

EDGAR_TICKER_API = "https://www.sec.gov/files/company_tickers.json"
EDGAR_FACTS_API = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"

# Fallback CIK lookup for commonly used tickers
# (used when the SEC ticker API is unreachable)
KNOWN_CIKS = {
    "PARR": "0001378590",  # Par Pacific Holdings Inc.
    "XOM":  "0000034088",  # Exxon Mobil Corp
    "CVX":  "0000093410",  # Chevron Corp
    "TSLA": "0001318605",  # Tesla Inc
    "AAPL": "0000320193",  # Apple Inc
    "MSFT": "0000789019",  # Microsoft Corp
    "AMZN": "0001018724",  # Amazon.com Inc
    "GOOGL":"0001652044",  # Alphabet Inc
    "META": "0001326801",  # Meta Platforms Inc
    "NVDA": "0001045810",  # NVIDIA Corp
    "JPM":  "0000019617",  # JPMorgan Chase & Co
    "BAC":  "0000070858",  # Bank of America Corp
    "WMT":  "0000104169",  # Walmart Inc
    "PFE":  "0000078003",  # Pfizer Inc
    "JNJ":  "0000200406",  # Johnson & Johnson
    "MRK":  "0000310158",  # Merck & Co Inc
    "PSX":  "0001534992",  # Phillips 66
    "VLO":  "0001035002",  # Valero Energy Corp
    "MPC":  "0001510295",  # Marathon Petroleum Corp
    "HFC":  "0000048039",  # HF Sinclair (formerly Holly Frontier)
    "DK":   "0000049600",  # Delek Group
}

# ---------------------------------------------------------------------------
# XBRL Tag Fallback Maps  (ordered by preference)
# ---------------------------------------------------------------------------

PL_TAGS = {
    "Revenues": [
        "Revenues",
        "SalesRevenueNet",
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        "SalesRevenueGoodsNet",
        "RevenueFromContractWithCustomerIncludingAssessedTax",
        "SalesAndRevenuesNet",
    ],
    "OperatingExpenses": [
        "CostOfGoodsAndServicesSold",
        "CostOfRevenue",
        "OperatingExpenses",
        "CostsAndExpenses",
        "CostOfGoodsSold",
    ],
    "OperatingIncomeLoss": [
        "OperatingIncomeLoss",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
    ],
    "InterestExpense": [
        "NonoperatingIncomeExpense",
        "InterestExpense",
        "InterestIncomeExpenseNet",
        "InterestAndDebtExpense",
        "OtherNonoperatingIncomeExpense",
    ],
    "IncomeLossBeforeTax": [
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesDomestic",
    ],
    "NetIncomeLoss": [
        "NetIncomeLoss",
        "ProfitLoss",
        "NetIncomeLossAvailableToCommonStockholdersBasic",
        "IncomeLossFromContinuingOperations",
        "NetIncomeLossAttributableToParentCompany",
    ],
}

BS_TAGS = {
    "Cash": [
        "CashAndCashEquivalentsAtCarryingValue",
        "CashCashEquivalentsAndShortTermInvestments",
        "CashAndShortTermInvestments",
        "Cash",
    ],
    "CurrentLiabilities": [
        "LiabilitiesCurrent",
        "LiabilitiesCurrentAndNoncurrent",
    ],
    "CurrentAssets": [
        "AssetsCurrent",
        "AssetsCurrentAndNoncurrent",
    ],
    "LongTermLiabilities": [
        "LiabilitiesNoncurrent",
        "LongTermDebtNoncurrent",
        "LongTermDebt",
        "LongTermDebtAndCapitalLeaseObligations",
    ],
    "NonCurrentAssets": [
        "AssetsNoncurrent",
        "PropertyPlantAndEquipmentNet",
        "PropertyPlantAndEquipmentAndIntangibleAssetsNet",
    ],
    "StockholdersEquity": [
        "StockholdersEquity",
        "StockholdersEquityAttributableToParent",
        "PartnersCapital",
        "MembersEquity",
        "LimitedLiabilityCompanyLlcMembersEquityIncludingPortionAttributableToNoncontrollingInterest",
        "PartnershipCapital",
    ],
}

# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------

RED_FILL = PatternFill("solid", fgColor="FFD2D2")
GREEN_FILL = PatternFill("solid", fgColor="D2FFD2")
YELLOW_FILL = PatternFill("solid", fgColor="FFFACD")
BLUE_FILL = PatternFill("solid", fgColor="D2E4FF")
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
SUBHEADER_FILL = PatternFill("solid", fgColor="2E75B6")
SECTION_FILL = PatternFill("solid", fgColor="BDD7EE")
LIGHT_GREY_FILL = PatternFill("solid", fgColor="F2F2F2")

WHITE_BOLD = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
DARK_BOLD = Font(name="Calibri", bold=True, color="1F4E79", size=11)
NORMAL_FONT = Font(name="Calibri", size=10)
SMALL_ITALIC = Font(name="Calibri", size=9, italic=True, color="595959")

THIN = Side(style="thin", color="BFBFBF")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")

FMT_USD = '#,##0_);[Red](#,##0)'
FMT_PCT = '0.0%;[Red]-0.0%'
FMT_RATIO = '0.00'


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _headers():
    return {"User-Agent": USER_AGENT, "Accept": "application/json"}


def _get(url: str, retries: int = 4, backoff: float = 2.0):
    """HTTP GET with exponential back-off."""
    for attempt in range(retries):
        try:
            resp = requests.get(url, headers=_headers(), timeout=30)
            if resp.status_code == 200:
                return resp
            if resp.status_code == 429:
                wait = backoff * (2 ** attempt)
                print(f"  Rate-limited. Retrying in {wait:.0f}s …")
                time.sleep(wait)
                continue
            print(f"  HTTP {resp.status_code} for {url}")
            return None
        except requests.RequestException as exc:
            wait = backoff * (2 ** attempt)
            print(f"  Network error: {exc}. Retrying in {wait:.0f}s …")
            time.sleep(wait)
    return None


def resolve_cik(ticker: str) -> str | None:
    """Return zero-padded 10-digit CIK for a ticker symbol."""
    ticker_upper = ticker.upper()
    print(f"Looking up CIK for {ticker_upper} …")

    # Try live SEC ticker API first
    resp = _get(EDGAR_TICKER_API)
    if resp is not None:
        data = resp.json()
        for _key, entry in data.items():
            if entry.get("ticker", "").upper() == ticker_upper:
                cik = str(entry["cik_str"]).zfill(10)
                print(f"  Found CIK: {cik}  ({entry.get('title', '')})")
                return cik
        print(f"  WARN: '{ticker_upper}' not found via live API. Trying fallback table …")

    # Fallback to hardcoded table
    if ticker_upper in KNOWN_CIKS:
        cik = KNOWN_CIKS[ticker_upper]
        print(f"  Using fallback CIK: {cik}")
        return cik

    print(f"  ERROR: Ticker '{ticker_upper}' not found in SEC database or fallback table.")
    return None


def fetch_facts(cik: str) -> dict:
    """Download and return the full company-facts JSON from EDGAR."""
    url = EDGAR_FACTS_API.format(cik=cik)
    print(f"Fetching EDGAR facts from {url} …")
    resp = _get(url)
    if resp is None:
        return {}
    return resp.json()


# ---------------------------------------------------------------------------
# Data Extraction
# ---------------------------------------------------------------------------

def _units(facts: dict, tag: str) -> list[dict]:
    """Return the list of unit-records for a US-GAAP tag, or []."""
    try:
        gaap = facts.get("facts", {}).get("us-gaap", {})
        tag_data = gaap.get(tag, {})
        units = tag_data.get("units", {})
        # prefer USD, fall back to pure (ratios etc.)
        return units.get("USD", units.get("pure", []))
    except Exception:
        return []


def _best_tag(facts: dict, candidates: list[str]) -> tuple[str, list[dict]]:
    """Try each candidate tag in order; return (tag_name, records)."""
    for tag in candidates:
        records = _units(facts, tag)
        if records:
            return tag, records
    return "", []


def _filter_quarterly(records: list[dict]) -> list[dict]:
    """Keep only 3-month (quarterly) period records."""
    out = []
    for r in records:
        start = r.get("start", "")
        end = r.get("end", "")
        if start and end:
            try:
                delta = (
                    datetime.strptime(end, "%Y-%m-%d")
                    - datetime.strptime(start, "%Y-%m-%d")
                ).days
                if 80 <= delta <= 100:
                    out.append(r)
            except ValueError:
                pass
    return out


def _filter_instantaneous(records: list[dict]) -> list[dict]:
    """Keep only point-in-time (balance sheet) records — no 'start'."""
    return [r for r in records if not r.get("start")]


def _latest_n_quarters(records: list[dict], n: int) -> list[dict]:
    """Return the n most recent unique quarters sorted newest-first."""
    seen_end = {}
    for r in records:
        end = r.get("end", "")
        form = r.get("form", "")
        if end not in seen_end:
            seen_end[end] = r
        elif form in ("10-Q", "10-K"):
            seen_end[end] = r
    sorted_recs = sorted(seen_end.values(), key=lambda x: x.get("end", ""), reverse=True)
    return sorted_recs[:n]


def _latest_n_instants(records: list[dict], n: int) -> list[dict]:
    """Return the n most recent balance-sheet snapshots."""
    seen_end = {}
    for r in records:
        end = r.get("end", "")
        form = r.get("form", "")
        if end not in seen_end:
            seen_end[end] = r
        elif form in ("10-Q", "10-K"):
            seen_end[end] = r
    sorted_recs = sorted(seen_end.values(), key=lambda x: x.get("end", ""), reverse=True)
    return sorted_recs[:n]


def extract_pl(facts: dict) -> dict:
    """
    Returns dict keyed by metric name:
        {"current": (value, date, tag), "prior": (value, date, tag)}
    'current' = most recent quarter, 'prior' = same quarter last year.
    """
    result = {}
    for metric, candidates in PL_TAGS.items():
        tag, records = _best_tag(facts, candidates)
        quarterly = _filter_quarterly(records)
        latest_quarters = _latest_n_quarters(quarterly, 8)

        current = prior = None
        if latest_quarters:
            current_rec = latest_quarters[0]
            current_end = datetime.strptime(current_rec["end"], "%Y-%m-%d")
            current = (current_rec.get("val"), current_rec["end"], tag)

            # find same quarter last year (±45 days around end-365)
            target_prior = current_end - timedelta(days=365)
            best = None
            best_diff = 999
            for rec in latest_quarters[1:]:
                rec_end = datetime.strptime(rec["end"], "%Y-%m-%d")
                diff = abs((rec_end - target_prior).days)
                if diff < 46 and diff < best_diff:
                    best = rec
                    best_diff = diff
            if best:
                prior = (best.get("val"), best["end"], tag)

        result[metric] = {"current": current, "prior": prior}
    return result


def extract_bs(facts: dict) -> dict:
    """
    Returns dict keyed by metric name:
        {"current": (value, date, tag), "prior": (value, date, tag)}
    'current' = latest quarter-end, 'prior' = one quarter earlier.
    """
    result = {}
    for metric, candidates in BS_TAGS.items():
        tag, records = _best_tag(facts, candidates)
        instants = _filter_instantaneous(records)
        latest = _latest_n_instants(instants, 4)

        current = prior = None
        if latest:
            current = (latest[0].get("val"), latest[0]["end"], tag)
        if len(latest) > 1:
            prior = (latest[1].get("val"), latest[1]["end"], tag)

        result[metric] = {"current": current, "prior": prior}
    return result


# ---------------------------------------------------------------------------
# Excel Generation
# ---------------------------------------------------------------------------

def _set_col_widths(ws, widths: list[float]):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _style_cell(cell, fill=None, font=None, alignment=None,
                border=None, number_format=None):
    if fill:
        cell.fill = fill
    if font:
        cell.font = font
    if alignment:
        cell.alignment = alignment
    if border:
        cell.border = border
    if number_format:
        cell.number_format = number_format


def _write_header_row(ws, row: int, texts: list, fills=None, fonts=None):
    for col, text in enumerate(texts, 1):
        cell = ws.cell(row=row, column=col, value=text)
        cell.fill = fills[col - 1] if fills else HEADER_FILL
        cell.font = fonts[col - 1] if fonts else WHITE_BOLD
        cell.alignment = CENTER
        cell.border = THIN_BORDER


def _fmt_val(val):
    """Return numeric value in millions (float) or None."""
    if val is None:
        return None
    try:
        return float(val) / 1_000_000
    except (TypeError, ValueError):
        return None


def build_financial_data_sheet(wb: Workbook, company_name: str, ticker: str,
                                pl: dict, bs: dict):
    ws = wb.create_sheet("Financial Data")
    ws.views.sheetView[0].showGridLines = True

    _set_col_widths(ws, [3, 36, 18, 18, 18, 14, 14])

    # ── Title ──────────────────────────────────────────────────────────────
    ws.merge_cells("A1:G1")
    title_cell = ws["A1"]
    title_cell.value = f"SEC EDGAR Financial Analysis  |  {company_name} ({ticker.upper()})"
    title_cell.fill = HEADER_FILL
    title_cell.font = Font(name="Calibri", bold=True, color="FFFFFF", size=14)
    title_cell.alignment = CENTER
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:G2")
    sub_cell = ws["A2"]
    sub_cell.value = (
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}   "
        "| Source: SEC EDGAR XBRL API   "
        "| Amounts in USD millions (M)"
    )
    sub_cell.fill = SUBHEADER_FILL
    sub_cell.font = Font(name="Calibri", color="FFFFFF", size=9, italic=True)
    sub_cell.alignment = CENTER
    ws.row_dimensions[2].height = 16

    # ── PL Section ─────────────────────────────────────────────────────────
    current_row = 4
    ws.merge_cells(f"A{current_row}:G{current_row}")
    sec_cell = ws.cell(row=current_row, column=1,
                       value="📊  INCOME STATEMENT (P&L) — Quarterly YoY Comparison")
    sec_cell.fill = SECTION_FILL
    sec_cell.font = DARK_BOLD
    sec_cell.alignment = LEFT
    ws.row_dimensions[current_row].height = 22
    current_row += 1

    # Determine period labels
    pl_current_date = "Current Q"
    pl_prior_date = "Prior Year Q"
    for _metric, data in pl.items():
        if data.get("current"):
            pl_current_date = data["current"][1]
            break
    for _metric, data in pl.items():
        if data.get("prior"):
            pl_prior_date = data["prior"][1]
            break

    headers = [
        "", "Metric",
        f"Current\n({pl_current_date})\n[USD M]",
        f"Prior Year\n({pl_prior_date})\n[USD M]",
        "Δ Amount\n[USD M]",
        "Δ %",
        "Alert"
    ]
    _write_header_row(ws, current_row, headers)
    ws.row_dimensions[current_row].height = 40
    header_row = current_row
    current_row += 1

    pl_display_order = [
        ("Revenues",          "Revenues (売上高)"),
        ("OperatingExpenses", "Operating Expenses (営業費用)"),
        ("OperatingIncomeLoss", "Operating Income/Loss (営業利益)"),
        ("InterestExpense",   "Non-Operating Income/Expense (金融損益)"),
        ("IncomeLossBeforeTax", "Income Before Tax (税引前利益)"),
        ("NetIncomeLoss",     "Net Income/Loss (純利益)"),
    ]

    pl_rows = {}
    for metric_key, label in pl_display_order:
        data = pl.get(metric_key, {})
        current_val = _fmt_val(data.get("current", (None,))[0])
        prior_val = _fmt_val(data.get("prior", (None,))[0])
        tag_used = data.get("current", (None, None, ""))[2] if data.get("current") else ""

        row_num = current_row
        pl_rows[metric_key] = row_num

        # Column A: index marker
        ws.cell(row=row_num, column=1).fill = LIGHT_GREY_FILL

        # Column B: label
        label_cell = ws.cell(row=row_num, column=2, value=label)
        label_cell.font = NORMAL_FONT
        label_cell.alignment = LEFT
        label_cell.border = THIN_BORDER

        # Column C: current value
        c_cell = ws.cell(row=row_num, column=3, value=current_val)
        c_cell.number_format = FMT_USD
        c_cell.alignment = RIGHT
        c_cell.border = THIN_BORDER
        if current_val is None:
            c_cell.value = "N/A"
            c_cell.font = SMALL_ITALIC

        # Column D: prior value
        d_cell = ws.cell(row=row_num, column=4, value=prior_val)
        d_cell.number_format = FMT_USD
        d_cell.alignment = RIGHT
        d_cell.border = THIN_BORDER
        if prior_val is None:
            d_cell.value = "N/A"
            d_cell.font = SMALL_ITALIC

        c_addr = f"C{row_num}"
        d_addr = f"D{row_num}"

        # Column E: Delta formula (Excel formula)
        e_cell = ws.cell(row=row_num, column=5)
        if current_val is not None and prior_val is not None:
            e_cell.value = f"={c_addr}-{d_addr}"
            e_cell.number_format = FMT_USD
        else:
            e_cell.value = "N/A"
            e_cell.font = SMALL_ITALIC
        e_cell.alignment = RIGHT
        e_cell.border = THIN_BORDER

        # Column F: % Change formula (Excel formula)
        f_cell = ws.cell(row=row_num, column=6)
        if current_val is not None and prior_val is not None and prior_val != 0:
            f_cell.value = f"=IF({d_addr}=0,\"\",({c_addr}-{d_addr})/{d_addr})"
            f_cell.number_format = FMT_PCT
        elif current_val is not None and prior_val is not None and prior_val == 0:
            f_cell.value = "∞"
        else:
            f_cell.value = "N/A"
            f_cell.font = SMALL_ITALIC
        f_cell.alignment = CENTER
        f_cell.border = THIN_BORDER

        # Column G: Alert logic
        g_cell = ws.cell(row=row_num, column=7)
        is_cost = metric_key == "OperatingExpenses"
        alert_triggered = False

        if current_val is not None and prior_val is not None and prior_val != 0:
            pct_change = (current_val - prior_val) / abs(prior_val)
            if is_cost:
                if pct_change > 0.20:
                    g_cell.value = "⚠️ Cost +20%↑"
                    alert_triggered = True
                else:
                    g_cell.value = "✅ Normal"
            else:
                if pct_change < -0.20:
                    g_cell.value = "⚠️ -20%↓ Drop"
                    alert_triggered = True
                elif pct_change > 0.20:
                    g_cell.value = "✅ +20%↑ Growth"
                else:
                    g_cell.value = "✅ Normal"
        else:
            g_cell.value = "—"
        g_cell.alignment = CENTER
        g_cell.border = THIN_BORDER

        # Row highlight on alert
        if alert_triggered:
            for col in range(2, 8):
                ws.cell(row=row_num, column=col).fill = RED_FILL

        # Tag footnote in tooltip-style (column A)
        a_cell = ws.cell(row=row_num, column=1)
        if tag_used:
            a_cell.value = "ℹ"
            a_cell.comment = None  # openpyxl comments need extra import; skip
        a_cell.alignment = CENTER
        a_cell.border = THIN_BORDER

        ws.row_dimensions[row_num].height = 18
        current_row += 1

    # Non-GAAP row for energy companies
    row_num = current_row
    ws.cell(row=row_num, column=1).fill = YELLOW_FILL
    ws.cell(row=row_num, column=1).value = "★"
    ws.cell(row=row_num, column=1).alignment = CENTER
    ws.cell(row=row_num, column=1).border = THIN_BORDER

    ng_label = ws.cell(row=row_num, column=2,
                       value="Operating Income ex-Inventory Adj. (在庫影響除き営業利益) [Non-GAAP]")
    ng_label.font = Font(name="Calibri", size=10, italic=True, color="7F6000")
    ng_label.alignment = LEFT
    ng_label.border = THIN_BORDER

    ng_note = ws.cell(row=row_num, column=3,
                      value="※ PARR等エネルギー企業では10-Q MD&Aの「Inventory Valuation Adjustment」を確認し、手動で調整してください。")
    ng_note.font = Font(name="Calibri", size=8, italic=True, color="7F6000")
    ng_note.alignment = LEFT
    ws.merge_cells(f"C{row_num}:G{row_num}")
    ng_note.border = THIN_BORDER
    ws.row_dimensions[row_num].height = 28
    current_row += 2

    # ── BS Section ─────────────────────────────────────────────────────────
    ws.merge_cells(f"A{current_row}:G{current_row}")
    sec2_cell = ws.cell(row=current_row, column=1,
                        value="🏦  BALANCE SHEET — Quarter-over-Quarter (QoQ) Comparison")
    sec2_cell.fill = SECTION_FILL
    sec2_cell.font = DARK_BOLD
    sec2_cell.alignment = LEFT
    ws.row_dimensions[current_row].height = 22
    current_row += 1

    bs_current_date = "Current Q"
    bs_prior_date = "Prior Q"
    for _metric, data in bs.items():
        if data.get("current"):
            bs_current_date = data["current"][1]
            break
    for _metric, data in bs.items():
        if data.get("prior"):
            bs_prior_date = data["prior"][1]
            break

    headers_bs = [
        "", "Metric",
        f"Current QE\n({bs_current_date})\n[USD M]",
        f"Prior QE\n({bs_prior_date})\n[USD M]",
        "Δ Amount\n[USD M]",
        "Δ %",
        "Alert"
    ]
    _write_header_row(ws, current_row, headers_bs)
    ws.row_dimensions[current_row].height = 40
    current_row += 1

    bs_display_order = [
        ("Cash",              "Cash & Equivalents (手元資金)",           False),
        ("CurrentLiabilities","Current Liabilities (流動負債)",          True),
        ("CurrentAssets",     "Current Assets (流動資産)",               False),
        ("LongTermLiabilities","Non-Current Liabilities / LT Debt (長期負債)", True),
        ("NonCurrentAssets",  "Non-Current Assets / PP&E (固定資産)",    False),
        ("StockholdersEquity","Stockholders' Equity (株主資本)",         False),
    ]

    bs_rows = {}
    for metric_key, label, is_liability in bs_display_order:
        data = bs.get(metric_key, {})
        current_val = _fmt_val(data.get("current", (None,))[0])
        prior_val = _fmt_val(data.get("prior", (None,))[0])
        tag_used = data.get("current", (None, None, ""))[2] if data.get("current") else ""

        row_num = current_row
        bs_rows[metric_key] = row_num

        ws.cell(row=row_num, column=1).fill = LIGHT_GREY_FILL
        ws.cell(row=row_num, column=1).border = THIN_BORDER

        label_cell = ws.cell(row=row_num, column=2, value=label)
        label_cell.font = NORMAL_FONT
        label_cell.alignment = LEFT
        label_cell.border = THIN_BORDER

        c_cell = ws.cell(row=row_num, column=3, value=current_val)
        c_cell.number_format = FMT_USD
        c_cell.alignment = RIGHT
        c_cell.border = THIN_BORDER
        if current_val is None:
            c_cell.value = "N/A"
            c_cell.font = SMALL_ITALIC

        d_cell = ws.cell(row=row_num, column=4, value=prior_val)
        d_cell.number_format = FMT_USD
        d_cell.alignment = RIGHT
        d_cell.border = THIN_BORDER
        if prior_val is None:
            d_cell.value = "N/A"
            d_cell.font = SMALL_ITALIC

        c_addr = f"C{row_num}"
        d_addr = f"D{row_num}"

        e_cell = ws.cell(row=row_num, column=5)
        if current_val is not None and prior_val is not None:
            e_cell.value = f"={c_addr}-{d_addr}"
            e_cell.number_format = FMT_USD
        else:
            e_cell.value = "N/A"
            e_cell.font = SMALL_ITALIC
        e_cell.alignment = RIGHT
        e_cell.border = THIN_BORDER

        f_cell = ws.cell(row=row_num, column=6)
        if current_val is not None and prior_val is not None and prior_val != 0:
            f_cell.value = f"=IF({d_addr}=0,\"\",({c_addr}-{d_addr})/{d_addr})"
            f_cell.number_format = FMT_PCT
        elif current_val is not None and prior_val is not None and prior_val == 0:
            f_cell.value = "∞"
        else:
            f_cell.value = "N/A"
            f_cell.font = SMALL_ITALIC
        f_cell.alignment = CENTER
        f_cell.border = THIN_BORDER

        g_cell = ws.cell(row=row_num, column=7)
        alert_triggered = False
        if current_val is not None and prior_val is not None and prior_val != 0:
            pct_change = (current_val - prior_val) / abs(prior_val)
            if is_liability:
                if pct_change > 0.20:
                    g_cell.value = "⚠️ Liab +20%↑"
                    alert_triggered = True
                else:
                    g_cell.value = "✅ Normal"
            else:
                if pct_change < -0.20:
                    g_cell.value = "⚠️ -20%↓ Drop"
                    alert_triggered = True
                elif pct_change > 0.20:
                    g_cell.value = "✅ +20%↑ Growth"
                else:
                    g_cell.value = "✅ Normal"
        else:
            g_cell.value = "—"
        g_cell.alignment = CENTER
        g_cell.border = THIN_BORDER

        if alert_triggered:
            for col in range(2, 8):
                ws.cell(row=row_num, column=col).fill = RED_FILL

        ws.row_dimensions[row_num].height = 18
        current_row += 1

    # Other Current Assets (derived)
    row_num = current_row
    cash_row = bs_rows.get("Cash")
    ca_row = bs_rows.get("CurrentAssets")
    ws.cell(row=row_num, column=1).fill = LIGHT_GREY_FILL
    ws.cell(row=row_num, column=1).border = THIN_BORDER
    label_cell = ws.cell(row=row_num, column=2,
                         value="Other Current Assets (その他流動資産) = CurrentAssets − Cash")
    label_cell.font = Font(name="Calibri", size=10, italic=True)
    label_cell.alignment = LEFT
    label_cell.border = THIN_BORDER

    if cash_row and ca_row:
        c_cell = ws.cell(row=row_num, column=3)
        c_cell.value = f"=C{ca_row}-C{cash_row}"
        c_cell.number_format = FMT_USD
        c_cell.alignment = RIGHT
        c_cell.border = THIN_BORDER

        d_cell = ws.cell(row=row_num, column=4)
        d_cell.value = f"=D{ca_row}-D{cash_row}"
        d_cell.number_format = FMT_USD
        d_cell.alignment = RIGHT
        d_cell.border = THIN_BORDER

        e_cell = ws.cell(row=row_num, column=5)
        e_cell.value = f"=C{row_num}-D{row_num}"
        e_cell.number_format = FMT_USD
        e_cell.alignment = RIGHT
        e_cell.border = THIN_BORDER

        f_cell = ws.cell(row=row_num, column=6)
        f_cell.value = f"=IF(D{row_num}=0,\"\",(C{row_num}-D{row_num})/D{row_num})"
        f_cell.number_format = FMT_PCT
        f_cell.alignment = CENTER
        f_cell.border = THIN_BORDER
    else:
        for col in range(3, 8):
            c = ws.cell(row=row_num, column=col, value="N/A")
            c.font = SMALL_ITALIC
            c.border = THIN_BORDER

    g_cell = ws.cell(row=row_num, column=7, value="(計算値)")
    g_cell.alignment = CENTER
    g_cell.border = THIN_BORDER
    ws.row_dimensions[row_num].height = 18
    current_row += 2

    # Legend
    ws.merge_cells(f"A{current_row}:G{current_row}")
    legend_cell = ws.cell(row=current_row, column=1,
                           value="[凡例] ⚠️ = Alert (20%以上の悪化)  |  ✅ = Normal  |  N/A = データ未取得  "
                                 "|  ★ = Non-GAAP（要手動確認）  |  Amounts in USD M (百万ドル単位)")
    legend_cell.font = SMALL_ITALIC
    legend_cell.fill = LIGHT_GREY_FILL
    legend_cell.alignment = LEFT
    ws.row_dimensions[current_row].height = 16

    return bs_rows


def build_dashboard_sheet(wb: Workbook, company_name: str, ticker: str,
                           cik: str, pl: dict, bs: dict, bs_rows: dict,
                           facts: dict):
    ws = wb.create_sheet("Dashboard", 0)
    ws.views.sheetView[0].showGridLines = True

    _set_col_widths(ws, [3, 28, 22, 22, 18, 14, 3])

    # Title
    ws.merge_cells("A1:G1")
    t = ws["A1"]
    t.value = f"Financial Health Dashboard  |  {company_name} ({ticker.upper()})"
    t.fill = HEADER_FILL
    t.font = Font(name="Calibri", bold=True, color="FFFFFF", size=16)
    t.alignment = CENTER
    ws.row_dimensions[1].height = 36

    ws.merge_cells("A2:G2")
    s = ws["A2"]
    s.value = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  CIK: {cik}  |  Source: SEC EDGAR XBRL"
    s.fill = SUBHEADER_FILL
    s.font = Font(name="Calibri", color="FFFFFF", size=9, italic=True)
    s.alignment = CENTER
    ws.row_dimensions[2].height = 16

    # ── Company Info ──
    row = 4
    ws.merge_cells(f"A{row}:G{row}")
    ci = ws.cell(row=row, column=1, value="🏢  Company Information")
    ci.fill = SECTION_FILL
    ci.font = DARK_BOLD
    ci.alignment = LEFT
    ws.row_dimensions[row].height = 22
    row += 1

    info_rows = [
        ("Company Name", company_name),
        ("Ticker Symbol", ticker.upper()),
        ("CIK (SEC)", cik),
        ("Data Source", "SEC EDGAR XBRL API (data.sec.gov)"),
        ("Analysis Date", datetime.now().strftime("%Y-%m-%d")),
        ("Reporting Currency", "USD"),
        ("Amount Unit", "USD Millions (M)"),
    ]
    for label, val in info_rows:
        l_cell = ws.cell(row=row, column=2, value=label)
        l_cell.font = Font(name="Calibri", bold=True, size=10)
        l_cell.alignment = LEFT
        l_cell.border = THIN_BORDER
        l_cell.fill = LIGHT_GREY_FILL

        v_cell = ws.cell(row=row, column=3, value=val)
        v_cell.font = NORMAL_FONT
        v_cell.alignment = LEFT
        v_cell.border = THIN_BORDER
        ws.merge_cells(f"C{row}:G{row}")
        ws.row_dimensions[row].height = 17
        row += 1

    row += 1

    # ── Key Ratios ──
    ws.merge_cells(f"A{row}:G{row}")
    kr = ws.cell(row=row, column=1, value="📐  Key Financial Ratios (数式連携・自動計算)")
    kr.fill = SECTION_FILL
    kr.font = DARK_BOLD
    kr.alignment = LEFT
    ws.row_dimensions[row].height = 22
    row += 1

    _write_header_row(ws, row, ["", "Ratio", "Current Value", "Prior Value", "Change", "Status", ""])
    ws.row_dimensions[row].height = 22
    row += 1

    cash_row = bs_rows.get("Cash")
    cl_row = bs_rows.get("CurrentLiabilities")
    ca_row = bs_rows.get("CurrentAssets")
    eq_row = bs_rows.get("StockholdersEquity")
    nca_row = bs_rows.get("NonCurrentAssets")

    fd = "'Financial Data'"

    # Current Ratio
    cr_row = row
    ws.cell(row=row, column=1).border = THIN_BORDER
    l = ws.cell(row=row, column=2, value="Current Ratio (流動比率)")
    l.font = NORMAL_FONT
    l.alignment = LEFT
    l.border = THIN_BORDER

    cur_formula = prior_formula = "N/A"
    if ca_row and cl_row:
        cur_formula = f"={fd}!C{ca_row}/{fd}!C{cl_row}"
        prior_formula = f"={fd}!D{ca_row}/{fd}!D{cl_row}"

    c3 = ws.cell(row=row, column=3)
    c3.value = cur_formula if ca_row and cl_row else "N/A"
    if ca_row and cl_row:
        c3.number_format = FMT_RATIO
    c3.alignment = RIGHT
    c3.border = THIN_BORDER

    c4 = ws.cell(row=row, column=4)
    c4.value = prior_formula if ca_row and cl_row else "N/A"
    if ca_row and cl_row:
        c4.number_format = FMT_RATIO
    c4.alignment = RIGHT
    c4.border = THIN_BORDER

    c5 = ws.cell(row=row, column=5)
    if ca_row and cl_row:
        c5.value = f"=C{row}-D{row}"
        c5.number_format = FMT_RATIO
    else:
        c5.value = "N/A"
    c5.alignment = RIGHT
    c5.border = THIN_BORDER

    # Status: Current Ratio < 1 = warning
    c6 = ws.cell(row=row, column=6)
    if ca_row and cl_row:
        c6.value = f'=IF(C{row}<1,"⚠️ <1.0 Low",IF(C{row}<1.5,"⚡ Moderate","✅ Healthy"))'
        bs_cash = bs.get("Cash", {})
        bs_cl = bs.get("CurrentLiabilities", {})
        bs_ca = bs.get("CurrentAssets", {})
        cur_ca = _fmt_val(bs_ca.get("current", (None,))[0])
        cur_cl = _fmt_val(bs_cl.get("current", (None,))[0])
        if cur_ca is not None and cur_cl is not None and cur_cl != 0:
            ratio = cur_ca / cur_cl
            if ratio < 1.0:
                c6.fill = RED_FILL
            elif ratio < 1.5:
                c6.fill = YELLOW_FILL
            else:
                c6.fill = GREEN_FILL
    else:
        c6.value = "N/A"
    c6.alignment = CENTER
    c6.border = THIN_BORDER
    ws.cell(row=row, column=7).border = THIN_BORDER
    ws.row_dimensions[row].height = 18
    row += 1

    # Equity Ratio
    eq_ratio_row = row
    l = ws.cell(row=row, column=2, value="Equity Ratio / 自己資本比率 (StockholdersEquity / TotalAssets)")
    l.font = NORMAL_FONT
    l.alignment = LEFT
    l.border = THIN_BORDER

    # Total assets = CA + NCA
    if ca_row and nca_row and eq_row:
        total_assets_cur = f"({fd}!C{ca_row}+{fd}!C{nca_row})"
        total_assets_pri = f"({fd}!D{ca_row}+{fd}!D{nca_row})"
        eq_cur_f = f"={fd}!C{eq_row}/{total_assets_cur}"
        eq_pri_f = f"={fd}!D{eq_row}/{total_assets_pri}"

        c3 = ws.cell(row=row, column=3, value=eq_cur_f)
        c3.number_format = FMT_PCT
        c3.alignment = RIGHT
        c3.border = THIN_BORDER

        c4 = ws.cell(row=row, column=4, value=eq_pri_f)
        c4.number_format = FMT_PCT
        c4.alignment = RIGHT
        c4.border = THIN_BORDER

        c5 = ws.cell(row=row, column=5, value=f"=C{row}-D{row}")
        c5.number_format = FMT_PCT
        c5.alignment = RIGHT
        c5.border = THIN_BORDER

        c6 = ws.cell(row=row, column=6)
        c6.value = f'=IF(C{row}<0.1,"⚠️ <10% Critical",IF(C{row}<0.3,"⚡ <30% Low","✅ Healthy"))'
        # Color by actual value
        bs_eq = bs.get("StockholdersEquity", {})
        bs_ca2 = bs.get("CurrentAssets", {})
        bs_nca = bs.get("NonCurrentAssets", {})
        cur_eq = _fmt_val(bs_eq.get("current", (None,))[0])
        cur_ca2 = _fmt_val(bs_ca2.get("current", (None,))[0])
        cur_nca = _fmt_val(bs_nca.get("current", (None,))[0])
        if cur_eq is not None and cur_ca2 is not None and cur_nca is not None:
            total = cur_ca2 + cur_nca
            if total != 0:
                ratio = cur_eq / total
                if ratio < 0.10:
                    c6.fill = RED_FILL
                elif ratio < 0.30:
                    c6.fill = YELLOW_FILL
                else:
                    c6.fill = GREEN_FILL
        c6.alignment = CENTER
        c6.border = THIN_BORDER
    else:
        for col in [3, 4, 5, 6]:
            c = ws.cell(row=row, column=col, value="N/A")
            c.border = THIN_BORDER
    ws.cell(row=row, column=1).border = THIN_BORDER
    ws.cell(row=row, column=7).border = THIN_BORDER
    ws.row_dimensions[row].height = 18
    row += 2

    # ── Cash Burn / Insolvency Alert ──
    ws.merge_cells(f"A{row}:G{row}")
    al = ws.cell(row=row, column=1, value="🚨  Bankruptcy / Cash Flow Risk Alert (倒産予兆・資金繰りリスク判定)")
    al.fill = PatternFill("solid", fgColor="C00000")
    al.font = Font(name="Calibri", bold=True, color="FFFFFF", size=12)
    al.alignment = LEFT
    ws.row_dimensions[row].height = 26
    row += 1

    # Sub-indicators
    _write_header_row(ws, row, ["", "Indicator", "Result", "Threshold", "Status", "", ""])
    ws.row_dimensions[row].height = 22
    row += 1

    bs_cash = bs.get("Cash", {})
    bs_cl2 = bs.get("CurrentLiabilities", {})
    bs_eq2 = bs.get("StockholdersEquity", {})

    cur_cash = _fmt_val(bs_cash.get("current", (None,))[0])
    pri_cash = _fmt_val(bs_cash.get("prior", (None,))[0])
    cur_cl2 = _fmt_val(bs_cl2.get("current", (None,))[0])
    pri_cl2 = _fmt_val(bs_cl2.get("prior", (None,))[0])
    cur_eq2 = _fmt_val(bs_eq2.get("current", (None,))[0])
    pri_eq2 = _fmt_val(bs_eq2.get("prior", (None,))[0])

    alert_indicators = []

    def _indicator_row(ws, row, name, result_text, threshold, status, is_bad):
        ws.cell(row=row, column=1).border = THIN_BORDER
        l = ws.cell(row=row, column=2, value=name)
        l.font = NORMAL_FONT
        l.alignment = LEFT
        l.border = THIN_BORDER
        ws.merge_cells(f"B{row}:B{row}")

        r = ws.cell(row=row, column=3, value=result_text)
        r.font = NORMAL_FONT
        r.alignment = CENTER
        r.border = THIN_BORDER

        t = ws.cell(row=row, column=4, value=threshold)
        t.font = SMALL_ITALIC
        t.alignment = CENTER
        t.border = THIN_BORDER

        s = ws.cell(row=row, column=5, value=status)
        s.alignment = CENTER
        s.border = THIN_BORDER
        s.fill = RED_FILL if is_bad else GREEN_FILL

        for col in [6, 7]:
            ws.cell(row=row, column=col).border = THIN_BORDER
        ws.row_dimensions[row].height = 18
        return is_bad

    # 1. Cash decline > 20%
    cash_declined = False
    if cur_cash is not None and pri_cash is not None and pri_cash != 0:
        pct = (cur_cash - pri_cash) / abs(pri_cash)
        result_text = f"{pct:+.1%}  ({pri_cash:,.1f}M → {cur_cash:,.1f}M)"
        threshold = "Warn if < -20%"
        is_bad = pct < -0.20
        cash_declined = is_bad
        status = "⚠️ Cash Drop > 20%!" if is_bad else "✅ Cash Stable"
        bad = _indicator_row(ws, row, "Cash QoQ Change (手元資金 前四半期比)", result_text, threshold, status, is_bad)
        if bad:
            alert_indicators.append("cash_drop")
    else:
        _indicator_row(ws, row, "Cash QoQ Change (手元資金 前四半期比)", "N/A", "Warn if < -20%", "— N/A", False)
    row += 1

    # 2. Current Liabilities increase > 20%
    cl_surge = False
    if cur_cl2 is not None and pri_cl2 is not None and pri_cl2 != 0:
        pct = (cur_cl2 - pri_cl2) / abs(pri_cl2)
        result_text = f"{pct:+.1%}  ({pri_cl2:,.1f}M → {cur_cl2:,.1f}M)"
        threshold = "Warn if > +20%"
        is_bad = pct > 0.20
        cl_surge = is_bad
        status = "⚠️ CL Surge > 20%!" if is_bad else "✅ CL Stable"
        bad = _indicator_row(ws, row, "Current Liabilities QoQ Change (流動負債 前四半期比)", result_text, threshold, status, is_bad)
        if bad:
            alert_indicators.append("cl_surge")
    else:
        _indicator_row(ws, row, "Current Liabilities QoQ Change (流動負債 前四半期比)", "N/A", "Warn if > +20%", "— N/A", False)
    row += 1

    # 3. Equity erosion > 20%
    eq_erosion = False
    if cur_eq2 is not None and pri_eq2 is not None and pri_eq2 != 0:
        pct = (cur_eq2 - pri_eq2) / abs(pri_eq2)
        result_text = f"{pct:+.1%}  ({pri_eq2:,.1f}M → {cur_eq2:,.1f}M)"
        threshold = "Warn if < -20%"
        is_bad = pct < -0.20
        eq_erosion = is_bad
        status = "⚠️ Equity Erosion > 20%!" if is_bad else "✅ Equity Stable"
        bad = _indicator_row(ws, row, "Stockholders' Equity QoQ Change (株主資本 前四半期比)", result_text, threshold, status, is_bad)
        if bad:
            alert_indicators.append("equity_erosion")
    else:
        _indicator_row(ws, row, "Stockholders' Equity QoQ Change (株主資本 前四半期比)", "N/A", "Warn if < -20%", "— N/A", False)
    row += 1

    row += 1

    # ── Overall Risk Verdict ──
    ws.merge_cells(f"A{row}:G{row}")
    vl = ws.cell(row=row, column=1, value="総合リスク判定 / Overall Risk Verdict")
    vl.fill = SUBHEADER_FILL
    vl.font = WHITE_BOLD
    vl.alignment = CENTER
    ws.row_dimensions[row].height = 22
    row += 1

    ws.merge_cells(f"A{row}:G{row}")
    verdict_cell = ws.cell(row=row, column=1)
    ws.row_dimensions[row].height = 40

    # Determine verdict
    red_flags = len(alert_indicators)
    cash_risk = cash_declined and (cl_surge or eq_erosion)

    if cash_risk:
        verdict = "【 ⚠️ 黒字倒産・資金繰り悪化の予兆あり 】  Cash shrinking while liabilities surge or equity is eroding. Immediate review recommended."
        verdict_cell.fill = RED_FILL
        verdict_cell.font = Font(name="Calibri", bold=True, color="C00000", size=13)
    elif red_flags >= 2:
        verdict = "【 ⚡ 要注意：複数の財務悪化シグナルを検知 】  Multiple deterioration signals detected. Monitor closely."
        verdict_cell.fill = YELLOW_FILL
        verdict_cell.font = Font(name="Calibri", bold=True, color="7F6000", size=13)
    elif red_flags == 1:
        verdict = "【 ⚡ 軽微なリスクシグナル 】  One deterioration signal detected. Continue monitoring."
        verdict_cell.fill = YELLOW_FILL
        verdict_cell.font = Font(name="Calibri", bold=True, color="7F6000", size=13)
    else:
        if all(v is None for v in [cur_cash, cur_cl2, cur_eq2]):
            verdict = "【 — データ不足：判定不可 】  Insufficient XBRL data to make a determination."
            verdict_cell.fill = LIGHT_GREY_FILL
            verdict_cell.font = Font(name="Calibri", bold=True, color="595959", size=13)
        else:
            verdict = "【 ✅ 現時点で重大なリスクシグナルなし 】  No major risk signals detected in the available data."
            verdict_cell.fill = GREEN_FILL
            verdict_cell.font = Font(name="Calibri", bold=True, color="1E6B2E", size=13)

    verdict_cell.value = verdict
    verdict_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    row += 2

    # ── Disclaimer ──
    ws.merge_cells(f"A{row}:G{row}")
    disc = ws.cell(row=row, column=1,
                   value="※ 免責事項：本ツールはSEC EDGAR公開データに基づく自動分析です。投資判断は必ず一次情報（10-Q/10-K）および専門家の意見をご確認ください。")
    disc.font = SMALL_ITALIC
    disc.fill = LIGHT_GREY_FILL
    disc.alignment = LEFT
    ws.row_dimensions[row].height = 16


# ---------------------------------------------------------------------------
# Demo / Offline Test Data
# ---------------------------------------------------------------------------

def _make_demo_record(end: str, val: float, start: str = "") -> dict:
    r = {"end": end, "val": val, "form": "10-Q", "accn": "0000000000-00-000000"}
    if start:
        r["start"] = start
    return r


def build_demo_data() -> tuple[dict, dict]:
    """Return synthetic PL and BS dicts for offline Excel-generation testing."""
    pl = {
        "Revenues": {
            "current": (2_150_000_000, "2024-09-30", "Revenues"),
            "prior":   (2_450_000_000, "2023-09-30", "Revenues"),
        },
        "OperatingExpenses": {
            "current": (2_050_000_000, "2024-09-30", "CostOfGoodsAndServicesSold"),
            "prior":   (2_200_000_000, "2023-09-30", "CostOfGoodsAndServicesSold"),
        },
        "OperatingIncomeLoss": {
            "current": (100_000_000,  "2024-09-30", "OperatingIncomeLoss"),
            "prior":   (250_000_000,  "2023-09-30", "OperatingIncomeLoss"),
        },
        "InterestExpense": {
            "current": (-45_000_000,  "2024-09-30", "NonoperatingIncomeExpense"),
            "prior":   (-40_000_000,  "2023-09-30", "NonoperatingIncomeExpense"),
        },
        "IncomeLossBeforeTax": {
            "current": (55_000_000,   "2024-09-30", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest"),
            "prior":   (210_000_000,  "2023-09-30", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest"),
        },
        "NetIncomeLoss": {
            "current": (35_000_000,   "2024-09-30", "NetIncomeLoss"),
            "prior":   (160_000_000,  "2023-09-30", "NetIncomeLoss"),
        },
    }
    bs = {
        "Cash": {
            "current": (180_000_000, "2024-09-30", "CashAndCashEquivalentsAtCarryingValue"),
            "prior":   (320_000_000, "2024-06-30", "CashAndCashEquivalentsAtCarryingValue"),
        },
        "CurrentLiabilities": {
            "current": (950_000_000, "2024-09-30", "LiabilitiesCurrent"),
            "prior":   (750_000_000, "2024-06-30", "LiabilitiesCurrent"),
        },
        "CurrentAssets": {
            "current": (820_000_000, "2024-09-30", "AssetsCurrent"),
            "prior":   (900_000_000, "2024-06-30", "AssetsCurrent"),
        },
        "LongTermLiabilities": {
            "current": (1_200_000_000, "2024-09-30", "LiabilitiesNoncurrent"),
            "prior":   (1_180_000_000, "2024-06-30", "LiabilitiesNoncurrent"),
        },
        "NonCurrentAssets": {
            "current": (1_850_000_000, "2024-09-30", "AssetsNoncurrent"),
            "prior":   (1_900_000_000, "2024-06-30", "AssetsNoncurrent"),
        },
        "StockholdersEquity": {
            "current": (420_000_000, "2024-09-30", "StockholdersEquity"),
            "prior":   (570_000_000, "2024-06-30", "StockholdersEquity"),
        },
    }
    return pl, bs


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    args = sys.argv[1:]
    demo_mode = "--demo" in args
    args = [a for a in args if a != "--demo"]
    ticker = args[0].upper() if args else "PARR"

    print(f"\n{'='*60}")
    print(f"  SEC EDGAR Financial Analyzer  |  Ticker: {ticker}")
    if demo_mode:
        print("  MODE: OFFLINE DEMO (synthetic data)")
    print(f"{'='*60}\n")

    if demo_mode:
        cik = KNOWN_CIKS.get(ticker, "0000000000")
        company_name = f"{ticker} Inc. [DEMO DATA]"
        pl, bs = build_demo_data()
        facts = {}
    else:
        # 1. Resolve CIK
        cik = resolve_cik(ticker)
        if cik is None:
            print("Cannot proceed without a valid CIK. Exiting.")
            sys.exit(1)

        # 2. Fetch XBRL facts
        facts = fetch_facts(cik)
        if not facts:
            print("Failed to fetch EDGAR company facts. Exiting.")
            sys.exit(1)

        company_name = facts.get("entityName", ticker)
        print(f"\nCompany: {company_name} (CIK: {cik})\n")

        # 3. Extract financials
        print("Extracting P&L data …")
        pl = extract_pl(facts)

        print("Extracting Balance Sheet data …")
        bs = extract_bs(facts)

    # Brief summary
    print("\n--- Data Availability ---")
    for metric, data in pl.items():
        cur = "✓" if data.get("current") else "✗"
        pri = "✓" if data.get("prior") else "✗"
        tag = data.get("current", (None, None, "?"))[2] if data.get("current") else "—"
        print(f"  PL/{metric:<30} current={cur} prior={pri}  tag={tag}")
    for metric, data in bs.items():
        cur = "✓" if data.get("current") else "✗"
        pri = "✓" if data.get("prior") else "✗"
        tag = data.get("current", (None, None, "?"))[2] if data.get("current") else "—"
        print(f"  BS/{metric:<30} current={cur} prior={pri}  tag={tag}")

    # 4. Build Excel workbook
    print("\nBuilding Excel workbook …")
    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    bs_rows = build_financial_data_sheet(wb, company_name, ticker, pl, bs)
    build_dashboard_sheet(wb, company_name, ticker, cik, pl, bs, bs_rows, facts)

    wb.move_sheet("Dashboard", offset=-len(wb.sheetnames) + 1)

    # 5. Save
    suffix = "_DEMO" if demo_mode else ""
    filename = f"{ticker}_financial_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}{suffix}.xlsx"
    wb.save(filename)
    print(f"\n{'='*60}")
    print(f"  ✅ Excel saved: {filename}")
    print(f"{'='*60}\n")
    print("Tabs:")
    print("  1. Dashboard      — Financial Health & Risk Verdict")
    print("  2. Financial Data — PL (YoY) and BS (QoQ) with Excel formulas\n")


if __name__ == "__main__":
    main()
