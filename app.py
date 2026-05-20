"""
SEC EDGAR Financial Analyzer — Streamlit Web App
Run: streamlit run app.py
Dependencies: pip install streamlit pandas requests openpyxl
"""

import io
import time
import requests
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────

USER_AGENT = "FinancialAnalyzer/1.0 (financial-analyzer@example.com)"
EDGAR_TICKER_API = "https://www.sec.gov/files/company_tickers.json"
EDGAR_FACTS_API  = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"

KNOWN_CIKS = {
    "PARR": "0001378590", "XOM":  "0000034088", "CVX":  "0000093410",
    "TSLA": "0001318605", "AAPL": "0000320193", "MSFT": "0000789019",
    "AMZN": "0001018724", "GOOGL":"0001652044", "META": "0001326801",
    "NVDA": "0001045810", "JPM":  "0000019617", "BAC":  "0000070858",
    "WMT":  "0000104169", "PFE":  "0000078003", "JNJ":  "0000200406",
    "MRK":  "0000310158", "PSX":  "0001534992", "VLO":  "0001035002",
    "MPC":  "0001510295", "HFC":  "0000048039", "DK":   "0000049600",
}

PL_TAGS = {
    "Revenues": [
        "Revenues", "SalesRevenueNet",
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        "SalesRevenueGoodsNet",
        "RevenueFromContractWithCustomerIncludingAssessedTax",
        "SalesAndRevenuesNet",
    ],
    "OperatingExpenses": [
        "CostOfGoodsAndServicesSold", "CostOfRevenue",
        "OperatingExpenses", "CostsAndExpenses", "CostOfGoodsSold",
    ],
    "OperatingIncomeLoss": [
        "OperatingIncomeLoss",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
    ],
    "InterestExpense": [
        "NonoperatingIncomeExpense", "InterestExpense",
        "InterestIncomeExpenseNet", "InterestAndDebtExpense",
        "OtherNonoperatingIncomeExpense",
    ],
    "IncomeLossBeforeTax": [
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesDomestic",
    ],
    "NetIncomeLoss": [
        "NetIncomeLoss", "ProfitLoss",
        "NetIncomeLossAvailableToCommonStockholdersBasic",
        "IncomeLossFromContinuingOperations",
    ],
}

BS_TAGS = {
    "Cash": [
        "CashAndCashEquivalentsAtCarryingValue",
        "CashCashEquivalentsAndShortTermInvestments",
        "CashAndShortTermInvestments", "Cash",
    ],
    "CurrentLiabilities": ["LiabilitiesCurrent", "LiabilitiesCurrentAndNoncurrent"],
    "CurrentAssets":      ["AssetsCurrent", "AssetsCurrentAndNoncurrent"],
    "LongTermLiabilities": [
        "LiabilitiesNoncurrent", "LongTermDebtNoncurrent",
        "LongTermDebt", "LongTermDebtAndCapitalLeaseObligations",
    ],
    "NonCurrentAssets": [
        "AssetsNoncurrent", "PropertyPlantAndEquipmentNet",
        "PropertyPlantAndEquipmentAndIntangibleAssetsNet",
    ],
    "StockholdersEquity": [
        "StockholdersEquity", "StockholdersEquityAttributableToParent",
        "PartnersCapital", "MembersEquity",
    ],
}

PL_LABELS = {
    "Revenues":            "売上高 / Revenues",
    "OperatingExpenses":   "営業費用 / Operating Expenses",
    "OperatingIncomeLoss": "営業利益 / Operating Income",
    "InterestExpense":     "金融損益 / Non-Op. Income/Exp.",
    "IncomeLossBeforeTax": "税引前利益 / Income Before Tax",
    "NetIncomeLoss":       "純利益 / Net Income",
}
BS_LABELS = {
    "Cash":               "手元資金 / Cash & Equivalents",
    "CurrentLiabilities": "流動負債 / Current Liabilities",
    "CurrentAssets":      "流動資産 / Current Assets",
    "LongTermLiabilities":"長期負債 / LT Liabilities",
    "NonCurrentAssets":   "固定資産 / Non-Current Assets",
    "StockholdersEquity": "株主資本 / Stockholders' Equity",
}

# ─────────────────────────────────────────────────────────────────────────────
# Data Fetching & Extraction  (shared with financial_analyzer.py logic)
# ─────────────────────────────────────────────────────────────────────────────

def _headers():
    return {"User-Agent": USER_AGENT, "Accept": "application/json"}

def _get(url: str, retries: int = 4, backoff: float = 2.0):
    for attempt in range(retries):
        try:
            r = requests.get(url, headers=_headers(), timeout=30)
            if r.status_code == 200:
                return r
            if r.status_code == 429:
                time.sleep(backoff * (2 ** attempt))
                continue
            return None
        except requests.RequestException:
            time.sleep(backoff * (2 ** attempt))
    return None

def resolve_cik(ticker: str) -> tuple[str | None, str]:
    t = ticker.upper()
    r = _get(EDGAR_TICKER_API)
    if r is not None:
        for _, entry in r.json().items():
            if entry.get("ticker", "").upper() == t:
                return str(entry["cik_str"]).zfill(10), entry.get("title", t)
    if t in KNOWN_CIKS:
        return KNOWN_CIKS[t], t
    return None, ""

def fetch_facts(cik: str) -> dict:
    r = _get(EDGAR_FACTS_API.format(cik=cik))
    return r.json() if r else {}

def _units(facts: dict, tag: str) -> list:
    try:
        gaap = facts.get("facts", {}).get("us-gaap", {})
        u = gaap.get(tag, {}).get("units", {})
        return u.get("USD", u.get("pure", []))
    except Exception:
        return []

def _best_tag(facts: dict, candidates: list) -> tuple[str, list]:
    for tag in candidates:
        recs = _units(facts, tag)
        if recs:
            return tag, recs
    return "", []

def _filter_quarterly(records: list) -> list:
    out = []
    for r in records:
        s, e = r.get("start", ""), r.get("end", "")
        if s and e:
            try:
                d = (datetime.strptime(e, "%Y-%m-%d") - datetime.strptime(s, "%Y-%m-%d")).days
                if 80 <= d <= 100:
                    out.append(r)
            except ValueError:
                pass
    return out

def _dedup_latest(records: list, n: int) -> list:
    seen = {}
    for r in records:
        e = r.get("end", "")
        if e not in seen or r.get("form", "") in ("10-Q", "10-K"):
            seen[e] = r
    return sorted(seen.values(), key=lambda x: x.get("end", ""), reverse=True)[:n]

def _filter_instant(records: list) -> list:
    return [r for r in records if not r.get("start")]

def extract_pl(facts: dict) -> dict:
    result = {}
    for metric, candidates in PL_TAGS.items():
        tag, recs = _best_tag(facts, candidates)
        quarterly = _dedup_latest(_filter_quarterly(recs), 8)
        current = prior = None
        if quarterly:
            cr = quarterly[0]
            current = (cr.get("val"), cr["end"], tag)
            target = datetime.strptime(cr["end"], "%Y-%m-%d") - timedelta(days=365)
            best, best_d = None, 999
            for r in quarterly[1:]:
                d = abs((datetime.strptime(r["end"], "%Y-%m-%d") - target).days)
                if d < 46 and d < best_d:
                    best, best_d = r, d
            if best:
                prior = (best.get("val"), best["end"], tag)
        result[metric] = {"current": current, "prior": prior}
    return result

def extract_bs(facts: dict) -> dict:
    result = {}
    for metric, candidates in BS_TAGS.items():
        tag, recs = _best_tag(facts, candidates)
        latest = _dedup_latest(_filter_instant(recs), 4)
        current = (latest[0].get("val"), latest[0]["end"], tag) if latest else None
        prior   = (latest[1].get("val"), latest[1]["end"], tag) if len(latest) > 1 else None
        result[metric] = {"current": current, "prior": prior}
    return result

def _m(val):
    """Convert raw value to USD millions float, or None."""
    try:
        return float(val) / 1_000_000 if val is not None else None
    except (TypeError, ValueError):
        return None

# ─────────────────────────────────────────────────────────────────────────────
# Build DataFrames for Streamlit display
# ─────────────────────────────────────────────────────────────────────────────

def build_pl_df(pl: dict) -> pd.DataFrame:
    rows = []
    for key, label in PL_LABELS.items():
        data = pl.get(key, {})
        cur  = _m(data["current"][0]) if data.get("current") else None
        pri  = _m(data["prior"][0])   if data.get("prior")   else None
        tag  = data["current"][2] if data.get("current") else "—"
        cdp  = data["current"][1] if data.get("current") else "—"
        pdp  = data["prior"][1]   if data.get("prior")   else "—"
        delta = cur - pri if cur is not None and pri is not None else None
        pct   = delta / abs(pri) if delta is not None and pri not in (None, 0) else None
        is_cost = key == "OperatingExpenses"
        alert = _alert(pct, is_cost)
        rows.append({
            "項目 / Metric":   label,
            f"当期 ({cdp})\n[USD M]": cur,
            f"前期 ({pdp})\n[USD M]": pri,
            "差額 [USD M]":   delta,
            "変化率 %":        pct,
            "アラート":        alert,
            "_pct":           pct,
            "_is_cost":       is_cost,
            "_tag":           tag,
        })
    return pd.DataFrame(rows)

def build_bs_df(bs: dict) -> pd.DataFrame:
    rows = []
    # Other Current Assets (derived)
    extra_labels = {**BS_LABELS}
    for key, label in extra_labels.items():
        data = bs.get(key, {})
        cur  = _m(data["current"][0]) if data.get("current") else None
        pri  = _m(data["prior"][0])   if data.get("prior")   else None
        tag  = data["current"][2] if data.get("current") else "—"
        cdp  = data["current"][1] if data.get("current") else "—"
        pdp  = data["prior"][1]   if data.get("prior")   else "—"
        delta = cur - pri if cur is not None and pri is not None else None
        pct   = delta / abs(pri) if delta is not None and pri not in (None, 0) else None
        is_liab = key in ("CurrentLiabilities", "LongTermLiabilities")
        alert = _alert(pct, is_liab)
        rows.append({
            "項目 / Metric":   label,
            f"当四半期末 ({cdp})\n[USD M]": cur,
            f"前四半期末 ({pdp})\n[USD M]": pri,
            "差額 [USD M]":   delta,
            "変化率 %":        pct,
            "アラート":        alert,
            "_pct":           pct,
            "_is_cost":       is_liab,
            "_tag":           tag,
        })
    # Other Current Assets = CurrentAssets - Cash
    ca  = _m(bs.get("CurrentAssets",  {}).get("current", (None,))[0])
    ca_p= _m(bs.get("CurrentAssets",  {}).get("prior",   (None,))[0])
    ca_d= bs.get("CurrentAssets",  {}).get("current", (None, "—"))[1]
    ca_dp=bs.get("CurrentAssets",  {}).get("prior",   (None, "—"))[1]
    cash= _m(bs.get("Cash",           {}).get("current", (None,))[0])
    cash_p=_m(bs.get("Cash",          {}).get("prior",   (None,))[0])
    if ca is not None and cash is not None:
        oca  = ca - cash
        oca_p= (ca_p - cash_p) if (ca_p is not None and cash_p is not None) else None
        delta = oca - oca_p if oca_p is not None else None
        pct   = delta / abs(oca_p) if delta is not None and oca_p not in (None, 0) else None
    else:
        oca = oca_p = delta = pct = None
    rows.append({
        "項目 / Metric":   "その他流動資産 / Other Current Assets (=CurrentAssets−Cash)",
        f"当四半期末 ({ca_d})\n[USD M]": oca,
        f"前四半期末 ({ca_dp})\n[USD M]": oca_p,
        "差額 [USD M]":   delta,
        "変化率 %":        pct,
        "アラート":        _alert(pct, False),
        "_pct":           pct,
        "_is_cost":       False,
        "_tag":           "(計算値)",
    })
    return pd.DataFrame(rows)

def _alert(pct, is_cost_or_liab: bool) -> str:
    if pct is None:
        return "—"
    if is_cost_or_liab:
        return "⚠️ +20%↑ 急増" if pct > 0.20 else "✅ 正常"
    return ("⚠️ -20%↓ 急減" if pct < -0.20 else
            ("✅ +20%↑ 成長" if pct > 0.20 else "✅ 正常"))

def _style_df(df: pd.DataFrame):
    display_cols = [c for c in df.columns if not c.startswith("_")]
    display_df = df[display_cols].reset_index(drop=True)

    def highlight_row(row):
        idx     = row.name
        pct     = df.iloc[idx]["_pct"]
        is_cost = df.iloc[idx]["_is_cost"]
        if pct is None:
            bg = ""
        else:
            triggered = (is_cost and pct > 0.20) or (not is_cost and pct < -0.20)
            bg = "background-color: #FFD2D2;" if triggered else ""
        return pd.Series([bg] * len(display_cols), index=display_cols)

    def fmt_usd(v):
        return f"{v:,.1f}" if isinstance(v, (int, float)) else "N/A"

    def fmt_pct(v):
        return f"{v:+.1%}" if isinstance(v, (int, float)) else "N/A"

    fmt = {}
    for c in display_cols:
        if "USD M" in c or "差額" in c:
            fmt[c] = fmt_usd
    fmt["変化率 %"] = fmt_pct

    right_cols = [c for c in display_cols if c not in ("項目 / Metric", "アラート")]

    styled = (
        display_df
        .style
        .apply(highlight_row, axis=1)
        .format(fmt)
        .set_properties(**{"text-align": "right"},  subset=right_cols)
        .set_properties(**{"text-align": "left"},   subset=["項目 / Metric"])
        .set_properties(**{"text-align": "center"}, subset=["アラート"])
    )
    return styled

# ─────────────────────────────────────────────────────────────────────────────
# Risk Calculations
# ─────────────────────────────────────────────────────────────────────────────

def compute_ratios(bs: dict) -> dict:
    def cur(k): return _m(bs.get(k, {}).get("current", (None,))[0])
    def pri(k): return _m(bs.get(k, {}).get("prior",   (None,))[0])

    ca   = cur("CurrentAssets");      ca_p  = pri("CurrentAssets")
    cl   = cur("CurrentLiabilities"); cl_p  = pri("CurrentLiabilities")
    cash = cur("Cash");               cash_p= pri("Cash")
    nca  = cur("NonCurrentAssets")
    eq   = cur("StockholdersEquity"); eq_p  = pri("StockholdersEquity")

    total  = (ca or 0) + (nca or 0)
    total_p= (ca_p or 0) + (pri("NonCurrentAssets") or 0)

    return {
        "current_ratio":      ca / cl          if (ca and cl and cl != 0) else None,
        "current_ratio_p":    ca_p / cl_p      if (ca_p and cl_p and cl_p != 0) else None,
        "equity_ratio":       eq / total        if (eq is not None and total != 0) else None,
        "equity_ratio_p":     eq_p / total_p   if (eq_p is not None and total_p != 0) else None,
        "cash_qoq":    (cash - cash_p) / abs(cash_p) if (cash is not None and cash_p not in (None, 0)) else None,
        "cl_qoq":      (cl - cl_p)     / abs(cl_p)   if (cl   is not None and cl_p   not in (None, 0)) else None,
        "eq_qoq":      (eq - eq_p)     / abs(eq_p)   if (eq   is not None and eq_p   not in (None, 0)) else None,
        "cash_cur": cash, "cash_pri": cash_p,
        "cl_cur":   cl,   "cl_pri":   cl_p,
        "eq_cur":   eq,   "eq_pri":   eq_p,
    }

def risk_verdict(r: dict) -> tuple[str, str, str]:
    """Returns (verdict_text, verdict_color, emoji)."""
    cash_drop  = r["cash_qoq"] is not None and r["cash_qoq"] < -0.20
    cl_surge   = r["cl_qoq"]   is not None and r["cl_qoq"]   > 0.20
    eq_erosion = r["eq_qoq"]   is not None and r["eq_qoq"]   < -0.20

    if cash_drop and (cl_surge or eq_erosion):
        return (
            "⚠️ 黒字倒産・資金繰り悪化の予兆あり\nCash shrinking while liabilities surge or equity erodes.",
            "#FFD2D2", "🔴"
        )
    n = sum([cash_drop, cl_surge, eq_erosion])
    if n >= 2:
        return "⚡ 要注意：複数の財務悪化シグナルを検知\nMultiple deterioration signals.", "#FFFACD", "🟡"
    if n == 1:
        return "⚡ 軽微なリスクシグナルあり\nOne deterioration signal detected.", "#FFFACD", "🟡"
    if all(v is None for v in [r["cash_qoq"], r["cl_qoq"], r["eq_qoq"]]):
        return "— データ不足：判定不可\nInsufficient data.", "#F2F2F2", "⚪"
    return "✅ 現時点で重大なリスクシグナルなし\nNo major risk signals detected.", "#D2FFD2", "🟢"

# ─────────────────────────────────────────────────────────────────────────────
# Excel Generation (in-memory BytesIO)
# ─────────────────────────────────────────────────────────────────────────────

THIN   = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

RED_F    = PatternFill("solid", fgColor="FFD2D2")
GREEN_F  = PatternFill("solid", fgColor="D2FFD2")
YELLOW_F = PatternFill("solid", fgColor="FFFACD")
BLUE_F   = PatternFill("solid", fgColor="1F4E79")
SUB_F    = PatternFill("solid", fgColor="2E75B6")
SEC_F    = PatternFill("solid", fgColor="BDD7EE")
GREY_F   = PatternFill("solid", fgColor="F2F2F2")

W_BOLD  = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
D_BOLD  = Font(name="Calibri", bold=True, color="1F4E79", size=11)
NORM    = Font(name="Calibri", size=10)
ITALIC  = Font(name="Calibri", size=9, italic=True, color="595959")
C       = Alignment(horizontal="center", vertical="center", wrap_text=True)
L       = Alignment(horizontal="left",   vertical="center", wrap_text=True)
R       = Alignment(horizontal="right",  vertical="center")

FMT_USD   = '#,##0_);[Red](#,##0)'
FMT_PCT   = '0.0%;[Red]-0.0%'
FMT_RATIO = '0.00'


def _cw(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def _hrow(ws, row, texts, fill=None):
    for col, t in enumerate(texts, 1):
        c = ws.cell(row=row, column=col, value=t)
        c.fill = fill or BLUE_F; c.font = W_BOLD
        c.alignment = C; c.border = BORDER

def _sc(cell, fill=None, font=None, align=None, fmt=None, border=True):
    if fill:   cell.fill   = fill
    if font:   cell.font   = font
    if align:  cell.alignment = align
    if fmt:    cell.number_format = fmt
    if border: cell.border = BORDER

def _mval(val):
    try:
        return float(val) / 1_000_000 if val is not None else None
    except:
        return None


def build_excel(company_name: str, ticker: str, cik: str,
                pl: dict, bs: dict) -> bytes:
    wb = Workbook()
    del wb["Sheet"]

    # ── Financial Data sheet ───────────────────────────────────────────────
    ws = wb.create_sheet("Financial Data")
    ws.views.sheetView[0].showGridLines = True
    _cw(ws, [3, 38, 18, 18, 18, 13, 16])

    ws.merge_cells("A1:G1")
    t = ws["A1"]
    t.value = f"財務分析レポート | {company_name} ({ticker.upper()}) | Source: SEC EDGAR XBRL"
    t.fill = BLUE_F; t.font = Font(name="Calibri", bold=True, color="FFFFFF", size=13)
    t.alignment = C; ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:G2")
    s = ws["A2"]
    s.value = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  CIK: {cik}  |  Amounts in USD Millions (M)"
    s.fill = SUB_F; s.font = Font(name="Calibri", color="FFFFFF", size=9, italic=True)
    s.alignment = C; ws.row_dimensions[2].height = 14

    # PL section
    row = 4
    ws.merge_cells(f"A{row}:G{row}")
    c = ws.cell(row=row, column=1, value="📊  損益計算書（P&L） — 前年同期比（YoY）")
    c.fill = SEC_F; c.font = D_BOLD; c.alignment = L
    ws.row_dimensions[row].height = 22; row += 1

    pl_cd = pl_pd = "—"
    for d in pl.values():
        if d.get("current"): pl_cd = d["current"][1]; break
    for d in pl.values():
        if d.get("prior"):   pl_pd = d["prior"][1];   break

    _hrow(ws, row, ["", "項目 / Metric",
                    f"当期\n({pl_cd})\n[USD M]",
                    f"前期\n({pl_pd})\n[USD M]",
                    "差額 [USD M]", "変化率 %", "アラート"])
    ws.row_dimensions[row].height = 40; row += 1

    pl_order = [
        ("Revenues",            "売上高 / Revenues"),
        ("OperatingExpenses",   "営業費用 / Operating Expenses"),
        ("OperatingIncomeLoss", "営業利益 / Operating Income"),
        ("InterestExpense",     "金融損益 / Non-Op. Income/Expense"),
        ("IncomeLossBeforeTax", "税引前利益 / Income Before Tax"),
        ("NetIncomeLoss",       "純利益 / Net Income"),
    ]
    bs_rows_map = {}

    for key, label in pl_order:
        data = pl.get(key, {})
        cur = _mval(data["current"][0]) if data.get("current") else None
        pri = _mval(data["prior"][0])   if data.get("prior")   else None
        tag = data["current"][2]        if data.get("current") else ""
        is_cost = key == "OperatingExpenses"

        ws.cell(row=row, column=1).fill = GREY_F
        _sc(ws.cell(row=row, column=1), border=True)

        lc = ws.cell(row=row, column=2, value=label); _sc(lc, font=NORM, align=L)

        cc = ws.cell(row=row, column=3, value=cur)
        _sc(cc, align=R, fmt=FMT_USD)
        if cur is None: cc.value = "N/A"; cc.font = ITALIC

        dc = ws.cell(row=row, column=4, value=pri)
        _sc(dc, align=R, fmt=FMT_USD)
        if pri is None: dc.value = "N/A"; dc.font = ITALIC

        ec = ws.cell(row=row, column=5)
        if cur is not None and pri is not None:
            ec.value = f"=C{row}-D{row}"; ec.number_format = FMT_USD
        else:
            ec.value = "N/A"; ec.font = ITALIC
        _sc(ec, align=R)

        fc = ws.cell(row=row, column=6)
        if cur is not None and pri is not None and pri != 0:
            fc.value = f"=IF(D{row}=0,\"\",(C{row}-D{row})/D{row})"; fc.number_format = FMT_PCT
        else:
            fc.value = "N/A"; fc.font = ITALIC
        _sc(fc, align=C)

        gc = ws.cell(row=row, column=7)
        alert_hit = False
        if cur is not None and pri is not None and pri != 0:
            pct = (cur - pri) / abs(pri)
            if is_cost and pct > 0.20:
                gc.value = "⚠️ +20%↑ 急増"; alert_hit = True
            elif not is_cost and pct < -0.20:
                gc.value = "⚠️ -20%↓ 急減"; alert_hit = True
            else:
                gc.value = "✅ 正常"
        else:
            gc.value = "—"
        _sc(gc, align=C)

        if alert_hit:
            for col in range(2, 8):
                ws.cell(row=row, column=col).fill = RED_F

        ws.row_dimensions[row].height = 18
        row += 1

    # Non-GAAP note
    ws.cell(row=row, column=1).value = "★"
    ws.cell(row=row, column=1).fill = YELLOW_F
    _sc(ws.cell(row=row, column=1), align=C)
    lc = ws.cell(row=row, column=2,
                 value="在庫影響除き営業利益 / Operating Income ex-Inventory Adj. [Non-GAAP]")
    lc.font = Font(name="Calibri", size=10, italic=True, color="7F6000")
    lc.alignment = L; lc.border = BORDER
    nc = ws.cell(row=row, column=3,
                 value="※ PARR等エネルギー企業は10-Q MD&Aの「Inventory Valuation Adjustment」を確認し手動で調整してください。")
    nc.font = Font(name="Calibri", size=8, italic=True, color="7F6000")
    nc.alignment = L
    ws.merge_cells(f"C{row}:G{row}"); nc.border = BORDER
    ws.row_dimensions[row].height = 26; row += 2

    # BS section
    ws.merge_cells(f"A{row}:G{row}")
    c = ws.cell(row=row, column=1, value="🏦  貸借対照表（B/S） — 前四半期比（QoQ）")
    c.fill = SEC_F; c.font = D_BOLD; c.alignment = L
    ws.row_dimensions[row].height = 22; row += 1

    bs_cd = bs_pd = "—"
    for d in bs.values():
        if d.get("current"): bs_cd = d["current"][1]; break
    for d in bs.values():
        if d.get("prior"):   bs_pd = d["prior"][1];   break

    _hrow(ws, row, ["", "項目 / Metric",
                    f"当四半期末\n({bs_cd})\n[USD M]",
                    f"前四半期末\n({bs_pd})\n[USD M]",
                    "差額 [USD M]", "変化率 %", "アラート"])
    ws.row_dimensions[row].height = 40; row += 1

    bs_order = [
        ("Cash",               "手元資金 / Cash & Equivalents",    False),
        ("CurrentLiabilities", "流動負債 / Current Liabilities",   True),
        ("CurrentAssets",      "流動資産 / Current Assets",        False),
        ("LongTermLiabilities","長期負債 / LT Liabilities",        True),
        ("NonCurrentAssets",   "固定資産 / Non-Current Assets",    False),
        ("StockholdersEquity", "株主資本 / Stockholders' Equity",  False),
    ]

    for key, label, is_liab in bs_order:
        data = bs.get(key, {})
        cur = _mval(data["current"][0]) if data.get("current") else None
        pri = _mval(data["prior"][0])   if data.get("prior")   else None
        bs_rows_map[key] = row

        ws.cell(row=row, column=1).fill = GREY_F
        _sc(ws.cell(row=row, column=1))

        lc = ws.cell(row=row, column=2, value=label); _sc(lc, font=NORM, align=L)

        cc = ws.cell(row=row, column=3, value=cur); _sc(cc, align=R, fmt=FMT_USD)
        if cur is None: cc.value = "N/A"; cc.font = ITALIC

        dc = ws.cell(row=row, column=4, value=pri); _sc(dc, align=R, fmt=FMT_USD)
        if pri is None: dc.value = "N/A"; dc.font = ITALIC

        ec = ws.cell(row=row, column=5)
        if cur is not None and pri is not None:
            ec.value = f"=C{row}-D{row}"; ec.number_format = FMT_USD
        else:
            ec.value = "N/A"; ec.font = ITALIC
        _sc(ec, align=R)

        fc = ws.cell(row=row, column=6)
        if cur is not None and pri is not None and pri != 0:
            fc.value = f"=IF(D{row}=0,\"\",(C{row}-D{row})/D{row})"; fc.number_format = FMT_PCT
        else:
            fc.value = "N/A"; fc.font = ITALIC
        _sc(fc, align=C)

        gc = ws.cell(row=row, column=7)
        alert_hit = False
        if cur is not None and pri is not None and pri != 0:
            pct = (cur - pri) / abs(pri)
            if is_liab and pct > 0.20:
                gc.value = "⚠️ +20%↑ 急増"; alert_hit = True
            elif not is_liab and pct < -0.20:
                gc.value = "⚠️ -20%↓ 急減"; alert_hit = True
            else:
                gc.value = "✅ 正常"
        else:
            gc.value = "—"
        _sc(gc, align=C)

        if alert_hit:
            for col in range(2, 8):
                ws.cell(row=row, column=col).fill = RED_F

        ws.row_dimensions[row].height = 18; row += 1

    # Other Current Assets (formula-derived)
    cash_r = bs_rows_map.get("Cash")
    ca_r   = bs_rows_map.get("CurrentAssets")
    ws.cell(row=row, column=1).fill = GREY_F; _sc(ws.cell(row=row, column=1))
    lc = ws.cell(row=row, column=2, value="その他流動資産 / Other Current Assets (=CurrentAssets−Cash)")
    lc.font = Font(name="Calibri", size=10, italic=True); lc.alignment = L; lc.border = BORDER
    if cash_r and ca_r:
        for col, (c_ref, d_ref) in enumerate([(f"C{ca_r}-C{cash_r}", f"D{ca_r}-D{cash_r}"),], 3):
            oc = ws.cell(row=row, column=col, value=f"={c_ref}")
            oc.number_format = FMT_USD; _sc(oc, align=R)
            od = ws.cell(row=row, column=col+1, value=f"={d_ref}")
            od.number_format = FMT_USD; _sc(od, align=R)
            oe = ws.cell(row=row, column=5, value=f"=C{row}-D{row}")
            oe.number_format = FMT_USD; _sc(oe, align=R)
            of_ = ws.cell(row=row, column=6, value=f"=IF(D{row}=0,\"\",(C{row}-D{row})/D{row})")
            of_.number_format = FMT_PCT; _sc(of_, align=C)
    else:
        for col in range(3, 8):
            c = ws.cell(row=row, column=col, value="N/A"); c.font = ITALIC; c.border = BORDER
    ws.cell(row=row, column=7, value="(計算値)").border = BORDER
    ws.row_dimensions[row].height = 18; row += 2

    # Legend
    ws.merge_cells(f"A{row}:G{row}")
    leg = ws.cell(row=row, column=1,
                  value="[凡例] ⚠️=アラート(20%以上悪化) ✅=正常 N/A=データ未取得 ★=Non-GAAP(要手動確認) | 金額はUSD百万単位")
    leg.font = ITALIC; leg.fill = GREY_F; leg.alignment = L
    ws.row_dimensions[row].height = 14

    # ── Dashboard sheet ────────────────────────────────────────────────────
    wd = wb.create_sheet("Dashboard", 0)
    wd.views.sheetView[0].showGridLines = True
    _cw(wd, [3, 30, 22, 22, 16, 14, 3])

    wd.merge_cells("A1:G1")
    t = wd["A1"]
    t.value = f"財務ヘルスダッシュボード | {company_name} ({ticker.upper()})"
    t.fill = BLUE_F; t.font = Font(name="Calibri", bold=True, color="FFFFFF", size=15)
    t.alignment = C; wd.row_dimensions[1].height = 34

    wd.merge_cells("A2:G2")
    s = wd["A2"]
    s.value = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  CIK: {cik}  |  Source: SEC EDGAR XBRL"
    s.fill = SUB_F; s.font = Font(name="Calibri", color="FFFFFF", size=9, italic=True)
    s.alignment = C; wd.row_dimensions[2].height = 14

    # Key Ratios (cross-sheet formulas)
    dr = 4
    wd.merge_cells(f"A{dr}:G{dr}")
    c = wd.cell(row=dr, column=1, value="📐  主要財務比率（Financial Dataシートと数式連携）")
    c.fill = SEC_F; c.font = D_BOLD; c.alignment = L; wd.row_dimensions[dr].height = 22; dr += 1

    _hrow(wd, dr, ["", "指標 / Ratio", "当期", "前期", "変化", "判定", ""]); wd.row_dimensions[dr].height = 22; dr += 1

    fd = "'Financial Data'"
    ca_r2 = bs_rows_map.get("CurrentAssets")
    cl_r2 = bs_rows_map.get("CurrentLiabilities")
    eq_r2 = bs_rows_map.get("StockholdersEquity")
    nca_r2= bs_rows_map.get("NonCurrentAssets")

    def _ratio_row(wd, dr, label, cur_f, pri_f, fmt, status_f, status_fill):
        wd.cell(row=dr, column=1).border = BORDER
        lc = wd.cell(row=dr, column=2, value=label); _sc(lc, font=NORM, align=L)
        cc = wd.cell(row=dr, column=3, value=cur_f); cc.number_format = fmt; _sc(cc, align=R)
        dc = wd.cell(row=dr, column=4, value=pri_f); dc.number_format = fmt; _sc(dc, align=R)
        ec = wd.cell(row=dr, column=5, value=f"=C{dr}-D{dr}"); ec.number_format = fmt; _sc(ec, align=R)
        sc = wd.cell(row=dr, column=6, value=status_f)
        _sc(sc, fill=status_fill, align=C)
        wd.cell(row=dr, column=7).border = BORDER
        wd.row_dimensions[dr].height = 18

    ratios = compute_ratios(bs)

    # Current Ratio
    if ca_r2 and cl_r2:
        cur_f = f"={fd}!C{ca_r2}/{fd}!C{cl_r2}"
        pri_f = f"={fd}!D{ca_r2}/{fd}!D{cl_r2}"
        sf    = f'=IF(C{dr}<1,"⚠️ <1.0 危険",IF(C{dr}<1.5,"⚡ 注意","✅ 良好"))'
        cr_v  = ratios["current_ratio"]
        sfill = RED_F if (cr_v and cr_v < 1) else (YELLOW_F if (cr_v and cr_v < 1.5) else GREEN_F)
    else:
        cur_f = pri_f = sf = "N/A"; sfill = GREY_F
    _ratio_row(wd, dr, "流動比率 / Current Ratio", cur_f, pri_f, FMT_RATIO, sf, sfill); dr += 1

    # Equity Ratio
    if ca_r2 and nca_r2 and eq_r2:
        ta_c  = f"({fd}!C{ca_r2}+{fd}!C{nca_r2})"
        ta_p  = f"({fd}!D{ca_r2}+{fd}!D{nca_r2})"
        cur_f = f"={fd}!C{eq_r2}/{ta_c}"
        pri_f = f"={fd}!D{eq_r2}/{ta_p}"
        sf    = f'=IF(C{dr}<0.1,"⚠️ <10% 危険",IF(C{dr}<0.3,"⚡ <30% 低水準","✅ 良好"))'
        er_v  = ratios["equity_ratio"]
        sfill = RED_F if (er_v is not None and er_v < 0.1) else (YELLOW_F if (er_v is not None and er_v < 0.3) else GREEN_F)
    else:
        cur_f = pri_f = sf = "N/A"; sfill = GREY_F
    _ratio_row(wd, dr, "自己資本比率 / Equity Ratio", cur_f, pri_f, FMT_PCT, sf, sfill); dr += 2

    # Risk Alert Section
    wd.merge_cells(f"A{dr}:G{dr}")
    c = wd.cell(row=dr, column=1, value="🚨  倒産予兆・資金繰りリスク判定")
    c.fill = PatternFill("solid", fgColor="C00000")
    c.font = Font(name="Calibri", bold=True, color="FFFFFF", size=12)
    c.alignment = L; wd.row_dimensions[dr].height = 26; dr += 1

    _hrow(wd, dr, ["", "チェック項目", "実績値", "判定基準", "状態", "", ""]); wd.row_dimensions[dr].height = 20; dr += 1

    def _risk_row(wd, dr, label, result, threshold, is_bad):
        wd.cell(row=dr, column=1).border = BORDER
        lc = wd.cell(row=dr, column=2, value=label); _sc(lc, font=NORM, align=L)
        rc = wd.cell(row=dr, column=3, value=result); _sc(rc, font=NORM, align=C)
        tc = wd.cell(row=dr, column=4, value=threshold); _sc(tc, font=ITALIC, align=C)
        status = "⚠️ 要注意" if is_bad else "✅ 正常"
        sc = wd.cell(row=dr, column=5, value=status)
        _sc(sc, fill=(RED_F if is_bad else GREEN_F), align=C)
        for col in [6, 7]: wd.cell(row=dr, column=col).border = BORDER
        wd.row_dimensions[dr].height = 18

    r = ratios
    cash_bad = r["cash_qoq"] is not None and r["cash_qoq"] < -0.20
    cl_bad   = r["cl_qoq"]   is not None and r["cl_qoq"]   > 0.20
    eq_bad   = r["eq_qoq"]   is not None and r["eq_qoq"]   < -0.20

    def _fmt_qoq(pct, cur, pri):
        if pct is None: return "N/A"
        return f"{pct:+.1%}  ({pri:,.0f}M → {cur:,.0f}M)" if (cur is not None and pri is not None) else f"{pct:+.1%}"

    _risk_row(wd, dr, "手元資金 QoQ変化 (Cash)", _fmt_qoq(r["cash_qoq"], r["cash_cur"], r["cash_pri"]), "< -20% で警告", cash_bad); dr += 1
    _risk_row(wd, dr, "流動負債 QoQ変化 (Current Liabilities)", _fmt_qoq(r["cl_qoq"], r["cl_cur"], r["cl_pri"]), "> +20% で警告", cl_bad); dr += 1
    _risk_row(wd, dr, "株主資本 QoQ変化 (Stockholders' Equity)", _fmt_qoq(r["eq_qoq"], r["eq_cur"], r["eq_pri"]), "< -20% で警告", eq_bad); dr += 2

    # Verdict
    wd.merge_cells(f"A{dr}:G{dr}")
    vt = wd.cell(row=dr, column=1, value="総合リスク判定 / Overall Risk Verdict")
    vt.fill = SUB_F; vt.font = W_BOLD; vt.alignment = C; wd.row_dimensions[dr].height = 22; dr += 1

    verdict_txt, verdict_hex, _ = risk_verdict(r)
    wd.merge_cells(f"A{dr}:G{dr}")
    vc = wd.cell(row=dr, column=1, value=verdict_txt)
    vc.fill = PatternFill("solid", fgColor=verdict_hex.lstrip("#"))
    vc.font = Font(name="Calibri", bold=True, size=13,
                   color="C00000" if "FFD2D2" in verdict_hex else
                         ("7F6000" if "FACD" in verdict_hex else "1E6B2E"))
    vc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wd.row_dimensions[dr].height = 44; dr += 2

    wd.merge_cells(f"A{dr}:G{dr}")
    dis = wd.cell(row=dr, column=1,
                  value="※ 免責：本ツールはSEC EDGAR公開データに基づく自動分析です。投資判断は必ず一次情報（10-Q/10-K）と専門家意見でご確認ください。")
    dis.font = ITALIC; dis.fill = GREY_F; dis.alignment = L; wd.row_dimensions[dr].height = 14

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ─────────────────────────────────────────────────────────────────────────────
# Streamlit UI
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="SEC財務分析ツール",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
/* Global */
[data-testid="stAppViewContainer"] { background: #F7F9FC; }
h1 { color: #1F4E79; }
h2 { color: #2E75B6; font-size:1.15rem; margin-top:1.2rem; margin-bottom:0.3rem; }
/* Metric cards */
[data-testid="metric-container"] {
    background: #fff; border-radius: 10px;
    border: 1px solid #D0E4F4; padding: 10px 14px;
    box-shadow: 0 1px 4px rgba(0,0,0,.07);
}
/* Alert box */
.alert-box {
    padding: 14px 18px; border-radius: 10px;
    font-size: 1.05rem; font-weight: 600;
    margin-bottom: 10px;
}
.alert-red    { background:#FFD2D2; color:#8B0000; border-left:5px solid #C00000; }
.alert-yellow { background:#FFFACD; color:#7F6000; border-left:5px solid #FFC000; }
.alert-green  { background:#D2FFD2; color:#1E6B2E; border-left:5px solid #00B050; }
.alert-grey   { background:#F2F2F2; color:#595959; border-left:5px solid #BFBFBF; }
/* Download btn */
[data-testid="stDownloadButton"] button {
    background: #1F4E79 !important; color: #fff !important;
    font-weight: 700 !important; border-radius: 8px !important;
    padding: 10px 24px !important; font-size: 1rem !important;
    border: none !important;
}
[data-testid="stDownloadButton"] button:hover {
    background: #2E75B6 !important;
}
</style>
""", unsafe_allow_html=True)

# Header
st.markdown("# 📊 SEC EDGAR 財務分析ツール")
st.markdown("米国上場企業のティッカーを入力すると、最新の財務データ（10-Q/10-K）を自動取得して分析します。")
st.markdown("---")

# Input
col_in, col_btn, col_demo = st.columns([3, 1.2, 1.2])
with col_in:
    ticker_input = st.text_input(
        "ティッカーシンボル（例: PARR, XOM, TSLA）",
        value="PARR",
        max_chars=10,
        placeholder="PARR",
        label_visibility="visible",
    )
with col_btn:
    st.markdown("<br>", unsafe_allow_html=True)
    run_btn = st.button("🔍 分析実行", type="primary", use_container_width=True)
with col_demo:
    st.markdown("<br>", unsafe_allow_html=True)
    demo_btn = st.button("🧪 デモデータ", use_container_width=True,
                         help="ネットワーク不要のサンプルデータで動作確認")

st.markdown("---")

# ── Demo data builder ──────────────────────────────────────────────────────

def demo_pl_bs():
    pl = {
        "Revenues":            {"current": (2_150_000_000, "2024-09-30", "Revenues"),
                                "prior":   (2_450_000_000, "2023-09-30", "Revenues")},
        "OperatingExpenses":   {"current": (2_050_000_000, "2024-09-30", "CostOfGoodsAndServicesSold"),
                                "prior":   (2_200_000_000, "2023-09-30", "CostOfGoodsAndServicesSold")},
        "OperatingIncomeLoss": {"current": (100_000_000,  "2024-09-30", "OperatingIncomeLoss"),
                                "prior":   (250_000_000,  "2023-09-30", "OperatingIncomeLoss")},
        "InterestExpense":     {"current": (-45_000_000,  "2024-09-30", "NonoperatingIncomeExpense"),
                                "prior":   (-40_000_000,  "2023-09-30", "NonoperatingIncomeExpense")},
        "IncomeLossBeforeTax": {"current": (55_000_000,   "2024-09-30", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest"),
                                "prior":   (210_000_000,  "2023-09-30", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest")},
        "NetIncomeLoss":       {"current": (35_000_000,   "2024-09-30", "NetIncomeLoss"),
                                "prior":   (160_000_000,  "2023-09-30", "NetIncomeLoss")},
    }
    bs = {
        "Cash":               {"current": (180_000_000,   "2024-09-30", "CashAndCashEquivalentsAtCarryingValue"),
                               "prior":   (320_000_000,   "2024-06-30", "CashAndCashEquivalentsAtCarryingValue")},
        "CurrentLiabilities": {"current": (950_000_000,   "2024-09-30", "LiabilitiesCurrent"),
                               "prior":   (750_000_000,   "2024-06-30", "LiabilitiesCurrent")},
        "CurrentAssets":      {"current": (820_000_000,   "2024-09-30", "AssetsCurrent"),
                               "prior":   (900_000_000,   "2024-06-30", "AssetsCurrent")},
        "LongTermLiabilities":{"current": (1_200_000_000, "2024-09-30", "LiabilitiesNoncurrent"),
                               "prior":   (1_180_000_000, "2024-06-30", "LiabilitiesNoncurrent")},
        "NonCurrentAssets":   {"current": (1_850_000_000, "2024-09-30", "AssetsNoncurrent"),
                               "prior":   (1_900_000_000, "2024-06-30", "AssetsNoncurrent")},
        "StockholdersEquity": {"current": (420_000_000,   "2024-09-30", "StockholdersEquity"),
                               "prior":   (570_000_000,   "2024-06-30", "StockholdersEquity")},
    }
    return pl, bs

# ── Render results ──────────────────────────────────────────────────────────

def render(company_name: str, ticker: str, cik: str, pl: dict, bs: dict):
    ratios  = compute_ratios(bs)
    verdict, v_hex, v_emoji = risk_verdict(ratios)

    # Verdict banner
    cls = ("alert-red"    if "FFD2D2" in v_hex else
           "alert-yellow" if "FACD"   in v_hex else
           "alert-grey"   if "F2F2F" in v_hex else "alert-green")
    st.markdown(f'<div class="alert-box {cls}">{v_emoji} {verdict.replace(chr(10), "  |  ")}</div>',
                unsafe_allow_html=True)

    # Company + ratio KPIs
    st.markdown(f"### 🏢 {company_name}  `{ticker.upper()}`  —  CIK: `{cik}`")
    k1, k2, k3, k4, k5, k6 = st.columns(6)

    def _pct_str(v):  return f"{v:.1%}"  if v is not None else "N/A"
    def _ratio_str(v): return f"{v:.2f}" if v is not None else "N/A"
    def _m_str(v):    return f"${v:,.0f}M" if v is not None else "N/A"

    cr  = ratios["current_ratio"]
    er  = ratios["equity_ratio"]
    k1.metric("流動比率", _ratio_str(cr),
              delta=f"{ratios['current_ratio']-ratios['current_ratio_p']:.2f}" if (cr and ratios["current_ratio_p"]) else None)
    k2.metric("自己資本比率", _pct_str(er),
              delta=f"{(ratios['equity_ratio']-ratios['equity_ratio_p']):.1%}" if (er and ratios["equity_ratio_p"]) else None)
    k3.metric("手元資金 QoQ", _pct_str(ratios["cash_qoq"]),
              delta=None, delta_color="inverse")
    k4.metric("流動負債 QoQ", _pct_str(ratios["cl_qoq"]))
    k5.metric("株主資本 QoQ", _pct_str(ratios["eq_qoq"]),
              delta=None, delta_color="inverse")
    k6.metric("現金残高", _m_str(ratios["cash_cur"]))

    st.markdown("---")

    # ── PL Table ──
    st.markdown("## 📊 損益計算書（P&L） — 前年同期比（YoY）")
    st.caption("差額・変化率の列はExcelダウンロード版では数式として埋め込まれます。")
    pl_df = build_pl_df(pl)
    st.dataframe(
        _style_df(pl_df),
        use_container_width=True,
        height=280,
    )

    # Non-GAAP note
    st.info("★ **在庫影響除き営業利益（Non-GAAP）** — PARR等エネルギー企業では10-Q MD&Aの「Inventory Valuation Adjustment」を確認し手動で調整してください。", icon="ℹ️")

    # ── BS Table ──
    st.markdown("## 🏦 貸借対照表（B/S） — 前四半期比（QoQ）")
    bs_df = build_bs_df(bs)
    st.dataframe(
        _style_df(bs_df),
        use_container_width=True,
        height=320,
    )

    st.markdown("---")

    # ── Risk detail ──
    st.markdown("## 🚨 リスク指標詳細")
    rc1, rc2, rc3 = st.columns(3)
    def _risk_card(col, label, pct, cur_val, pri_val, warn_cond):
        is_bad = warn_cond(pct) if pct is not None else False
        bg = "#FFD2D2" if is_bad else "#D2FFD2"
        icon = "⚠️" if is_bad else "✅"
        txt = f"{pct:+.1%}" if pct is not None else "N/A"
        detail = f"{pri_val:,.0f}M → {cur_val:,.0f}M" if (cur_val is not None and pri_val is not None) else ""
        col.markdown(f"""
        <div style="background:{bg};border-radius:10px;padding:14px 16px;border:1px solid #ccc;">
        <div style="font-weight:700;font-size:.95rem;">{icon} {label}</div>
        <div style="font-size:1.6rem;font-weight:800;margin:4px 0;">{txt}</div>
        <div style="font-size:.8rem;color:#555;">{detail}</div>
        </div>""", unsafe_allow_html=True)

    _risk_card(rc1, "手元資金 QoQ", ratios["cash_qoq"], ratios["cash_cur"], ratios["cash_pri"],
               lambda p: p < -0.20)
    _risk_card(rc2, "流動負債 QoQ", ratios["cl_qoq"],   ratios["cl_cur"],   ratios["cl_pri"],
               lambda p: p > 0.20)
    _risk_card(rc3, "株主資本 QoQ", ratios["eq_qoq"],   ratios["eq_cur"],   ratios["eq_pri"],
               lambda p: p < -0.20)

    st.markdown("---")

    # ── Excel download ──
    st.markdown("## 📥 Excelレポートのダウンロード")
    st.markdown("数式連携・カラーハイライト付きの **.xlsx** ファイルをダウンロードできます。")
    with st.spinner("Excelファイルを生成中…"):
        excel_bytes = build_excel(company_name, ticker, cik, pl, bs)
    filename = f"{ticker.upper()}_financial_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    st.download_button(
        label="📥 このレポートをExcelでダウンロード",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption(f"ファイル名: `{filename}`  |  シート構成: Dashboard / Financial Data")


# ── Main execution logic ───────────────────────────────────────────────────

if demo_btn:
    pl, bs = demo_pl_bs()
    render("PARR Inc. [DEMO DATA]", "PARR", KNOWN_CIKS["PARR"], pl, bs)

elif run_btn:
    ticker = ticker_input.strip().upper()
    if not ticker:
        st.warning("ティッカーを入力してください。")
        st.stop()

    with st.spinner(f"SEC EDGAR から {ticker} のデータを取得中…"):
        cik, company_name = resolve_cik(ticker)
        if cik is None:
            st.error(f"ティッカー `{ticker}` のCIKが見つかりませんでした。\n"
                     "スペルを確認するか、`--demo` ボタンで動作確認してください。")
            st.stop()

        facts = fetch_facts(cik)
        if not facts:
            st.error("SEC EDGAR からデータを取得できませんでした。\n"
                     "ネットワーク環境を確認するか、しばらく時間をおいて再試行してください。")
            st.stop()

        if not company_name or company_name == ticker:
            company_name = facts.get("entityName", ticker)

        pl = extract_pl(facts)
        bs = extract_bs(facts)

    render(company_name, ticker, cik, pl, bs)

else:
    st.markdown("""
    <div style="text-align:center;padding:60px 20px;color:#8B9BB4;">
    <div style="font-size:4rem;">📊</div>
    <div style="font-size:1.2rem;font-weight:600;margin:12px 0;">ティッカーを入力して「分析実行」を押してください</div>
    <div style="font-size:.9rem;">対応例: PARR · XOM · CVX · TSLA · AAPL · MSFT · AMZN など</div>
    <div style="font-size:.85rem;margin-top:8px;color:#AAB8C8;">
    ネットワーク不要のデモは「🧪 デモデータ」ボタンで確認できます
    </div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")
st.caption("データソース: SEC EDGAR XBRL API (data.sec.gov) | 本ツールは情報提供目的のみです。投資判断には必ず一次情報をご確認ください。")
