"""
SEC EDGAR Financial Analyzer — Streamlit Web App (v2)
Run: streamlit run app.py
Dependencies: pip install streamlit pandas requests openpyxl
"""

import io
import re
import time
import requests
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
from html.parser import HTMLParser
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import yfinance as yf  # noqa: F401  (imported lazily in helpers)
except ImportError:
    yf = None

# ─────────────────────────────────────────────────────────────────────────────
# yfinance Helpers (gracefully degrade when Yahoo Finance is blocked)
# ─────────────────────────────────────────────────────────────────────────────

def _search_tickers(query: str) -> list:
    """Search for tickers using yfinance. Returns list of {ticker, name} dicts."""
    try:
        import yfinance as yf
        results = yf.Search(query, max_results=8)
        out = []
        for q in results.quotes:
            sym = q.get("symbol", "")
            name = q.get("shortname") or q.get("longname") or sym
            if sym and "." not in sym:  # skip non-US tickers (e.g. 7203.T)
                out.append({"ticker": sym, "name": name})
        return out
    except Exception:
        return []


def _get_chart_data(ticker: str):
    """Fetch 5Y daily close for ticker and optionally S&P500 (^GSPC).
    Returns (stock_df, sp500_df_or_None). stock_df is None only on total failure.
    """
    import pandas as pd

    def _fetch(sym: str):
        """Return a tz-naive daily Close DataFrame, or None on any failure."""
        try:
            import yfinance as yf
            df = yf.Ticker(sym).history(period="5y", interval="1d")
            if df.empty:
                return None
            # Case-insensitive column search (yfinance may vary by version)
            col = next((c for c in df.columns if c.lower() == "close"), None)
            if col is None:
                return None
            s = df[[col]].copy()
            s.columns = ["Close"]
            # Normalise to tz-naive date index so both series align safely
            if getattr(s.index, "tzinfo", None) is not None:
                s.index = s.index.tz_convert("UTC").tz_localize(None)
            s.index = pd.to_datetime(s.index.date)
            return s
        except Exception:
            return None

    stock = _fetch(ticker)
    if stock is None or stock.empty:
        return None, None

    sp500 = _fetch("^GSPC")
    if sp500 is not None and not sp500.empty:
        try:
            start = max(stock.index[0], sp500.index[0])
            stock = stock[stock.index >= start]
            sp500 = sp500[sp500.index >= start]
        except Exception:
            sp500 = None  # alignment failed — show stock only

    return stock, sp500


def _compute_valuation(ticker: str, facts: dict, bs: dict) -> dict:
    """Compute PER/PBR with 3-level fallback.

    1. yfinance info direct fields (trailingPE, priceToBook, EPS, bookValue)
    2. yfinance EPS/bookValue + price from history()
    3. EDGAR fallback:
       - PER: TTM diluted EPS from last 4 quarterly records (matches Yahoo Finance
              methodology); falls back to latest 10-K annual EPS
       - PBR: price × shares / equity

    Price is always fetched from history() first (same path as stock chart).
    A single Ticker object is used for all yfinance calls to share session state.
    """
    import math
    from datetime import datetime as _dt

    def _f(v):
        try:
            f = float(v)
            return None if math.isnan(f) else f
        except (TypeError, ValueError):
            return None

    def _fpos(v):
        f = _f(v)
        return f if f and f > 0 else None

    def _ttm_eps(eps_records: list) -> float | None:
        """Sum 4 most recent individual quarterly EPS for TTM.
        Uses a wider window (60-120 days) to handle fiscal quarters that
        don't fall in the standard 80-100 day range (e.g. Tesla Q4).
        """
        quarterly = []
        for r in eps_records:
            s, e, v = r.get("start"), r.get("end"), r.get("val")
            if not (s and e and v is not None):
                continue
            try:
                d = (_dt.strptime(e, "%Y-%m-%d") - _dt.strptime(s, "%Y-%m-%d")).days
                if 60 <= d <= 120:          # wider window: quarter only (not semi-annual)
                    quarterly.append((e, _f(v)))
            except (ValueError, TypeError):
                pass
        seen, uniq = set(), []
        for e, v in sorted(quarterly, key=lambda x: x[0], reverse=True):
            if e not in seen and v is not None:
                seen.add(e)
                uniq.append(v)
                if len(uniq) == 4:
                    break
        return sum(uniq) if len(uniq) == 4 else None

    try:
        import yfinance as yf
    except ImportError:
        return {}

    # Single Ticker object — shares crumb/session across history() and info
    t = yf.Ticker(ticker)

    # Price from history() — most reliable yfinance endpoint
    price = None
    try:
        h = t.history(period="5d")
        if not h.empty:
            price = float(h["Close"].iloc[-1])
    except Exception:
        pass

    # info for fundamental fields
    info = {}
    try:
        info = t.info or {}
    except Exception:
        pass

    if not price:
        price = _fpos(info.get("currentPrice")) or _fpos(info.get("regularMarketPrice"))

    pe, pe_label = None, "PER"
    pb = None

    # ── PER ──────────────────────────────────────────────────────────────────
    # Priority 1: trailingPE direct (= Yahoo Finance's displayed PE value)
    # This is the most reliable match — avoids GAAP EPS distortions from
    # inventory adjustments, impairments, or one-time items (e.g. PSX, PARR).
    t_pe  = _fpos(info.get("trailingPE"))
    t_eps = _f(info.get("trailingEps"))
    f_pe  = _fpos(info.get("forwardPE"))
    f_eps = _f(info.get("forwardEps"))

    if t_pe:
        pe, pe_label = t_pe, "PER（実績）"
    elif price and t_eps and t_eps > 0:
        pe, pe_label = price / t_eps, "PER（実績）"

    # Priority 2: EDGAR TTM diluted EPS (fallback when yfinance returns no data)
    if not pe and price:
        for _eps_tag in ("EarningsPerShareDiluted", "EarningsPerShareBasic"):
            _eps_recs = facts.get("facts", {}).get("us-gaap", {}) \
                             .get(_eps_tag, {}).get("units", {}).get("USD/shares", [])
            if not _eps_recs:
                continue
            ttm = _ttm_eps(_eps_recs)
            if ttm and ttm > 0:
                pe, pe_label = price / ttm, "PER（実績）"
                break
            # Annual 10-K EPS as last resort
            _annual = sorted(
                [r for r in _eps_recs
                 if r.get("form") in ("10-K", "10-K/A") and r.get("val") is not None],
                key=lambda r: r.get("end", ""), reverse=True
            )
            if _annual:
                eps_val = _f(_annual[0].get("val"))
                if eps_val and eps_val > 0:
                    pe, pe_label = price / eps_val, "PER（実績）"
                    break

    # Forward PER only when no trailing data at all
    if not pe:
        if f_pe:
            pe, pe_label = f_pe, "PER（予想）"
        elif price and f_eps and f_eps > 0:
            pe, pe_label = price / f_eps, "PER（予想）"

    # ── PBR ──────────────────────────────────────────────────────────────────
    # Priority 1: yfinance priceToBook (= Yahoo Finance's own PBR value)
    pb = _fpos(info.get("priceToBook"))
    if not pb and price:
        bv = _f(info.get("bookValue"))
        if bv and bv > 0:
            pb = price / bv

    # Priority 2: EDGAR equity / shares (fallback)
    if not pb and price:
        shares = None
        for _ns in ("dei", "us-gaap"):
            for _tag in ("EntityCommonStockSharesOutstanding", "CommonStockSharesOutstanding"):
                _sh = facts.get("facts", {}).get(_ns, {}).get(_tag, {}) \
                           .get("units", {}).get("shares", [])
                if _sh:
                    _sh = sorted(_sh, key=lambda r: r.get("end", ""), reverse=True)
                    shares = _fpos(_sh[0].get("val"))
                    if shares:
                        break
            if shares:
                break
        if shares:
            _eq = bs.get("StockholdersEquity", {}).get("current")
            eq = _f(_eq[0]) if _eq and _eq[0] is not None else None
            if eq and eq > 0:
                pb = (price * shares) / eq

    return {"pe": pe, "pe_label": pe_label, "pb": pb}



# ─────────────────────────────────────────────────────────────────────────────
# Constants & Tag Maps
# ─────────────────────────────────────────────────────────────────────────────

USER_AGENT = "FinancialWebAnalyzer/1.0 (financial-analyzer@example.com)"
EDGAR_TICKER_API   = "https://www.sec.gov/files/company_tickers.json"
EDGAR_SUBMISSIONS  = "https://data.sec.gov/submissions/CIK{cik}.json"
EDGAR_FACTS_API    = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"
EDGAR_ARCHIVES     = "https://www.sec.gov/Archives/edgar/data/{cik_int}/{acc_nodash}/{doc}"

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
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        "RevenueFromContractWithCustomerIncludingAssessedTax",
        "Revenues", "SalesRevenueNet", "SalesRevenueGoodsNet",
        "SalesAndRevenuesNet", "RevenuesNetOfInterestExpense",
    ],
    "OperatingExpenses": [
        "OperatingExpenses",
        "CostsAndExpenses",
        "OperatingCostsAndExpenses",
        "CostAndExpenses",
        "CostOfRevenue",
        "CostOfGoodsAndServicesSold",
        "CostOfGoodsSold",
        "CostOfSales",
        "CostsOfRevenue",
        "CostOfGoodsAndServiceExcludingDepreciationDepletionAndAmortization",
    ],
    "OperatingIncomeLoss": [
        "OperatingIncomeLoss",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
    ],
    "InterestExpense": [
        "NonoperatingIncomeExpense",
        "OtherNonoperatingIncomeExpense",
        "OtherNonoperatingExpense",
        "NonoperatingExpense",
        "InterestIncomeExpenseNet",
        "InterestAndDebtExpense",
        "InterestExpense",
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
        "NetIncomeLossAttributableToParentCompany",
    ],
}

BS_TAGS = {
    "Cash": [
        "CashAndCashEquivalentsAtCarryingValue",
        "CashCashEquivalentsAndShortTermInvestments",
        "CashAndShortTermInvestments", "Cash",
        "CashAndCashEquivalentsAndRestrictedCashAndRestrictedCashEquivalents",
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
        "LongTermDebt",
        "LongTermDebtNoncurrent",
        "LiabilitiesNoncurrent",
        "LongTermDebtAndCapitalLeaseObligations",
        "LongTermDebtAndFinanceLeaseLiabilities",
        "FinanceLeaseLiabilityNoncurrent",
        "LongTermLineOfCredit",
        "SeniorLongTermNotes",
        "DebtAndCapitalLeaseObligations",
    ],
    "NonCurrentAssets": [
        "PropertyPlantAndEquipmentNet",
        "AssetsNoncurrent",
        "PropertyPlantAndEquipmentAndIntangibleAssetsNet",
        "PropertyPlantAndEquipmentNetIncludingDiscontinuedOperations",
        "NoncurrentAssets",
        "PropertyPlantAndEquipmentGross",
        "PropertyPlantAndEquipmentAndFinanceLeaseRightOfUseAssetAfterAccumulatedDepreciationAndAmortization",
        "RightOfUseAsset",
    ],
    "TotalLiabilities": [
        "Liabilities",
        "LiabilitiesTotal",
    ],
    "TotalAssets": [
        "Assets",
        "AssetsNet",
    ],
    "StockholdersEquity": [
        "StockholdersEquity",
        "StockholdersEquityAttributableToParent",
        "PartnersCapital", "MembersEquity",
        "LiabilitiesAndStockholdersEquity",
    ],
}

PL_LABELS = {
    "Revenues":            "売上高 / Revenues",
    "OperatingExpenses":   "営業費用 / Operating Expenses",
    "OperatingIncomeLoss": "営業利益 / Operating Income",
    "InterestExpense":     "営業外費用 / Total Other Expense, net",
    "IncomeLossBeforeTax": "税引前利益 / Income Before Tax",
    "NetIncomeLoss":       "純利益 / Net Income",
}

# PARR以外（一般企業）向け。標準的な損益計算書の並びを網羅し、候補タグも大幅に拡充。
# 直接タグが取得できない項目は extract_pl() 内で会計恒等式から導出する。
SIMPLE_PL_TAGS = {
    "Revenues": [
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        "RevenueFromContractWithCustomerIncludingAssessedTax",
        "Revenues", "SalesRevenueNet", "SalesRevenueGoodsNet",
        "SalesAndRevenuesNet", "RevenuesNetOfInterestExpense",
        "RevenueFromContractWithCustomerExcludingAssessedTaxProductAndService",
    ],
    "CostOfRevenue": [
        "CostOfRevenue",
        "CostOfGoodsAndServicesSold",
        "CostOfGoodsSold",
        "CostOfSales",
        "CostsOfRevenue",
        "CostOfGoodsAndServiceExcludingDepreciationDepletionAndAmortization",
    ],
    "GrossProfit": ["GrossProfit"],
    "OperatingExpenses": [
        "OperatingExpenses",
        "OperatingCostsAndExpenses",
        "CostsAndExpenses",
        "CostAndExpenses",
    ],
    "OperatingIncomeLoss": [
        "OperatingIncomeLoss",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
    ],
    # Excel出力(_build_fd_sheet)と同じキー名を使うためキーは InterestExpense のまま
    "InterestExpense": [
        "NonoperatingIncomeExpense",
        "OtherNonoperatingIncomeExpense",
        "OtherNonoperatingExpense",
        "NonoperatingExpense",
        "InterestIncomeExpenseNet",
        "InterestAndDebtExpense",
        "InterestExpense",
    ],
    "IncomeLossBeforeTax": [
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesDomestic",
    ],
    "IncomeTaxExpense": [
        "IncomeTaxExpenseBenefit",
        "CurrentIncomeTaxExpenseBenefit",
    ],
    "NetIncomeLoss": [
        "NetIncomeLoss", "ProfitLoss",
        "NetIncomeLossAvailableToCommonStockholdersBasic",
        "IncomeLossFromContinuingOperations",
        "NetIncomeLossAttributableToParentCompany",
    ],
}

SIMPLE_PL_LABELS = {
    "Revenues":            "売上高 / Revenues",
    "CostOfRevenue":       "売上原価 / Cost of Revenue",
    "GrossProfit":         "売上総利益 / Gross Profit",
    "OperatingExpenses":   "営業費用 / Operating Expenses",
    "OperatingIncomeLoss": "営業利益 / Operating Income",
    "InterestExpense":     "営業外損益 / Non-Operating Income (Expense)",
    "IncomeLossBeforeTax": "税引前利益 / Income Before Tax",
    "IncomeTaxExpense":    "法人税等 / Income Tax Expense",
    "NetIncomeLoss":       "純利益 / Net Income",
}

# 四半期トレンドグラフ用（全企業対象）: PARR用と一般企業用の候補タグを統合したもの
_TREND_REVENUE_TAGS = list(dict.fromkeys(
    PL_TAGS["Revenues"] + SIMPLE_PL_TAGS["Revenues"]))
_TREND_NETINCOME_TAGS = list(dict.fromkeys(
    PL_TAGS["NetIncomeLoss"] + SIMPLE_PL_TAGS["NetIncomeLoss"]))
BS_LABELS = {
    "Cash":               "手元資金 / Cash & Equivalents",
    "CurrentLiabilities": "流動負債 / Current Liabilities",
    "LongTermLiabilities":"長期負債 / LT Liabilities",
    "NonCurrentAssets":   "固定資産 / Non-Current Assets",
    "StockholdersEquity": "株主資本 / Stockholders' Equity",
}

# ─────────────────────────────────────────────────────────────────────────────
# HTTP Helper
# ─────────────────────────────────────────────────────────────────────────────

def _headers():
    return {"User-Agent": USER_AGENT, "Accept": "application/json, text/html"}

def _get(url: str, retries: int = 4, backoff: float = 2.0, as_text: bool = False):
    for attempt in range(retries):
        try:
            r = requests.get(url, headers=_headers(), timeout=30)
            if r.status_code == 200:
                return r.text if as_text else r
            if r.status_code == 429:
                time.sleep(backoff * (2 ** attempt))
                continue
            return None
        except requests.RequestException:
            time.sleep(backoff * (2 ** attempt))
    return None

# ─────────────────────────────────────────────────────────────────────────────
# CIK Resolution
# ─────────────────────────────────────────────────────────────────────────────

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

# ─────────────────────────────────────────────────────────────────────────────
# Filings List (10-Q / 10-K)
# ─────────────────────────────────────────────────────────────────────────────

def get_filings_list(cik: str) -> list[dict]:
    """Return list of dicts for recent 10-Q and 10-K filings."""
    r = _get(EDGAR_SUBMISSIONS.format(cik=cik))
    if not r:
        return []
    try:
        data = r.json()
    except Exception:
        return []
    fy_end  = data.get("fiscalYearEnd", "1231")   # e.g. "1231" or "0331"
    recent  = data.get("filings", {}).get("recent", {})
    forms      = recent.get("form", [])
    dates      = recent.get("filingDate", [])
    periods    = recent.get("reportDate", [])
    accessions = recent.get("accessionNumber", [])
    prim_docs  = recent.get("primaryDocument", [])

    result = []
    for form, date, period, acc, doc in zip(forms, dates, periods, accessions, prim_docs):
        if form in ("10-Q", "10-K") and period:
            result.append({
                "form": form,
                "date": date,
                "period": period,
                "accession": acc,
                "primary_doc": doc,
                "fy_end": fy_end,
                "label": f"{form} — {period}（提出日: {date}）",
            })
    return result[:24]

# ─────────────────────────────────────────────────────────────────────────────
# MD&A Extraction
# ─────────────────────────────────────────────────────────────────────────────

class _HTMLStripper(HTMLParser):
    _SKIP = {"script", "style", "head", "meta", "link", "noscript"}
    _BLOCK = {"p", "div", "br", "li", "tr", "td", "th",
              "h1", "h2", "h3", "h4", "h5", "h6", "table", "section"}

    def __init__(self):
        super().__init__()
        self._parts = []
        self._depth = 0

    def handle_starttag(self, tag, attrs):
        if tag in self._SKIP:
            self._depth += 1
        if self._depth == 0 and tag in self._BLOCK:
            self._parts.append("\n")

    def handle_endtag(self, tag):
        if tag in self._SKIP and self._depth > 0:
            self._depth -= 1

    def handle_data(self, data):
        if self._depth == 0:
            self._parts.append(data)

    def get_text(self):
        return "".join(self._parts)


def _strip_html(html: str) -> str:
    s = _HTMLStripper()
    try:
        s.feed(html)
        return s.get_text()
    except Exception:
        return re.sub(r"<[^>]+>", " ", html)



def _find_toc_end(text: str) -> int:
    """Estimate where the Table of Contents ends.
    TOC entries look like: "Item N. Some Title.........32"
    Returns char position after the last detected TOC entry
    (searched only within first 40% of the document).
    """
    toc_pat = re.compile(r'(?i)\bITEM\s+\d+[A-Za-z]?[^\n]{0,120}\s+\d{1,3}\s*\n')
    limit = int(len(text) * 0.40)
    last_pos = 0
    for m in toc_pat.finditer(text, 0, limit):
        last_pos = m.end()
    return last_pos


def extract_mda(html: str, form_type: str = "10-Q", max_chars: int = 30_000) -> str:
    """Extract MD&A prose from a SEC filing HTML.

    10-K: Item 7 → Item 8  (Item 7Aコンテンツも含む)
    10-Q: Item 2 → Item 3

    Strategy:
    1. Strip all HTML tables (removes financial tables + most TOC tables).
    2. For 10-K: detect TOC end position, then search for Item 7 only in
       the document body (post-TOC region).
    3. Extract to the next end-marker (Item 8 for 10-K, Item 3 for 10-Q).
    4. Fallback: if TOC detection misses, search full text with best-match.
    """
    # ── Strip tables ────────────────────────────────────────────────────────
    html = re.sub(r'<table[\s>].*?</table>', ' ', html,
                  flags=re.IGNORECASE | re.DOTALL)
    text = _strip_html(html)
    text = re.sub(r'[ \t]{2,}', ' ', text)
    text = re.sub(r'\n{3,}', '\n\n', text)

    if form_type == "10-K":
        # Full section title (most specific → avoids TOC shorthand)
        full_title = (
            r'(?i)ITEM[\s.]*7[.\s]+'
            "MANAGEMENT[\u2019\u2018'\u0060]?S[\s]+"
            r'DISCUSSION[\s]+AND[\s]+ANALYSIS[\s]+OF[\s]+'
            r'FINANCIAL[\s]+CONDITION[\s]+AND[\s]+RESULTS[\s]+OF[\s]+OPERATIONS'
        )
        short_title = r'(?i)ITEM[\s.]*7[.\s]+MANAGEMENT'
        bare        = r'(?i)\bITEM\s*7\b'
        start_pats  = [full_title, short_title, bare]
        # Stop at Item 8 only (include Item 7A in the extract)
        end_pats = [
            r'(?i)\bITEM\s*8[.\s]',
            r'(?i)\bITEM\s*8\b',
        ]
    else:
        full_title = (
            r'(?i)ITEM[\s.]*2[.\s]+'
            "MANAGEMENT[\u2019\u2018'\u0060]?S[\s]+"
            r'DISCUSSION[\s]+AND[\s]+ANALYSIS[\s]+OF[\s]+'
            r'FINANCIAL[\s]+CONDITION[\s]+AND[\s]+RESULTS[\s]+OF[\s]+OPERATIONS'
        )
        short_title = r'(?i)ITEM[\s.]*2[.\s]+MANAGEMENT'
        bare        = r'(?i)\bITEM\s*2\b'
        start_pats  = [full_title, short_title, bare]
        end_pats = [
            r'(?i)\bITEM\s*3[.\s]',
            r'(?i)\bITEM\s*3\b',
            r'(?i)\bITEM\s*4\b',
        ]

    # ── Primary: search only in the post-TOC body region ─────────────────
    body_offset = _find_toc_end(text) if form_type == "10-K" else 0
    search_text = text[body_offset:]

    section_start = -1
    for pat in start_pats:
        m = re.search(pat, search_text)
        if m:
            section_start = body_offset + m.start()
            break

    # ── Fallback: best-match across full text ────────────────────────────
    if section_start == -1:
        best_start, best_len = -1, 0
        for pat in start_pats:
            matched = False
            for m in re.finditer(pat, text):
                matched = True
                pos = m.start(); hl = len(m.group())
                probe = text[pos + hl: pos + hl + 120_000]
                eoff = len(probe)
                for ep in end_pats:
                    em = re.search(ep, probe)
                    if em: eoff = min(eoff, em.start())
                if eoff > best_len:
                    best_len, best_start = eoff, pos
            if matched: break
        section_start = best_start

    if section_start == -1:
        return text[:max_chars].strip()

    # ── Cut from section_start to the end marker ────────────────────────
    tail = text[section_start:]
    end_offset = len(tail)
    for pat in end_pats:
        m = re.search(pat, tail[80:])   # skip past the section header itself
        if m:
            end_offset = min(end_offset, m.start() + 80)

    return tail[:end_offset].strip()[:max_chars]


def fetch_mda(cik: str, accession: str, primary_doc: str, form_type: str = "10-Q") -> str:
    cik_int    = int(cik)
    acc_nodash = accession.replace("-", "")
    url = EDGAR_ARCHIVES.format(cik_int=cik_int, acc_nodash=acc_nodash, doc=primary_doc)
    html = _get(url, as_text=True)
    if not html:
        return "（MD&Aテキストを取得できませんでした。SECのネットワーク制限またはファイル形式の問題の可能性があります。）"
    return extract_mda(html, form_type=form_type)

# ─────────────────────────────────────────────────────────────────────────────
# XBRL Fact Fetching & Extraction
# ─────────────────────────────────────────────────────────────────────────────

def fetch_facts(cik: str) -> dict:
    r = _get(EDGAR_FACTS_API.format(cik=cik))
    return r.json() if r else {}


def _units(facts: dict, tag: str) -> list:
    try:
        u = facts.get("facts", {}).get("us-gaap", {}).get(tag, {}).get("units", {})
        return u.get("USD", u.get("pure", []))
    except Exception:
        return []


def _nearest_days(records: list, target: datetime | None) -> int | None:
    """レコード群の中で target に最も近い end 日付との差（日数）。"""
    if target is None or not records:
        return None
    best = None
    for r in records:
        try:
            d = abs((datetime.strptime(r["end"], "%Y-%m-%d") - target).days)
            if best is None or d < best:
                best = d
        except (ValueError, KeyError, TypeError):
            pass
    return best


def _best_tag(facts: dict, candidates: list, target: datetime | None = None,
              tolerance_days: int = 65) -> tuple[str, list]:
    """対象期をカバーしている候補タグを選ぶ。

    企業は年度によって使用するXBRLタグを変える（例: Revenues →
    RevenueFromContractWithCustomerExcludingAssessedTax）。「データが存在する
    最初の候補」を無条件に採用すると、対象期のデータを持たない旧タグを掴んで
    N/Aになることがあるため、target が与えられた場合は対象期に最も近い
    レコードを持つ候補を優先する。target が無い場合は従来通りの挙動。
    """
    first_tag, first_recs = "", []
    best_tag, best_recs, best_d = "", [], None
    for tag in candidates:
        recs = _units(facts, tag)
        if not recs:
            continue
        if not first_recs:
            first_tag, first_recs = tag, recs
        d = _nearest_days(_filter_instant(recs) or recs, target)
        if d is not None and d <= tolerance_days and (best_d is None or d < best_d):
            best_tag, best_recs, best_d = tag, recs, d
    if best_recs:
        return best_tag, best_recs
    return first_tag, first_recs


# Duration ranges per fiscal quarter: (strict_lo, strict_hi, relaxed_lo, relaxed_hi) in days
_QUARTER_RANGES = {
    1: (75,  110, 60,  125),   # Q1  ~3 months
    2: (165, 195, 150, 210),   # Q2  ~6 months YTD
    3: (255, 285, 240, 300),   # Q3  ~9 months YTD
    4: (340, 400, 325, 415),   # Q4/10-K  ~12 months
}


def _filter_duration(records: list, lo: int, hi: int) -> list:
    out = []
    for r in records:
        s, e = r.get("start", ""), r.get("end", "")
        if s and e:
            try:
                d = (datetime.strptime(e, "%Y-%m-%d") - datetime.strptime(s, "%Y-%m-%d")).days
                if lo <= d <= hi:
                    out.append(r)
            except ValueError:
                pass
    return out


def _quarter_num(period_str: str, fy_end_mmdd: str) -> int:
    """Compute fiscal quarter (1–4) from period end date and FY-end MMDD string (e.g. '1231')."""
    try:
        p_month = datetime.strptime(period_str, "%Y-%m-%d").month
        fy_month = int(fy_end_mmdd[:2])
        fy_start_month = (fy_month % 12) + 1
        months_in = (p_month - fy_start_month) % 12 + 1
        return min(4, (months_in + 2) // 3)
    except Exception:
        return 1


def _best_ytd_tag(facts: dict, candidates: list, quarter_num: int,
                  target: datetime | None = None,
                  tolerance_days: int = 65) -> tuple[str, list, list]:
    """Return (tag, all_recs, period_recs) for the candidate whose records match
    the YTD duration for the given fiscal quarter (1=3M, 2=6M, 3=9M, 4=12M).

    target が与えられた場合は「対象期に最も近いレコードを持つ候補」を優先する
    （企業が年度途中でタグを切り替えているケースで、旧タグを掴んでN/Aになるのを防ぐ）。
    厳密な期間レンジで一致する候補を、緩いレンジより優先する。
    target が無い場合は従来通り「最初に一致した候補」を返す。
    """
    lo, hi, rlo, rhi = _QUARTER_RANGES.get(quarter_num, _QUARTER_RANGES[1])
    any_tag, any_recs = "", []
    strict_first = None
    relaxed_first = None
    best = None                      # (rank, tag, recs, period_recs)
    for tag in candidates:
        recs = _units(facts, tag)
        if not recs:
            continue
        if not any_recs:
            any_tag, any_recs = tag, recs
        f  = _filter_duration(recs, lo, hi)
        fw = f or _filter_duration(recs, rlo, rhi)
        if f and strict_first is None:
            strict_first = (tag, recs, f)
        if fw and relaxed_first is None:
            relaxed_first = (tag, recs, fw)
        if target is not None and fw:
            d = _nearest_days(fw, target)
            if d is not None and d <= tolerance_days:
                rank = (d, 0 if f else 1)     # 近さ優先、同距離なら厳密レンジ優先
                if best is None or rank < best[0]:
                    best = (rank, tag, recs, fw)
    if best is not None:
        return best[1], best[2], best[3]
    if strict_first:
        return strict_first
    if relaxed_first:
        return relaxed_first
    return any_tag, any_recs, []


def _filter_instant(records: list) -> list:
    return [r for r in records if not r.get("start")]


def _dedup_latest(records: list, n: int) -> list:
    seen = {}
    for r in records:
        e = r.get("end", "")
        if e not in seen or r.get("form", "") in ("10-Q", "10-K"):
            seen[e] = r
    return sorted(seen.values(), key=lambda x: x.get("end", ""), reverse=True)[:n]


def _find_closest(records: list, target: datetime, max_days: int = 45) -> dict | None:
    best, best_d = None, max_days + 1
    for r in records:
        try:
            diff = abs((datetime.strptime(r["end"], "%Y-%m-%d") - target).days)
            if diff < best_d:
                best, best_d = r, diff
        except (ValueError, KeyError):
            pass
    return best


_Q_PERIOD_LABELS = {1: "3ヶ月", 2: "6ヶ月累積", 3: "9ヶ月累積", 4: "通期（12ヶ月）"}


def extract_pl(facts: dict, target_period: str | None = None, form_type: str = "10-Q",
               quarter_num: int = 1, is_parr: bool = False) -> dict:
    target = datetime.strptime(target_period, "%Y-%m-%d") if target_period else None
    effective_q = 4 if (form_type == "10-K") else quarter_num
    result = {}

    # Choose tag set based on company type
    tag_set = PL_TAGS if is_parr else SIMPLE_PL_TAGS

    for metric, candidates in tag_set.items():
        tag, _recs, period_recs = _best_ytd_tag(facts, candidates, effective_q, target)
        period_recs = _dedup_latest(period_recs, 100)
        current = prior = None
        if period_recs:
            cur_rec = _find_closest(period_recs, target, 65) if target else period_recs[0]
            if cur_rec is not None:
                current = (cur_rec.get("val"), cur_rec["end"], tag)
                cur_end = datetime.strptime(cur_rec["end"], "%Y-%m-%d")
                # YoY: go back exactly one year for both quarterly and annual
                prior_target = cur_end - timedelta(days=365)
                others = [r for r in period_recs if r["end"] != cur_rec["end"]]
                prior_rec = _find_closest(others, prior_target, 65)
                if prior_rec:
                    prior = (prior_rec.get("val"), prior_rec["end"], tag)
        result[metric] = {"current": current, "prior": prior}

    if is_parr:
        # Derive OperatingExpenses = Revenues − OperatingIncomeLoss per period independently.
        # Runs for each period where the XBRL tags above yielded no value.
        rev = result.get("Revenues", {})
        opi = result.get("OperatingIncomeLoss", {})
        for which in ("current", "prior"):
            t = result.get("OperatingExpenses", {}).get(which)
            if t is None or t[0] is None:
                r_t = rev.get(which)
                o_t = opi.get(which)
                if r_t and o_t and r_t[0] is not None and o_t[0] is not None:
                    result.setdefault("OperatingExpenses", {})[which] = (
                        r_t[0] - o_t[0], r_t[1], "※導出値: Revenues − OperatingIncomeLoss"
                    )
    else:
        # 一般企業向け: 直接タグが取得できなかった項目を会計恒等式から補完する。
        # 企業ごとに開示する科目が異なるため、取得できた科目だけを使って
        # 補完可能なものを順に埋めていく（補完できない場合は N/A のまま）。
        def _t(metric, which):
            v = result.get(metric, {}).get(which)
            return v if (v and v[0] is not None) else None

        def _set(metric, which, val, ref, note):
            result.setdefault(metric, {"current": None, "prior": None})[which] = (val, ref[1], note)

        for which in ("current", "prior"):
            rev = _t("Revenues", which)
            cor = _t("CostOfRevenue", which)
            gp  = _t("GrossProfit", which)
            opi = _t("OperatingIncomeLoss", which)
            nop = _t("InterestExpense", which)
            ibt = _t("IncomeLossBeforeTax", which)
            ni  = _t("NetIncomeLoss", which)

            # 売上総利益 = 売上高 − 売上原価
            if gp is None and rev and cor:
                _set("GrossProfit", which, rev[0] - cor[0], rev,
                     "※導出値: Revenues − CostOfRevenue")
                gp = _t("GrossProfit", which)
            # 売上原価 = 売上高 − 売上総利益
            if cor is None and rev and gp:
                _set("CostOfRevenue", which, rev[0] - gp[0], rev,
                     "※導出値: Revenues − GrossProfit")
            # 営業費用 = 売上高 − 営業利益
            if _t("OperatingExpenses", which) is None and rev and opi:
                _set("OperatingExpenses", which, rev[0] - opi[0], rev,
                     "※導出値: Revenues − OperatingIncomeLoss")
            # 税引前利益 = 営業利益 + 営業外損益
            if ibt is None and opi and nop:
                _set("IncomeLossBeforeTax", which, opi[0] + nop[0], opi,
                     "※導出値: OperatingIncome + NonOperating")
                ibt = _t("IncomeLossBeforeTax", which)
            # 法人税等 = 税引前利益 − 純利益
            if _t("IncomeTaxExpense", which) is None and ibt and ni:
                _set("IncomeTaxExpense", which, ibt[0] - ni[0], ibt,
                     "※導出値: IncomeBeforeTax − NetIncome")

    return result


def extract_quarterly_trend(facts: dict, filings: list, years: int = 3) -> pd.DataFrame:
    """決算期リスト（最大24期）から、過去N年分の「四半期単独」の売上高・純利益を算出する
    （全企業対象）。10-Qは累積(YTD)値、10-Kは通期値として開示されるため、前の四半期までの
    累積値を差し引く de-cumulation を行う（例: Q2単独 = 6ヶ月累積 − Q1）。
    タグ候補はPARR用・一般企業用を統合したものを使い、期ごとに対象期に最も近い
    レコードを持つタグを選ぶため、年度途中でタグを切り替えた企業にも追従できる。
    直近フィリングの期末日を基準に過去N年分に絞り込んで返す。
    """
    ordered = sorted(
        [f for f in filings if f.get("form") in ("10-Q", "10-K") and f.get("period")],
        key=lambda f: f["period"],
    )
    rows = []
    prev_rev = prev_ni = None
    for f in ordered:
        period = f["period"]
        form   = f.get("form", "10-Q")
        fy_end = f.get("fy_end", "1231")
        q_num  = _quarter_num(period, fy_end)
        eff_q  = 4 if form == "10-K" else q_num
        target = datetime.strptime(period, "%Y-%m-%d")

        _, _, rev_recs = _best_ytd_tag(facts, _TREND_REVENUE_TAGS, eff_q, target)
        _, _, ni_recs  = _best_ytd_tag(facts, _TREND_NETINCOME_TAGS, eff_q, target)
        rev_recs = _dedup_latest(rev_recs, 100)
        ni_recs  = _dedup_latest(ni_recs, 100)
        rev_rec  = _find_closest(rev_recs, target, 65) if rev_recs else None
        ni_rec   = _find_closest(ni_recs, target, 65) if ni_recs else None
        ytd_rev  = rev_rec.get("val") if rev_rec else None
        ytd_ni   = ni_rec.get("val")  if ni_rec  else None

        # 四半期単独値 = 当YTD − 前四半期までのYTD（Q1、または直前データが無い場合はYTDそのまま）
        if eff_q == 1 or prev_rev is None:
            q_rev = ytd_rev
        else:
            q_rev = (ytd_rev - prev_rev) if (ytd_rev is not None and prev_rev is not None) else None
        if eff_q == 1 or prev_ni is None:
            q_ni = ytd_ni
        else:
            q_ni = (ytd_ni - prev_ni) if (ytd_ni is not None and prev_ni is not None) else None

        prev_rev, prev_ni = ytd_rev, ytd_ni

        rows.append({
            "period": period, "period_dt": target, "form": form, "quarter_num": eff_q,
            "quarter_label": f"{target.year}Q{eff_q}",
            "revenue": _m(q_rev) if q_rev is not None else None,
            "net_income": _m(q_ni) if q_ni is not None else None,
        })

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    latest = df["period_dt"].max()
    cutoff = latest - timedelta(days=365 * years + 45)
    df = df[df["period_dt"] >= cutoff].sort_values("period_dt").reset_index(drop=True)
    return df


def extract_bs(facts: dict, target_period: str | None = None) -> dict:
    target = datetime.strptime(target_period, "%Y-%m-%d") if target_period else None
    result = {}
    for metric, candidates in BS_TAGS.items():
        tag, recs = _best_tag(facts, candidates, target)
        instants = _dedup_latest(_filter_instant(recs), 100)
        current = prior = None
        if instants:
            cur_rec = _find_closest(instants, target, 65) if target else instants[0]
            if cur_rec is None and not target:
                cur_rec = instants[0]
            if cur_rec is not None:
                current = (cur_rec.get("val"), cur_rec["end"], tag)
                cur_end = datetime.strptime(cur_rec["end"], "%Y-%m-%d")
                prior_target = cur_end - timedelta(days=92)
                others = [r for r in instants if r["end"] != cur_rec["end"]]
                prior_rec = _find_closest(others, prior_target, 65)
                if prior_rec:
                    prior = (prior_rec.get("val"), prior_rec["end"], tag)
        result[metric] = {"current": current, "prior": prior}

    # Always derive LongTermLiabilities = TotalLiabilities - CurrentLiabilities
    tl = result.get("TotalLiabilities", {})
    cl = result.get("CurrentLiabilities", {})
    for which in ("current", "prior"):
        tl_t = tl.get(which); cl_t = cl.get(which)
        if tl_t and cl_t and tl_t[0] is not None and cl_t[0] is not None:
            result["LongTermLiabilities"][which] = (
                tl_t[0] - cl_t[0], tl_t[1], "※導出: TotalLiabilities − CurrentLiabilities"
            )

    # Always derive NonCurrentAssets = TotalAssets - CurrentAssets
    tot = result.get("TotalAssets", {})
    ca  = result.get("CurrentAssets", {})
    for which in ("current", "prior"):
        t_t = tot.get(which); c_t = ca.get(which)
        if t_t and c_t and t_t[0] is not None and c_t[0] is not None:
            result["NonCurrentAssets"][which] = (
                t_t[0] - c_t[0], t_t[1], "※導出: TotalAssets − CurrentAssets"
            )

    result.pop("TotalLiabilities", None)
    result.pop("TotalAssets", None)
    return result

# ─────────────────────────────────────────────────────────────────────────────
# DataFrame Builders
# ─────────────────────────────────────────────────────────────────────────────

def _m(val):
    try:
        return float(val) / 1_000_000 if val is not None else None
    except (TypeError, ValueError):
        return None


def _safe_val(data: dict, which: str):
    """Extract float value from a (val, date, tag) tuple safely."""
    v = data.get(which)
    return _m(v[0]) if v is not None else None


def _canon_date(dataset: dict, which: str) -> str:
    """Find the most-represented period date across all metrics."""
    from collections import Counter
    dates = []
    for data in dataset.values():
        v = data.get(which)
        if v is not None:
            dates.append(v[1])
    if not dates:
        return "—"
    return Counter(dates).most_common(1)[0][0]


def _period_ok(data: dict, which: str, canon: str, tolerance_days: int = 50) -> bool:
    """Return True if a metric's period date is within tolerance of canonical date."""
    v = data.get(which)
    if v is None or canon == "—":
        return False
    try:
        diff = abs((datetime.strptime(v[1], "%Y-%m-%d")
                    - datetime.strptime(canon, "%Y-%m-%d")).days)
        return diff <= tolerance_days
    except ValueError:
        return False


def _pct_label(pct, cur, pri, is_bad_if_high: bool) -> str:
    """Human-readable % change with sign-flip awareness."""
    if pct is None:
        return "N/A"
    if cur is not None and pri is not None and cur < 0 and pri < 0:
        return "N/A"
    # Sign-change cases
    if cur is not None and pri is not None:
        if pri < 0 and cur > 0:
            return "黒字転換" if not is_bad_if_high else f"{pct:+.0%}"
        if pri > 0 and cur < 0:
            return "赤字転落" if not is_bad_if_high else f"{pct:+.0%}"
    return f"{pct:+.0%}"


def _alert(pct, cur, pri, is_bad_if_high: bool) -> str:
    if pct is None:
        return "—"
    # Sign-flip cases for income/asset metrics
    if not is_bad_if_high and cur is not None and pri is not None:
        if pri < 0 and cur > 0:
            return "✅ 黒字転換"
        if pri > 0 and cur < 0:
            return "⚠️ 赤字転落"
    if is_bad_if_high:
        return "⚠️ +20%↑ 急増" if pct > 0.20 else "✅ 正常"
    return "⚠️ -20%↓ 急減" if pct < -0.20 else ("✅ +20%↑ 成長" if pct > 0.20 else "✅ 正常")





def build_pl_df(pl: dict, labels: dict | None = None) -> pd.DataFrame:
    # ① Determine canonical period dates (most common across metrics)
    canon_cur = _canon_date(pl, "current")
    canon_pri = _canon_date(pl, "prior")
    cur_col   = f"当期 ({canon_cur})\n[USD M]"
    pri_col   = f"前期 ({canon_pri})\n[USD M]"

    use_labels = labels if labels is not None else PL_LABELS
    rows = []
    for key, label in use_labels.items():
        data = pl.get(key, {})
        # ② Only accept values whose period aligns with the canonical date
        cur = _safe_val(data, "current") if _period_ok(data, "current", canon_cur) else None
        pri = _safe_val(data, "prior")   if _period_ok(data, "prior",   canon_pri) else None
        # Round to integer millions
        cur = round(cur) if cur is not None else None
        pri = round(pri) if pri is not None else None
        tag = data["current"][2]         if data.get("current") else "—"

        delta = (cur - pri) if (cur is not None and pri is not None) else None
        pct   = delta / abs(pri) if (delta is not None and pri not in (None, 0)) else None
        is_cost = key == "OperatingExpenses"
        rows.append({
            "項目 / Metric": label,
            cur_col:         cur,
            pri_col:         pri,
            "差額 [USD M]":  delta,
            "変化率 %":       _pct_label(pct, cur, pri, is_cost),
            "_pct": pct, "_is_cost": is_cost, "_cur": cur, "_pri": pri, "_tag": tag,
        })
    return pd.DataFrame(rows)


def build_bs_df(bs: dict) -> pd.DataFrame:
    canon_cur = _canon_date(bs, "current")
    canon_pri = _canon_date(bs, "prior")
    cur_col   = f"当四半期末 ({canon_cur})\n[USD M]"
    pri_col   = f"前四半期末 ({canon_pri})\n[USD M]"

    def _bv(k, w):
        data = bs.get(k, {})
        v = _safe_val(data, w) if _period_ok(data, w, canon_cur if w == "current" else canon_pri) else None
        return round(v) if v is not None else None

    def _item(key, label, is_liab):
        data   = bs.get(key, {})
        cur    = _bv(key, "current"); pri = _bv(key, "prior")
        tag    = data["current"][2] if data.get("current") else "—"
        delta  = (cur - pri) if (cur is not None and pri is not None) else None
        pct    = delta / abs(pri) if (delta is not None and pri not in (None, 0)) else None
        return {"項目 / Metric": label, cur_col: cur, pri_col: pri,
                "差額 [USD M]": delta, "変化率 %": _pct_label(pct, cur, pri, is_liab),
                "_pct": pct, "_is_cost": is_liab, "_cur": cur, "_pri": pri,
                "_tag": tag, "_is_subtotal": False}

    def _subtotal(label, vals_c, vals_p):
        cur = sum(v for v in vals_c if v is not None) if any(v is not None for v in vals_c) else None
        pri = sum(v for v in vals_p if v is not None) if any(v is not None for v in vals_p) else None
        delta = (cur - pri) if (cur is not None and pri is not None) else None
        return {"項目 / Metric": label, cur_col: cur, pri_col: pri,
                "差額 [USD M]": delta, "変化率 %": "",
                "_pct": None, "_is_cost": False, "_cur": cur, "_pri": pri,
                "_tag": "(合計)", "_is_subtotal": True}

    cash_c = _bv("Cash", "current");          cash_p = _bv("Cash", "prior")
    ca_c   = _bv("CurrentAssets", "current"); ca_p   = _bv("CurrentAssets", "prior")
    oca_c  = (ca_c - cash_c) if (ca_c is not None and cash_c is not None) else None
    oca_p  = (ca_p - cash_p) if (ca_p is not None and cash_p is not None) else None
    nca_c  = _bv("NonCurrentAssets", "current"); nca_p = _bv("NonCurrentAssets", "prior")
    cl_c   = _bv("CurrentLiabilities", "current"); cl_p = _bv("CurrentLiabilities", "prior")
    ltl_c  = _bv("LongTermLiabilities", "current"); ltl_p = _bv("LongTermLiabilities", "prior")
    eq_c   = _bv("StockholdersEquity", "current"); eq_p  = _bv("StockholdersEquity", "prior")

    oca_d   = (oca_c - oca_p) if (oca_c is not None and oca_p is not None) else None
    oca_pct = oca_d / abs(oca_p) if (oca_d is not None and oca_p not in (None, 0)) else None

    rows = [
        # ── 資産の部 ──────────────────────────────────
        _item("Cash", "手元資金 / Cash & Equivalents", False),
        {"項目 / Metric": "その他流動資産 / Other Current Assets",
         cur_col: oca_c, pri_col: oca_p, "差額 [USD M]": oca_d,
         "変化率 %": _pct_label(oca_pct, oca_c, oca_p, False),
         "_pct": oca_pct, "_is_cost": False, "_cur": oca_c, "_pri": oca_p,
         "_tag": "(計算値)", "_is_subtotal": False},
        _item("NonCurrentAssets", "固定資産 / Non-Current Assets", False),
        _subtotal("▶ 資産合計 / Total Assets",
                  [cash_c, oca_c, nca_c], [cash_p, oca_p, nca_p]),
        # ── 負債・資本の部 ──────────────────────────────
        _item("CurrentLiabilities", "流動負債 / Current Liabilities", True),
        _item("LongTermLiabilities", "長期負債 / LT Liabilities", True),
        _item("StockholdersEquity", "株主資本 / Stockholders' Equity", False),
        _subtotal("▶ 負債・資本合計 / Total L+E",
                  [cl_c, ltl_c, eq_c], [cl_p, ltl_p, eq_p]),
    ]
    return pd.DataFrame(rows)


def _bs_line_item(facts: dict, candidates: list, target: datetime | None):
    """extract_bs() と同じロジックで、単一のXBRLタグ候補リストから当期・前四半期の
    生値(raw USD, 未換算)を1組取得する（CONDENSED向けの個別科目ルックアップ用）。"""
    tag, recs = _best_tag(facts, candidates, target)
    instants = _dedup_latest(_filter_instant(recs), 100)
    if not instants:
        return None, None
    cur_rec = _find_closest(instants, target, 65) if target else instants[0]
    if cur_rec is None:
        return None, None
    cur_val = cur_rec.get("val")
    cur_end = datetime.strptime(cur_rec["end"], "%Y-%m-%d")
    prior_target = cur_end - timedelta(days=92)
    others = [r for r in instants if r["end"] != cur_rec["end"]]
    prior_rec = _find_closest(others, prior_target, 65)
    pri_val = prior_rec.get("val") if prior_rec else None
    return cur_val, pri_val


# (section_key, 表示ラベル, 候補タグ) — CONDENSED CONSOLIDATED BALANCE SHEETS の
# 標準的な並び順・粒度を模した固定リスト。企業に存在しない科目は自動的にスキップされ、
# 各セクションの「その他（未分類）」行に差額として吸収される。
_CONDENSED_BS_LAYOUT = [
    ("ca",  "売掛金 / Accounts Receivable, net",
     ["AccountsReceivableNetCurrent", "ReceivablesNetCurrent", "AccountsReceivableNet"]),
    ("ca",  "棚卸資産 / Inventories",
     ["InventoryNet", "InventoryNetCurrent"]),
    ("ca",  "前払費用・その他流動資産 / Prepaid & Other Current Assets",
     ["PrepaidExpenseAndOtherAssetsCurrent", "PrepaidExpenseCurrent", "OtherAssetsCurrent"]),
    ("nca", "有形固定資産 / Property, Plant & Equipment, net",
     ["PropertyPlantAndEquipmentNet"]),
    ("nca", "使用権資産（オペレーティングリース）/ Operating Lease ROU Assets",
     ["OperatingLeaseRightOfUseAsset"]),
    ("nca", "のれん / Goodwill",
     ["Goodwill"]),
    ("nca", "無形資産 / Intangible Assets, net",
     ["FiniteLivedIntangibleAssetsNet", "IntangibleAssetsNetExcludingGoodwill"]),
    ("cl",  "買掛金 / Accounts Payable",
     ["AccountsPayableCurrent", "AccountsPayableTradeCurrent"]),
    ("cl",  "未払費用 / Accrued Liabilities",
     ["AccruedLiabilitiesCurrent", "AccountsPayableAndAccruedLiabilitiesCurrent"]),
    ("cl",  "長期負債の流動部分 / Current Portion of LT Debt",
     ["LongTermDebtCurrent", "DebtCurrent"]),
    ("cl",  "流動リース負債 / Current Operating Lease Liabilities",
     ["OperatingLeaseLiabilityCurrent"]),
    ("ncl", "長期負債 / Long-Term Debt",
     ["LongTermDebtNoncurrent", "LongTermDebt"]),
    ("ncl", "固定リース負債 / Non-Current Operating Lease Liabilities",
     ["OperatingLeaseLiabilityNoncurrent"]),
    ("ncl", "繰延税金負債 / Deferred Tax Liabilities",
     ["DeferredIncomeTaxLiabilitiesNet", "DeferredTaxLiabilitiesNoncurrent"]),
    ("eq",  "資本金 / Common Stock",
     ["CommonStockValue"]),
    ("eq",  "資本剰余金 / Additional Paid-in Capital",
     ["AdditionalPaidInCapital", "AdditionalPaidInCapitalCommonStock"]),
    ("eq",  "利益剰余金 / Retained Earnings (Accumulated Deficit)",
     ["RetainedEarningsAccumulatedDeficit"]),
    ("eq",  "その他包括利益累計額 / Accumulated OCI",
     ["AccumulatedOtherComprehensiveIncomeLossNetOfTax"]),
    ("eq",  "自己株式 / Treasury Stock",
     ["TreasuryStockValue", "TreasuryStockCommonValue"]),
]


def extract_bs_condensed(facts: dict, bs: dict, target_period: str | None = None) -> pd.DataFrame:
    """PARR限定機能: 実際の「CONDENSED CONSOLIDATED BALANCE SHEETS」に近い粒度・並び順で
    B/Sを表示する（現金→売掛金→棚卸資産→…→流動資産合計→固定資産→資産合計→…の順）。

    個別科目は専用のXBRLタグ候補から取得し、各セクションの小計は extract_bs() が
    既に算出している公式の合計値（Cash / CurrentAssets / LongTermLiabilities /
    StockholdersEquity 等）をそのまま使う。個別科目の合計と公式合計との差額は
    「その他（未分類）」行に計上するため、小計は常に公式の値と一致する。
    """
    target = datetime.strptime(target_period, "%Y-%m-%d") if target_period else None

    canon_cur = _canon_date(bs, "current")
    canon_pri = _canon_date(bs, "prior")
    cur_col = f"当四半期末 ({canon_cur})\n[USD M]"
    pri_col = f"前四半期末 ({canon_pri})\n[USD M]"

    def _official(key, which):
        data = bs.get(key, {})
        canon = canon_cur if which == "current" else canon_pri
        return _safe_val(data, which) if _period_ok(data, which, canon) else None

    def _row(label, cur, pri, is_bad_if_high, tag_note, is_subtotal=False):
        cur = round(cur) if cur is not None else None
        pri = round(pri) if pri is not None else None
        delta = (cur - pri) if (cur is not None and pri is not None) else None
        pct = delta / abs(pri) if (delta is not None and pri not in (None, 0)) else None
        return {
            "項目 / Metric": label, cur_col: cur, pri_col: pri,
            "差額 [USD M]": delta,
            "変化率 %": "" if is_subtotal else _pct_label(pct, cur, pri, is_bad_if_high),
            "_pct": None if is_subtotal else pct, "_is_cost": is_bad_if_high,
            "_cur": cur, "_pri": pri, "_tag": tag_note, "_is_subtotal": is_subtotal,
        }

    # 個別科目をセクション別に集計（データが無い科目は自動的にスキップ）
    sections = {"ca": [], "nca": [], "cl": [], "ncl": [], "eq": []}
    for sec, label, candidates in _CONDENSED_BS_LAYOUT:
        cur_raw, pri_raw = _bs_line_item(facts, candidates, target)
        if cur_raw is None and pri_raw is None:
            continue
        cur_m = _m(cur_raw) if cur_raw is not None else None
        pri_m = _m(pri_raw) if pri_raw is not None else None
        sections[sec].append((label, cur_m, pri_m))

    def _sum(items, idx):
        return sum((it[idx] or 0) for it in items)

    rows = []

    # ── 資産の部 ──────────────────────────────────────────
    cash_c, cash_p = _official("Cash", "current"), _official("Cash", "prior")
    rows.append(_row("手元資金 / Cash & Equivalents", cash_c, cash_p, False, "Cash"))
    for label, c, p in sections["ca"]:
        rows.append(_row(label, c, p, False, "(内訳)"))
    ca_c, ca_p = _official("CurrentAssets", "current"), _official("CurrentAssets", "prior")
    named_ca_c = (cash_c or 0) + _sum(sections["ca"], 1)
    named_ca_p = (cash_p or 0) + _sum(sections["ca"], 2)
    plug_ca_c = (ca_c - named_ca_c) if ca_c is not None else None
    plug_ca_p = (ca_p - named_ca_p) if ca_p is not None else None
    rows.append(_row("その他流動資産（未分類）/ Other Current Assets",
                      plug_ca_c, plug_ca_p, False, "(算出差額)"))
    rows.append(_row("▶ 流動資産合計 / Total Current Assets",
                      ca_c, ca_p, False, "(小計)", is_subtotal=True))

    for label, c, p in sections["nca"]:
        rows.append(_row(label, c, p, False, "(内訳)"))
    nca_c, nca_p = _official("NonCurrentAssets", "current"), _official("NonCurrentAssets", "prior")
    plug_nca_c = (nca_c - _sum(sections["nca"], 1)) if nca_c is not None else None
    plug_nca_p = (nca_p - _sum(sections["nca"], 2)) if nca_p is not None else None
    rows.append(_row("その他固定資産（未分類）/ Other Non-Current Assets",
                      plug_nca_c, plug_nca_p, False, "(算出差額)"))

    total_assets_c = (ca_c + nca_c) if (ca_c is not None and nca_c is not None) else None
    total_assets_p = (ca_p + nca_p) if (ca_p is not None and nca_p is not None) else None
    rows.append(_row("▶ 資産合計 / Total Assets",
                      total_assets_c, total_assets_p, False, "(合計)", is_subtotal=True))

    # ── 負債の部 ──────────────────────────────────────────
    for label, c, p in sections["cl"]:
        rows.append(_row(label, c, p, True, "(内訳)"))
    cl_c, cl_p = _official("CurrentLiabilities", "current"), _official("CurrentLiabilities", "prior")
    plug_cl_c = (cl_c - _sum(sections["cl"], 1)) if cl_c is not None else None
    plug_cl_p = (cl_p - _sum(sections["cl"], 2)) if cl_p is not None else None
    rows.append(_row("その他流動負債（未分類）/ Other Current Liabilities",
                      plug_cl_c, plug_cl_p, True, "(算出差額)"))
    rows.append(_row("▶ 流動負債合計 / Total Current Liabilities",
                      cl_c, cl_p, True, "(小計)", is_subtotal=True))

    for label, c, p in sections["ncl"]:
        rows.append(_row(label, c, p, True, "(内訳)"))
    ncl_c, ncl_p = _official("LongTermLiabilities", "current"), _official("LongTermLiabilities", "prior")
    plug_ncl_c = (ncl_c - _sum(sections["ncl"], 1)) if ncl_c is not None else None
    plug_ncl_p = (ncl_p - _sum(sections["ncl"], 2)) if ncl_p is not None else None
    rows.append(_row("その他固定負債（未分類）/ Other Non-Current Liabilities",
                      plug_ncl_c, plug_ncl_p, True, "(算出差額)"))

    total_liab_c = (cl_c + ncl_c) if (cl_c is not None and ncl_c is not None) else None
    total_liab_p = (cl_p + ncl_p) if (cl_p is not None and ncl_p is not None) else None
    rows.append(_row("▶ 負債合計 / Total Liabilities",
                      total_liab_c, total_liab_p, True, "(合計)", is_subtotal=True))

    # ── 資本の部 ──────────────────────────────────────────
    for label, c, p in sections["eq"]:
        rows.append(_row(label, c, p, False, "(内訳)"))
    eq_c, eq_p = _official("StockholdersEquity", "current"), _official("StockholdersEquity", "prior")
    plug_eq_c = (eq_c - _sum(sections["eq"], 1)) if eq_c is not None else None
    plug_eq_p = (eq_p - _sum(sections["eq"], 2)) if eq_p is not None else None
    rows.append(_row("その他資本（未分類）/ Other Equity Items",
                      plug_eq_c, plug_eq_p, False, "(算出差額)"))
    rows.append(_row("▶ 資本合計 / Total Stockholders' Equity",
                      eq_c, eq_p, False, "(合計)", is_subtotal=True))

    total_le_c = (total_liab_c + eq_c) if (total_liab_c is not None and eq_c is not None) else None
    total_le_p = (total_liab_p + eq_p) if (total_liab_p is not None and eq_p is not None) else None
    rows.append(_row("▶ 負債・資本合計 / Total Liabilities and Equity",
                      total_le_c, total_le_p, False, "(合計)", is_subtotal=True))

    return pd.DataFrame(rows)


def _bs_imbalance_note(df: pd.DataFrame) -> str:
    """Return a warning note when Assets total ≠ Liabilities+Equity total."""
    sub = df[df["_is_subtotal"] == True]
    if len(sub) < 2:
        return ""
    cur_cols = [c for c in df.columns if "当" in c and "USD" in c]
    if not cur_cols:
        return ""
    try:
        a  = float(sub.iloc[0][cur_cols[0]])
        le = float(sub.iloc[1][cur_cols[0]])
        if abs(a - le) > 1.0:   # > $1M 差異で表示
            return (
                "※ 資産合計と負債・資本合計が一致していません。"
                "株主資本（StockholdersEquity）は親会社株主に帰属する持分のみであり、"
                "非支配株主持分（Non-controlling Interests）が含まれていない可能性があります。"
            )
    except (TypeError, ValueError):
        pass
    return ""

# ─────────────────────────────────────────────────────────────────────────────
# Pandas Styler
# ─────────────────────────────────────────────────────────────────────────────

def _style_df(df: pd.DataFrame):
    display_cols = [c for c in df.columns if not c.startswith("_")]
    display_df   = df[display_cols].reset_index(drop=True).copy()

    # 数値列は数値型のまま保持する（st.dataframe が右寄せ表示するため）。
    # 欠損セルは Streamlit 側が薄いグレーの "None" プレースホルダで描画する。
    for col in display_cols:
        if "USD M" in col or "差額" in col:
            display_df[col] = pd.to_numeric(display_df[col], errors="coerce")

    def highlight_row(row):
        idx     = row.name
        is_sub  = df.iloc[idx].get("_is_subtotal", False)
        if is_sub:
            return pd.Series(["background-color: #BDD7EE; font-weight: bold;"] * len(display_cols),
                             index=display_cols)
        pct     = df.iloc[idx]["_pct"]
        is_cost = df.iloc[idx]["_is_cost"]
        cur     = df.iloc[idx].get("_cur")
        pri     = df.iloc[idx].get("_pri")
        if pct is None or (isinstance(pct, float) and pd.isna(pct)):
            bg = ""
        elif not is_cost and cur is not None and pri is not None and pri > 0 and cur < 0:
            bg = "background-color: #FFD2D2;"
        elif not is_cost and cur is not None and pri is not None and pri < 0 and cur > 0:
            bg = ""
        else:
            bad = (is_cost and pct > 0.20) or (not is_cost and pct < -0.20)
            bg  = "background-color: #FFD2D2;" if bad else ""
        return pd.Series([bg] * len(display_cols), index=display_cols)

    def fmt_usd(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return "N/A"
        return f"{round(v):,}"

    fmt       = {c: fmt_usd for c in display_cols if "USD M" in c or "差額" in c}
    right_cols = [c for c in display_cols if c not in ("項目 / Metric", "変化率 %")]

    return (
        display_df.style
        .apply(highlight_row, axis=1)
        .format(fmt, na_rep="N/A")
        .set_properties(**{"text-align": "right"}, subset=right_cols)
        .set_properties(**{"text-align": "right"}, subset=["変化率 %"])
        .set_properties(**{"text-align": "left"},  subset=["項目 / Metric"])
    )

# ─────────────────────────────────────────────────────────────────────────────
# Risk Calculations
# ─────────────────────────────────────────────────────────────────────────────

def compute_ratios(bs: dict) -> dict:
    def cv(k):
        v = bs.get(k, {}).get("current")
        return _m(v[0]) if v is not None else None
    def pv(k):
        v = bs.get(k, {}).get("prior")
        return _m(v[0]) if v is not None else None

    ca, ca_p = cv("CurrentAssets"),      pv("CurrentAssets")
    cl, cl_p = cv("CurrentLiabilities"), pv("CurrentLiabilities")
    cash, cash_p = cv("Cash"),           pv("Cash")
    nca  = cv("NonCurrentAssets")
    eq, eq_p = cv("StockholdersEquity"), pv("StockholdersEquity")

    total   = (ca or 0)   + (nca or 0)
    total_p = (ca_p or 0) + (pv("NonCurrentAssets") or 0)

    def _qoq(cur, pri): return (cur - pri) / abs(pri) if (cur is not None and pri not in (None, 0)) else None

    return {
        "current_ratio":   ca / cl       if (ca and cl and cl != 0) else None,
        "current_ratio_p": ca_p / cl_p   if (ca_p and cl_p and cl_p != 0) else None,
        "equity_ratio":    eq / total     if (eq is not None and total != 0) else None,
        "equity_ratio_p":  eq_p / total_p if (eq_p is not None and total_p != 0) else None,
        "cash_qoq":  _qoq(cash, cash_p),
        "cl_qoq":    _qoq(cl,   cl_p),
        "eq_qoq":    _qoq(eq,   eq_p),
        "cash_cur": cash, "cash_pri": cash_p,
        "cl_cur":   cl,   "cl_pri":   cl_p,
        "eq_cur":   eq,   "eq_pri":   eq_p,
    }


def risk_verdict(r: dict) -> tuple[str, str, str]:
    """Return (message, bg_hex, emoji) based on KPI threshold checks.
    Any single triggered alert results in at least a yellow warning.
    """
    alerts = []

    # ── 絶対水準チェック ──────────────────────────────────────────────────────
    cr = r.get("current_ratio")
    if cr is not None and cr < 1.0:
        alerts.append(f"⚠️ 流動比率 {cr:.2f} < 1.0（短期支払い能力不足）")

    er = r.get("equity_ratio")
    if er is not None and er < 0.30:
        alerts.append(f"⚠️ 自己資本比率 {er:.1%} < 30%（財務基盤が脆弱）")

    # ── QoQ変化率チェック ─────────────────────────────────────────────────────
    if r.get("cash_qoq") is not None and r["cash_qoq"] < -0.20:
        alerts.append(f"⚠️ 現金 QoQ {r['cash_qoq']:+.1%}（急減）")
    if r.get("cl_qoq") is not None and r["cl_qoq"] > 0.20:
        alerts.append(f"⚠️ 流動負債 QoQ {r['cl_qoq']:+.1%}（急増）")
    if r.get("eq_qoq") is not None and r["eq_qoq"] < -0.20:
        alerts.append(f"⚠️ 株主資本 QoQ {r['eq_qoq']:+.1%}（急減）")

    # ── データ不足 ─────────────────────────────────────────────────────────────
    all_none = all(r.get(k) is None for k in
                   ["current_ratio", "equity_ratio", "cash_qoq", "cl_qoq", "eq_qoq"])
    if all_none:
        return "— データ不足：判定不可 / Insufficient data.", "#F2F2F2", "⚪"

    # ── 結果 ───────────────────────────────────────────────────────────────────
    n = len(alerts)
    if n == 0:
        return "✅ 現時点で重大なリスクシグナルなし / No major risk signals.", "#D2FFD2", "🟢"

    detail = "  |  ".join(alerts)
    if n >= 3 or (r.get("cash_qoq") is not None and r["cash_qoq"] < -0.20
                  and r.get("cl_qoq") is not None and r["cl_qoq"] > 0.20):
        return f"🔴 重大なリスクシグナルを検知\n{detail}", "#FFD2D2", "🔴"
    if n >= 2:
        return f"🟡 要注意：複数の財務悪化シグナルを検知\n{detail}", "#FFFACD", "🟡"
    return f"🟡 軽微なリスクシグナルあり\n{detail}", "#FFFACD", "🟡"


_KPI_NOTES_MD = """
| 指標 | 見方・チェック理由 |
|------|------------------|
| **流動比率** | 短期の支払い能力を測定。**1.0未満**は1年以内の債務に対して現金化できる資産が不足＝黒字倒産リスクの警戒サイン。 |
| **自己資本比率** | 中長期の倒産リスク（企業の頑丈さ）を測定。市況変動が激しいエネルギーセクターでは**30%以上**が健全目安。 |
| **現金 QoQ** | 企業のリアルな体力ゲージ。QoQで**20%以上急減**している場合は、手元資金が急速に流出している警告シグナル。 |
| **流動負債 QoQ** | 目の前に迫る支払いの増減。現金が減っている局面でここが激増＝短期資金繰りの**デッドクロス**リスク。 |
| **株主資本 QoQ** | 基礎体力の増減。マイナスは本業赤字 or 身の丈に合わない配当・自社株買いで会社が細っているサイン。 |
"""

# ─────────────────────────────────────────────────────────────────────────────
# Excel Builder (3 tabs)
# ─────────────────────────────────────────────────────────────────────────────

_THIN   = Side(style="thin", color="BFBFBF")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

_RED_F  = PatternFill("solid", fgColor="FFD2D2")
_GRN_F  = PatternFill("solid", fgColor="D2FFD2")
_YLW_F  = PatternFill("solid", fgColor="FFFACD")
_BLU_F  = PatternFill("solid", fgColor="1F4E79")
_SUB_F  = PatternFill("solid", fgColor="2E75B6")
_SEC_F  = PatternFill("solid", fgColor="BDD7EE")
_GRY_F  = PatternFill("solid", fgColor="F2F2F2")

_W_BOLD = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
_D_BOLD = Font(name="Calibri", bold=True, color="1F4E79", size=11)
_NORM   = Font(name="Calibri", size=10)
_ITAL   = Font(name="Calibri", size=9,  italic=True, color="595959")
_C      = Alignment(horizontal="center", vertical="center", wrap_text=True)
_L      = Alignment(horizontal="left",   vertical="center", wrap_text=True)
_R      = Alignment(horizontal="right",  vertical="center")

FMT_USD   = '#,##0_);[Red](#,##0)'
FMT_PCT   = '0%;[Red]-0%'
FMT_RATIO = '0.00'


def _cw(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _hrow(ws, row, texts, fill=None):
    for col, t in enumerate(texts, 1):
        c = ws.cell(row=row, column=col, value=t)
        c.fill = fill or _BLU_F
        c.font = _W_BOLD
        c.alignment = _C
        c.border = _BORDER


def _sc(cell, fill=None, font=None, align=None, fmt=None):
    if fill:  cell.fill   = fill
    if font:  cell.font   = font
    if align: cell.alignment = align
    if fmt:   cell.number_format = fmt
    cell.border = _BORDER


def _mval(val):
    try:
        return float(val) / 1_000_000 if val is not None else None
    except Exception:
        return None


def _data_row(ws, row_num, label, cur, pri, is_bad_if_high: bool, formula: bool = True):
    """Write one data row; return whether alert was triggered."""
    ws.cell(row=row_num, column=1).fill = _GRY_F
    ws.cell(row=row_num, column=1).border = _BORDER

    lc = ws.cell(row=row_num, column=2, value=label); _sc(lc, font=_NORM, align=_L)
    cc = ws.cell(row=row_num, column=3, value=cur);   _sc(cc, align=_R, fmt=FMT_USD)
    dc = ws.cell(row=row_num, column=4, value=pri);   _sc(dc, align=_R, fmt=FMT_USD)
    if cur  is None: cc.value = "N/A"; cc.font = _ITAL
    if pri  is None: dc.value = "N/A"; dc.font = _ITAL

    ec = ws.cell(row=row_num, column=5)
    if cur is not None and pri is not None:
        ec.value = f"=C{row_num}-D{row_num}"; ec.number_format = FMT_USD
    else:
        ec.value = "N/A"; ec.font = _ITAL
    _sc(ec, align=_R)

    fc = ws.cell(row=row_num, column=6)
    if cur is not None and pri is not None:
        pct_raw = (cur - pri) / abs(pri) if pri != 0 else None
        pct_text = _pct_label(pct_raw, cur, pri, is_bad_if_high)
        fc.value = pct_text
        if pct_text in ("N/A", "黒字転換", "赤字転落"):
            fc.font = _ITAL
    else:
        fc.value = "N/A"; fc.font = _ITAL
    _sc(fc, align=_C)

    ws.row_dimensions[row_num].height = 18
    return False


def _build_fd_sheet(wb, company_name, ticker, cik, pl, bs, quarter_num: int = 1) -> dict:
    ws = wb.create_sheet("Financial Data")
    ws.views.sheetView[0].showGridLines = True
    _cw(ws, [3, 46, 18, 18, 18, 14])

    ws.merge_cells("A1:F1")
    t = ws["A1"]
    t.value = f"財務分析レポート | {company_name} ({ticker.upper()}) | Source: SEC EDGAR XBRL"
    t.fill = _BLU_F; t.font = Font(name="Calibri", bold=True, color="FFFFFF", size=13)
    t.alignment = _C; ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:F2")
    s = ws["A2"]
    s.value = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  CIK: {cik}  |  Amounts in USD Millions (M)"
    s.fill = _SUB_F; s.font = Font(name="Calibri", color="FFFFFF", size=9, italic=True)
    s.alignment = _C; ws.row_dimensions[2].height = 14

    # ── PL section
    row = 4
    ws.merge_cells(f"A{row}:F{row}")
    _qlabel = _Q_PERIOD_LABELS.get(quarter_num, "")
    c = ws.cell(row=row, column=1,
                value=f"📊  損益計算書（P&L） — 前年同期比（YoY）  [{_qlabel}]")
    c.fill = _SEC_F; c.font = _D_BOLD; c.alignment = _L; ws.row_dimensions[row].height = 22; row += 1

    pl_cd = pl_pd = "—"
    for d in pl.values():
        if d.get("current"): pl_cd = d["current"][1]; break
    for d in pl.values():
        if d.get("prior"):   pl_pd = d["prior"][1];   break

    _hrow(ws, row, ["", "項目 / Metric", f"当期\n({pl_cd})\n{_qlabel}\n[USD M]",
                    f"前期\n({pl_pd})\n{_qlabel}\n[USD M]", "差額 [USD M]", "変化率 %"])
    ws.row_dimensions[row].height = 40; row += 1

    pl_order = [
        ("Revenues",            "売上高 / Revenues",                          False),
        ("OperatingExpenses",   "営業費用 / Operating Expenses",              True),
        ("OperatingIncomeLoss", "営業利益 / Operating Income",                False),
        ("InterestExpense",     "営業外費用 / Total Other Expense, net",      False),
        ("IncomeLossBeforeTax", "税引前利益 / Income Before Tax",             False),
        ("NetIncomeLoss",       "純利益 / Net Income",                        False),
    ]
    bs_rows_map = {}
    for key, label, is_liab in pl_order:
        data = pl.get(key, {})
        cur  = _mval(data["current"][0]) if data.get("current") else None
        pri  = _mval(data["prior"][0])   if data.get("prior")   else None
        _data_row(ws, row, label, cur, pri, is_liab)
        row += 1

    # Non-GAAP note
    ws.cell(row=row, column=1).value = "★"; ws.cell(row=row, column=1).fill = _YLW_F
    _sc(ws.cell(row=row, column=1), align=_C)
    lc = ws.cell(row=row, column=2, value="在庫影響除き営業利益 / Operating Income ex-Inventory Adj. [Non-GAAP]")
    lc.font = Font(name="Calibri", size=10, italic=True, color="7F6000"); lc.alignment = _L; lc.border = _BORDER
    nc = ws.cell(row=row, column=3,
                 value="※ PARR等エネルギー企業は10-Q MD&Aの「Inventory Valuation Adjustment」を確認し手動で調整してください。")
    nc.font = Font(name="Calibri", size=8, italic=True, color="7F6000"); nc.alignment = _L
    ws.merge_cells(f"C{row}:F{row}"); nc.border = _BORDER
    ws.row_dimensions[row].height = 26; row += 2

    # ── BS section
    ws.merge_cells(f"A{row}:F{row}")
    c = ws.cell(row=row, column=1, value="🏦  貸借対照表（B/S） — 前四半期比（QoQ）")
    c.fill = _SEC_F; c.font = _D_BOLD; c.alignment = _L; ws.row_dimensions[row].height = 22; row += 1

    bs_cd = bs_pd = "—"
    for d in bs.values():
        if d.get("current"): bs_cd = d["current"][1]; break
    for d in bs.values():
        if d.get("prior"):   bs_pd = d["prior"][1];   break

    _hrow(ws, row, ["", "項目 / Metric", f"当四半期末\n({bs_cd})\n[USD M]",
                    f"前四半期末\n({bs_pd})\n[USD M]", "差額 [USD M]", "変化率 %"])
    ws.row_dimensions[row].height = 40; row += 1

    def _bsv(key, which):
        d = bs.get(key, {})
        return _mval(d[which][0]) if d.get(which) else None

    def _subtotal_row(label, cur_f, pri_f):
        nonlocal row
        ws.cell(row=row, column=1).fill = _SEC_F; ws.cell(row=row, column=1).border = _BORDER
        lc = ws.cell(row=row, column=2, value=label)
        lc.fill = _SEC_F; lc.font = _D_BOLD; lc.alignment = _L; lc.border = _BORDER
        cc = ws.cell(row=row, column=3, value=cur_f)
        cc.number_format = FMT_USD; _sc(cc, fill=_SEC_F, align=_R); cc.font = _D_BOLD
        dc = ws.cell(row=row, column=4, value=pri_f)
        dc.number_format = FMT_USD; _sc(dc, fill=_SEC_F, align=_R); dc.font = _D_BOLD
        ec = ws.cell(row=row, column=5, value=f"=C{row}-D{row}")
        ec.number_format = FMT_USD; _sc(ec, fill=_SEC_F, align=_R); ec.font = _D_BOLD
        fc = ws.cell(row=row, column=6); fc.value = ""; fc.fill = _SEC_F; fc.border = _BORDER
        ws.row_dimensions[row].height = 20; row += 1

    # ── Assets section ──────────────────────────────────────────────────────
    # Cash
    cash_c = _bsv("Cash", "current"); cash_p = _bsv("Cash", "prior")
    bs_rows_map["Cash"] = row
    _data_row(ws, row, "手元資金 / Cash & Equivalents", cash_c, cash_p, False); row += 1

    # Other Current Assets (computed from CurrentAssets - Cash)
    ca_c = _bsv("CurrentAssets", "current"); ca_p = _bsv("CurrentAssets", "prior")
    oca_c = (ca_c - cash_c) if (ca_c is not None and cash_c is not None) else None
    oca_p = (ca_p - cash_p) if (ca_p is not None and cash_p is not None) else None
    bs_rows_map["OCA"] = row
    _data_row(ws, row, "その他流動資産 / Other Current Assets", oca_c, oca_p, False); row += 1

    # Non-Current Assets
    nca_c = _bsv("NonCurrentAssets", "current"); nca_p = _bsv("NonCurrentAssets", "prior")
    bs_rows_map["NonCurrentAssets"] = row
    _data_row(ws, row, "固定資産 / Non-Current Assets", nca_c, nca_p, False); row += 1

    # Total Assets subtotal
    ta_rows = [bs_rows_map["Cash"], bs_rows_map["OCA"], bs_rows_map["NonCurrentAssets"]]
    ta_f_c = "+".join(f"C{r}" for r in ta_rows)
    ta_f_p = "+".join(f"D{r}" for r in ta_rows)
    bs_rows_map["TotalAssetsRow"] = row
    _subtotal_row("▶ 資産合計 / Total Assets", f"={ta_f_c}", f"={ta_f_p}")

    # ── Liabilities & Equity section ────────────────────────────────────────
    cl_c = _bsv("CurrentLiabilities", "current"); cl_p = _bsv("CurrentLiabilities", "prior")
    bs_rows_map["CurrentLiabilities"] = row
    _data_row(ws, row, "流動負債 / Current Liabilities", cl_c, cl_p, True); row += 1

    ltl_c = _bsv("LongTermLiabilities", "current"); ltl_p = _bsv("LongTermLiabilities", "prior")
    bs_rows_map["LongTermLiabilities"] = row
    _data_row(ws, row, "長期負債 / LT Liabilities", ltl_c, ltl_p, True); row += 1

    eq_c = _bsv("StockholdersEquity", "current"); eq_p = _bsv("StockholdersEquity", "prior")
    bs_rows_map["StockholdersEquity"] = row
    _data_row(ws, row, "株主資本 / Stockholders' Equity", eq_c, eq_p, False); row += 1

    # Total L+E subtotal
    le_rows = [bs_rows_map["CurrentLiabilities"], bs_rows_map["LongTermLiabilities"], bs_rows_map["StockholdersEquity"]]
    le_f_c = "+".join(f"C{r}" for r in le_rows)
    le_f_p = "+".join(f"D{r}" for r in le_rows)
    bs_rows_map["TotalLERow"] = row
    _subtotal_row("▶ 負債・資本合計 / Total L+E", f"={le_f_c}", f"={le_f_p}")
    row += 1

    ws.merge_cells(f"A{row}:F{row}")
    note = ws.cell(row=row, column=1,
                   value="※ 資産合計≠負債・資本合計の場合、株主資本は親会社帰属分のみで非支配株主持分が未計上の可能性があります。")
    note.font = _ITAL; note.fill = _GRY_F; note.alignment = _L; ws.row_dimensions[row].height = 14; row += 1

    ws.merge_cells(f"A{row}:F{row}")
    leg = ws.cell(row=row, column=1,
                  value="[凡例] 変化率 N/A=計算不可（データ欠損 or 両期マイナス）  ★=Non-GAAP要手動確認  | 金額はUSD百万単位")
    leg.font = _ITAL; leg.fill = _GRY_F; leg.alignment = _L; ws.row_dimensions[row].height = 14

    return bs_rows_map


def _build_dashboard_sheet(wb, company_name, ticker, cik, pl, bs, bs_rows_map):
    wd = wb.create_sheet("Dashboard", 0)
    wd.views.sheetView[0].showGridLines = True
    _cw(wd, [3, 34, 18, 18, 14, 14, 60])

    wd.merge_cells("A1:G1")
    t = wd["A1"]
    t.value = f"財務ヘルスダッシュボード | {company_name} ({ticker.upper()})"
    t.fill = _BLU_F; t.font = Font(name="Calibri", bold=True, color="FFFFFF", size=15)
    t.alignment = _C; wd.row_dimensions[1].height = 34

    wd.merge_cells("A2:G2")
    s = wd["A2"]
    s.value = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  CIK: {cik}  |  Source: SEC EDGAR XBRL"
    s.fill = _SUB_F; s.font = Font(name="Calibri", color="FFFFFF", size=9, italic=True)
    s.alignment = _C; wd.row_dimensions[2].height = 14

    dr = 4
    wd.merge_cells(f"A{dr}:G{dr}")
    c = wd.cell(row=dr, column=1, value="📐  主要財務比率（Financial Dataシートと数式連携）")
    c.fill = _SEC_F; c.font = _D_BOLD; c.alignment = _L; wd.row_dimensions[dr].height = 22; dr += 1
    _hrow(wd, dr, ["", "指標 / Ratio", "当期", "前期", "変化", "判定", "見方・チェック理由"]); wd.row_dimensions[dr].height = 22; dr += 1

    fd = "'Financial Data'"
    ca_r = bs_rows_map.get("CurrentAssets"); cl_r = bs_rows_map.get("CurrentLiabilities")
    eq_r = bs_rows_map.get("StockholdersEquity"); nca_r = bs_rows_map.get("NonCurrentAssets")

    _RATIO_NOTES = {
        "流動比率 / Current Ratio":    "短期の支払い能力。1.0未満＝黒字倒産リスク警戒。1.5以上が健全目安。",
        "自己資本比率 / Equity Ratio": "中長期の倒産リスク（企業の頑丈さ）。エネルギーセクターは30%以上が健全目安。",
    }

    def _ratio_row(label, cur_f, pri_f, fmt, status_f, sfill):
        nonlocal dr
        wd.cell(row=dr, column=1).border = _BORDER
        lc = wd.cell(row=dr, column=2, value=label); _sc(lc, font=_NORM, align=_L)
        cc = wd.cell(row=dr, column=3, value=cur_f); cc.number_format = fmt; _sc(cc, align=_R)
        dc = wd.cell(row=dr, column=4, value=pri_f); dc.number_format = fmt; _sc(dc, align=_R)
        ec = wd.cell(row=dr, column=5, value=f"=C{dr}-D{dr}"); ec.number_format = fmt; _sc(ec, align=_R)
        sc = wd.cell(row=dr, column=6, value=status_f); _sc(sc, fill=sfill, align=_C)
        nc = wd.cell(row=dr, column=7, value=_RATIO_NOTES.get(label, ""))
        nc.font = Font(name="Calibri", size=9, italic=True, color="595959")
        nc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        nc.border = _BORDER
        wd.row_dimensions[dr].height = 28; dr += 1

    ratios = compute_ratios(bs)
    cr_v   = ratios["current_ratio"];   cr_p   = ratios["current_ratio_p"]
    er_v   = ratios["equity_ratio"];    er_p   = ratios["equity_ratio_p"]
    cr_fill = _RED_F if (cr_v and cr_v < 1) else (_YLW_F if (cr_v and cr_v < 1.5) else _GRN_F)
    er_fill = _RED_F if (er_v is not None and er_v < 0.1) else (_YLW_F if (er_v is not None and er_v < 0.3) else _GRN_F)

    def _cr_status(v): return '⚠️ <1.0 危険' if (v and v < 1) else ('⚡ 注意' if (v and v < 1.5) else '✅ 良好')
    def _er_status(v): return '⚠️ <10% 危険' if (v is not None and v < 0.1) else ('⚡ <30% 低水準' if (v is not None and v < 0.3) else '✅ 良好')

    if ca_r and cl_r:
        _ratio_row("流動比率 / Current Ratio",
                   f"={fd}!C{ca_r}/{fd}!C{cl_r}", f"={fd}!D{ca_r}/{fd}!D{cl_r}", FMT_RATIO,
                   f'=IF(C{dr-1}<1,"⚠️ <1.0 危険",IF(C{dr-1}<1.5,"⚡ 注意","✅ 良好"))', cr_fill)
    else:
        _ratio_row("流動比率 / Current Ratio",
                   cr_v, cr_p, FMT_RATIO, _cr_status(cr_v), cr_fill)
    if ca_r and nca_r and eq_r:
        ta_c = f"({fd}!C{ca_r}+{fd}!C{nca_r})"; ta_p = f"({fd}!D{ca_r}+{fd}!D{nca_r})"
        _ratio_row("自己資本比率 / Equity Ratio",
                   f"={fd}!C{eq_r}/{ta_c}", f"={fd}!D{eq_r}/{ta_p}", FMT_PCT,
                   f'=IF(C{dr-1}<0.1,"⚠️ <10% 危険",IF(C{dr-1}<0.3,"⚡ <30% 低水準","✅ 良好"))', er_fill)
    else:
        _ratio_row("自己資本比率 / Equity Ratio",
                   er_v, er_p, FMT_PCT, _er_status(er_v), er_fill)

    dr += 2
    wd.merge_cells(f"A{dr}:G{dr}")
    rk = wd.cell(row=dr, column=1, value="📉  主要BS項目 QoQ変化チェック")
    rk.fill = PatternFill("solid", fgColor="385723")
    rk.font = Font(name="Calibri", bold=True, color="FFFFFF", size=12)
    rk.alignment = _L; wd.row_dimensions[dr].height = 26; dr += 1
    _hrow(wd, dr, ["", "チェック指標", "実績値（当期→前期）", "判定基準", "状態", "", "見方・チェック理由"]); wd.row_dimensions[dr].height = 20; dr += 1

    r = ratios

    _KPI_NOTES = {
        "手元資金 QoQ変化": "企業のリアルな体力ゲージ。QoQ -20%以上の急減は、手元資金が急速に流出している警告シグナル。",
        "流動負債 QoQ変化": "目の前に迫る支払いの増減。現金減少局面でここが激増する場合、短期資金繰りのデッドクロスリスク。",
        "株主資本 QoQ変化": "基礎体力の増減。マイナスは本業赤字 or 身の丈に合わない配当・自社株買いで会社が細っているサイン。",
    }

    def _risk_row_ex(label, pct, cur_v, pri_v, threshold, bad_cond):
        nonlocal dr
        is_bad = bad_cond(pct) if pct is not None else False
        wd.cell(row=dr, column=1).border = _BORDER
        lc = wd.cell(row=dr, column=2, value=label); _sc(lc, font=_NORM, align=_L)
        result_txt = f"{pct:+.1%}  ({pri_v:,.0f}M \u2192 {cur_v:,.0f}M)" if (pct is not None and cur_v is not None and pri_v is not None) else "N/A"
        rc = wd.cell(row=dr, column=3, value=result_txt); _sc(rc, font=_NORM, align=_C)
        tc = wd.cell(row=dr, column=4, value=threshold); _sc(tc, font=_ITAL, align=_C)
        sc = wd.cell(row=dr, column=5, value="⚠️ 要注意" if is_bad else "✅ 正常")
        _sc(sc, fill=(_RED_F if is_bad else _GRN_F), align=_C)
        wd.cell(row=dr, column=6).border = _BORDER
        nc = wd.cell(row=dr, column=7, value=_KPI_NOTES.get(label, ""))
        nc.font = Font(name="Calibri", size=9, italic=True, color="595959")
        nc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        nc.border = _BORDER
        wd.row_dimensions[dr].height = 30; dr += 1

    _risk_row_ex("手元資金 QoQ変化", r["cash_qoq"], r["cash_cur"], r["cash_pri"], "< -20% で警告", lambda p: p < -0.20)
    _risk_row_ex("流動負債 QoQ変化", r["cl_qoq"],   r["cl_cur"],   r["cl_pri"],   "> +20% で警告", lambda p: p > 0.20)
    _risk_row_ex("株主資本 QoQ変化", r["eq_qoq"],   r["eq_cur"],   r["eq_pri"],   "< -20% で警告", lambda p: p < -0.20)

    wd.merge_cells(f"A{dr}:G{dr}")
    dis = wd.cell(row=dr, column=1,
                  value="※ 免責：本ツールはSEC EDGAR公開データに基づく自動分析です。投資判断は必ず一次情報（10-Q/10-K）と専門家意見でご確認ください。")
    dis.font = _ITAL; dis.fill = _GRY_F; dis.alignment = _L; wd.row_dimensions[dr].height = 14


def _build_mda_sheet(wb, mda_text: str, company_name: str, period: str):
    wm = wb.create_sheet("MD&A_Text")
    wm.views.sheetView[0].showGridLines = True
    wm.column_dimensions["A"].width = 4
    wm.column_dimensions["B"].width = 120

    wm.merge_cells("A1:B1")
    t = wm["A1"]
    t.value = f"Management's Discussion and Analysis (MD&A)  |  {company_name}  |  Period: {period}"
    t.fill = _BLU_F; t.font = Font(name="Calibri", bold=True, color="FFFFFF", size=13)
    t.alignment = _C; wm.row_dimensions[1].height = 28

    wm.merge_cells("A2:B2")
    s = wm["A2"]
    s.value = "※ 以下のテキストをClaude/ChatGPT等のAIにそのままコピー&ペーストしてMD&Aの要約・分析を依頼できます。"
    s.fill = _YLW_F; s.font = Font(name="Calibri", size=10, italic=True, color="7F6000"); s.alignment = _L
    wm.row_dimensions[2].height = 18

    # Split text into chunks of 30000 chars (Excel cell limit is 32767)
    CHUNK = 30000
    chunks = [mda_text[i:i+CHUNK] for i in range(0, len(mda_text), CHUNK)] if mda_text else ["（テキストなし）"]
    for idx, chunk in enumerate(chunks, start=3):
        wm.merge_cells(f"A{idx}:B{idx}")
        c = wm.cell(row=idx, column=1, value=chunk)
        c.font = Font(name="Calibri", size=10)
        c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        c.border = _BORDER
        wm.row_dimensions[idx].height = max(60, min(400, len(chunk) // 80 * 15))


def build_excel(company_name: str, ticker: str, cik: str,
                pl: dict, bs: dict, mda_text: str, period: str,
                quarter_num: int = 1) -> bytes:
    wb = Workbook()
    del wb["Sheet"]
    bs_rows_map = _build_fd_sheet(wb, company_name, ticker, cik, pl, bs, quarter_num)
    _build_dashboard_sheet(wb, company_name, ticker, cik, pl, bs, bs_rows_map)
    _build_mda_sheet(wb, mda_text, company_name, period)
    wb.move_sheet("Dashboard", offset=-len(wb.sheetnames) + 1)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ─────────────────────────────────────────────────────────────────────────────
# Demo Data
# ─────────────────────────────────────────────────────────────────────────────

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

DEMO_MDA = """\
ITEM 2. MANAGEMENT'S DISCUSSION AND ANALYSIS OF FINANCIAL CONDITION AND RESULTS OF OPERATIONS

Overview

The following discussion and analysis should be read in conjunction with our unaudited condensed consolidated financial statements and the related notes included elsewhere in this Quarterly Report on Form 10-Q.

We are an independent energy company focused on acquiring and developing value-added downstream energy businesses. Our primary operations consist of refining and distribution of petroleum products, retail fuel sales, and logistics operations.

Results of Operations — Three Months Ended September 30, 2024

Revenues decreased $300.0 million, or 12.2%, to $2,150.0 million for the three months ended September 30, 2024, compared to $2,450.0 million for the same period in 2023. The decrease was primarily attributable to lower refined product prices driven by declining crude oil markets and tightened crack spreads across our refining segments.

Operating expenses of $2,050.0 million for the three months ended September 30, 2024 decreased by $150.0 million compared to the prior year period, driven by lower feedstock costs partially offsetting volume declines.

Operating income of $100.0 million decreased by $150.0 million, or 60.0%, compared to operating income of $250.0 million in the prior year period. The decrease reflects the impact of inventory valuation adjustments of approximately $(85) million related to lower crude oil prices during the quarter. Excluding inventory valuation adjustments, adjusted operating income would have been approximately $185 million.

Liquidity and Capital Resources

Cash and cash equivalents decreased by $140.0 million to $180.0 million as of September 30, 2024, from $320.0 million as of June 30, 2024, primarily due to debt repayments and capital expenditures. We had $950.0 million in current liabilities as of September 30, 2024.

[DEMO DATA — このテキストはデモ用サンプルです。実際のSEC取得時は英文の10-Q/10-K MD&Aテキストがここに表示されます。]
"""

# ─────────────────────────────────────────────────────────────────────────────
# Streamlit UI
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="SEC財務分析ツール",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""<style>
#MainMenu {visibility: hidden;}
header {visibility: hidden;}
footer {visibility: hidden;}
[data-testid="stDeployButton"] {display: none;}
[data-testid="stToolbar"] {display: none;}
</style>""", unsafe_allow_html=True)

st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #F7F9FC; }
h1 { color: #1F4E79; }
h2 { color: #2E75B6; font-size:1.1rem; margin-top:1.2rem; margin-bottom:0.3rem; }
[data-testid="metric-container"] {
    background:#fff; border-radius:10px;
    border:1px solid #D0E4F4; padding:10px 14px;
    box-shadow:0 1px 4px rgba(0,0,0,.07);
}
.alert-box { padding:14px 18px; border-radius:10px; font-size:1.05rem;
             font-weight:600; margin-bottom:10px; }
.alert-red    { background:#FFD2D2; color:#8B0000; border-left:5px solid #C00000; }
.alert-yellow { background:#FFFACD; color:#7F6000; border-left:5px solid #FFC000; }
.alert-green  { background:#D2FFD2; color:#1E6B2E; border-left:5px solid #00B050; }
.alert-grey   { background:#F2F2F2; color:#595959; border-left:5px solid #BFBFBF; }
[data-testid="stDownloadButton"] button {
    background:#1F4E79 !important; color:#fff !important;
    font-weight:700 !important; border-radius:8px !important;
    padding:10px 24px !important; font-size:1rem !important; border:none !important;
}
[data-testid="stDownloadButton"] button:hover { background:#2E75B6 !important; }
</style>
""", unsafe_allow_html=True)

# Session state init
for k in ("filings", "last_ticker", "cik", "company_name", "facts", "search_results"):
    if k not in st.session_state:
        st.session_state[k] = None

st.markdown("# 📊 Par Pacific Holdings — SEC EDGAR 財務分析ツール")
st.caption(
    "※ 本ツールはPar Pacific Holdingsを基準として設計されています。"
    "他社のP/Lは売上・費用の計上区分や勘定科目の定義が異なる場合があり、"
    "一部項目が欠損またはズレが生じる可能性があります。他社データは原本データでバックチェックください。"
)
st.markdown("""
米国上場企業の財務情報を SEC EDGAR から自動取得し、損益計算書・貸借対照表の推移、
主要KPI（流動比率・自己資本比率・現金増減など）、バリュエーション指標（PER・PBR）、
および株価と S&P500 のパフォーマンス比較をワンストップで確認できる財務分析ツールです。
対象レポートの原本ファイリングへのリンクも提供しており、数値のバックチェックが可能です。

**使い方**: 会社名（英語）を入力して検索 → ティッカーと対象決算期を選択 → 「財務分析を実行」
""")
st.markdown("---")

# ── Step 1: Company name search ────────────────────────────────────────────
col_search, col_btn = st.columns([3, 1.4])
with col_search:
    company_query = st.text_input("会社名を入力（英語）", value="Par Pacific", placeholder="例: Par Pacific, ExxonMobil, Tesla")
with col_btn:
    st.markdown("<br>", unsafe_allow_html=True)
    search_btn = st.button("🔍 企業を検索", use_container_width=True)

# Search results as selectbox
if search_btn or st.session_state.get("search_results"):
    if search_btn:
        candidates = _search_tickers(company_query)
        st.session_state.search_results = candidates
    else:
        candidates = st.session_state.get("search_results", [])

    if candidates:
        options = [f"{c['ticker']} — {c['name']}" for c in candidates]
        # Add manual entry option
        options.append("✏️ 手動でティッカーを入力")
        col_sel, col_fetch = st.columns([3, 1.4])
        with col_sel:
            chosen = st.selectbox("企業を選択してください", options)
        with col_fetch:
            st.markdown("<br>", unsafe_allow_html=True)
            fetch_btn = st.button("📋 決算期リストを取得", use_container_width=True)

        if chosen == "✏️ 手動でティッカーを入力":
            ticker_input = st.text_input("ティッカーシンボルを直接入力", value="PARR")
        else:
            ticker_input = chosen.split(" — ")[0]
    else:
        st.warning("候補が見つかりませんでした。ティッカーを直接入力してください。")
        col_t, col_f = st.columns([3, 1.4])
        with col_t:
            ticker_input = st.text_input("ティッカーシンボル（例: PARR, XOM, TSLA）", value="PARR")
        with col_f:
            st.markdown("<br>", unsafe_allow_html=True)
            fetch_btn = st.button("📋 決算期リストを取得", use_container_width=True)
else:
    # Default: show search box without results yet
    col_t, col_f = st.columns([3, 1.4])
    with col_t:
        ticker_input = st.text_input("ティッカーシンボル（例: PARR, XOM, TSLA）", value="PARR")
    with col_f:
        st.markdown("<br>", unsafe_allow_html=True)
        fetch_btn = st.button("📋 決算期リストを取得", use_container_width=True)

if fetch_btn:
    ticker = ticker_input.strip().upper()
    with st.spinner(f"{ticker} のCIKと決算期リストを取得中…"):
        cik, company_name = resolve_cik(ticker)
        if cik is None:
            st.error(f"ティッカー `{ticker}` が見つかりません。スペルを確認してください。")
        else:
            filings = get_filings_list(cik)
            if not filings:
                st.warning("決算期リストを取得できませんでした。ネットワーク環境を確認してください。")
            else:
                st.session_state.filings      = filings
                st.session_state.last_ticker  = ticker
                st.session_state.cik          = cik
                st.session_state.company_name = company_name
                st.session_state.facts        = None
                st.success(f"✅ {company_name} — {len(filings)} 件の決算期を取得しました。")

st.markdown("---")

# ── Step 2: Period selection + Analysis ────────────────────────────────────
if st.session_state.get("filings"):
    filings = st.session_state.filings
    labels  = [f["label"] for f in filings]
    col_sel, col_run = st.columns([4, 1.4])
    with col_sel:
        selected_label = st.selectbox("対象決算期（10-Q / 10-K）", labels,
                                      help="最新期がデフォルトで選択されています")
    with col_run:
        st.markdown("<br>", unsafe_allow_html=True)
        run_btn = st.button("🔍 財務分析を実行", type="primary", use_container_width=True)

    selected = next(f for f in filings if f["label"] == selected_label)

    if run_btn:
        ticker       = st.session_state.last_ticker
        cik          = st.session_state.cik
        company_name = st.session_state.company_name
        period       = selected["period"]

        with st.spinner("財務データとMD&Aテキストを取得中（10〜30秒かかる場合があります）…"):
            if st.session_state.facts is None:
                facts = fetch_facts(cik)
                st.session_state.facts = facts
            else:
                facts = st.session_state.facts

            if not facts:
                st.error("EDGAR XBRLデータを取得できませんでした。しばらく時間をおいて再試行してください。")
                st.stop()

            if not company_name or company_name == ticker:
                company_name = facts.get("entityName", ticker)
                st.session_state.company_name = company_name

            _q_num = _quarter_num(period, selected.get("fy_end", "1231")) \
                     if selected.get("form") != "10-K" else 4
            is_parr = (ticker.upper() == "PARR")
            pl  = extract_pl(facts, period, form_type=selected.get("form", "10-Q"),
                             quarter_num=_q_num, is_parr=is_parr)
            bs  = extract_bs(facts, period)
            mda = fetch_mda(cik, selected["accession"], selected["primary_doc"],
                            form_type=selected.get("form", "10-Q"))

        ratios  = compute_ratios(bs)
        verdict, v_hex, v_emoji = risk_verdict(ratios)

        cls = ("alert-red" if "FFD2D2" in v_hex else
               "alert-yellow" if "FACD" in v_hex else
               "alert-grey" if "F2F2" in v_hex else "alert-green")
        st.markdown(f'<div class="alert-box {cls}">{v_emoji} {verdict.replace(chr(10),"  |  ")}</div>',
                    unsafe_allow_html=True)
        st.markdown(f"### 🏢 {company_name}  `{ticker.upper()}`  —  期間: `{period}`  ({selected['form']})")

        k1,k2,k3,k4,k5 = st.columns(5)
        def _ps(v): return f"{v*100:.1f}%" if v is not None else "N/A"
        def _rs(v): return f"{v:.1f}" if v is not None else "N/A"
        k1.metric("流動比率", _rs(ratios["current_ratio"]))
        k2.metric("自己資本比率", _ps(ratios["equity_ratio"]))
        k3.metric("現金 QoQ", _ps(ratios["cash_qoq"]))
        k4.metric("流動負債 QoQ", _ps(ratios["cl_qoq"]))
        k5.metric("株主資本 QoQ", _ps(ratios["eq_qoq"]))
        with st.expander("📖 各指標の見方・チェック理由", expanded=False):
            st.markdown(_KPI_NOTES_MD)
        st.markdown("---")

        _ql = _Q_PERIOD_LABELS.get(_q_num, "")
        st.markdown(f"## 📊 損益計算書（P&L） — 前年同期比（YoY）  `{_ql}`")
        _pl_labels = PL_LABELS if is_parr else SIMPLE_PL_LABELS
        _pl_df = build_pl_df(pl, labels=_pl_labels)
        st.dataframe(_style_df(_pl_df), width="stretch",
                     height=min(430, 60 + 35 * len(_pl_df)))
        st.info("★ **Non-GAAP**: PARR等エネルギー企業は在庫影響除き営業利益をMD&Aで確認してください。", icon="ℹ️")
        if is_parr:
            st.caption(
                "※ 本ツールはPar Pacific Holdings（PARR）を基準として設計されています。"
                "他社では売上・費用の計上区分や勘定科目の定義が異なる場合があり、"
                "一部項目が欠損またはズレが生じる可能性があります。他社データは参考程度でご利用ください。"
            )
        else:
            st.caption(
                "※ 他社の場合、対象期に最も近いデータを持つXBRLタグを自動選択し、"
                "直接開示されていない項目は会計恒等式から導出しています"
                "（例: 売上総利益 = 売上高 − 売上原価、法人税等 = 税引前利益 − 純利益）。"
                "業種により計上区分・勘定科目の定義が異なるため、重要な判断の際は"
                "原本ファイリングでのバックチェックを推奨します。"
            )

        # ── 四半期別 売上高・純利益トレンド（過去3年・全企業対象） ──────────
        st.markdown("---")
        st.markdown("## 📈 四半期別 売上高・純利益の推移（過去3年）")
        with st.spinner("決算期リストから四半期トレンドを計算中…"):
            trend_df = extract_quarterly_trend(facts, filings, years=3)
        if trend_df is not None and not trend_df.empty and (
            trend_df["revenue"].notna().any() or trend_df["net_income"].notna().any()
        ):
            import plotly.graph_objects as go
            _tfig = go.Figure()
            _tfig.add_trace(go.Scatter(
                x=trend_df["quarter_label"], y=trend_df["revenue"].round(1),
                name="売上高 (USD M)", mode="lines+markers",
                line=dict(color="#2563EB", width=2), marker=dict(size=6),
                yaxis="y1",
                hovertemplate="%{x}<br>売上高: $%{y:,.1f}M<extra></extra>",
            ))
            _tfig.add_trace(go.Scatter(
                x=trend_df["quarter_label"], y=trend_df["net_income"].round(1),
                name="純利益 (USD M)", mode="lines+markers",
                line=dict(color="#16A34A", width=2, dash="dot"), marker=dict(size=6),
                yaxis="y2",
                hovertemplate="%{x}<br>純利益: $%{y:,.1f}M<extra></extra>",
            ))
            _tfig.update_layout(
                height=360,
                margin=dict(l=0, r=0, t=10, b=0),
                legend=dict(orientation="h", yanchor="bottom", y=1.01,
                            xanchor="right", x=1),
                hovermode="x unified",
                plot_bgcolor="white",
                paper_bgcolor="white",
                xaxis=dict(showgrid=False, zeroline=False, title="決算期"),
                yaxis=dict(
                    title="売上高 (USD M)", title_font_color="#2563EB",
                    tickfont=dict(color="#2563EB"), showgrid=True,
                    gridcolor="#F3F4F6", zeroline=False,
                ),
                yaxis2=dict(
                    title="純利益 (USD M)", title_font_color="#16A34A",
                    tickfont=dict(color="#16A34A"), overlaying="y", side="right",
                    showgrid=False, zeroline=True, zerolinecolor="#E5E7EB",
                ),
            )
            st.plotly_chart(_tfig, use_container_width=True)
            st.caption(
                "※ 10-Qは累積(YTD)値、10-Kは通期値として開示されるため、各四半期単独の値は"
                "前の四半期までの累積値を差し引いて算出（de-cumulation）しています。"
                "決算期リストの取得範囲（最大24期）を超える過去データが必要な場合、"
                "最も古い四半期の値は正しく算出できないことがあります。"
            )
        else:
            st.info("四半期トレンドデータを算出できませんでした。")

        st.markdown("## 🏦 貸借対照表（B/S） — 前四半期比（QoQ）")
        _bs_df = build_bs_df(bs)
        st.dataframe(_style_df(_bs_df), width="stretch", height=340)
        _bs_note = _bs_imbalance_note(_bs_df)
        if _bs_note:
            st.caption(_bs_note)

        # ── PARR限定: B/S 詳細版（Condensed Consolidated Balance Sheets相当） ──
        if is_parr:
            st.markdown("---")
            st.markdown("## 🏦 貸借対照表（B/S）— 詳細版（PARR限定・前四半期比較）")
            st.caption(
                "※ 実際のCONDENSED CONSOLIDATED BALANCE SHEETSに近い粒度・並び順で表示しています。"
                "個別科目に含まれない残差は各セクションの「その他（未分類）」行に計上されるため、"
                "各小計は上のB/Sサマリーと一致します。原本ファイリングでのバックチェックを推奨します。"
            )
            with st.spinner("B/S詳細版を集計中…"):
                bs_condensed_df = extract_bs_condensed(facts, bs, period)
            if bs_condensed_df is not None and not bs_condensed_df.empty:
                st.dataframe(_style_df(bs_condensed_df), width="stretch",
                             height=min(650, 60 + 35 * len(bs_condensed_df)))
            else:
                st.info("B/S詳細版データを算出できませんでした。")

        st.markdown("---")
        st.markdown("## 📝 Management's Discussion and Analysis (MD&A)")
        st.caption("以下のテキストをそのままClaude等のAIにコピー＆ペーストして要約・分析できます。")
        if mda:
            st.text_area("MD&A テキスト", value=mda, height=420, label_visibility="collapsed")
        else:
            st.warning("MD&Aテキストを取得できませんでした。")
        st.markdown("---")

        st.markdown("## 📥 分析結果をExcelでダウンロード")
        st.markdown("Dashboard / Financial Data / MD&A_Text の3シート構成、数式・ハイライト付き。")
        with st.spinner("Excelファイルを生成中…"):
            xls = build_excel(company_name, ticker, cik, pl, bs, mda or "", period, _q_num)
        fname = f"{ticker.upper()}_financial_{period}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.download_button("📥 分析結果をExcelでダウンロード", data=xls, file_name=fname,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
        st.caption(f"ファイル名: `{fname}`  |  シート: Dashboard / Financial Data / MD&A_Text")

        # ── SEC EDGAR 原本リンク ──────────────────────────────────
        st.markdown("---")
        st.markdown("## 📄 SEC EDGAR 原本ファイリング")
        _acc       = selected["accession"]
        _doc       = selected["primary_doc"]
        _form_type = selected.get("form", "10-Q")
        _cik_int   = int(cik)
        _acc_nodash = _acc.replace("-", "")
        _doc_url   = f"https://www.sec.gov/Archives/edgar/data/{_cik_int}/{_acc_nodash}/{_doc}"
        _search_url = (
            f"https://www.sec.gov/cgi-bin/browse-edgar"
            f"?action=getcompany&CIK={_cik_int}&type={_form_type}&dateb=&owner=include&count=10"
        )
        _col1, _col2 = st.columns(2)
        with _col1:
            st.link_button(
                f"📑 {_form_type} 原本を開く（HTML）",
                _doc_url,
                use_container_width=True,
            )
        with _col2:
            st.link_button(
                f"🔍 EDGAR {_form_type} 一覧",
                _search_url,
                use_container_width=True,
            )
        st.caption(
            f"Accession: `{_acc}`  |  File: `{_doc}`  |  "
            f"[直接URL]({_doc_url})"
        )

        st.markdown("---")
        st.markdown("## 📈 5年間株価推移 & バリュエーション指標")

        # Two columns: chart on left, valuation on right
        col_chart, col_val = st.columns([3, 2])

        with col_chart:
            st.markdown("### 📈 過去5年間の株価推移（vs S&P500）")
            with st.spinner("株価データを取得中..."):
                _stock, _sp500 = _get_chart_data(ticker)
            if _stock is not None and not _stock.empty:
                import plotly.graph_objects as go
                _fig = go.Figure()
                # 両系列とも「表示期間の開始日=0%」で指数化し、同一軸上の騰落率として
                # 表示する（絶対株価とS&P500の%騰落率を別軸で表示すると、下落局面での
                # 相対的な下落幅＝暴落率の比較ができないため、単位を統一）。
                _stock_pct = (_stock["Close"] / _stock["Close"].iloc[0] - 1) * 100
                _fig.add_trace(go.Scatter(
                    x=_stock.index,
                    y=_stock_pct.round(2),
                    name=f"{ticker.upper()} 騰落率",
                    line=dict(color="#2563EB", width=1.8),
                    customdata=_stock[["Close"]].values,
                    hovertemplate=("%{x|%Y-%m-%d}<br>" + ticker.upper() +
                                    ": %{y:+.1f}%（$%{customdata[0]:,.2f}）<extra></extra>"),
                ))
                # ── S&P500 騰落率（同一軸）※取得できた場合のみ ──
                if _sp500 is not None and not _sp500.empty:
                    _sp_pct = (_sp500["Close"] / _sp500["Close"].iloc[0] - 1) * 100
                    _fig.add_trace(go.Scatter(
                        x=_sp500.index,
                        y=_sp_pct.round(2),
                        name="S&P500 騰落率",
                        line=dict(color="#9CA3AF", width=1.5, dash="dot"),
                        hovertemplate="%{x|%Y-%m-%d}<br>S&P500: %{y:+.1f}%<extra></extra>",
                    ))
                _fig.update_layout(
                    height=340,
                    margin=dict(l=0, r=0, t=10, b=0),
                    legend=dict(orientation="h", yanchor="bottom", y=1.01,
                                xanchor="right", x=1),
                    hovermode="x unified",
                    plot_bgcolor="white",
                    paper_bgcolor="white",
                    xaxis=dict(showgrid=False, zeroline=False),
                    yaxis=dict(
                        title="騰落率 (%)　※期間開始日を0%として指数化",
                        showgrid=True,
                        gridcolor="#F3F4F6",
                        zeroline=True,
                        zerolinecolor="#9CA3AF",
                        zerolinewidth=1,
                        ticksuffix="%",
                    ),
                )
                st.plotly_chart(_fig, use_container_width=True)
                st.caption(
                    f"※ {ticker.upper()}とS&P500はいずれも表示期間の開始日を基準（0%）とした"
                    "騰落率で表示しています。同じ軸で比較できるため、下落局面での相対的な"
                    "下落幅（暴落率の比較）を直接読み取れます。実際の株価（$）はグラフに"
                    "マウスを合わせると表示されます。"
                )
            else:
                st.info("株価データを取得できませんでした（ネットワーク制限の可能性があります）")

        with col_val:
            st.markdown("### 💹 バリュエーション指標")
            val_info = _compute_valuation(ticker, facts, bs)

            pe       = val_info.get("pe")
            pe_label = val_info.get("pe_label", "PER（株価収益率）")
            pb       = val_info.get("pb")

            st.metric(pe_label, f"{pe:.1f}倍" if pe else "N/A")
            st.metric("PBR（株価純資産倍率）", f"{pb:.1f}倍" if pb else "N/A")

            with st.expander("📖 各指標の見方", expanded=False):
                st.markdown("""
| 指標 | 見方 |
|------|------|
| **PER（実績）** | 株価÷EPS（直近12ヶ月TTM）。**15〜20倍**が標準。30倍超は割高警戒。赤字時はN/A→予想PERで代替表示。 |
| **PER（予想）** | 株価÷来期予想EPS。赤字期など実績PERが取得できない場合に自動切替。 |
| **PBR** | 株価÷1株純資産。**1倍割れ**は理論上割安。エネルギーは1〜2倍が標準。 |
""")

else:
    st.markdown("""
    <div style="text-align:center;padding:60px 20px;color:#8B9BB4;">
    <div style="font-size:4rem;">📊</div>
    <div style="font-size:1.2rem;font-weight:600;margin:12px 0;">
        ① ティッカーを入力 → 「📋 決算期リストを取得」<br>
        ② 決算期を選択 → 「🔍 財務分析を実行」
    </div>
    <div style="font-size:.9rem;margin-top:8px;">対応例: PARR · XOM · CVX · TSLA · AAPL · MSFT · AMZN</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")
st.caption("Source: SEC EDGAR XBRL API (data.sec.gov) | 本ツールは情報提供目的のみです。投資判断には必ず一次情報をご確認ください。")
