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
    recent = data.get("filings", {}).get("recent", {})
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


def extract_mda(html: str, max_chars: int = 30000) -> str:
    # Remove table HTML — keep narrative prose only, skip financial data tables
    html = re.sub(r'<table[\s>].*?</table>', ' ', html, flags=re.IGNORECASE | re.DOTALL)
    text = _strip_html(html)
    text = re.sub(r"[ \t]{2,}", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)

    start_pats = [
        r"(?i)ITEM\s*2[\.\:\-\u2014\s]+MANAGEMENT[''\s]*S\s+DISCUSSION\s+AND\s+ANALYSIS",
        r"(?i)ITEM\s*2\b.{0,10}MANAGEMENT.{0,60}DISCUSSION",
        r"(?i)Management[''\s]*s\s+Discussion\s+and\s+Analysis\s+of\s+Financial\s+Condition",
        r"(?i)MANAGEMENT[''\s]*S\s+DISCUSSION\s+AND\s+ANALYSIS",
        r"(?i)ITEM\s*2\b",
    ]
    start = -1
    for pat in start_pats:
        m = re.search(pat, text)
        if m:
            start = m.start()
            break

    if start == -1:
        return text[:max_chars].strip()

    end_pats = [
        r"(?i)ITEM\s*3[\.\:\-\u2014\s]+QUANTITATIVE",
        r"(?i)ITEM\s*3[\.\:\-\u2014\s]+MARKET\s+RISK",
        r"(?i)\bITEM\s*3[\.\:\-\u2014\s]",
        r"(?i)\bITEM\s*4\b",
        r"(?i)PART\s+II\b",
    ]
    tail = text[start + 200:]
    end_offset = len(tail)
    for pat in end_pats:
        m = re.search(pat, tail)
        if m:
            end_offset = min(end_offset, m.start())

    mda = text[start: start + 200 + end_offset].strip()
    return mda[:max_chars]
def fetch_mda(cik: str, accession: str, primary_doc: str) -> str:
    cik_int    = int(cik)
    acc_nodash = accession.replace("-", "")
    url = EDGAR_ARCHIVES.format(cik_int=cik_int, acc_nodash=acc_nodash, doc=primary_doc)
    html = _get(url, as_text=True)
    if not html:
        return "（MD&Aテキストを取得できませんでした。SECのネットワーク制限またはファイル形式の問題の可能性があります。）"
    return extract_mda(html)

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


def _best_tag(facts: dict, candidates: list) -> tuple[str, list]:
    for tag in candidates:
        recs = _units(facts, tag)
        if recs:
            return tag, recs
    return "", []


def _filter_quarterly(records: list, lo: int = 75, hi: int = 110) -> list:
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


def _filter_annual(records: list) -> list:
    """Filter for full-year duration records (340-400 days)."""
    out = []
    for r in records:
        s, e = r.get("start", ""), r.get("end", "")
        if s and e:
            try:
                d = (datetime.strptime(e, "%Y-%m-%d") - datetime.strptime(s, "%Y-%m-%d")).days
                if 340 <= d <= 400:
                    out.append(r)
            except ValueError:
                pass
    return out


def _best_quarterly_tag(facts: dict, candidates: list) -> tuple[str, list, list]:
    """Return (tag, all_recs, quarterly_recs) for the first candidate that has quarterly data.

    Strategy:
    1. Scan every candidate; pick the first whose records contain at least one
       quarter-length duration (75-110 days, then relaxed 60-125 days as fallback).
    2. If nothing qualifies, return the first candidate with any records so the
       caller can still attempt the derivation fallback.
    """
    any_tag, any_recs = "", []
    relaxed_tag, relaxed_recs, relaxed_q = "", [], []

    for tag in candidates:
        recs = _units(facts, tag)
        if not recs:
            continue
        q = _filter_quarterly(recs, 75, 110)
        if q:
            return tag, recs, q
        # Widen to ±2 weeks to catch fiscal-week-calendar quarters (e.g. 13×7=91±14)
        q_wide = _filter_quarterly(recs, 60, 125)
        if q_wide and not relaxed_tag:
            relaxed_tag, relaxed_recs, relaxed_q = tag, recs, q_wide
        if not any_recs:
            any_tag, any_recs = tag, recs

    if relaxed_tag:
        return relaxed_tag, relaxed_recs, relaxed_q
    return any_tag, any_recs, []


def _best_annual_tag(facts: dict, candidates: list) -> tuple[str, list, list]:
    """Return (tag, all_recs, annual_recs) for the first candidate with full-year duration records."""
    any_tag, any_recs = "", []
    for tag in candidates:
        recs = _units(facts, tag)
        if not recs:
            continue
        a = _filter_annual(recs)
        if a:
            return tag, recs, a
        if not any_recs:
            any_tag, any_recs = tag, recs
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


def extract_pl(facts: dict, target_period: str | None = None, form_type: str = "10-Q") -> dict:
    target = datetime.strptime(target_period, "%Y-%m-%d") if target_period else None
    is_annual = (form_type == "10-K")
    result = {}
    for metric, candidates in PL_TAGS.items():
        # For 10-K use annual-duration records; for 10-Q use quarterly.
        if is_annual:
            tag, _recs, period_recs = _best_annual_tag(facts, candidates)
        else:
            tag, _recs, period_recs = _best_quarterly_tag(facts, candidates)
        period_recs = _dedup_latest(period_recs, 30)
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

    return result


def extract_bs(facts: dict, target_period: str | None = None) -> dict:
    target = datetime.strptime(target_period, "%Y-%m-%d") if target_period else None
    result = {}
    for metric, candidates in BS_TAGS.items():
        tag, recs = _best_tag(facts, candidates)
        instants = _dedup_latest(_filter_instant(recs), 12)
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
            return "黒字転換" if not is_bad_if_high else f"{pct:+.1%}"
        if pri > 0 and cur < 0:
            return "赤字転落" if not is_bad_if_high else f"{pct:+.1%}"
    return f"{pct:+.1%}"


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


def build_pl_df(pl: dict) -> pd.DataFrame:
    # ① Determine canonical period dates (most common across metrics)
    canon_cur = _canon_date(pl, "current")
    canon_pri = _canon_date(pl, "prior")
    cur_col   = f"当期 ({canon_cur})\n[USD M]"
    pri_col   = f"前期 ({canon_pri})\n[USD M]"

    rows = []
    for key, label in PL_LABELS.items():
        data = pl.get(key, {})
        # ② Only accept values whose period aligns with the canonical date
        cur = _safe_val(data, "current") if _period_ok(data, "current", canon_cur) else None
        pri = _safe_val(data, "prior")   if _period_ok(data, "prior",   canon_pri) else None
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
        return _safe_val(data, w) if _period_ok(data, w, canon_cur if w == "current" else canon_pri) else None

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
    display_df   = df[display_cols].reset_index(drop=True)

    # Convert None → NaN in numeric cols so na_rep="N/A" applies correctly
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
        return f"{v:,.1f}"

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
    cash_drop  = r["cash_qoq"] is not None and r["cash_qoq"] < -0.20
    cl_surge   = r["cl_qoq"]   is not None and r["cl_qoq"]   > 0.20
    eq_erosion = r["eq_qoq"]   is not None and r["eq_qoq"]   < -0.20
    if cash_drop and (cl_surge or eq_erosion):
        return "⚠️ 黒字倒産・資金繰り悪化の予兆あり\nCash shrinking while liabilities surge or equity erodes.", "#FFD2D2", "🔴"
    n = sum([cash_drop, cl_surge, eq_erosion])
    if n >= 2:
        return "⚡ 要注意：複数の財務悪化シグナルを検知\nMultiple deterioration signals.", "#FFFACD", "🟡"
    if n == 1:
        return "⚡ 軽微なリスクシグナルあり\nOne deterioration signal detected.", "#FFFACD", "🟡"
    if all(v is None for v in [r["cash_qoq"], r["cl_qoq"], r["eq_qoq"]]):
        return "— データ不足：判定不可\nInsufficient data.", "#F2F2F2", "⚪"
    return "✅ 現時点で重大なリスクシグナルなし\nNo major risk signals detected.", "#D2FFD2", "🟢"


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
FMT_PCT   = '0.0%;[Red]-0.0%'
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
    if cur is not None and pri is not None and pri != 0:
        fc.value = f"=IF(D{row_num}=0,\"\",(C{row_num}-D{row_num})/D{row_num})"
        fc.number_format = FMT_PCT
    else:
        fc.value = "N/A"; fc.font = _ITAL
    _sc(fc, align=_C)

    ws.row_dimensions[row_num].height = 18
    return False


def _build_fd_sheet(wb, company_name, ticker, cik, pl, bs) -> dict:
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
    c = ws.cell(row=row, column=1, value="📊  損益計算書（P&L） — 前年同期比（YoY）")
    c.fill = _SEC_F; c.font = _D_BOLD; c.alignment = _L; ws.row_dimensions[row].height = 22; row += 1

    pl_cd = pl_pd = "—"
    for d in pl.values():
        if d.get("current"): pl_cd = d["current"][1]; break
    for d in pl.values():
        if d.get("prior"):   pl_pd = d["prior"][1];   break

    _hrow(ws, row, ["", "項目 / Metric", f"当期\n({pl_cd})\n[USD M]",
                    f"前期\n({pl_pd})\n[USD M]", "差額 [USD M]", "変化率 %"])
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
                pl: dict, bs: dict, mda_text: str, period: str) -> bytes:
    wb = Workbook()
    del wb["Sheet"]
    bs_rows_map = _build_fd_sheet(wb, company_name, ticker, cik, pl, bs)
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
for k in ("filings", "last_ticker", "cik", "company_name", "facts"):
    if k not in st.session_state:
        st.session_state[k] = None

st.markdown("# 📊 SEC EDGAR 財務分析ツール")
st.markdown("米国上場企業のティッカーと対象決算期を選択して「財務分析を実行」してください。")
st.markdown("---")

# ── Step 1: Ticker + period fetch ──────────────────────────────────────────
col_t, col_f, col_d = st.columns([3, 1.4, 1.4])
with col_t:
    ticker_input = st.text_input("ティッカーシンボル（例: PARR, XOM, TSLA）",
                                 value="PARR", max_chars=10)
with col_f:
    st.markdown("<br>", unsafe_allow_html=True)
    fetch_btn = st.button("📋 決算期リストを取得", use_container_width=True)
with col_d:
    st.markdown("<br>", unsafe_allow_html=True)
    demo_btn = st.button("🧪 デモデータ", use_container_width=True,
                         help="ネットワーク不要のサンプルデータで動作確認")

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
if demo_btn:
    # --- Demo mode ---
    pl, bs = demo_pl_bs()
    ratios = compute_ratios(bs)
    verdict, v_hex, v_emoji = risk_verdict(ratios)
    mda_text  = DEMO_MDA
    company_name = "PARR Inc. [DEMO DATA]"
    ticker    = "PARR"
    cik       = KNOWN_CIKS["PARR"]
    period    = "2024-09-30"

    cls = ("alert-red" if "FFD2D2" in v_hex else
           "alert-yellow" if "FACD" in v_hex else
           "alert-grey" if "F2F2" in v_hex else "alert-green")
    st.markdown(f'<div class="alert-box {cls}">{v_emoji} {verdict.replace(chr(10),"  |  ")}</div>',
                unsafe_allow_html=True)
    st.markdown(f"### 🏢 {company_name}  `{ticker}`  —  期間: `{period}`")

    k1,k2,k3,k4,k5 = st.columns(5)
    def _ps(v): return f"{v:.1%}" if v is not None else "N/A"
    def _rs(v): return f"{v:.2f}" if v is not None else "N/A"
    k1.metric("流動比率", _rs(ratios["current_ratio"]))
    k2.metric("自己資本比率", _ps(ratios["equity_ratio"]))
    k3.metric("現金 QoQ", _ps(ratios["cash_qoq"]))
    k4.metric("流動負債 QoQ", _ps(ratios["cl_qoq"]))
    k5.metric("株主資本 QoQ", _ps(ratios["eq_qoq"]))
    with st.expander("📖 各指標の見方・チェック理由", expanded=False):
        st.markdown(_KPI_NOTES_MD)
    st.markdown("---")

    st.markdown("## 📊 損益計算書（P&L） — 前年同期比（YoY）")
    st.dataframe(_style_df(build_pl_df(pl)), width="stretch", height=270)
    st.info("★ **Non-GAAP**: PARR等エネルギー企業は在庫影響除き営業利益をMD&Aで確認してください。", icon="ℹ️")

    st.markdown("## 🏦 貸借対照表（B/S） — 前四半期比（QoQ）")
    _bs_df = build_bs_df(bs)
    st.dataframe(_style_df(_bs_df), width="stretch", height=340)
    _bs_note = _bs_imbalance_note(_bs_df)
    if _bs_note:
        st.caption(_bs_note)

    st.markdown("## 📝 Management's Discussion and Analysis (MD&A)")
    st.caption("以下のテキストをそのままClaude等のAIにコピー＆ペーストして要約・分析できます。")
    st.text_area("MD&A テキスト", value=mda_text, height=400, label_visibility="collapsed")
    st.markdown("---")

    st.markdown("## 📥 分析結果をExcelでダウンロード")
    with st.spinner("Excelファイルを生成中…"):
        xls = build_excel(company_name, ticker, cik, pl, bs, mda_text, period)
    fname = f"{ticker}_financial_{period}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    st.download_button("📥 分析結果をExcelでダウンロード", data=xls, file_name=fname,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)
    st.caption(f"ファイル名: `{fname}`  |  シート: Dashboard / Financial Data / MD&A_Text")

elif st.session_state.get("filings"):
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

            pl  = extract_pl(facts, period, form_type=selected.get("form", "10-Q"))
            bs  = extract_bs(facts, period)
            mda = fetch_mda(cik, selected["accession"], selected["primary_doc"])

        ratios  = compute_ratios(bs)
        verdict, v_hex, v_emoji = risk_verdict(ratios)

        cls = ("alert-red" if "FFD2D2" in v_hex else
               "alert-yellow" if "FACD" in v_hex else
               "alert-grey" if "F2F2" in v_hex else "alert-green")
        st.markdown(f'<div class="alert-box {cls}">{v_emoji} {verdict.replace(chr(10),"  |  ")}</div>',
                    unsafe_allow_html=True)
        st.markdown(f"### 🏢 {company_name}  `{ticker.upper()}`  —  期間: `{period}`  ({selected['form']})")

        k1,k2,k3,k4,k5 = st.columns(5)
        def _ps(v): return f"{v:.1%}" if v is not None else "N/A"
        def _rs(v): return f"{v:.2f}" if v is not None else "N/A"
        k1.metric("流動比率", _rs(ratios["current_ratio"]))
        k2.metric("自己資本比率", _ps(ratios["equity_ratio"]))
        k3.metric("現金 QoQ", _ps(ratios["cash_qoq"]))
        k4.metric("流動負債 QoQ", _ps(ratios["cl_qoq"]))
        k5.metric("株主資本 QoQ", _ps(ratios["eq_qoq"]))
        with st.expander("📖 各指標の見方・チェック理由", expanded=False):
            st.markdown(_KPI_NOTES_MD)
        st.markdown("---")

        st.markdown("## 📊 損益計算書（P&L） — 前年同期比（YoY）")
        st.dataframe(_style_df(build_pl_df(pl)), width="stretch", height=270)
        st.info("★ **Non-GAAP**: PARR等エネルギー企業は在庫影響除き営業利益をMD&Aで確認してください。", icon="ℹ️")

        st.markdown("## 🏦 貸借対照表（B/S） — 前四半期比（QoQ）")
        _bs_df = build_bs_df(bs)
    st.dataframe(_style_df(_bs_df), width="stretch", height=340)
    _bs_note = _bs_imbalance_note(_bs_df)
    if _bs_note:
        st.caption(_bs_note)

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
            xls = build_excel(company_name, ticker, cik, pl, bs, mda or "", period)
        fname = f"{ticker.upper()}_financial_{period}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.download_button("📥 分析結果をExcelでダウンロード", data=xls, file_name=fname,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
        st.caption(f"ファイル名: `{fname}`  |  シート: Dashboard / Financial Data / MD&A_Text")

else:
    st.markdown("""
    <div style="text-align:center;padding:60px 20px;color:#8B9BB4;">
    <div style="font-size:4rem;">📊</div>
    <div style="font-size:1.2rem;font-weight:600;margin:12px 0;">
        ① ティッカーを入力 → 「📋 決算期リストを取得」<br>
        ② 決算期を選択 → 「🔍 財務分析を実行」
    </div>
    <div style="font-size:.9rem;margin-top:8px;">対応例: PARR · XOM · CVX · TSLA · AAPL · MSFT · AMZN</div>
    <div style="font-size:.85rem;margin-top:6px;color:#AAB8C8;">
        ネットワーク不要の動作確認は「🧪 デモデータ」ボタン
    </div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")
st.caption("Source: SEC EDGAR XBRL API (data.sec.gov) | 本ツールは情報提供目的のみです。投資判断には必ず一次情報をご確認ください。")
