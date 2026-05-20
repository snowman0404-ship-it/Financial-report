#!/usr/bin/env python3
"""Financial Analysis Web App — Streamlit / SEC EDGAR API"""

import io
import re
import time
from collections import Counter
from datetime import datetime, timedelta
from html.parser import HTMLParser

import pandas as pd
import requests
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════

USER_AGENT      = "FinancialWebAnalyzer/1.0 (financial-analysis@example.com)"
EDGAR_TICKERS   = "https://www.sec.gov/files/company_tickers.json"
EDGAR_FACTS_API = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"
EDGAR_SUBS_API  = "https://data.sec.gov/submissions/CIK{cik}.json"
EDGAR_ARCHIVES  = "https://www.sec.gov/Archives/edgar/data/{cik_int}/{acc_nodash}/{doc}"

KNOWN_CIKS = {
    "PARR": "0001378590", "AAPL": "0000320193", "MSFT": "0000789019",
    "GOOGL": "0001652044", "GOOG": "0001652044", "AMZN": "0001018724",
    "META": "0001326801", "TSLA": "0001318605", "NVDA": "0001045810",
    "JPM":  "0000019617", "BAC":  "0000070858", "XOM":  "0000034088",
    "CVX":  "0000093410", "WMT":  "0000104169", "JNJ":  "0000200406",
    "PG":   "0000080424", "KO":   "0000021344", "PFE":  "0000078003",
    "MRK":  "0000310158", "DIS":  "0001001039", "NFLX": "0001065280",
    "INTC": "0000050863", "AMD":  "0000002488", "COP":  "0001163165",
    "MPC":  "0001510295", "VLO":  "0001035002", "PSX":  "0001534701",
}

# ── PL XBRL Tags (priority order – first match with data wins) ──────────────
PL_TAGS = {
    "Revenues": [
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        "RevenueFromContractWithCustomerIncludingAssessedTax",
        "Revenues", "SalesRevenueNet", "SalesRevenueGoodsNet",
        "SalesAndRevenuesNet", "RevenuesNetOfInterestExpense",
        "NetRevenues", "TotalRevenues",
    ],
    "OperatingExpenses": [
        "CostOfRevenue", "CostOfGoodsAndServicesSold", "CostOfGoodsSold",
        "OperatingExpenses", "CostsAndExpenses", "OperatingCostsAndExpenses",
        "CostAndExpenses", "CostOfSales", "CostsOfRevenue",
        "CostOfGoodsAndServiceExcludingDepreciationDepletionAndAmortization",
    ],
    "OperatingIncomeLoss": [
        "OperatingIncomeLoss", "OperatingIncome",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
    ],
    "InterestExpense": [
        "NonoperatingIncomeExpense", "InterestExpense",
        "InterestIncomeExpenseNet", "InterestAndDebtExpense",
        "OtherNonoperatingIncomeExpense", "InterestAndOtherIncome",
    ],
    "IncomeLossBeforeTax": [
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesDomestic",
        "IncomeLossBeforeIncomeTaxes",
    ],
    "NetIncomeLoss": [
        "NetIncomeLoss", "ProfitLoss",
        "NetIncomeLossAvailableToCommonStockholdersBasic",
        "IncomeLossFromContinuingOperations",
        "ComprehensiveIncomeNetOfTax",
    ],
}

# ── BS XBRL Tags ─────────────────────────────────────────────────────────────
BS_TAGS = {
    "Cash": [
        "CashAndCashEquivalentsAtCarryingValue",
        "CashCashEquivalentsAndShortTermInvestments",
        "CashAndShortTermInvestments", "Cash",
        "CashAndCashEquivalentsAndRestrictedCashAndRestrictedCashEquivalents",
    ],
    "CurrentLiabilities": ["LiabilitiesCurrent"],
    "CurrentAssets":      ["AssetsCurrent"],
    "LongTermLiabilities": [
        "LongTermDebt", "LongTermDebtNoncurrent", "LiabilitiesNoncurrent",
        "LongTermDebtAndCapitalLeaseObligations",
        "LongTermDebtAndFinanceLeaseLiabilities",
        "FinanceLeaseLiabilityNoncurrent", "LongTermLineOfCredit",
        "SeniorLongTermNotes", "DebtAndCapitalLeaseObligations",
        "LongTermNotesPayable",
    ],
    "NonCurrentAssets": [
        "PropertyPlantAndEquipmentNet", "AssetsNoncurrent",
        "PropertyPlantAndEquipmentAndIntangibleAssetsNet",
        "NoncurrentAssets", "PropertyPlantAndEquipmentGross",
        "PropertyPlantAndEquipmentAndFinanceLeaseRightOfUseAssetAfterAccumulatedDepreciationAndAmortization",
        "RightOfUseAsset",
    ],
    "TotalAssets": ["Assets", "AssetsNet", "AssetsTotal"],
    "StockholdersEquity": [
        "StockholdersEquity", "StockholdersEquityAttributableToParent",
        "PartnersCapital", "MembersEquity",
        "CommonStockholdersEquity", "ShareholdersEquity",
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
    "OtherCurrentAssets": "その他流動資産 / Other Current Assets (=CurrentAssets−Cash)",
}

PL_ORDER = [
    "Revenues", "OperatingExpenses", "OperatingIncomeLoss",
    "InterestExpense", "IncomeLossBeforeTax", "NetIncomeLoss",
]
BS_ORDER = [
    "Cash", "CurrentLiabilities", "CurrentAssets",
    "LongTermLiabilities", "NonCurrentAssets", "StockholdersEquity",
    "OtherCurrentAssets",
]
COST_KEYS = {"OperatingExpenses", "CurrentLiabilities", "LongTermLiabilities"}

# ══════════════════════════════════════════════════════════════════════════════
# HTTP HELPER
# ══════════════════════════════════════════════════════════════════════════════

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
            elif r.status_code in (403, 404):
                return None
        except requests.RequestException:
            if attempt < retries - 1:
                time.sleep(backoff * (2 ** attempt))
    return None

# ══════════════════════════════════════════════════════════════════════════════
# CIK RESOLUTION
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data(ttl=3600, show_spinner=False)
def resolve_cik(ticker: str) -> str | None:
    t = ticker.upper().strip()
    if t in KNOWN_CIKS:
        return KNOWN_CIKS[t]
    r = _get(EDGAR_TICKERS)
    if r:
        try:
            for v in r.json().values():
                if v.get("ticker", "").upper() == t:
                    return str(v["cik_str"]).zfill(10)
        except Exception:
            pass
    return None

# ══════════════════════════════════════════════════════════════════════════════
# FILING LIST
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data(ttl=3600, show_spinner=False)
def get_filings_list(cik: str) -> list[dict]:
    r = _get(EDGAR_SUBS_API.format(cik=cik))
    if not r:
        return []
    try:
        data = r.json()
    except Exception:
        return []

    recent  = data.get("filings", {}).get("recent", {})
    forms   = recent.get("form",            [])
    dates   = recent.get("filingDate",      [])
    periods = recent.get("reportDate",      [])
    accnos  = recent.get("accessionNumber", [])
    docs    = recent.get("primaryDocument", [])

    filings = []
    for form, date, period, acc, doc in zip(forms, dates, periods, accnos, docs):
        if form in ("10-Q", "10-K"):
            filings.append({
                "label":      f"{period}  [{form}]  (filed {date})",
                "form":       form,
                "date":       date,
                "period":     period,
                "accession":  acc,
                "acc_nodash": acc.replace("-", ""),
                "primary_doc": doc,
            })
    return sorted(filings, key=lambda x: x["period"], reverse=True)

# ══════════════════════════════════════════════════════════════════════════════
# MD&A EXTRACTION
# ══════════════════════════════════════════════════════════════════════════════

class _HTMLStripper(HTMLParser):
    def __init__(self):
        super().__init__()
        self._parts = []
        self._skip_tags = {"script", "style", "head", "meta", "link"}
        self._depth = 0

    def handle_starttag(self, tag, attrs):
        if tag in self._skip_tags:
            self._depth += 1

    def handle_endtag(self, tag):
        if tag in self._skip_tags and self._depth > 0:
            self._depth -= 1

    def handle_data(self, data):
        if self._depth == 0:
            self._parts.append(data)

    def get_text(self):
        return " ".join(self._parts)


def extract_mda(html: str, max_chars: int = 30_000) -> str:
    stripper = _HTMLStripper()
    try:
        stripper.feed(html)
        text = stripper.get_text()
    except Exception:
        text = re.sub(r"<[^>]+>", " ", html)

    text = re.sub(r"\s+", " ", text).strip()

    start_pats = [
        r"ITEM\s+2[\.\s]+MANAGEMENT.S\s+DISCUSSION",
        r"Item\s+2[\.\s]+Management.s\s+Discussion",
        r"MANAGEMENT.S\s+DISCUSSION\s+AND\s+ANALYSIS",
    ]
    end_pats = [
        r"ITEM\s+3[\.\s]+",
        r"Item\s+3[\.\s]+",
        r"QUANTITATIVE\s+AND\s+QUALITATIVE\s+DISCLOSURES",
    ]

    start = -1
    for pat in start_pats:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            start = m.start()
            break

    if start == -1:
        return text[:max_chars]

    end = len(text)
    for pat in end_pats:
        m = re.search(pat, text[start + 200:], re.IGNORECASE)
        if m:
            end = start + 200 + m.start()
            break

    return text[start:end][:max_chars]


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_mda(cik: str, accession: str, primary_doc: str) -> str:
    cik_int = str(int(cik))
    acc_nodash = accession.replace("-", "")
    url = EDGAR_ARCHIVES.format(cik_int=cik_int, acc_nodash=acc_nodash, doc=primary_doc)
    html = _get(url, as_text=True)
    if html:
        return extract_mda(html)
    return "(MD&A テキストを取得できませんでした)"

# ══════════════════════════════════════════════════════════════════════════════
# XBRL DATA EXTRACTION
# ══════════════════════════════════════════════════════════════════════════════

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


def _filter_quarterly(records: list) -> list:
    out = []
    for r in records:
        s, e = r.get("start", ""), r.get("end", "")
        if s and e:
            try:
                d = (datetime.strptime(e, "%Y-%m-%d") - datetime.strptime(s, "%Y-%m-%d")).days
                if 75 <= d <= 110:
                    out.append(r)
            except ValueError:
                pass
    return out


def _filter_instant(records: list) -> list:
    return [r for r in records if not r.get("start")]


def _dedup_latest(records: list, n: int = 30) -> list:
    seen = {}
    for r in records:
        e = r.get("end", "")
        if e not in seen or r.get("form", "") in ("10-Q", "10-K"):
            seen[e] = r
    return sorted(seen.values(), key=lambda x: x.get("end", ""), reverse=True)[:n]


def _find_closest(records: list, target: datetime, max_days: int = 55) -> dict | None:
    best, best_d = None, max_days + 1
    for r in records:
        try:
            diff = abs((datetime.strptime(r["end"], "%Y-%m-%d") - target).days)
            if diff < best_d:
                best, best_d = r, diff
        except (ValueError, KeyError):
            pass
    return best


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_facts(cik: str) -> dict:
    r = _get(EDGAR_FACTS_API.format(cik=cik))
    return r.json() if r else {}


def _tuple_has_val(d: dict, key: str, which: str) -> bool:
    t = d.get(key, {}).get(which)
    return t is not None and t[0] is not None


def extract_pl(facts: dict, target_period: str | None = None) -> dict:
    target = datetime.strptime(target_period, "%Y-%m-%d") if target_period else None
    result = {}

    for metric, candidates in PL_TAGS.items():
        tag, recs = _best_tag(facts, candidates)
        quarterly = _dedup_latest(_filter_quarterly(recs))
        current = prior = None
        if quarterly:
            cur_rec = _find_closest(quarterly, target, 55) if target else quarterly[0]
            if cur_rec is None:
                cur_rec = quarterly[0]
            current = (cur_rec.get("val"), cur_rec["end"], tag)
            cur_end = datetime.strptime(cur_rec["end"], "%Y-%m-%d")
            prior_target = cur_end - timedelta(days=365)
            others = [r for r in quarterly if r["end"] != cur_rec["end"]]
            prior_rec = _find_closest(others, prior_target, 55)
            if prior_rec:
                prior = (prior_rec.get("val"), prior_rec["end"], tag)
        result[metric] = {"current": current, "prior": prior}

    # Derive OperatingExpenses = Revenues − OperatingIncomeLoss
    if not _tuple_has_val(result, "OperatingExpenses", "current"):
        rev = result.get("Revenues", {})
        opi = result.get("OperatingIncomeLoss", {})
        for which in ("current", "prior"):
            r_t = rev.get(which)
            o_t = opi.get(which)
            if r_t and o_t and r_t[0] is not None and o_t[0] is not None:
                result["OperatingExpenses"][which] = (
                    r_t[0] - o_t[0], r_t[1], "※導出: Revenues − OperatingIncomeLoss"
                )
    return result


def extract_bs(facts: dict, target_period: str | None = None) -> dict:
    target = datetime.strptime(target_period, "%Y-%m-%d") if target_period else None
    result = {}

    for metric, candidates in BS_TAGS.items():
        tag, recs = _best_tag(facts, candidates)
        instants = _dedup_latest(_filter_instant(recs))
        current = prior = None
        if instants:
            cur_rec = _find_closest(instants, target, 55) if target else instants[0]
            if cur_rec is None:
                cur_rec = instants[0]
            current = (cur_rec.get("val"), cur_rec["end"], tag)
            cur_end = datetime.strptime(cur_rec["end"], "%Y-%m-%d")
            prior_target = cur_end - timedelta(days=92)
            others = [r for r in instants if r["end"] != cur_rec["end"]]
            prior_rec = _find_closest(others, prior_target, 55)
            if prior_rec:
                prior = (prior_rec.get("val"), prior_rec["end"], tag)
        result[metric] = {"current": current, "prior": prior}

    # Derive NonCurrentAssets = TotalAssets − CurrentAssets
    if not _tuple_has_val(result, "NonCurrentAssets", "current"):
        tot = result.get("TotalAssets", {})
        ca  = result.get("CurrentAssets", {})
        for which in ("current", "prior"):
            t_t = tot.get(which)
            c_t = ca.get(which)
            if t_t and c_t and t_t[0] is not None and c_t[0] is not None:
                result["NonCurrentAssets"][which] = (
                    t_t[0] - c_t[0], t_t[1], "※導出: TotalAssets − CurrentAssets"
                )

    result.pop("TotalAssets", None)
    return result

# ══════════════════════════════════════════════════════════════════════════════
# VALUE HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _m(val) -> float | None:
    try:
        return float(val) / 1_000_000 if val is not None else None
    except (TypeError, ValueError):
        return None


def _gv(data: dict, key: str, which: str = "current") -> float | None:
    t = data.get(key, {}).get(which)
    if t is None:
        return None
    return _m(t[0])


def _gd(data: dict, key: str, which: str = "current") -> str:
    t = data.get(key, {}).get(which)
    return t[1] if t is not None else ""


def _pct(cur, pri) -> float | None:
    if cur is None or pri is None:
        return None
    try:
        if abs(float(pri)) < 1e-9:
            return None
        return (float(cur) - float(pri)) / abs(float(pri))
    except (TypeError, ValueError):
        return None


def _fmt_pct(pct: float | None, cur=None, pri=None) -> str:
    if pct is None:
        return "N/A"
    if cur is not None and pri is not None:
        if pri < 0 and cur >= 0:
            return "黒字転換"
        if pri >= 0 and cur < 0:
            return "赤字転落"
    return f"{pct:+.1%}"


def _canon_date(data: dict, which: str) -> str:
    dates = [v.get(which, (None, None))[1] for v in data.values()
             if v.get(which) and v[which][1]]
    if not dates:
        return ""
    return Counter(dates).most_common(1)[0][0]

# ══════════════════════════════════════════════════════════════════════════════
# ALERT LOGIC
# ══════════════════════════════════════════════════════════════════════════════

def compute_alerts(pl: dict, bs: dict) -> list[dict]:
    rev_c  = _gv(pl, "Revenues",            "current")
    rev_p  = _gv(pl, "Revenues",            "prior")
    opex_c = _gv(pl, "OperatingExpenses",   "current")
    opex_p = _gv(pl, "OperatingExpenses",   "prior")
    opi_c  = _gv(pl, "OperatingIncomeLoss", "current")
    opi_p  = _gv(pl, "OperatingIncomeLoss", "prior")
    ibt_c  = _gv(pl, "IncomeLossBeforeTax", "current")
    ibt_p  = _gv(pl, "IncomeLossBeforeTax", "prior")

    cash_c = _gv(bs, "Cash",               "current")
    cash_p = _gv(bs, "Cash",               "prior")
    cl_c   = _gv(bs, "CurrentLiabilities", "current")
    cl_p   = _gv(bs, "CurrentLiabilities", "prior")
    ca_c   = _gv(bs, "CurrentAssets",      "current")
    ca_p   = _gv(bs, "CurrentAssets",      "prior")
    eq_c   = _gv(bs, "StockholdersEquity", "current")
    eq_p   = _gv(bs, "StockholdersEquity", "prior")

    oca_c = (ca_c - cash_c) if (ca_c is not None and cash_c is not None) else None
    oca_p = (ca_p - cash_p) if (ca_p is not None and cash_p is not None) else None

    rev_chg  = _pct(rev_c,  rev_p)
    opex_chg = _pct(opex_c, opex_p)
    opi_chg  = _pct(opi_c,  opi_p)
    ibt_chg  = _pct(ibt_c,  ibt_p)
    cash_chg = _pct(cash_c, cash_p)
    cl_chg   = _pct(cl_c,   cl_p)
    oca_chg  = _pct(oca_c,  oca_p)
    eq_chg   = _pct(eq_c,   eq_p)

    alerts = []

    # PL-1: 収益性悪化
    if rev_chg is not None and opex_chg is not None and rev_chg < opex_chg:
        alerts.append({
            "type": "PL", "key": "OperatingExpenses",
            "title": "⚠️ 収益性悪化",
            "reason": (f"売上高変化率 {_fmt_pct(rev_chg, rev_c, rev_p)}"
                       f" < 営業費用変化率 {_fmt_pct(opex_chg, opex_c, opex_p)}"),
        })

    # PL-2: コストコントロール不全
    if rev_chg is not None and rev_chg >= 0 and opi_chg is not None and opi_chg < 0:
        alerts.append({
            "type": "PL", "key": "OperatingIncomeLoss",
            "title": "⚠️ コストコントロール不全",
            "reason": (f"増収（売上 {_fmt_pct(rev_chg, rev_c, rev_p)}）に"
                       f"もかかわらず営業利益 {_fmt_pct(opi_chg, opi_c, opi_p)}"),
        })

    # PL-3: 金融・本業外リスク
    if opi_c is not None and opi_c > 0 and ibt_c is not None:
        if ibt_c < 0:
            alerts.append({
                "type": "PL", "key": "IncomeLossBeforeTax",
                "title": "⚠️ 金融・本業外リスク",
                "reason": f"営業利益 +{opi_c:,.1f}M なのに税引前利益がマイナス（{ibt_c:,.1f}M）",
            })
        elif ibt_chg is not None and ibt_chg <= -0.20:
            alerts.append({
                "type": "PL", "key": "IncomeLossBeforeTax",
                "title": "⚠️ 金融・本業外リスク",
                "reason": f"営業利益プラスなのに税引前利益が大幅減少（{_fmt_pct(ibt_chg, ibt_c, ibt_p)}）",
            })

    # BS-1: 資金繰りショート懸念
    if cash_chg is not None and cash_chg <= -0.20 and cl_chg is not None and cl_chg >= 0.10:
        alerts.append({
            "type": "BS", "key": "Cash",
            "title": "⚠️ 資金繰りショート懸念",
            "reason": (f"手元資金 {_fmt_pct(cash_chg, cash_c, cash_p)}"
                       f" かつ 流動負債 {_fmt_pct(cl_chg, cl_c, cl_p)}"),
        })

    # BS-2: 在庫・売掛金の滞留リスク
    if cash_chg is not None and cash_chg < 0 and oca_chg is not None and oca_chg >= 0.20:
        alerts.append({
            "type": "BS", "key": "OtherCurrentAssets",
            "title": "⚠️ 在庫・売掛金の滞留リスク",
            "reason": (f"手元資金 {_fmt_pct(cash_chg, cash_c, cash_p)}"
                       f" かつ その他流動資産 {_fmt_pct(oca_chg, oca_c, oca_p)}"),
        })

    # BS-3: 自己資本の減少
    if eq_chg is not None and eq_chg < 0:
        alerts.append({
            "type": "BS", "key": "StockholdersEquity",
            "title": "⚠️ 自己資本の減少",
            "reason": f"株主資本変化率 {_fmt_pct(eq_chg, eq_c, eq_p)}",
        })

    return alerts


def compute_ratios(bs: dict) -> dict:
    ca  = _gv(bs, "CurrentAssets",      "current")
    cl  = _gv(bs, "CurrentLiabilities", "current")
    eq  = _gv(bs, "StockholdersEquity", "current")
    nca = _gv(bs, "NonCurrentAssets",   "current")
    ta  = (ca + nca) if (ca is not None and nca is not None) else None

    return {
        "current_ratio": (ca / cl)      if (ca and cl and cl != 0)  else None,
        "equity_ratio":  (eq / ta * 100) if (eq and ta and ta != 0) else None,
        "current_assets": ca,
        "current_liabilities": cl,
        "equity": eq,
        "total_assets": ta,
    }

# ══════════════════════════════════════════════════════════════════════════════
# DATAFRAME BUILDERS
# ══════════════════════════════════════════════════════════════════════════════

def _row_alert_label(key: str, cur, pri, pct: float | None) -> str:
    if pct is None:
        return "—"
    if cur is not None and pri is not None:
        if pri < 0 and cur > 0:
            return "✅ 黒字転換"
        if pri > 0 and cur < 0:
            return "⚠️ 赤字転落"
    is_cost = key in COST_KEYS
    if is_cost:
        return "⚠️ +20%↑ 悪化" if pct >= 0.20 else "✅ 正常"
    else:
        if pct <= -0.20:
            return "⚠️ -20%↓ 悪化"
        if pct >= 0.20:
            return "✅ +20%↑ 成長"
        return "✅ 正常"


def _build_table_rows(data: dict, order: list, labels: dict) -> tuple[pd.DataFrame, str, str]:
    cur_date = _canon_date(data, "current")
    pri_date = _canon_date(data, "prior")
    cur_col  = f"当期 ({cur_date}) [USD M]" if cur_date else "当期 [USD M]"
    pri_col  = f"前期 ({pri_date}) [USD M]" if pri_date else "前期 [USD M]"

    rows = []
    for key in order:
        if key == "OtherCurrentAssets":
            ca_c   = _gv(data, "CurrentAssets", "current")
            cash_c = _gv(data, "Cash",          "current")
            ca_p   = _gv(data, "CurrentAssets", "prior")
            cash_p = _gv(data, "Cash",          "prior")
            cur_v  = (ca_c - cash_c) if (ca_c is not None and cash_c is not None) else None
            pri_v  = (ca_p - cash_p) if (ca_p is not None and cash_p is not None) else None
        else:
            cur_v = _gv(data, key, "current")
            pri_v = _gv(data, key, "prior")

        pct  = _pct(cur_v, pri_v)
        diff = (cur_v - pri_v) if (cur_v is not None and pri_v is not None) else None

        rows.append({
            "項目 / Metric": labels.get(key, key),
            cur_col:         cur_v,
            pri_col:         pri_v,
            "差額 [USD M]":  diff,
            "変化率 %":      _fmt_pct(pct, cur_v, pri_v),
            "アラート":      _row_alert_label(key, cur_v, pri_v, pct),
        })

    return pd.DataFrame(rows), cur_col, pri_col


def build_pl_df(pl: dict) -> pd.DataFrame:
    df, _, _ = _build_table_rows(pl, PL_ORDER, PL_LABELS)
    return df


def build_bs_df(bs: dict) -> pd.DataFrame:
    df, _, _ = _build_table_rows(bs, BS_ORDER, BS_LABELS)
    return df


def _style_df(df: pd.DataFrame):
    numeric_cols = [c for c in df.columns if "[USD M]" in c]
    display_df = df.copy()
    for col in numeric_cols:
        display_df[col] = pd.to_numeric(display_df[col], errors="coerce")

    def highlight_row(row):
        alert = str(row["アラート"]) if "アラート" in row.index else ""
        bg = "background-color: #FFD2D2" if "⚠️" in alert else ""
        return pd.Series([bg] * len(row), index=row.index)

    fmt = {c: "{:,.1f}" for c in numeric_cols}
    return display_df.style.apply(highlight_row, axis=1).format(fmt, na_rep="N/A")

# ══════════════════════════════════════════════════════════════════════════════
# EXCEL BUILDER
# ══════════════════════════════════════════════════════════════════════════════

_RED_FILL   = PatternFill("solid", fgColor="FFD2D2")
_PALE_FILL  = PatternFill("solid", fgColor="DDEEFF")
_HEAD_FILL  = PatternFill("solid", fgColor="2F5496")
_DASH_FILL  = PatternFill("solid", fgColor="1F3864")
_GRAY_FILL  = PatternFill("solid", fgColor="F2F2F2")
_WHITE_FONT = Font(color="FFFFFF", bold=True)
_BOLD       = Font(bold=True)
_THIN       = Side(style="thin", color="CCCCCC")
_BORDER     = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_CENTER     = Alignment(horizontal="center", vertical="center", wrap_text=True)
_LEFT       = Alignment(horizontal="left",   vertical="center", wrap_text=True)
_RIGHT      = Alignment(horizontal="right",  vertical="center")


def _w(ws, row, col, value, fill=None, font=None, align=None, border=True, num_fmt=None):
    c = ws.cell(row=row, column=col, value=value)
    if fill:
        c.fill = fill
    if font:
        c.font = font
    if align:
        c.alignment = align
    if border:
        c.border = _BORDER
    if num_fmt:
        c.number_format = num_fmt
    return c


def _build_dashboard_sheet(ws, ticker: str, period: str,
                            ratios: dict, alerts: list[dict]) -> None:
    ws.sheet_view.showGridLines = True
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 60

    # Title
    ws.merge_cells("A1:C1")
    c = ws["A1"]
    c.value = f"📊 Financial Analysis Dashboard — {ticker.upper()}  [{period}]"
    c.fill  = _DASH_FILL
    c.font  = Font(color="FFFFFF", bold=True, size=14)
    c.alignment = _CENTER
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:C2")
    ws["A2"].value = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  Data: SEC EDGAR XBRL"
    ws["A2"].alignment = _CENTER
    ws["A2"].font = Font(italic=True, color="888888")
    ws.row_dimensions[2].height = 18

    # ── 財務健全性指標 ──────────────────────────────────────────────
    r = 4
    ws.merge_cells(f"A{r}:C{r}")
    ws[f"A{r}"].value = "■ 財務健全性指標"
    ws[f"A{r}"].fill = _HEAD_FILL
    ws[f"A{r}"].font = _WHITE_FONT
    ws[f"A{r}"].alignment = _LEFT
    r += 1

    cr = ratios.get("current_ratio")
    er = ratios.get("equity_ratio")
    ca = ratios.get("current_assets")
    cl = ratios.get("current_liabilities")
    eq = ratios.get("equity")
    ta = ratios.get("total_assets")

    def ratio_row(label, value, fmt, ok_thresh, bad_thresh, higher_is_better=True):
        nonlocal r
        ok = value is not None and (value >= ok_thresh if higher_is_better else value <= ok_thresh)
        fill = _GRAY_FILL if value is None else (None if ok else _RED_FILL)
        _w(ws, r, 1, label, fill=fill, font=_BOLD, align=_LEFT)
        _w(ws, r, 2, (fmt % value) if value is not None else "N/A", fill=fill, align=_RIGHT)
        _w(ws, r, 3, "", fill=fill, align=_LEFT)
        ws.row_dimensions[r].height = 18
        r += 1

    ratio_row("流動比率 (Current Ratio)",    cr, "%.2f x",  1.5, 1.0)
    ratio_row("自己資本比率 (Equity Ratio)", er, "%.1f %%", 30,  15)

    metrics = [
        ("流動資産 (Current Assets)",      ca,  "USD M"),
        ("流動負債 (Current Liabilities)", cl,  "USD M"),
        ("株主資本 (Stockholders' Equity)", eq, "USD M"),
        ("総資産 (Total Assets)",          ta,  "USD M"),
    ]
    for label, val, unit in metrics:
        _w(ws, r, 1, label,  align=_LEFT)
        _w(ws, r, 2, (f"{val:,.1f} {unit}" if val is not None else "N/A"), align=_RIGHT)
        _w(ws, r, 3, "", align=_LEFT)
        ws.row_dimensions[r].height = 18
        r += 1

    # ── 複合アラート判定 ────────────────────────────────────────────
    r += 1
    ws.merge_cells(f"A{r}:C{r}")
    ws[f"A{r}"].value = "■ 複合アラート判定"
    ws[f"A{r}"].fill = _HEAD_FILL
    ws[f"A{r}"].font = _WHITE_FONT
    ws[f"A{r}"].alignment = _LEFT
    r += 1

    if not alerts:
        ws.merge_cells(f"A{r}:C{r}")
        ws[f"A{r}"].value = "✅ アラートなし — 財務指標に重大な異常は検出されませんでした"
        ws[f"A{r}"].alignment = _LEFT
        r += 1
    else:
        for al in alerts:
            _w(ws, r, 1, al["title"],  fill=_RED_FILL, font=_BOLD, align=_LEFT)
            _w(ws, r, 2, al["type"],   fill=_RED_FILL, align=_CENTER)
            _w(ws, r, 3, al["reason"], fill=_RED_FILL, align=_LEFT)
            ws.row_dimensions[r].height = 20
            r += 1

    # ── 判定基準メモ ────────────────────────────────────────────────
    r += 1
    ws.merge_cells(f"A{r}:C{r}")
    ws[f"A{r}"].value = "■ 判定基準（参考）"
    ws[f"A{r}"].fill = _HEAD_FILL
    ws[f"A{r}"].font = _WHITE_FONT
    ws[f"A{r}"].alignment = _LEFT
    r += 1
    notes = [
        ("PL-1", "収益性悪化",           "売上高変化率 < 営業費用変化率"),
        ("PL-2", "コストコントロール不全", "売上高変化率 ≧ 0% かつ 営業利益変化率 < 0%"),
        ("PL-3", "金融・本業外リスク",    "営業利益プラスなのに税引前利益がマイナス or ≦ -20%"),
        ("BS-1", "資金繰りショート懸念", "手元資金変化率 ≦ -20% かつ 流動負債変化率 ≧ +10%"),
        ("BS-2", "在庫・売掛金の滞留リスク", "手元資金変化率 < 0% かつ その他流動資産変化率 ≧ +20%"),
        ("BS-3", "自己資本の減少",        "株主資本変化率 < 0%"),
    ]
    for code, name, cond in notes:
        _w(ws, r, 1, f"{code} {name}", fill=_GRAY_FILL, font=_BOLD, align=_LEFT)
        ws.merge_cells(f"B{r}:C{r}")
        _w(ws, r, 2, cond, fill=_GRAY_FILL, align=_LEFT)
        ws.row_dimensions[r].height = 16
        r += 1


def _build_fd_sheet(ws, pl: dict, bs: dict,
                    alert_keys: set[str]) -> None:
    ws.sheet_view.showGridLines = True
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 20

    cur_pl = _canon_date(pl, "current")
    pri_pl = _canon_date(pl, "prior")
    cur_bs = _canon_date(bs, "current")
    pri_bs = _canon_date(bs, "prior")

    def write_header(row, labels):
        for col, label in enumerate(labels, 1):
            _w(ws, row, col, label, fill=_HEAD_FILL, font=_WHITE_FONT, align=_CENTER)
        ws.row_dimensions[row].height = 22

    def write_section_title(row, title):
        ws.merge_cells(f"A{row}:F{row}")
        ws[f"A{row}"].value = title
        ws[f"A{row}"].fill  = _DASH_FILL
        ws[f"A{row}"].font  = _WHITE_FONT
        ws[f"A{row}"].alignment = _LEFT
        ws.row_dimensions[row].height = 20

    def write_data_row(row_idx, key, label, cur_v, pri_v, is_alert):
        fill = _RED_FILL if is_alert else None

        _w(ws, row_idx, 1, label,
           fill=fill, font=(_BOLD if is_alert else None), align=_LEFT)

        c_cur = ws.cell(row=row_idx, column=2, value=cur_v)
        c_cur.number_format = '#,##0.0'
        if fill:
            c_cur.fill = fill
        c_cur.border  = _BORDER
        c_cur.alignment = _RIGHT

        c_pri = ws.cell(row=row_idx, column=3, value=pri_v)
        c_pri.number_format = '#,##0.0'
        if fill:
            c_pri.fill = fill
        c_pri.border  = _BORDER
        c_pri.alignment = _RIGHT

        b_ref = get_column_letter(2)
        c_ref = get_column_letter(3)

        c_diff = ws.cell(row=row_idx, column=4,
                         value=f"={b_ref}{row_idx}-{c_ref}{row_idx}")
        c_diff.number_format = '#,##0.0'
        if fill:
            c_diff.fill = fill
        c_diff.border = _BORDER
        c_diff.alignment = _RIGHT

        c_chg = ws.cell(
            row=row_idx, column=5,
            value=(f'=IF({c_ref}{row_idx}=0,"N/A",'
                   f'({b_ref}{row_idx}-{c_ref}{row_idx})/ABS({c_ref}{row_idx}))'),
        )
        c_chg.number_format = '0.0%'
        if fill:
            c_chg.fill = fill
        c_chg.border = _BORDER
        c_chg.alignment = _RIGHT

        alert_txt = ("⚠️ アラート対象" if is_alert else "✅ 正常")
        _w(ws, row_idx, 6, alert_txt,
           fill=fill, align=_CENTER,
           font=Font(bold=is_alert, color=("CC0000" if is_alert else "006600")))

        ws.row_dimensions[row_idx].height = 16

    r = 1

    # ── PL Section ──────────────────────────────────────────────────
    write_section_title(r, "【損益計算書 (P&L) — 前年同期比 (YoY) 3ヶ月実績】")
    r += 1
    write_header(r, [
        "項目 / Metric",
        f"当期 ({cur_pl}) [USD M]",
        f"前期 ({pri_pl}) [USD M]",
        "差額 [USD M]",
        "変化率 %",
        "アラート",
    ])
    r += 1

    for key in PL_ORDER:
        cur_v = _gv(pl, key, "current")
        pri_v = _gv(pl, key, "prior")
        write_data_row(r, key, PL_LABELS[key], cur_v, pri_v, key in alert_keys)
        r += 1

    # Note row
    ws.merge_cells(f"A{r}:F{r}")
    ws[f"A{r}"].value = (
        "★ Non-GAAP注記: 石油精製等のエネルギー企業では在庫影響除き営業利益をMD&Aでご確認ください。"
    )
    ws[f"A{r}"].font = Font(italic=True, color="555555")
    ws[f"A{r}"].fill = _PALE_FILL
    ws.row_dimensions[r].height = 16
    r += 2

    # ── BS Section ──────────────────────────────────────────────────
    write_section_title(r, "【貸借対照表 (B/S) — 前四半期比 (QoQ)】")
    r += 1
    write_header(r, [
        "項目 / Metric",
        f"当四半期末 ({cur_bs}) [USD M]",
        f"前四半期末 ({pri_bs}) [USD M]",
        "差額 [USD M]",
        "変化率 %",
        "アラート",
    ])
    r += 1

    for key in BS_ORDER:
        if key == "OtherCurrentAssets":
            ca_c   = _gv(bs, "CurrentAssets", "current")
            cash_c = _gv(bs, "Cash",          "current")
            ca_p   = _gv(bs, "CurrentAssets", "prior")
            cash_p = _gv(bs, "Cash",          "prior")
            cur_v  = (ca_c - cash_c) if (ca_c is not None and cash_c is not None) else None
            pri_v  = (ca_p - cash_p) if (ca_p is not None and cash_p is not None) else None
        else:
            cur_v = _gv(bs, key, "current")
            pri_v = _gv(bs, key, "prior")
        write_data_row(r, key, BS_LABELS[key], cur_v, pri_v, key in alert_keys)
        r += 1


def _build_mda_sheet(ws, mda_text: str) -> None:
    ws.sheet_view.showGridLines = True
    ws.column_dimensions["A"].width = 120

    ws.merge_cells("A1:A2")
    ws["A1"].value = "Management's Discussion and Analysis (MD&A) — Extracted from SEC Filing"
    ws["A1"].fill  = _DASH_FILL
    ws["A1"].font  = _WHITE_FONT
    ws["A1"].alignment = _LEFT
    ws.row_dimensions[1].height = 24

    chunk_size = 2000
    chunks = [mda_text[i:i + chunk_size] for i in range(0, len(mda_text), chunk_size)]
    for idx, chunk in enumerate(chunks, start=3):
        c = ws.cell(row=idx, column=1, value=chunk)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[idx].height = 80


def build_excel(pl: dict, bs: dict, mda_text: str,
                ticker: str, period: str,
                alerts: list[dict], ratios: dict) -> bytes:
    wb = Workbook()
    alert_keys = {a["key"] for a in alerts}

    ws_dash = wb.active
    ws_dash.title = "Dashboard"
    _build_dashboard_sheet(ws_dash, ticker, period, ratios, alerts)

    ws_fd = wb.create_sheet("Financial Data")
    _build_fd_sheet(ws_fd, pl, bs, alert_keys)

    ws_mda = wb.create_sheet("MD&A_Text")
    _build_mda_sheet(ws_mda, mda_text)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ══════════════════════════════════════════════════════════════════════════════
# DEMO DATA
# ══════════════════════════════════════════════════════════════════════════════

def demo_data() -> tuple[dict, dict, str]:
    pl = {
        "Revenues":            {"current": (1_823_800_000, "2026-03-31", "Revenues"),
                                "prior":   (1_745_000_000, "2025-03-31", "Revenues")},
        "OperatingExpenses":   {"current": (1_758_500_000, "2026-03-31", "CostOfRevenue"),
                                "prior":   (1_760_800_000, "2025-03-31", "CostOfRevenue")},
        "OperatingIncomeLoss": {"current": (  65_300_000,  "2026-03-31", "OperatingIncomeLoss"),
                                "prior":   ( -15_800_000,  "2025-03-31", "OperatingIncomeLoss")},
        "InterestExpense":     {"current": (  -6_800_000,  "2026-03-31", "NonoperatingIncomeExpense"),
                                "prior":   ( -21_500_000,  "2025-03-31", "NonoperatingIncomeExpense")},
        "IncomeLossBeforeTax": {"current": (  58_500_000,  "2026-03-31", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest"),
                                "prior":   ( -37_300_000,  "2025-03-31", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest")},
        "NetIncomeLoss":       {"current": (  54_500_000,  "2026-03-31", "NetIncomeLoss"),
                                "prior":   ( -30_400_000,  "2025-03-31", "NetIncomeLoss")},
    }
    bs = {
        "Cash":               {"current": (172_200_000, "2026-03-31", "CashAndCashEquivalentsAtCarryingValue"),
                               "prior":   (164_100_000, "2025-12-31", "CashAndCashEquivalentsAtCarryingValue")},
        "CurrentLiabilities": {"current": (1_324_700_000, "2026-03-31", "LiabilitiesCurrent"),
                               "prior":   (1_106_100_000, "2025-12-31", "LiabilitiesCurrent")},
        "CurrentAssets":      {"current": (2_150_900_000, "2026-03-31", "AssetsCurrent"),
                               "prior":   (1_776_100_000, "2025-12-31", "AssetsCurrent")},
        "LongTermLiabilities":{"current": (  947_600_000, "2026-03-31", "LongTermDebt"),
                               "prior":   (  802_900_000, "2025-12-31", "LongTermDebt")},
        "NonCurrentAssets":   {"current": (1_900_000_000, "2026-03-31", "AssetsNoncurrent"),
                               "prior":   (1_850_000_000, "2025-12-31", "AssetsNoncurrent")},
        "StockholdersEquity": {"current": (1_515_800_000, "2026-03-31", "StockholdersEquity"),
                               "prior":   (1_511_500_000, "2025-12-31", "StockholdersEquity")},
    }
    mda = (
        "ITEM 2. MANAGEMENT'S DISCUSSION AND ANALYSIS OF FINANCIAL CONDITION "
        "AND RESULTS OF OPERATIONS\n\n"
        "[DEMO] Par Pacific Holdings reported revenues of $1,823.8M for Q1 2026, "
        "representing a 4.5% increase year-over-year. Operating income improved "
        "significantly from -$15.8M to $65.3M, driven by improved refining margins "
        "and operational efficiencies across our Hawaii and Mountain West refineries. "
        "Net income of $54.5M compares favorably to a net loss of $30.4M in Q1 2025. "
        "Cash position increased modestly to $172.2M. The company continues to focus "
        "on operational excellence and balance sheet strength. "
        "※ This is demo data for illustration purposes."
    )
    return pl, bs, mda

# ══════════════════════════════════════════════════════════════════════════════
# STREAMLIT UI
# ══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="SEC EDGAR 財務分析",
    page_icon="📊",
    layout="wide",
)

st.title("📊 SEC EDGAR 財務分析ツール")
st.caption("US上場企業のティッカーを入力 → 決算期を選択 → 財務分析を実行")

# ── Session State Initialisation ─────────────────────────────────────────────
for key in ("filings", "ticker_loaded", "result"):
    if key not in st.session_state:
        st.session_state[key] = None

# ── Step 1: Ticker + Get Filings ─────────────────────────────────────────────
col_tick, col_btn, col_demo = st.columns([3, 1.4, 1.4])
with col_tick:
    ticker_input = st.text_input(
        "ティッカー（例: PARR, AAPL, MSFT）",
        value="PARR",
        placeholder="US ticker symbol",
    )
with col_btn:
    st.write("")
    get_btn = st.button("📋 決算期リストを取得", use_container_width=True)
with col_demo:
    st.write("")
    demo_btn = st.button("🧪 デモデータで試す", use_container_width=True)

if get_btn:
    ticker = ticker_input.upper().strip()
    with st.spinner(f"{ticker} の決算期リストを取得中…"):
        cik = resolve_cik(ticker)
        if not cik:
            st.error(f"❌ ティッカー '{ticker}' のCIKが見つかりませんでした。")
        else:
            filings = get_filings_list(cik)
            if not filings:
                st.warning("決算期リストを取得できませんでした（ネットワーク制限の可能性あり）。デモデータをお試しください。")
            else:
                st.session_state.filings      = filings
                st.session_state.ticker_loaded = ticker
                st.session_state.result        = None
                st.success(f"✅ {len(filings)} 件の決算期を取得しました。")

if demo_btn:
    st.session_state.result = {"demo": True}
    st.session_state.filings = None

# ── Step 2: Period Selector + Analyse ────────────────────────────────────────
if st.session_state.get("filings"):
    filings = st.session_state.filings
    labels  = [f["label"] for f in filings]

    col_sel, col_run = st.columns([4, 1.4])
    with col_sel:
        sel_idx = st.selectbox(
            "対象決算期（Form 10-Q / 10-K）",
            range(len(labels)),
            format_func=lambda i: labels[i],
        )
    with col_run:
        st.write("")
        run_btn = st.button("🔍 財務分析を実行", type="primary", use_container_width=True)

    if run_btn:
        selected = filings[sel_idx]
        ticker   = st.session_state.ticker_loaded
        cik      = resolve_cik(ticker)

        with st.spinner("SEC EDGAR からデータを取得中…"):
            facts  = fetch_facts(cik)
            period = selected["period"]
            pl     = extract_pl(facts, period)
            bs     = extract_bs(facts, period)
            mda    = fetch_mda(cik, selected["accession"], selected["primary_doc"])

        alerts = compute_alerts(pl, bs)
        ratios = compute_ratios(bs)
        st.session_state.result = {
            "demo": False,
            "ticker": ticker,
            "period": period,
            "form":   selected["form"],
            "pl": pl, "bs": bs,
            "mda": mda,
            "alerts": alerts,
            "ratios": ratios,
        }

# ── Results Display ───────────────────────────────────────────────────────────
res = st.session_state.get("result")

if res and res.get("demo"):
    # Load demo
    pl, bs, mda = demo_data()
    alerts  = compute_alerts(pl, bs)
    ratios  = compute_ratios(bs)
    ticker  = "PARR"
    period  = "2026-03-31"
    form    = "10-Q"
    is_demo = True
elif res and not res.get("demo"):
    pl, bs, mda = res["pl"], res["bs"], res["mda"]
    alerts  = res["alerts"]
    ratios  = res["ratios"]
    ticker  = res["ticker"]
    period  = res["period"]
    form    = res["form"]
    is_demo = False
else:
    res = None

if res is not None:
    if is_demo:
        st.info("🧪 デモデータを表示中（実際のSECデータではありません）")

    st.divider()

    # ── Dashboard ────────────────────────────────────────────────────────────
    st.subheader(f"🏦 ダッシュボード — {ticker}  {period}  [{form}]")

    # KPI Metrics
    cr  = ratios.get("current_ratio")
    er  = ratios.get("equity_ratio")
    ca  = ratios.get("current_assets")
    cl  = ratios.get("current_liabilities")
    eq  = ratios.get("equity")
    ta  = ratios.get("total_assets")

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("流動比率",     f"{cr:.2f} x"    if cr is not None else "N/A",
              delta=("良好" if cr and cr >= 1.5 else ("注意" if cr and cr >= 1.0 else "要警戒")))
    k2.metric("自己資本比率", f"{er:.1f} %"    if er is not None else "N/A",
              delta=("良好" if er and er >= 30 else ("注意" if er and er >= 15 else "要警戒")))
    k3.metric("流動資産",     f"{ca:,.1f} M"   if ca is not None else "N/A")
    k4.metric("流動負債",     f"{cl:,.1f} M"   if cl is not None else "N/A")

    # ── Composite Alerts ─────────────────────────────────────────────────────
    st.markdown("### 🚨 複合アラート判定")
    if not alerts:
        st.success("✅ アラートなし — 財務指標に重大な異常は検出されませんでした。")
    else:
        for al in alerts:
            with st.container():
                col_ic, col_body = st.columns([0.06, 0.94])
                with col_body:
                    st.error(f"**{al['title']}** [{al['type']}]\n\n{al['reason']}")

    st.divider()

    # ── PL Table ─────────────────────────────────────────────────────────────
    st.markdown("## 📈 損益計算書（P&L）— 前年同期比（YoY）3ヶ月実績")
    pl_df = build_pl_df(pl)
    st.dataframe(_style_df(pl_df), width="stretch", height=270)

    st.info("★ **Non-GAAP注記**: PARR等エネルギー企業は在庫影響除き営業利益をMD&Aで確認してください。",
            icon="ℹ️")

    # ── BS Table ─────────────────────────────────────────────────────────────
    st.markdown("## 🏛️ 貸借対照表（B/S）— 前四半期比（QoQ）")
    bs_df = build_bs_df(bs)
    st.dataframe(_style_df(bs_df), width="stretch", height=310)

    st.divider()

    # ── MD&A ─────────────────────────────────────────────────────────────────
    st.markdown("## 📄 Management's Discussion and Analysis (MD&A)")
    st.caption("以下のテキストをそのままコピーしてClaudeなどの生成AIに貼り付けると、業績の日本語要約が得られます。")
    st.text_area(
        label="MD&A テキスト（英文・SEC提出原文）",
        value=mda,
        height=420,
        help="SEC EDGAR の Filing HTML から自動抽出した Item 2 のテキストです。",
    )

    st.divider()

    # ── Excel Download ────────────────────────────────────────────────────────
    st.markdown("### 💾 Excelダウンロード")
    with st.spinner("Excelファイルを生成中…"):
        excel_bytes = build_excel(pl, bs, mda, ticker, period, alerts, ratios)

    fname = f"financial_analysis_{ticker}_{period}.xlsx"
    st.download_button(
        label="📥 分析結果をExcelでダウンロード（3タブ構成）",
        data=excel_bytes,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption(f"ファイル名: {fname}  |  タブ: Dashboard / Financial Data / MD&A_Text")

else:
    st.info("👆 上のボタンで「決算期リストを取得」してから分析を実行するか、「デモデータ」をお試しください。")
