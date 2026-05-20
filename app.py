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
        "CostOfRevenue",
        "CostOfGoodsAndServicesSold",
        "CostOfGoodsSold",
        "OperatingExpenses",
        "CostsAndExpenses",
        "OperatingCostsAndExpenses",
        "CostAndExpenses",
        "CostOfRevenueExcludingDepreciation",
        "OperatingExpensesExcludingDepreciation",
        "CostOfGoodsAndServiceExcludingDepreciationDepletionAndAmortization",
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
    def __init__(self):
        super().__init__()
        self._parts = []
        self._skip = False

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style", "head"):
            self._skip = True
        if tag in ("p", "div", "br", "li", "tr", "h1", "h2", "h3", "h4"):
            self._parts.append("\n")

    def handle_endtag(self, tag):
        if tag in ("script", "style", "head"):
            self._skip = False

    def handle_data(self, data):
        if not self._skip:
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
    text = _strip_html(html)
    text = re.sub(r"[ \t]{2,}", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)

    start_pats = [
        r"(?im)^[\s ]*(ITEM\s*2[\.\:\-—\s]+MANAGEMENT.{0,80}DISCUSSION)",
        r"(?im)(Management[’\']s\s+Discussion\s+and\s+Analysis\s+of\s+Financial)",
        r"(?im)(MANAGEMENT[’\']S\s+DISCUSSION\s+AND\s+ANALYSIS)",
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
        r"(?im)^[\s ]*(ITEM\s*3[\.\:\-—\s]+QUANTITATIVE)",
        r"(?im)^[\s ]*(ITEM\s*3[\.\:\-—\s]+MARKET\s+RISK)",
        r"(?im)^[\s ]*(ITEM\s*3[\.\:\-—\s])",
    ]
    tail = text[start + 200:]
    end_offset = len(tail)
    for pat in end_pats:
        m = re.search(pat, tail)
        if m:
            end_offset = min(end_offset, m.start())
            break

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


def extract_pl(facts: dict, target_period: str | None = None) -> dict:
    target = datetime.strptime(target_period, "%Y-%m-%d") if target_period else None
    result = {}
    for metric, candidates in PL_TAGS.items():
        tag, recs = _best_tag(facts, candidates)
        quarterly = _dedup_latest(_filter_quarterly(recs), 30)
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

    # Derive OperatingExpenses = Revenues - OperatingIncomeLoss (if not found or val is None)
    def _val_is_missing(d, key):
        t = d.get(key, {}).get("current")
        return t is None or t[0] is None

    if _val_is_missing(result, "OperatingExpenses"):
        rev = result.get("Revenues", {})
        opi = result.get("OperatingIncomeLoss", {})
        for which in ("current", "prior"):
            r_t = rev.get(which)
            o_t = opi.get(which)
            if r_t and o_t:
                r_val, r_end, _ = r_t
                o_val, o_end, _ = o_t
                if r_val is not None and o_val is not None:
                    if not result.get("OperatingExpenses"):
                        result["OperatingExpenses"] = {}
                    result["OperatingExpenses"][which] = (r_val - o_val, r_end, "※導出値: Revenues − OperatingIncomeLoss")

    return result


def extract_bs(facts: dict, target_period: str | None = None) -> dict:
    target = datetime.strptime(target_period, "%Y-%m-%d") if target_period else None
    result = {}
    for metric, candidates in BS_TAGS.items():
        tag, recs = _best_tag(facts, candidates)
        instants = _dedup_latest(_filter_instant(recs), 8)
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

    # Derive NonCurrentAssets = TotalAssets - CurrentAssets (if not found or val is None)
    def _bs_val_missing(d, key):
        t = d.get(key, {}).get("current")
        return t is None or t[0] is None

    if _bs_val_missing(result, "NonCurrentAssets"):
        tot = result.get("TotalAssets", {})
        ca  = result.get("CurrentAssets", {})
        for which in ("current", "prior"):
            t_t = tot.get(which)
            c_t = ca.get(which)
            if t_t and c_t:
                t_val, t_end, _ = t_t
                c_val, c_end, _ = c_t
                if t_val is not None and c_val is not None:
                    if not result.get("NonCurrentAssets"):
                        result["NonCurrentAssets"] = {}
                    result["NonCurrentAssets"][which] = (t_val - c_val, t_end, "※導出値: TotalAssets − CurrentAssets")

    # Remove TotalAssets from result (it's a helper, not displayed)
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
            "アラート":       _alert(pct, cur, pri, is_cost),
            "_pct": pct, "_is_cost": is_cost, "_cur": cur, "_pri": pri, "_tag": tag,
        })
    return pd.DataFrame(rows)


def build_bs_df(bs: dict) -> pd.DataFrame:
    # ① Canonical dates
    canon_cur = _canon_date(bs, "current")
    canon_pri = _canon_date(bs, "prior")
    cur_col   = f"当四半期末 ({canon_cur})\n[USD M]"
    pri_col   = f"前四半期末 ({canon_pri})\n[USD M]"

    rows = []
    for key, label in BS_LABELS.items():
        data = bs.get(key, {})
        cur  = _safe_val(data, "current") if _period_ok(data, "current", canon_cur) else None
        pri  = _safe_val(data, "prior")   if _period_ok(data, "prior",   canon_pri) else None
        tag  = data["current"][2]         if data.get("current") else "—"

        delta   = (cur - pri) if (cur is not None and pri is not None) else None
        pct     = delta / abs(pri) if (delta is not None and pri not in (None, 0)) else None
        is_liab = key in ("CurrentLiabilities", "LongTermLiabilities")
        rows.append({
            "項目 / Metric": label,
            cur_col:         cur,
            pri_col:         pri,
            "差額 [USD M]":  delta,
            "変化率 %":       _pct_label(pct, cur, pri, is_liab),
            "アラート":       _alert(pct, cur, pri, is_liab),
            "_pct": pct, "_is_cost": is_liab, "_cur": cur, "_pri": pri, "_tag": tag,
        })
    # ③ Other Current Assets (derived)
    def _bv(k, w): return _safe_val(bs.get(k, {}), w) if _period_ok(bs.get(k, {}), w, canon_cur if w == "current" else canon_pri) else None
    ca  = _bv("CurrentAssets", "current"); ca_p  = _bv("CurrentAssets", "prior")
    cash= _bv("Cash", "current");          cashp = _bv("Cash", "prior")
    oca = (ca - cash)   if (ca   is not None and cash  is not None) else None
    ocap= (ca_p - cashp)if (ca_p is not None and cashp is not None) else None
    oca_d  = (oca - ocap) if (oca is not None and ocap is not None) else None
    oca_pct= oca_d / abs(ocap) if (oca_d is not None and ocap not in (None, 0)) else None
    rows.append({
        "項目 / Metric": "その他流動資産 / Other Current Assets (=CurrentAssets−Cash)",
        cur_col: oca, pri_col: ocap,
        "差額 [USD M]": oca_d,
        "変化率 %":      _pct_label(oca_pct, oca, ocap, False),
        "アラート":      _alert(oca_pct, oca, ocap, False),
        "_pct": oca_pct, "_is_cost": False, "_cur": oca, "_pri": ocap, "_tag": "(計算値)",
    })
    return pd.DataFrame(rows)

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
        pct     = df.iloc[idx]["_pct"]
        is_cost = df.iloc[idx]["_is_cost"]
        cur     = df.iloc[idx].get("_cur")
        pri     = df.iloc[idx].get("_pri")
        if pct is None or (isinstance(pct, float) and pd.isna(pct)):
            bg = ""
        elif not is_cost and cur is not None and pri is not None and pri > 0 and cur < 0:
            bg = "background-color: #FFD2D2;"   # 赤字転落
        elif not is_cost and cur is not None and pri is not None and pri < 0 and cur > 0:
            bg = ""                              # 黒字転換はハイライトなし
        else:
            bad = (is_cost and pct > 0.20) or (not is_cost and pct < -0.20)
            bg  = "background-color: #FFD2D2;" if bad else ""
        return pd.Series([bg] * len(display_cols), index=display_cols)

    def fmt_usd(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return "N/A"
        return f"{v:,.1f}"

    fmt       = {c: fmt_usd for c in display_cols if "USD M" in c or "差額" in c}
    right_cols = [c for c in display_cols if c not in ("項目 / Metric", "アラート", "変化率 %")]

    return (
        display_df.style
        .apply(highlight_row, axis=1)
        .format(fmt, na_rep="N/A")
        .set_properties(**{"text-align": "right"},  subset=right_cols)
        .set_properties(**{"text-align": "right"},  subset=["変化率 %"])
        .set_properties(**{"text-align": "left"},   subset=["項目 / Metric"])
        .set_properties(**{"text-align": "center"}, subset=["アラート"])
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


def compute_alerts(pl: dict, bs: dict) -> list[dict]:
    """6条件の複合アラート（PL×3 + BS×3）を返す。"""
    def _v(data, key, which):
        t = data.get(key, {}).get(which)
        return _m(t[0]) if (t is not None and t[0] is not None) else None

    def _pct(cur, pri):
        if cur is None or pri is None:
            return None
        try:
            if abs(float(pri)) < 1e-9:
                return None
            return (float(cur) - float(pri)) / abs(float(pri))
        except (TypeError, ValueError):
            return None

    def _fp(pct, cur, pri):
        if pct is None:
            return "N/A"
        if cur is not None and pri is not None:
            if pri < 0 and cur >= 0:
                return "黒字転換"
            if pri >= 0 and cur < 0:
                return "赤字転落"
        return f"{pct:+.1%}"

    rev_c  = _v(pl, "Revenues",            "current"); rev_p  = _v(pl, "Revenues",            "prior")
    opex_c = _v(pl, "OperatingExpenses",   "current"); opex_p = _v(pl, "OperatingExpenses",   "prior")
    opi_c  = _v(pl, "OperatingIncomeLoss", "current"); opi_p  = _v(pl, "OperatingIncomeLoss", "prior")
    ibt_c  = _v(pl, "IncomeLossBeforeTax", "current"); ibt_p  = _v(pl, "IncomeLossBeforeTax", "prior")
    cash_c = _v(bs, "Cash",               "current");  cash_p = _v(bs, "Cash",               "prior")
    cl_c   = _v(bs, "CurrentLiabilities", "current");  cl_p   = _v(bs, "CurrentLiabilities", "prior")
    ca_c   = _v(bs, "CurrentAssets",      "current");  ca_p   = _v(bs, "CurrentAssets",      "prior")
    eq_c   = _v(bs, "StockholdersEquity", "current");  eq_p   = _v(bs, "StockholdersEquity", "prior")

    oca_c = (ca_c - cash_c) if (ca_c is not None and cash_c is not None) else None
    oca_p = (ca_p - cash_p) if (ca_p is not None and cash_p is not None) else None

    rev_chg  = _pct(rev_c,  rev_p);  opex_chg = _pct(opex_c, opex_p)
    opi_chg  = _pct(opi_c,  opi_p);  ibt_chg  = _pct(ibt_c,  ibt_p)
    cash_chg = _pct(cash_c, cash_p); cl_chg   = _pct(cl_c,   cl_p)
    oca_chg  = _pct(oca_c,  oca_p);  eq_chg   = _pct(eq_c,   eq_p)

    alerts = []

    # PL-1: 収益性悪化
    if rev_chg is not None and opex_chg is not None and rev_chg < opex_chg:
        alerts.append({"code": "PL-1", "type": "PL", "title": "収益性悪化",
                        "reason": f"売上高変化率 {_fp(rev_chg, rev_c, rev_p)} < 営業費用変化率 {_fp(opex_chg, opex_c, opex_p)}"})

    # PL-2: コストコントロール不全
    if rev_chg is not None and rev_chg >= 0 and opi_chg is not None and opi_chg < 0:
        alerts.append({"code": "PL-2", "type": "PL", "title": "コストコントロール不全",
                        "reason": f"増収（売上 {_fp(rev_chg, rev_c, rev_p)}）にもかかわらず営業利益 {_fp(opi_chg, opi_c, opi_p)}"})

    # PL-3: 金融・本業外リスク
    if opi_c is not None and opi_c > 0 and ibt_c is not None:
        if ibt_c < 0:
            alerts.append({"code": "PL-3", "type": "PL", "title": "金融・本業外リスク",
                            "reason": f"営業利益 +{opi_c:,.1f}M なのに税引前利益がマイナス（{ibt_c:,.1f}M）"})
        elif ibt_chg is not None and ibt_chg <= -0.20:
            alerts.append({"code": "PL-3", "type": "PL", "title": "金融・本業外リスク",
                            "reason": f"営業利益プラスなのに税引前利益が大幅減少（{_fp(ibt_chg, ibt_c, ibt_p)}）"})

    # BS-1: 資金繰りショート懸念
    if cash_chg is not None and cash_chg <= -0.20 and cl_chg is not None and cl_chg >= 0.10:
        alerts.append({"code": "BS-1", "type": "BS", "title": "資金繰りショート懸念",
                        "reason": f"手元資金 {_fp(cash_chg, cash_c, cash_p)} かつ 流動負債 {_fp(cl_chg, cl_c, cl_p)}"})

    # BS-2: 在庫・売掛金の滞留リスク
    if cash_chg is not None and cash_chg < 0 and oca_chg is not None and oca_chg >= 0.20:
        alerts.append({"code": "BS-2", "type": "BS", "title": "在庫・売掛金の滞留リスク",
                        "reason": f"手元資金 {_fp(cash_chg, cash_c, cash_p)} かつ その他流動資産 {_fp(oca_chg, oca_c, oca_p)}"})

    # BS-3: 自己資本の減少
    if eq_chg is not None and eq_chg < 0:
        alerts.append({"code": "BS-3", "type": "BS", "title": "自己資本の減少",
                        "reason": f"株主資本変化率 {_fp(eq_chg, eq_c, eq_p)}"})

    return alerts


def _show_composite_alerts(alerts: list[dict]) -> None:
    """Streamlit UIに複合アラートをレンダリングする。"""
    st.markdown("### 🚨 複合アラート判定（PL×3 / BS×3）")
    if not alerts:
        st.success("✅ アラートなし — 6条件すべてで重大な財務悪化シグナルは検出されませんでした。")
        return
    for al in alerts:
        st.markdown(
            f'<div class="alert-box alert-red">⚠️ [{al["code"]}] {al["title"]} '
            f'<span style="font-weight:400;font-size:.95rem;">— {al["reason"]}</span></div>',
            unsafe_allow_html=True,
        )

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

    gc = ws.cell(row=row_num, column=7)
    alert = False
    if cur is not None and pri is not None and pri != 0:
        pct = (cur - pri) / abs(pri)
        if is_bad_if_high and pct > 0.20:
            gc.value = "⚠️ +20%↑ 急増"; alert = True
        elif not is_bad_if_high and pct < -0.20:
            gc.value = "⚠️ -20%↓ 急減"; alert = True
        else:
            gc.value = "✅ 正常"
    else:
        gc.value = "—"
    _sc(gc, align=_C)

    if alert:
        for col in range(2, 8):
            ws.cell(row=row_num, column=col).fill = _RED_F
    ws.row_dimensions[row_num].height = 18
    return alert


def _build_fd_sheet(wb, company_name, ticker, cik, pl, bs) -> dict:
    ws = wb.create_sheet("Financial Data")
    ws.views.sheetView[0].showGridLines = True
    _cw(ws, [3, 42, 18, 18, 18, 13, 16])

    ws.merge_cells("A1:G1")
    t = ws["A1"]
    t.value = f"財務分析レポート | {company_name} ({ticker.upper()}) | Source: SEC EDGAR XBRL"
    t.fill = _BLU_F; t.font = Font(name="Calibri", bold=True, color="FFFFFF", size=13)
    t.alignment = _C; ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:G2")
    s = ws["A2"]
    s.value = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  CIK: {cik}  |  Amounts in USD Millions (M)"
    s.fill = _SUB_F; s.font = Font(name="Calibri", color="FFFFFF", size=9, italic=True)
    s.alignment = _C; ws.row_dimensions[2].height = 14

    # ── PL section
    row = 4
    ws.merge_cells(f"A{row}:G{row}")
    c = ws.cell(row=row, column=1, value="📊  損益計算書（P&L） — 前年同期比（YoY）3ヶ月実績")
    c.fill = _SEC_F; c.font = _D_BOLD; c.alignment = _L; ws.row_dimensions[row].height = 22; row += 1

    pl_cd = pl_pd = "—"
    for d in pl.values():
        if d.get("current"): pl_cd = d["current"][1]; break
    for d in pl.values():
        if d.get("prior"):   pl_pd = d["prior"][1];   break

    _hrow(ws, row, ["", "項目 / Metric", f"当期\n({pl_cd})\n[USD M]",
                    f"前期\n({pl_pd})\n[USD M]", "差額 [USD M]", "変化率 %", "アラート"])
    ws.row_dimensions[row].height = 40; row += 1

    pl_order = [
        ("Revenues",            "売上高 / Revenues",               False),
        ("OperatingExpenses",   "営業費用 / Operating Expenses",   True),
        ("OperatingIncomeLoss", "営業利益 / Operating Income",     False),
        ("InterestExpense",     "金融損益 / Non-Op. Income/Exp.", False),
        ("IncomeLossBeforeTax", "税引前利益 / Income Before Tax",  False),
        ("NetIncomeLoss",       "純利益 / Net Income",             False),
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
    ws.merge_cells(f"C{row}:G{row}"); nc.border = _BORDER
    ws.row_dimensions[row].height = 26; row += 2

    # ── BS section
    ws.merge_cells(f"A{row}:G{row}")
    c = ws.cell(row=row, column=1, value="🏦  貸借対照表（B/S） — 前四半期比（QoQ）")
    c.fill = _SEC_F; c.font = _D_BOLD; c.alignment = _L; ws.row_dimensions[row].height = 22; row += 1

    bs_cd = bs_pd = "—"
    for d in bs.values():
        if d.get("current"): bs_cd = d["current"][1]; break
    for d in bs.values():
        if d.get("prior"):   bs_pd = d["prior"][1];   break

    _hrow(ws, row, ["", "項目 / Metric", f"当四半期末\n({bs_cd})\n[USD M]",
                    f"前四半期末\n({bs_pd})\n[USD M]", "差額 [USD M]", "変化率 %", "アラート"])
    ws.row_dimensions[row].height = 40; row += 1

    bs_order = [
        ("Cash",               "手元資金 / Cash & Equivalents",   False),
        ("CurrentLiabilities", "流動負債 / Current Liabilities",  True),
        ("CurrentAssets",      "流動資産 / Current Assets",       False),
        ("LongTermLiabilities","長期負債 / LT Liabilities",       True),
        ("NonCurrentAssets",   "固定資産 / Non-Current Assets",   False),
        ("StockholdersEquity", "株主資本 / Stockholders' Equity", False),
    ]
    for key, label, is_liab in bs_order:
        data = bs.get(key, {})
        cur  = _mval(data["current"][0]) if data.get("current") else None
        pri  = _mval(data["prior"][0])   if data.get("prior")   else None
        bs_rows_map[key] = row
        _data_row(ws, row, label, cur, pri, is_liab)
        row += 1

    # Other Current Assets
    cash_r = bs_rows_map.get("Cash"); ca_r = bs_rows_map.get("CurrentAssets")
    ws.cell(row=row, column=1).fill = _GRY_F; _sc(ws.cell(row=row, column=1))
    lc = ws.cell(row=row, column=2, value="その他流動資産 / Other Current Assets (=CurrentAssets−Cash)")
    lc.font = Font(name="Calibri", size=10, italic=True); lc.alignment = _L; lc.border = _BORDER
    if cash_r and ca_r:
        for col_idx, formula in [(3, f"=C{ca_r}-C{cash_r}"), (4, f"=D{ca_r}-D{cash_r}")]:
            c2 = ws.cell(row=row, column=col_idx, value=formula); c2.number_format = FMT_USD; _sc(c2, align=_R)
        e2 = ws.cell(row=row, column=5, value=f"=C{row}-D{row}"); e2.number_format = FMT_USD; _sc(e2, align=_R)
        f2 = ws.cell(row=row, column=6, value=f"=IF(D{row}=0,\"\",(C{row}-D{row})/D{row})"); f2.number_format = FMT_PCT; _sc(f2, align=_C)
    else:
        for col in range(3, 8):
            c2 = ws.cell(row=row, column=col, value="N/A"); c2.font = _ITAL; c2.border = _BORDER
    ws.cell(row=row, column=7, value="(計算値)").border = _BORDER
    ws.row_dimensions[row].height = 18; row += 2

    ws.merge_cells(f"A{row}:G{row}")
    leg = ws.cell(row=row, column=1,
                  value="[凡例] ⚠️=アラート(20%以上悪化) ✅=正常 N/A=データ未取得 ★=Non-GAAP(要手動確認) | 金額はUSD百万単位")
    leg.font = _ITAL; leg.fill = _GRY_F; leg.alignment = _L; ws.row_dimensions[row].height = 14

    return bs_rows_map


def _build_dashboard_sheet(wb, company_name, ticker, cik, pl, bs, bs_rows_map):
    wd = wb.create_sheet("Dashboard", 0)
    wd.views.sheetView[0].showGridLines = True
    _cw(wd, [3, 32, 22, 22, 16, 14, 3])

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
    _hrow(wd, dr, ["", "指標 / Ratio", "当期", "前期", "変化", "判定", ""]); wd.row_dimensions[dr].height = 22; dr += 1

    fd = "'Financial Data'"
    ca_r = bs_rows_map.get("CurrentAssets"); cl_r = bs_rows_map.get("CurrentLiabilities")
    eq_r = bs_rows_map.get("StockholdersEquity"); nca_r = bs_rows_map.get("NonCurrentAssets")

    def _ratio_row(label, cur_f, pri_f, fmt, status_f, sfill):
        nonlocal dr
        wd.cell(row=dr, column=1).border = _BORDER
        lc = wd.cell(row=dr, column=2, value=label); _sc(lc, font=_NORM, align=_L)
        cc = wd.cell(row=dr, column=3, value=cur_f); cc.number_format = fmt; _sc(cc, align=_R)
        dc = wd.cell(row=dr, column=4, value=pri_f); dc.number_format = fmt; _sc(dc, align=_R)
        ec = wd.cell(row=dr, column=5, value=f"=C{dr}-D{dr}"); ec.number_format = fmt; _sc(ec, align=_R)
        sc = wd.cell(row=dr, column=6, value=status_f); _sc(sc, fill=sfill, align=_C)
        wd.cell(row=dr, column=7).border = _BORDER; wd.row_dimensions[dr].height = 18; dr += 1

    ratios = compute_ratios(bs)
    cr_v = ratios["current_ratio"]; er_v = ratios["equity_ratio"]
    cr_fill = _RED_F if (cr_v and cr_v < 1) else (_YLW_F if (cr_v and cr_v < 1.5) else _GRN_F)
    er_fill = _RED_F if (er_v is not None and er_v < 0.1) else (_YLW_F if (er_v is not None and er_v < 0.3) else _GRN_F)

    if ca_r and cl_r:
        _ratio_row("流動比率 / Current Ratio",
                   f"={fd}!C{ca_r}/{fd}!C{cl_r}", f"={fd}!D{ca_r}/{fd}!D{cl_r}", FMT_RATIO,
                   f'=IF(C{dr-1}<1,"⚠️ <1.0 危険",IF(C{dr-1}<1.5,"⚡ 注意","✅ 良好"))', cr_fill)
    if ca_r and nca_r and eq_r:
        ta_c = f"({fd}!C{ca_r}+{fd}!C{nca_r})"; ta_p = f"({fd}!D{ca_r}+{fd}!D{nca_r})"
        _ratio_row("自己資本比率 / Equity Ratio",
                   f"={fd}!C{eq_r}/{ta_c}", f"={fd}!D{eq_r}/{ta_p}", FMT_PCT,
                   f'=IF(C{dr-1}<0.1,"⚠️ <10% 危険",IF(C{dr-1}<0.3,"⚡ <30% 低水準","✅ 良好"))', er_fill)

    dr += 1
    wd.merge_cells(f"A{dr}:G{dr}")
    c = wd.cell(row=dr, column=1, value="🚨  複合アラート判定（PL×3 / BS×3）")
    c.fill = PatternFill("solid", fgColor="C00000")
    c.font = Font(name="Calibri", bold=True, color="FFFFFF", size=12)
    c.alignment = _L; wd.row_dimensions[dr].height = 26; dr += 1
    _hrow(wd, dr, ["", "コード", "アラート名", "判定理由", "状態", "", ""]); wd.row_dimensions[dr].height = 20; dr += 1

    composite_alerts = compute_alerts(pl, bs)
    all_codes = ["PL-1", "PL-2", "PL-3", "BS-1", "BS-2", "BS-3"]
    all_titles = {
        "PL-1": "収益性悪化",           "PL-2": "コストコントロール不全",
        "PL-3": "金融・本業外リスク",   "BS-1": "資金繰りショート懸念",
        "BS-2": "在庫・売掛金の滞留リスク", "BS-3": "自己資本の減少",
    }
    triggered = {a["code"]: a["reason"] for a in composite_alerts}
    for code in all_codes:
        is_bad = code in triggered
        wd.cell(row=dr, column=1).border = _BORDER
        _sc(wd.cell(row=dr, column=2, value=code), font=Font(name="Calibri", bold=True, size=10), align=_C)
        _sc(wd.cell(row=dr, column=3, value=all_titles[code]), font=_NORM, align=_L)
        reason_cell = wd.cell(row=dr, column=4, value=triggered.get(code, "—条件非該当—"))
        _sc(reason_cell, font=_NORM, align=_L)
        sc = wd.cell(row=dr, column=5, value="⚠️ アラート" if is_bad else "✅ 正常")
        _sc(sc, fill=(_RED_F if is_bad else _GRN_F), align=_C)
        if is_bad:
            for col in [2, 3, 4]:
                wd.cell(row=dr, column=col).fill = PatternFill("solid", fgColor="FFD2D2")
        for col in [6, 7]: wd.cell(row=dr, column=col).border = _BORDER
        wd.row_dimensions[dr].height = 18; dr += 1

    dr += 1
    wd.merge_cells(f"A{dr}:G{dr}")
    n_alerts = len(composite_alerts)
    if n_alerts == 0:
        overall_txt = "✅ アラートなし — 6条件すべてで重大な財務悪化シグナルは検出されませんでした"
        overall_fill = PatternFill("solid", fgColor="D2FFD2")
        overall_font = Font(name="Calibri", bold=True, size=12, color="1E6B2E")
    else:
        overall_txt = f"⚠️ {n_alerts}件のアラートを検出 — 詳細は上記の判定理由を確認してください"
        overall_fill = PatternFill("solid", fgColor="FFD2D2")
        overall_font = Font(name="Calibri", bold=True, size=12, color="C00000")
    vc = wd.cell(row=dr, column=1, value=overall_txt)
    vc.fill = overall_fill; vc.font = overall_font
    vc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wd.row_dimensions[dr].height = 30; dr += 2

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

    k1,k2,k3,k4,k5,k6 = st.columns(6)
    def _ps(v): return f"{v:.1%}" if v is not None else "N/A"
    def _rs(v): return f"{v:.2f}" if v is not None else "N/A"
    def _ms(v): return f"${v:,.0f}M" if v is not None else "N/A"
    k1.metric("流動比率", _rs(ratios["current_ratio"]))
    k2.metric("自己資本比率", _ps(ratios["equity_ratio"]))
    k3.metric("現金 QoQ", _ps(ratios["cash_qoq"]))
    k4.metric("流動負債 QoQ", _ps(ratios["cl_qoq"]))
    k5.metric("株主資本 QoQ", _ps(ratios["eq_qoq"]))
    k6.metric("現金残高", _ms(ratios["cash_cur"]))
    _show_composite_alerts(compute_alerts(pl, bs))
    st.markdown("---")

    st.markdown("## 📊 損益計算書（P&L） — 前年同期比（YoY）")
    st.dataframe(_style_df(build_pl_df(pl)), width="stretch", height=270)
    st.info("★ **Non-GAAP**: PARR等エネルギー企業は在庫影響除き営業利益をMD&Aで確認してください。", icon="ℹ️")

    st.markdown("## 🏦 貸借対照表（B/S） — 前四半期比（QoQ）")
    st.dataframe(_style_df(build_bs_df(bs)), width="stretch", height=310)

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

            pl  = extract_pl(facts, period)
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

        k1,k2,k3,k4,k5,k6 = st.columns(6)
        def _ps(v): return f"{v:.1%}" if v is not None else "N/A"
        def _rs(v): return f"{v:.2f}" if v is not None else "N/A"
        def _ms(v): return f"${v:,.0f}M" if v is not None else "N/A"
        k1.metric("流動比率", _rs(ratios["current_ratio"]))
        k2.metric("自己資本比率", _ps(ratios["equity_ratio"]))
        k3.metric("現金 QoQ", _ps(ratios["cash_qoq"]))
        k4.metric("流動負債 QoQ", _ps(ratios["cl_qoq"]))
        k5.metric("株主資本 QoQ", _ps(ratios["eq_qoq"]))
        k6.metric("現金残高", _ms(ratios["cash_cur"]))
        _show_composite_alerts(compute_alerts(pl, bs))
        st.markdown("---")

        st.markdown("## 📊 損益計算書（P&L） — 前年同期比（YoY）3ヶ月実績")
        st.dataframe(_style_df(build_pl_df(pl)), width="stretch", height=270)
        st.info("★ **Non-GAAP**: PARR等エネルギー企業は在庫影響除き営業利益をMD&Aで確認してください。", icon="ℹ️")

        st.markdown("## 🏦 貸借対照表（B/S） — 前四半期比（QoQ）")
        st.dataframe(_style_df(build_bs_df(bs)), width="stretch", height=310)

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
