"""
PowerPointレポート生成 — templates/parr_report_template.pptx をそのまま使い、
プレースホルダ（"X" や "A年BQ"）を実データで置き換える。

テンプレートの構成:
  スライド1 (P&L)  : 7x5表（売上高／営業費用／営業利益／営業外費用／税引き前利益／純利益）
                     ＋「売上高の推移（3年）」「株価の推移（3年）」のグラフ枠
  スライド2 (B/S)  : 5x4表（前四半期 と 今期（差異））＋ 内訳表 5x3

テンプレートの書式（フォント・色・罫線）を壊さないため、セルの text をまとめて
代入するのではなく、既存 run の text だけを差し替える方針をとる。
"""

from __future__ import annotations

import io
from datetime import datetime
from pathlib import Path

TEMPLATE_PATH = Path(__file__).with_name("templates") / "parr_report_template.pptx"

# テンプレート注記「成績良化を青字、悪化を赤字で記載」に対応する色
_PLAIN_RGB = (0x00, 0x00, 0x00)     # 前年度比の文字色（黒で統一）

# P&L 表の行の並び（テンプレートの行順と対応）
_PL_ROWS = [
    ("Revenues",            "売上高"),
    ("OperatingExpenses",   "営業費用"),
    ("OperatingIncomeLoss", "営業利益"),
    ("InterestExpense",     "営業外費用"),
    ("IncomeLossBeforeTax", "税引き前利益"),
    ("NetIncomeLoss",       "純利益"),
]


def is_available() -> bool:
    """python-pptx とテンプレートが揃っていれば True。"""
    try:
        import pptx  # noqa: F401
    except ImportError:
        return False
    return TEMPLATE_PATH.exists()


# ─────────────────────────────────────────────────────────────────────────────
# 低レベルヘルパー（書式を保ったままテキストを差し替える）
# ─────────────────────────────────────────────────────────────────────────────

def _set_runs(paragraph, text: str, rgb: tuple | None = None) -> None:
    """段落内の最初の run に text を入れ、残りの run は空にする。

    python-pptx で paragraph.text / cell.text に代入すると run が1つに潰れて
    フォントや色の指定が失われるため、run 単位で書き換える。
    """
    runs = paragraph.runs
    if not runs:
        return
    runs[0].text = text
    for r in runs[1:]:
        r.text = ""
    if rgb is not None:
        from pptx.dml.color import RGBColor
        runs[0].font.color.rgb = RGBColor(*rgb)


def _clear_table_fill(table, keep: set | None = None) -> None:
    """表の塗りつぶしを解除して背景を透明にする。

    セル個別に「塗りつぶしなし」を指定するだけでなく、テーブルスタイル側の
    見出し行/縞模様/先頭列の強調フラグも落とす。これらが立っていると
    スタイル由来の色がセル指定より優先されて残ることがあるため。

    keep に (行, 列) を渡したセルは触らず、テンプレートの塗りをそのまま残す
    （「今期」の紺、「合計」のオレンジなど、意図的に色を付けている見出し）。
    keep 対象の色は tcPr に直接書かれた solidFill なので、上のスタイルフラグを
    落としても影響を受けない。
    """
    keep = keep or set()
    tbl_pr = table._tbl.find(
        "{http://schemas.openxmlformats.org/drawingml/2006/main}tblPr")
    if tbl_pr is not None:
        for flag in ("firstRow", "lastRow", "firstCol", "lastCol",
                     "bandRow", "bandCol"):
            tbl_pr.set(flag, "0")
    for ri, row in enumerate(table.rows):
        for ci, cell in enumerate(row.cells):
            if (ri, ci) in keep:
                continue
            cell.fill.background()          # <a:noFill/> を書き込む


def _set_cell(table, row: int, col: int, text: str,
              para: int = 0, rgb: tuple | None = None) -> None:
    cell = table.cell(row, col)
    paras = cell.text_frame.paragraphs
    if para < len(paras):
        _set_runs(paras[para], text, rgb)


def _replace_in_shape(shape, mapping: dict) -> None:
    """図形内のテキストを run 単位で置換する（複数 run にまたがる語にも対応）。"""
    if not shape.has_text_frame:
        return
    for p in shape.text_frame.paragraphs:
        joined = "".join(r.text for r in p.runs)
        if not joined:
            continue
        new = joined
        for old, rep in mapping.items():
            new = new.replace(old, rep)
        if new != joined:
            _set_runs(p, new)


def _fmt(v, unit: str = "") -> str:
    """USD百万単位の整数表記。None は "N/A"。"""
    if v is None:
        return "N/A"
    try:
        return f"{round(float(v)):,}{unit}"
    except (TypeError, ValueError):
        return "N/A"


def _fmt_signed(v) -> str:
    if v is None:
        return "N/A"
    try:
        return f"{round(float(v)):+,}"
    except (TypeError, ValueError):
        return "N/A"


def _m(val):
    """生の USD を百万単位へ。"""
    try:
        return None if val is None else float(val) / 1_000_000
    except (TypeError, ValueError):
        return None


def _pl_value(pl: dict, key: str, which: str):
    t = pl.get(key, {}).get(which)
    return _m(t[0]) if t and t[0] is not None else None


def _bs_value(bs: dict, key: str, which: str):
    t = bs.get(key, {}).get(which)
    return _m(t[0]) if t and t[0] is not None else None


def _quarter_label(period: str, form: str, quarter_num: int) -> tuple:
    """'2024-09-30' → ('2024年', '3Q') のような (年, 四半期) 表記を返す。"""
    try:
        y = datetime.strptime(period, "%Y-%m-%d").year
    except (ValueError, TypeError):
        return ("", "")
    q = "通期" if form == "10-K" else f"{quarter_num}Q"
    return (f"{y}年", q)


# ─────────────────────────────────────────────────────────────────────────────
# グラフ（PowerPointネイティブチャートとして挿入）
# ─────────────────────────────────────────────────────────────────────────────

def _add_native_chart(slide, anchor_shape, categories, series_name, values,
                      chart_type_name: str, number_format: str):
    """anchor_shape の位置・サイズにネイティブチャートを差し込み、枠図形は削除する。

    画像ではなくネイティブチャートにしておくと、PowerPoint側で色や軸を
    後から編集できる。
    """
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
    from pptx.util import Pt

    left, top = anchor_shape.left, anchor_shape.top
    width, height = anchor_shape.width, anchor_shape.height

    data = CategoryChartData()
    data.categories = list(categories)
    data.add_series(series_name, list(values), number_format)

    ctype = {"line": XL_CHART_TYPE.LINE_MARKERS,
             "bar": XL_CHART_TYPE.COLUMN_CLUSTERED}[chart_type_name]
    gframe = slide.shapes.add_chart(ctype, left, top, width, height, data)
    chart = gframe.chart

    chart.has_title = False
    chart.has_legend = False
    try:
        chart.font.size = Pt(9)
        chart.category_axis.tick_labels.font.size = Pt(8)
        chart.value_axis.tick_labels.font.size = Pt(8)
        chart.value_axis.has_major_gridlines = True
    except Exception:
        pass

    # プレースホルダの枠図形を削除（チャートと二重に見えるのを防ぐ）
    el = anchor_shape._element
    el.getparent().remove(el)
    return chart


def _find_shape_by_text(slide, needle: str):
    for sh in slide.shapes:
        if sh.has_text_frame and needle in sh.text_frame.text:
            return sh
    return None


def _find_chart_frames(slide):
    """スライド1にある2つのグラフ枠（テキストの無い大きめの四角）を上から順に返す。

    表やタイトル帯を巻き込まないよう、テキストが空・十分な大きさ・表/グラフでない、
    の3条件で絞り込む。
    """
    frames = []
    for sh in slide.shapes:
        if sh.has_table or sh.has_chart or not sh.has_text_frame:
            continue
        if sh.text_frame.text.strip():
            continue
        if not (sh.width and sh.height):
            continue
        if sh.width > 3_000_000 and sh.height > 1_500_000:
            frames.append(sh)
    return sorted(frames, key=lambda s: s.top or 0)


# ─────────────────────────────────────────────────────────────────────────────
# メイン
# ─────────────────────────────────────────────────────────────────────────────

def build_report(company_name: str, ticker: str, period: str, form: str,
                 quarter_num: int, pl: dict, bs: dict,
                 trend_df=None, price_df=None) -> bytes:
    """テンプレートに実データを流し込んだ .pptx を bytes で返す。

    pl / bs は extract_pl() / extract_bs() の戻り値、
    trend_df は extract_quarterly_trend() の戻り値（売上高グラフ用）、
    price_df は 'Close' 列を持つ日次株価DataFrame（株価グラフ用）。
    """
    from pptx import Presentation

    prs = Presentation(str(TEMPLATE_PATH))
    cur_y, cur_q = _quarter_label(period, form, quarter_num)
    prior_y = f"{int(cur_y[:-1]) - 1}年" if cur_y[:-1].isdigit() else ""

    # ── スライド1: P&L ──────────────────────────────────────
    s1 = prs.slides[0]
    _short_name = (company_name or ticker).split(",")[0].strip()
    title_map = {
        "A年BQ": f"{cur_y}{cur_q}",
        "(A-1)YBQ": f"{prior_y}{cur_q}",
        "Par Pacific": _short_name,     # 他社を分析した場合は社名を差し替える
    }
    for sh in s1.shapes:
        _replace_in_shape(sh, title_map)

    # 列の並びは「左=前年度、右=最新」（時系列で左→右に読める向き）
    t1 = next(sh.table for sh in s1.shapes if sh.has_table)
    _set_cell(t1, 0, 1, prior_y, para=0)
    _set_cell(t1, 0, 1, cur_q,   para=1)
    _set_cell(t1, 0, 2, cur_y,   para=0)
    _set_cell(t1, 0, 2, cur_q,   para=1)

    for i, (key, label) in enumerate(_PL_ROWS, start=1):
        cur = _pl_value(pl, key, "current")
        pri = _pl_value(pl, key, "prior")
        diff = (cur - pri) if (cur is not None and pri is not None) else None
        _set_cell(t1, i, 0, label)
        _set_cell(t1, i, 1, _fmt(pri))      # 左: 前年度
        _set_cell(t1, i, 2, _fmt(cur))      # 右: 最新
        # 前年度比は良化/悪化で色分けせず、黒で統一する
        _set_cell(t1, i, 3, _fmt_signed(diff), rgb=_PLAIN_RGB)

    # グラフ枠に売上高・株価チャートを差し込む
    frames = _find_chart_frames(s1)
    if len(frames) >= 1 and trend_df is not None and not trend_df.empty:
        _t = trend_df.dropna(subset=["revenue"])
        if not _t.empty:
            _add_native_chart(
                s1, frames[0], _t["quarter_label"].tolist(), "売上高 (MUSD)",
                [round(float(v)) for v in _t["revenue"]], "line", '#,##0')
    if len(frames) >= 2 and price_df is not None and not price_df.empty:
        _p = price_df.dropna()
        if not _p.empty:
            try:                                    # 月末リサンプルで点数を絞る
                _p = _p.resample("ME").last().dropna()
            except ValueError:
                _p = _p.resample("M").last().dropna()   # 旧pandas
            except Exception:
                pass
            _add_native_chart(
                s1, frames[1], [d.strftime("%y/%m") for d in _p.index],
                f"{ticker.upper()} 株価 (USD)",
                [round(float(v), 2) for v in _p["Close"]], "line", '#,##0.00')

    # ── スライド2: B/S ──────────────────────────────────────
    s2 = prs.slides[1]
    for sh in s2.shapes:
        _replace_in_shape(sh, title_map)

    # B/S側の表は原則として背景を塗らない。ただし「今期」ヘッダ（紺）と
    # 「合計」（オレンジ）はテンプレートの色を残す。
    # 表1(BS概要) の r0 は c0+c1 / c2+c3 が結合されているため両方を指定する。
    _keep = {0: {(0, 2), (0, 3)},           # BS概要: 「今期」ヘッダ
             1: {(0, 2), (4, 0)}}           # 内訳:   「今期」ヘッダ・「合計」
    tables2 = [sh.table for sh in s2.shapes if sh.has_table]
    for _i, _t in enumerate(tables2):
        _clear_table_fill(_t, keep=_keep.get(_i, set()))
    t2 = tables2[0]

    cash_c = _bs_value(bs, "Cash", "current")
    cash_p = _bs_value(bs, "Cash", "prior")
    ca_c   = _bs_value(bs, "CurrentAssets", "current")
    ca_p   = _bs_value(bs, "CurrentAssets", "prior")
    oca_c  = (ca_c - cash_c) if (ca_c is not None and cash_c is not None) else None
    oca_p  = (ca_p - cash_p) if (ca_p is not None and cash_p is not None) else None
    nca_c  = _bs_value(bs, "NonCurrentAssets", "current")
    nca_p  = _bs_value(bs, "NonCurrentAssets", "prior")
    cl_c   = _bs_value(bs, "CurrentLiabilities", "current")
    cl_p   = _bs_value(bs, "CurrentLiabilities", "prior")
    ltl_c  = _bs_value(bs, "LongTermLiabilities", "current")
    ltl_p  = _bs_value(bs, "LongTermLiabilities", "prior")
    eq_c   = _bs_value(bs, "StockholdersEquity", "current")
    eq_p   = _bs_value(bs, "StockholdersEquity", "prior")

    ta_c = sum(v for v in (cash_c, oca_c, nca_c) if v is not None) \
        if any(v is not None for v in (cash_c, oca_c, nca_c)) else None
    ta_p = sum(v for v in (cash_p, oca_p, nca_p) if v is not None) \
        if any(v is not None for v in (cash_p, oca_p, nca_p)) else None

    def _with_diff(cur, pri):
        if cur is None:
            return "N/A"
        if pri is None:
            return _fmt(cur)
        return f"{_fmt(cur)} ({_fmt_signed(cur - pri)})"

    # 見出し行（前四半期 / 今期）
    _set_cell(t2, 0, 0, f"　前四半期（{_period_short(bs, 'prior')}）")
    _set_cell(t2, 0, 2, f"今期（{_period_short(bs, 'current')}）")
    _set_cell(t2, 1, 0, f"総資産 {_fmt(ta_p)} MUSD")
    _set_cell(t2, 1, 2, f"総資産 {_with_diff(ta_c, ta_p)} MUSD")

    # 列の意味: 0=前四半期(資産) 1=前四半期(負債・資本) 2=今期(資産,差異) 3=今期(負債・資本,差異)
    for row, (asset_p, asset_c, liab_p, liab_c) in enumerate([
        (cash_p, cash_c, cl_p,  cl_c),      # 手元資金 / 流動負債
        (oca_p,  oca_c,  ltl_p, ltl_c),     # その他の流動資産 / 長期負債
        (nca_p,  nca_c,  eq_p,  eq_c),      # 固定資産・長期資産 / 株主資本
    ], start=2):
        _set_cell(t2, row, 0, _fmt(asset_p), para=1)
        _set_cell(t2, row, 1, _fmt(liab_p),  para=1)
        _set_cell(t2, row, 2, _with_diff(asset_c, asset_p), para=1)
        _set_cell(t2, row, 3, _with_diff(liab_c,  liab_p),  para=1)

    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()


def _period_short(bs: dict, which: str) -> str:
    """B/S から代表的な基準日を拾って 'YYYY/MM' 表記にする。"""
    for key in ("Cash", "CurrentAssets", "StockholdersEquity"):
        t = bs.get(key, {}).get(which)
        if t and t[1]:
            try:
                return datetime.strptime(t[1], "%Y-%m-%d").strftime("%Y/%m")
            except ValueError:
                pass
    return "—"
