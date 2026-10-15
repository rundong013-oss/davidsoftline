#!/usr/bin/env python3
"""
update_monitor.py — 增量更新 Apparel_updating file.xlsx

读取现有 Excel，从最后一行日期起增量拉取数据，追加新行/新列。
替代 apparel_luxury_etf.py（全量重建）+ macro_update.py（独立宏观）。

用法：
    python3 scripts/update_monitor.py
    python3 scripts/update_monitor.py --demo
    python3 scripts/update_monitor.py --full   # 强制全量重建
"""
import argparse
import datetime as dt
import io
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from fmp_client import FMPClient, FMPError

# ══════════════════════════════════════════════════════════════════════
#  Constants — 从 apparel_luxury_etf.py / macro_update.py 提取
# ══════════════════════════════════════════════════════════════════════
XLSX_PATH = "output/monitor/Apparel_updating file.xlsx"

BASE_DATE = "2025-12-01"        # 运动服饰 & 奢侈品基期
ETF_BASE_DATE = "2025-01-01"    # ETF & 宏观基期
MACRO_BASE = "2025-01-01"

# ── 运动服饰 ─────────────────────────────────────────────────────────
SPORTSWEAR = [
    # (展示代码, Yahoo代码, Bloomberg, 区域, 备注)
    ("NKE.N",      "NKE",       "NKE US",    "US/EU",    "Nike"),
    ("ADDY",       "ADS.DE",    "ADS GY",    "US/EU",    "Adidas 德国本地股 EUR"),
    ("ZGN.N",      "ZGN",       "ZGN US",    "US/EU",    "Zegna"),
    ("WWW.N",      "WWW",       "WWW US",    "US/EU",    "Wolverine"),
    ("VSCO.N",     None,        None,        "US/EU",    "疑似退市"),
    ("UAA.N",      "UAA",       "UAA US",    "US/EU",    "Under Armour"),
    ("SKX.N",      None,        None,        "US/EU",    "私有化退市"),
    ("SHOO.O",     "SHOO",      "SHOO US",   "US/EU",    "Steven Madden"),
    ("RL.N",       "RL",        "RL US",     "US/EU",    "Ralph Lauren"),
    ("PUM.DF",     "PUM.DE",    "PUM GY",    "US/EU",    "Puma EUR"),
    ("ONON.N",     "ONON",      "ONON US",   "US/EU",    "On Holding"),
    ("NWL.O",      "NWL",       "NWL US",    "US/EU",    "Newell Brands"),
    ("LULU.O",     "LULU",      "LULU US",   "US/EU",    "Lululemon"),
    ("LEVI.N",     "LEVI",      "LEVI US",   "US/EU",    "Levi Strauss"),
    ("KTB.N",      "KTB",       "KTB US",    "US/EU",    "Kontoor Brands"),
    ("HBI.N",      None,        None,        "US/EU",    "被 Gildan 收购退市"),
    ("GIL.N",      "GIL",       "GIL US",    "US/EU",    "Gildan"),
    ("GAP.N",      "GAP",       "GAP US",    "US/EU",    "Gap"),
    ("DECK.N",     "DECK",      "DECK US",   "US/EU",    "Deckers"),
    ("CROX.O",     "CROX",      "CROX US",   "US/EU",    "Crocs"),
    ("COLM.O",     "COLM",      "COLM US",   "US/EU",    "Columbia"),
    ("BZUN.O",     "BZUN",      "BZUN US",   "US/EU",    "宝尊电商 ADR"),
    ("AS.N",       "AS",        "AS US",     "US/EU",    "Amer Sports"),
    ("ANF.N",      "ANF",       "ANF US",    "US/EU",    "Abercrombie"),
    ("9983.T",     "9983.T",    "9983 JT",   "Japan",    "迅销"),
    ("8111.T",     "8111.T",    "8111 JT",   "Japan",    "Goldwin"),
    ("7936.T",     "7936.T",    "7936 JT",   "Japan",    "亚瑟士"),
    ("3998.HK",    "3998.HK",   "3998 HK",   "China/HK", "波司登"),
    ("2331.HK",    "2331.HK",   "2331 HK",   "China/HK", "李宁"),
    ("2020.HK",    "2020.HK",   "2020 HK",   "China/HK", "安踏"),
    ("1836.HK",    "1836.HK",   "1836 HK",   "China/HK", "九兴控股"),
    ("1368.HK",    "1368.HK",   "1368 HK",   "China/HK", "特步"),
    ("1361.HK",    "1361.HK",   "1361 HK",   "China/HK", "361度"),
    ("002832.SZ",  "002832.SZ", "002832 CH", "China/HK", "比音勒芬"),
]

# ── 奢侈品 ───────────────────────────────────────────────────────────
LUXURY = [
    ("VFC",        "VFC",       "VFC US",    "US",       "V.F. Corporation"),
    ("CFR",        "CFR.SW",    "CFR SW",    "Europe",   "Richemont 瑞士"),
    ("1913.HK",    "1913.HK",   "1913 HK",  "Asia",     "Prada"),
    ("MONC",       "MONC.MI",   "MONC IM",  "Europe",   "Moncler 米兰"),
    ("KER",        "KER.PA",    "KER FP",   "Europe",   "Kering 巴黎"),
    ("SFER",       "SFER.MI",   "SFER IM",  "Europe",   "Ferragamo 米兰"),
    ("RACE",       "RACE",      "RACE US",  "US",       "Ferrari"),
    ("RMS",        "RMS.PA",    "RMS FP",   "Europe",   "Hermes 巴黎"),
    ("MC",         "MC.PA",     "MC FP",    "Europe",   "LVMH 巴黎"),
    ("BRBY",       "BRBY.L",    "BRBY LN",  "Europe",   "Burberry 伦敦"),
    ("CPRI",       "CPRI",      "CPRI US",  "US",       "Capri Holdings"),
    ("PVH",        "PVH",       "PVH US",   "US",       "PVH Corp"),
    ("TPR",        "TPR",       "TPR US",   "US",       "Tapestry"),
    ("GOOS",       "GOOS.TO",   "GOOS CT",  "North America", "Canada Goose 多伦多"),
    ("ZGN",        "ZGN",       "ZGN US",   "US",       "Zegna"),
]

# ── ETFs ─────────────────────────────────────────────────────────────
ETFS = [
    ("XLY",  "XLY",  "可选消费 ETF 市值加权",       "被亚马逊、特斯拉主导"),
    ("RSPD", "RSPD", "标普500可选消费等权 ETF",      "去掉巨头集中度"),
    ("XRT",  "XRT",  "标普零售精选 ETF 等权",        "零售属性强，含服装零售"),
]

# ── 周涨跌 sheet 行布局（固定） ──────────────────────────────────────
# Row 1: 标题 "周涨跌幅"
# Row 2: header — "Ticker", date-range labels...
# Row 3: "运动服饰" section label
# Row 4–36: sportswear tickers (34 rows, matching SPORTSWEAR order)
# Row 37: blank
# Row 38: "奢侈品" section label
# Row 39–53: luxury tickers (15 rows, matching LUXURY order)
# Row 54: blank
# Row 55: "宏观指标" section label
# Row 56: date-range header row for macro
# Row 57–72: macro data rows

WC_ROW_SW_START = 4       # first sportswear ticker
WC_ROW_LUX_START = 39     # first luxury ticker
WC_ROW_MACRO_LABEL = 55
WC_ROW_MACRO_HDR = 56
WC_ROW_BRENT = 57
WC_ROW_WTI = 58
WC_ROW_GAS = 59
WC_ROW_BRENT_CHG = 60
WC_ROW_WTI_CHG = 61
WC_ROW_GAS_CHG = 62
# Row 63: blank
WC_ROW_YIELD_HDR = 64
WC_ROW_2Y = 65
WC_ROW_10Y = 66
WC_ROW_30Y = 67
WC_ROW_2Y_CHG = 68
WC_ROW_10Y_CHG = 69
WC_ROW_30Y_CHG = 70
# Row 71: blank
WC_ROW_CLAIMS = 72

# ── FRED 序列 ────────────────────────────────────────────────────────
FRED_WEEKLY = [
    ("ICSA",    "初请失业金(千人)"),
    ("GASREGW", "全美汽油均价($/gal)"),
    ("DGS2",    "2Y(%)"),
]

FRED_MONTHLY = [
    ("CPIAUCSL", "CPI YoY%",       "yoy",      "0.0"),
    ("CPILFESL", "核心CPI YoY%",   "yoy",      "0.0"),
    ("PPIACO",   "PPI YoY%",       "yoy",      "0.0"),
    ("PCEPI",    "PCE YoY%",       "yoy",      "0.0"),
    ("PCEPILFE", "核心PCE YoY%",   "yoy",      "0.0"),
    ("PAYEMS",   "非农delta(千人)",    "mom_diff",  "#,##0"),
    ("UNRATE",   "失业率(%)",      "level",     "0.0"),
    ("RSAFS",    "零售销售MoM%",   "mom_pct",   "0.0"),
    ("UMCSENT",  "消费者信心",     "level",     "0.0"),
    ("REVOLSL",  "信用卡余额YoY%", "yoy",      "0.0"),
]


# ══════════════════════════════════════════════════════════════════════
#  数据获取 — 复用自 apparel_luxury_etf.py
# ══════════════════════════════════════════════════════════════════════
def _fetch_daily_fmp(yahoo_codes, start_date):
    """FMP 拉取日线收盘价，返回 (成功dict, 失败list)。"""
    client = FMPClient()
    start = (pd.Timestamp(start_date) - pd.Timedelta(days=10)).strftime("%Y-%m-%d")
    end = (dt.date.today() + dt.timedelta(days=1)).strftime("%Y-%m-%d")
    YF_TO_FMP = {"BZ=F": "BZUSD", "CL=F": "CLUSD"}
    out, failed = {}, []
    for code in yahoo_codes:
        fmp_sym = YF_TO_FMP.get(code, code)
        try:
            df = client.historical_prices(fmp_sym, start, end)
            if not df.empty and "close" in df.columns:
                out[code] = df["close"].dropna()
            else:
                failed.append(code)
        except FMPError:
            failed.append(code)
    return out, failed


def _fetch_daily_yf(yahoo_codes, start_date):
    """yfinance fallback 批量拉取。"""
    import yfinance as yf
    start = (pd.Timestamp(start_date) - pd.Timedelta(days=10)).strftime("%Y-%m-%d")
    end = (dt.date.today() + dt.timedelta(days=1)).strftime("%Y-%m-%d")
    raw = yf.download(yahoo_codes, start=start, end=end, interval="1d",
                      auto_adjust=False, group_by="ticker", progress=False, threads=True)
    out = {}
    if len(yahoo_codes) == 1:
        try:
            out[yahoo_codes[0]] = raw["Close"].dropna()
        except KeyError:
            out[yahoo_codes[0]] = pd.Series(dtype=float)
    else:
        for code in yahoo_codes:
            try:
                out[code] = raw[code]["Close"].dropna()
            except KeyError:
                out[code] = pd.Series(dtype=float)
    return out


def fetch_daily(yahoo_codes, start_date):
    """FMP 优先，失败回退 yfinance。"""
    fmp_data, failed = _fetch_daily_fmp(yahoo_codes, start_date)
    if failed:
        print(f"  FMP 失败 {len(failed)} 只，yfinance fallback: {', '.join(failed)}")
        try:
            yf_data = _fetch_daily_yf(failed, start_date)
            fmp_data.update(yf_data)
        except Exception as e:
            print(f"  yfinance fallback 也失败: {e}")
    for code in yahoo_codes:
        if code not in fmp_data:
            fmp_data[code] = pd.Series(dtype=float)
    return pd.DataFrame(fmp_data)


def to_weekly(daily, base_date):
    """日线转周度（W-FRI），含基期行。"""
    daily = daily.sort_index()
    after = daily[daily.index >= base_date]
    if after.empty:
        return after
    base = after.apply(lambda s: s.dropna().iloc[0] if s.notna().any() else np.nan)
    weekly = after.resample("W-FRI").last()
    weekly = weekly[weekly.index > pd.Timestamp(base_date)]
    base_row = pd.DataFrame([base.values], columns=daily.columns,
                            index=[pd.Timestamp(base_date)])
    return pd.concat([base_row, weekly]).ffill()


# ══════════════════════════════════════════════════════════════════════
#  FRED 数据 — 复用自 macro_update.py
# ══════════════════════════════════════════════════════════════════════
def _fred_one(series_id, start):
    """拉单个 FRED 序列。"""
    from curl_cffi import requests as cr
    end_str = (dt.date.today() + dt.timedelta(days=1)).isoformat()
    url = (f"https://fred.stlouisfed.org/graph/fredgraph.csv"
           f"?id={series_id}&cosd={start}&coed={end_str}")
    try:
        r = cr.get(url, timeout=15, impersonate="chrome")
        r.raise_for_status()
        df = pd.read_csv(io.StringIO(r.text))
        date_col = df.columns[0]
        df[date_col] = pd.to_datetime(df[date_col])
        df = df.set_index(date_col)
        s = df.iloc[:, 0].replace(".", np.nan).astype(float).dropna()
        s.name = series_id
        return series_id, s
    except Exception as e:
        print(f"  FRED {series_id}: {e}")
        return series_id, pd.Series(dtype=float, name=series_id)


def fetch_fred_parallel(series_list, start="2024-01-01"):
    """并行拉取多个 FRED 序列。"""
    results = {}
    with ThreadPoolExecutor(max_workers=6) as pool:
        futures = {pool.submit(_fred_one, sid, start): sid
                   for sid in series_list}
        for f in as_completed(futures):
            sid, s = f.result()
            results[sid] = s
    return results


def _fetch_oil_yield_fmp(start, end):
    """FMP 拉取油价和收益率。"""
    client = FMPClient()
    oil, yields = {}, {}
    for fmp_sym, name in [("BZUSD", "Brent($/bbl)"), ("CLUSD", "WTI($/bbl)")]:
        try:
            df = client.historical_prices(fmp_sym, start, end)
            if not df.empty and "close" in df.columns:
                oil[name] = df["close"].dropna()
        except FMPError as e:
            print(f"  FMP {fmp_sym} 失败: {e}")
    for fmp_sym, name in [("^TNX", "10Y(%)"), ("^TYX", "30Y(%)")]:
        try:
            df = client.historical_prices(fmp_sym, start, end)
            if not df.empty and "close" in df.columns:
                yields[name] = df["close"].dropna()
        except FMPError as e:
            print(f"  FMP {fmp_sym} 失败: {e}")
    return oil, yields


def fetch_macro():
    """拉取全部宏观数据，返回 dict。"""
    data = {}
    start = (pd.Timestamp(MACRO_BASE) - pd.Timedelta(days=10)).strftime("%Y-%m-%d")
    end = (dt.date.today() + dt.timedelta(days=1)).strftime("%Y-%m-%d")

    print("  Brent/WTI/Yields (FMP)...")
    oil, yields = _fetch_oil_yield_fmp(start, end)

    # yfinance fallback
    need_oil = [n for n in ["Brent($/bbl)", "WTI($/bbl)"] if n not in oil]
    need_yields = [n for n in ["10Y(%)", "30Y(%)"] if n not in yields]
    if need_oil or need_yields:
        print(f"  yfinance fallback: oil={need_oil}, yields={need_yields}")
        try:
            import yfinance as yf
            oil_map = {"BZ=F": "Brent($/bbl)", "CL=F": "WTI($/bbl)"}
            yield_map = {"^TNX": "10Y(%)", "^TYX": "30Y(%)"}
            codes = ([c for c, n in oil_map.items() if n in need_oil] +
                     [c for c, n in yield_map.items() if n in need_yields])
            if codes:
                raw = yf.download(codes, start=start, end=end, interval="1d",
                                  auto_adjust=False, group_by="ticker",
                                  progress=False, threads=True)
                for code, name in oil_map.items():
                    if name in need_oil:
                        try:
                            oil[name] = (raw[code]["Close"].dropna() if len(codes) > 1
                                         else raw["Close"].dropna())
                        except KeyError:
                            pass
                for code, name in yield_map.items():
                    if name in need_yields:
                        try:
                            yields[name] = (raw[code]["Close"].dropna() if len(codes) > 1
                                            else raw["Close"].dropna())
                        except KeyError:
                            pass
        except Exception as e:
            print(f"  yfinance fallback 失败: {e}")

    data["oil"] = to_weekly(pd.DataFrame(oil), MACRO_BASE) if oil else pd.DataFrame()
    data["yield"] = to_weekly(pd.DataFrame(yields), MACRO_BASE) if yields else pd.DataFrame()

    # FRED
    all_fred_ids = ([sid for sid, _ in FRED_WEEKLY] +
                    [sid for sid, _, _, _ in FRED_MONTHLY])
    print(f"  FRED {len(all_fred_ids)} 个序列 (并行)...")
    fred = fetch_fred_parallel(all_fred_ids, start="2023-12-01")

    weekly_parts = {}
    for sid, name in FRED_WEEKLY:
        s = fred.get(sid, pd.Series(dtype=float))
        if not s.empty:
            s_cut = s[s.index >= MACRO_BASE]
            if not s_cut.empty:
                weekly_parts[name] = s_cut

    gas_key = "全美汽油均价($/gal)"
    if gas_key in weekly_parts:
        g = weekly_parts.pop(gas_key).to_frame(gas_key)
        data["gas"] = g.resample("W-FRI").last().ffill().dropna()
    else:
        data["gas"] = pd.DataFrame()

    y2_key = "2Y(%)"
    if y2_key in weekly_parts:
        y2 = weekly_parts.pop(y2_key)
        y2_df = y2.to_frame(y2_key)
        y2_weekly = to_weekly(y2_df, MACRO_BASE)
        data["yield"] = data["yield"].join(y2_weekly, how="outer").ffill()
        cols = [c for c in ["2Y(%)", "10Y(%)", "30Y(%)"] if c in data["yield"].columns]
        data["yield"] = data["yield"][cols]

    data["claims"] = pd.DataFrame(weekly_parts) if weekly_parts else pd.DataFrame()

    monthly = {}
    for sid, name, transform, _ in FRED_MONTHLY:
        s = fred.get(sid, pd.Series(dtype=float))
        if s.empty:
            continue
        if transform == "yoy":
            c = s.pct_change(12) * 100
        elif transform == "mom_pct":
            c = s.pct_change(1) * 100
        elif transform == "mom_diff":
            c = s.diff(1)
        else:
            c = s
        c = c[c.index >= MACRO_BASE].dropna()
        if not c.empty:
            monthly[name] = c
    data["monthly"] = pd.DataFrame(monthly) if monthly else pd.DataFrame()
    return data


# ══════════════════════════════════════════════════════════════════════
#  读取 Excel 中的现有数据，确定增量起始日期
# ══════════════════════════════════════════════════════════════════════
def read_last_date(wb, sheet_name="运动服饰 Prices"):
    """从 Prices sheet 读取最后一行日期。"""
    ws = wb[sheet_name]
    last_row = ws.max_row
    for r in range(last_row, 2, -1):
        v = ws.cell(r, 1).value
        if v is not None:
            if isinstance(v, dt.datetime):
                return v.date()
            if isinstance(v, dt.date):
                return v
            try:
                return pd.Timestamp(v).date()
            except Exception:
                pass
    return None


def read_existing_prices(wb, sheet_name, ticker_codes):
    """从 Prices sheet 读取已有数据为 DataFrame。"""
    ws = wb[sheet_name]
    # 读取 header (row 2)
    headers = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(2, c).value
        headers.append(v)

    dates, rows_data = [], []
    for r in range(3, ws.max_row + 1):
        d = ws.cell(r, 1).value
        if d is None:
            continue
        if isinstance(d, dt.datetime):
            d = d.date()
        dates.append(pd.Timestamp(d))
        row = {}
        for c in range(2, ws.max_column + 1):
            col_name = headers[c - 1]
            if col_name:
                v = ws.cell(r, c).value
                row[col_name] = float(v) if v is not None else np.nan
        rows_data.append(row)

    if not dates:
        return pd.DataFrame()
    df = pd.DataFrame(rows_data, index=dates)
    return df


# ══════════════════════════════════════════════════════════════════════
#  增量写入 — Prices / Normalized sheets
# ══════════════════════════════════════════════════════════════════════
def append_prices_rows(wb, sheet_name, new_rows_df):
    """追加新行到 Prices sheet。new_rows_df: index=dates, columns=ticker codes。"""
    from openpyxl.styles import Font
    ws = wb[sheet_name]
    font = Font(name="Arial", size=10)
    blue = Font(name="Arial", size=10, color="0000FF")

    # 读取 header 映射
    col_map = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(2, c).value
        if v:
            col_map[v] = c

    start_row = ws.max_row + 1
    for i, (d, row) in enumerate(new_rows_df.iterrows()):
        r = start_row + i
        cell = ws.cell(r, 1, value=d.date() if hasattr(d, 'date') else d)
        cell.number_format = "yyyy-mm-dd"
        cell.font = font
        for ticker, col_idx in col_map.items():
            if ticker == "Date":
                continue
            v = row.get(ticker)
            if v is not None and not (isinstance(v, float) and np.isnan(v)):
                cell = ws.cell(r, col_idx, value=float(v))
                cell.font = blue
                cell.number_format = "#,##0.00"
    return start_row, start_row + len(new_rows_df) - 1


def append_normalized_rows(wb, sheet_name, prices_sheet_name, new_count):
    """追加新行到 Normalized sheet，使用公式引用 Prices。"""
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
    ws = wb[sheet_name]
    ws_p = wb[prices_sheet_name]
    font = Font(name="Arial", size=10)
    prefix = prices_sheet_name  # e.g. "运动服饰 Prices"

    # 从 Prices sheet 读日期 + 确定列数
    p_max_row = ws_p.max_row
    p_max_col = ws_p.max_column

    # Normalized 当前最后行
    n_start = ws.max_row + 1
    # Prices 新行起始 = p_max_row - new_count + 1
    p_start = p_max_row - new_count + 1

    for i in range(new_count):
        r_n = n_start + i
        r_p = p_start + i
        # 日期 — 引用 Prices 的日期
        date_val = ws_p.cell(r_p, 1).value
        cell = ws.cell(r_n, 1, value=date_val)
        cell.number_format = "yyyy-mm-dd"
        cell.font = font
        # 公式列
        for j in range(2, p_max_col + 1):
            cl = get_column_letter(j)
            ws.cell(r_n, j,
                    value=f"=IF(OR('{prefix}'!{cl}{r_p}=\"\","
                          f"'{prefix}'!{cl}$3=\"\"),\"\","
                          f"'{prefix}'!{cl}{r_p}/'{prefix}'!{cl}$3)")
            ws.cell(r_n, j).font = font
            ws.cell(r_n, j).number_format = "0.000"


# ══════════════════════════════════════════════════════════════════════
#  增量写入 — 行业指数 sheet
# ══════════════════════════════════════════════════════════════════════
def append_index_rows(wb, sw_weekly_new, lux_weekly_new, sw_weekly_all, lux_weekly_all):
    """追加新行到行业指数 sheet。"""
    from openpyxl.styles import Font
    ws = wb["行业指数"]
    font = Font(name="Arial", size=10)
    blue = Font(name="Arial", size=10, color="0000FF")

    def make_index(weekly_all, tickers):
        cols = [t[0] for t in tickers]
        df = weekly_all.reindex(columns=cols)
        norm = df / df.iloc[0]
        return norm.mean(axis=1, skipna=True)

    sw_idx = make_index(sw_weekly_all, SPORTSWEAR)
    lux_idx = make_index(lux_weekly_all, LUXURY)

    # 只取新增日期
    new_dates = sw_weekly_new.index
    start_row = ws.max_row + 1

    for i, d in enumerate(new_dates):
        r = start_row + i
        cell = ws.cell(r, 1, value=d.date() if hasattr(d, 'date') else d)
        cell.number_format = "yyyy-mm-dd"
        cell.font = font
        sv = sw_idx.get(d)
        if sv is not None and not np.isnan(sv):
            c = ws.cell(r, 2, value=round(float(sv), 4))
            c.font = blue; c.number_format = "0.0000"
        lv = lux_idx.get(d)
        if lv is not None and not np.isnan(lv):
            c = ws.cell(r, 3, value=round(float(lv), 4))
            c.font = blue; c.number_format = "0.0000"


# ══════════════════════════════════════════════════════════════════════
#  增量写入 — ETFs sheet
# ══════════════════════════════════════════════════════════════════════
def append_etf_rows(wb, etf_weekly_new, etf_weekly_all):
    """追加新行到 可选消费ETFs sheet（价格 + 归一化列）。"""
    from openpyxl.styles import Font
    ws_name = "可选消费ETFs"
    if ws_name not in wb.sheetnames:
        print(f"  警告：找不到 {ws_name} sheet，跳过")
        return
    ws = wb[ws_name]
    font = Font(name="Arial", size=10)
    blue = Font(name="Arial", size=10, color="0000FF")

    etf_codes = [t[0] for t in ETFS]
    etf_norm_all = etf_weekly_all.reindex(columns=etf_codes)
    etf_norm_all = etf_norm_all / etf_norm_all.iloc[0]

    # 找到数据区域的最后一行（在统计/说明之前）
    # header in row 2, data starts row 3
    # 找最后一个有日期的行
    last_data_row = 2
    for r in range(3, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v is None:
            break
        if isinstance(v, (dt.datetime, dt.date)):
            last_data_row = r
        else:
            try:
                pd.Timestamp(v)
                last_data_row = r
            except Exception:
                break

    sep = len(ETFS) + 2  # 空列位置

    start_row = last_data_row + 1
    for i, d in enumerate(etf_weekly_new.index):
        r = start_row + i
        cell = ws.cell(r, 1, value=d.date() if hasattr(d, 'date') else d)
        cell.number_format = "yyyy-mm-dd"
        cell.font = font
        for j, code in enumerate(etf_codes, start=2):
            v = etf_weekly_new.loc[d, code] if code in etf_weekly_new.columns else np.nan
            if not np.isnan(v):
                c = ws.cell(r, j, value=float(v))
                c.font = blue; c.number_format = "#,##0.00"
        for j, code in enumerate(etf_codes):
            nv = etf_norm_all.loc[d, code] if (code in etf_norm_all.columns
                                                and d in etf_norm_all.index) else np.nan
            if not np.isnan(nv):
                c = ws.cell(r, sep + j, value=round(float(nv), 4))
                c.font = font; c.number_format = "0.0000"


# ══════════════════════════════════════════════════════════════════════
#  增量写入 — 周涨跌 sheet
# ══════════════════════════════════════════════════════════════════════
def _week_label(d):
    """生成周涨跌列标签：YYMMDD-MMDD（Mon-Sun，与原始 Excel 格式一致）。"""
    fri = pd.Timestamp(d)
    mon = fri - pd.Timedelta(days=4)
    sun = fri + pd.Timedelta(days=2)
    return f"{mon.strftime('%y%m%d')}-{sun.strftime('%m%d')}"


def append_weekly_changes(wb, sw_prices_all, lux_prices_all, macro):
    """追加新列到周涨跌 sheet。"""
    from openpyxl.styles import Font
    ws = wb["周涨跌"]
    font = Font(name="Arial", size=10)

    # 现有列数
    existing_cols = ws.max_column

    # 读取已有的日期标签，用 Monday 起始日（前6字符）做去重，
    # 避免 Mon-Fri vs Mon-Sun 格式差异导致误判
    existing_mondays = set()
    for c in range(2, existing_cols + 1):
        v = ws.cell(2, c).value
        if v:
            existing_mondays.add(str(v)[:6])

    # 计算所有股票的周回报
    sw_codes = [t[0] for t in SPORTSWEAR]
    lux_codes = [t[0] for t in LUXURY]

    sw_ret = sw_prices_all.reindex(columns=sw_codes).pct_change()
    lux_ret = lux_prices_all.reindex(columns=lux_codes).pct_change()

    # 确定新增周（按 Monday 起始日去重）
    all_dates = sw_ret.index[1:]  # skip first (NaN from pct_change)
    new_dates = []
    for d in all_dates:
        label = _week_label(d)
        if label[:6] not in existing_mondays:
            new_dates.append(d)

    if not new_dates:
        print("  周涨跌：无新增周")
        return 0

    # 准备宏观数据
    oil_weekly = macro.get("oil", pd.DataFrame())
    gas_weekly = macro.get("gas", pd.DataFrame())
    yield_weekly = macro.get("yield", pd.DataFrame())
    claims = macro.get("claims", pd.DataFrame())

    # 计算宏观周变动
    oil_all = oil_weekly.copy()
    if not gas_weekly.empty:
        oil_all = oil_all.join(gas_weekly, how="outer").ffill()

    for col_offset, d in enumerate(new_dates):
        c = existing_cols + 1 + col_offset
        label = _week_label(d)

        # Row 2: date label
        ws.cell(2, c, value=label).font = Font(name="Arial", size=10, bold=True)
        # Row 56: macro date label
        ws.cell(WC_ROW_MACRO_HDR, c, value=label).font = Font(name="Arial", size=10, bold=True)
        # Row 64: yield date label
        ws.cell(WC_ROW_YIELD_HDR, c, value=label).font = Font(name="Arial", size=10, bold=True)

        # Sportswear returns
        for i, code in enumerate(sw_codes):
            row = WC_ROW_SW_START + i
            v = sw_ret.loc[d, code] if (code in sw_ret.columns and d in sw_ret.index) else np.nan
            if not np.isnan(v):
                ws.cell(row, c, value=round(float(v), 6)).font = font

        # Luxury returns
        for i, code in enumerate(lux_codes):
            row = WC_ROW_LUX_START + i
            v = lux_ret.loc[d, code] if (code in lux_ret.columns and d in lux_ret.index) else np.nan
            if not np.isnan(v):
                ws.cell(row, c, value=round(float(v), 6)).font = font

        # Macro — oil/gas levels
        for row_n, col_name in [(WC_ROW_BRENT, "Brent($/bbl)"),
                                 (WC_ROW_WTI, "WTI($/bbl)"),
                                 (WC_ROW_GAS, "全美汽油均价($/gal)")]:
            src = oil_all if col_name != "全美汽油均价($/gal)" else gas_weekly
            if not src.empty and col_name in src.columns and d in src.index:
                v = src.loc[d, col_name]
                if not np.isnan(v):
                    ws.cell(row_n, c, value=round(float(v), 2)).font = font

        # Oil/gas weekly changes
        if not oil_all.empty:
            for row_n, col_name in [(WC_ROW_BRENT_CHG, "Brent($/bbl)"),
                                     (WC_ROW_WTI_CHG, "WTI($/bbl)")]:
                if col_name in oil_all.columns and d in oil_all.index:
                    prev_idx = oil_all.index.get_loc(d)
                    if prev_idx > 0:
                        prev_v = oil_all.iloc[prev_idx - 1][col_name]
                        cur_v = oil_all.loc[d, col_name]
                        if not np.isnan(prev_v) and not np.isnan(cur_v) and prev_v != 0:
                            chg = (cur_v - prev_v) / prev_v * 100
                            ws.cell(row_n, c, value=round(float(chg), 2)).font = font

        if not gas_weekly.empty and "全美汽油均价($/gal)" in gas_weekly.columns and d in gas_weekly.index:
            prev_idx = gas_weekly.index.get_loc(d)
            if prev_idx > 0:
                prev_v = gas_weekly.iloc[prev_idx - 1]["全美汽油均价($/gal)"]
                cur_v = gas_weekly.loc[d, "全美汽油均价($/gal)"]
                if not np.isnan(prev_v) and not np.isnan(cur_v) and prev_v != 0:
                    chg = (cur_v - prev_v) / prev_v * 100
                    ws.cell(WC_ROW_GAS_CHG, c, value=round(float(chg), 2)).font = font

        # Yields levels
        if not yield_weekly.empty:
            for row_n, col_name in [(WC_ROW_2Y, "2Y(%)"),
                                     (WC_ROW_10Y, "10Y(%)"),
                                     (WC_ROW_30Y, "30Y(%)")]:
                if col_name in yield_weekly.columns and d in yield_weekly.index:
                    v = yield_weekly.loc[d, col_name]
                    if not np.isnan(v):
                        ws.cell(row_n, c, value=round(float(v), 2)).font = font

            # Yield weekly changes in bps
            for row_n, col_name in [(WC_ROW_2Y_CHG, "2Y(%)"),
                                     (WC_ROW_10Y_CHG, "10Y(%)"),
                                     (WC_ROW_30Y_CHG, "30Y(%)")]:
                if col_name in yield_weekly.columns and d in yield_weekly.index:
                    idx = yield_weekly.index.get_loc(d)
                    if idx > 0:
                        prev_v = yield_weekly.iloc[idx - 1][col_name]
                        cur_v = yield_weekly.loc[d, col_name]
                        if not np.isnan(prev_v) and not np.isnan(cur_v):
                            bps = round((cur_v - prev_v) * 100)
                            ws.cell(row_n, c, value=int(bps)).font = font

        # Claims
        if not claims.empty and "初请失业金(千人)" in claims.columns:
            # 找最接近的日期
            claim_dates = claims.index
            closest = claim_dates[claim_dates <= d]
            if len(closest) > 0:
                cd = closest[-1]
                # 只取这一周范围内的
                if (d - cd).days <= 7:
                    v = claims.loc[cd, "初请失业金(千人)"]
                    if not np.isnan(v):
                        ws.cell(WC_ROW_CLAIMS, c, value=round(float(v), 0)).font = font

    print(f"  周涨跌：追加 {len(new_dates)} 列")
    return len(new_dates)


# ══════════════════════════════════════════════════════════════════════
#  更新宏观详细 sheets（油价汽油、美债收益率、宏观高频数据）
# ══════════════════════════════════════════════════════════════════════
def update_macro_sheets(wb, macro):
    """重写宏观详细 sheets（全量覆盖，数据量不大）。"""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.chart import LineChart, Reference
    from openpyxl.utils import get_column_letter

    font = Font(name="Arial", size=10)
    bold = Font(name="Arial", size=10, bold=True)
    blue = Font(name="Arial", size=10, color="0000FF")
    hd = PatternFill("solid", start_color="DDEBF7")
    green = PatternFill("solid", start_color="E2EFDA")
    orange = PatternFill("solid", start_color="FCE4D6")

    def ensure(name):
        if name in wb.sheetnames:
            del wb[name]
        return wb.create_sheet(name)

    def write_ts(ws, df, start_row, num_fmt="#,##0.00"):
        cols = list(df.columns)
        ws.cell(row=start_row, column=1, value="Date").font = bold
        ws.cell(row=start_row, column=1).fill = hd
        for j, c in enumerate(cols, 2):
            cell = ws.cell(row=start_row, column=j, value=c)
            cell.font = bold; cell.fill = hd
            cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions["A"].width = 12
        for j in range(2, len(cols) + 2):
            ws.column_dimensions[get_column_letter(j)].width = 18
        for i, (d, row) in enumerate(df.iterrows(), start=start_row + 1):
            ws.cell(row=i, column=1, value=d.date()).number_format = "yyyy-mm-dd"
            ws.cell(row=i, column=1).font = font
            for j, c in enumerate(cols, 2):
                v = row[c]
                cell = ws.cell(row=i, column=j,
                               value=None if pd.isna(v) else round(float(v), 2))
                cell.font = blue; cell.number_format = num_fmt
        return start_row + len(df)

    def add_chart(ws, title, n_cols, hdr_row, last_row, anchor):
        ch = LineChart()
        ch.title = title; ch.height, ch.width = 12, 22
        ch.y_axis.number_format = "0.0"
        for j in range(2, 2 + n_cols):
            ref = Reference(ws, min_col=j, min_row=hdr_row, max_row=last_row)
            ch.add_data(ref, titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=1, min_row=hdr_row + 1,
                                     max_row=last_row))
        for s in ch.series:
            s.smooth = False; s.graphicalProperties.line.width = 20000
        ws.add_chart(ch, anchor)

    # 油价汽油
    ws = ensure("油价汽油")
    ws["A1"] = "油价 & 汽油价格（周度，2025-01-01 起）"; ws["A1"].font = bold
    oil = macro["oil"]; gas = macro["gas"]
    combined = oil.copy()
    if not gas.empty:
        combined = combined.join(gas, how="outer").ffill()
    if not combined.empty:
        last = write_ts(ws, combined, 2)
        nc = len(combined.columns)
        add_chart(ws, "油价 & 汽油", nc, 2, last,
                  f"{get_column_letter(nc + 3)}2")

    # 美债收益率
    ws = ensure("美债收益率")
    ws["A1"] = "美债收益率 2Y/10Y/30Y（周度，2025-01-01 起）"; ws["A1"].font = bold
    yld = macro["yield"]
    if not yld.empty:
        last = write_ts(ws, yld, 2, num_fmt="0.000")
        nc_yld = len(yld.columns)
        add_chart(ws, "US Treasury Yields", nc_yld, 2, last,
                  f"{get_column_letter(nc_yld + 3)}2")
        r = last + 3
    else:
        r = 5
    ws.cell(row=r, column=1, value="加息/降息重要变化").font = bold
    ws.cell(row=r, column=1).fill = orange

    # 地缘政治 — 保留现有内容不覆盖
    if "地缘政治" not in wb.sheetnames:
        ws = wb.create_sheet("地缘政治")
        ws["A1"] = "重要地缘政治事件"; ws["A1"].font = bold; ws["A1"].fill = orange
        ws.cell(row=2, column=1, value="时间点/时间段").font = bold
        ws.cell(row=3, column=1, value="详情").font = bold

    # 宏观高频数据
    ws = ensure("宏观高频数据")
    ws["A1"] = "美国宏观高频数据"; ws["A1"].font = bold

    r = 3
    ws.cell(row=r, column=1, value="一、周度数据（初请失业金）").font = bold
    ws.cell(row=r, column=1).fill = green; r += 1
    claims = macro["claims"]
    if not claims.empty:
        r = write_ts(ws, claims, r, num_fmt="#,##0") + 1
    else:
        ws.cell(row=r, column=1, value="（FRED 数据不可用）").font = font; r += 1

    r += 2
    ws.cell(row=r, column=1, value="二、月度数据").font = bold
    ws.cell(row=r, column=1).fill = green; r += 1
    monthly = macro["monthly"]
    if not monthly.empty:
        fmt_map = {n: f for _, n, _, f in FRED_MONTHLY}
        cols = list(monthly.columns)
        ws.cell(row=r, column=1, value="Date").font = bold
        ws.cell(row=r, column=1).fill = hd
        for j, c in enumerate(cols, 2):
            cell = ws.cell(row=r, column=j, value=c)
            cell.font = bold; cell.fill = hd
        for j in range(2, len(cols) + 2):
            ws.column_dimensions[get_column_letter(j)].width = 16
        r += 1
        for d, row in monthly.iterrows():
            ws.cell(row=r, column=1, value=d.date()).number_format = "yyyy-mm-dd"
            ws.cell(row=r, column=1).font = font
            for j, c in enumerate(cols, 2):
                v = row[c]
                cell = ws.cell(row=r, column=j,
                               value=None if pd.isna(v) else round(float(v), 1))
                cell.font = blue; cell.number_format = fmt_map.get(c, "0.0")
            r += 1

    r += 2
    ws.cell(row=r, column=1,
            value="说明：CPI/PPI/PCE/信用卡余额为同比%；零售销售为环比%；非农为环比变化千人。"
                  "来源 FRED，日期为统计期非发布日。"
                  "信用卡余额(REVOLSL)=循环信贷总额，反映消费者信用卡支出趋势。").font = font


# ══════════════════════════════════════════════════════════════════════
#  更新 Charts 数据范围
# ══════════════════════════════════════════════════════════════════════
def update_charts(wb):
    """重建 Charts sheet（图表引用范围可能变了）。"""
    from openpyxl.chart import LineChart, Reference
    from openpyxl.styles import Font

    font = Font(name="Arial", size=10)

    if "Charts" in wb.sheetnames:
        del wb["Charts"]
    ws_ch = wb.create_sheet("Charts")

    sw_n = wb["运动服饰 Normalized"]
    lux_n = wb["奢侈品 Normalized"]
    sw_last = sw_n.max_row
    lux_last = lux_n.max_row

    sw_weekly_cols = []
    for c in range(2, sw_n.max_column + 1):
        v = sw_n.cell(2, c).value
        if v:
            sw_weekly_cols.append(v)

    lux_weekly_cols = []
    for c in range(2, lux_n.max_column + 1):
        v = lux_n.cell(2, c).value
        if v:
            lux_weekly_cols.append(v)

    anchor = 1

    # 运动服饰分区域
    for reg in ["US/EU", "Japan", "China/HK"]:
        ch = LineChart()
        ch.title = f"运动服饰 {reg}（{BASE_DATE} = 1）"
        ch.height, ch.width = 11, 26
        ch.y_axis.number_format = "0.00"
        for j, (u, y, b, r, _) in enumerate(SPORTSWEAR, start=2):
            if r != reg:
                continue
            if u not in sw_weekly_cols:
                continue
            ref = Reference(sw_n, min_col=j, min_row=2, max_row=sw_last)
            ch.add_data(ref, titles_from_data=True)
        ch.set_categories(Reference(sw_n, min_col=1, min_row=3, max_row=sw_last))
        for s in ch.series:
            s.smooth = False; s.marker.symbol = "none"
        ws_ch.add_chart(ch, f"A{anchor}")
        anchor += 22

    # 奢侈品整体
    ch = LineChart()
    ch.title = f"奢侈品（{BASE_DATE} = 1）"
    ch.height, ch.width = 11, 26
    ch.y_axis.number_format = "0.00"
    for j, (u, y, b, r, _) in enumerate(LUXURY, start=2):
        if u not in lux_weekly_cols:
            continue
        ref = Reference(lux_n, min_col=j, min_row=2, max_row=lux_last)
        ch.add_data(ref, titles_from_data=True)
    ch.set_categories(Reference(lux_n, min_col=1, min_row=3, max_row=lux_last))
    for s in ch.series:
        s.smooth = False; s.marker.symbol = "none"
    ws_ch.add_chart(ch, f"A{anchor}")

    # 走势大图
    for label, tickers_list, ws_norm_name in [
        ("运动服饰走势", SPORTSWEAR, "运动服饰 Normalized"),
        ("奢侈品走势", LUXURY, "奢侈品 Normalized"),
    ]:
        if label in wb.sheetnames:
            del wb[label]
        ws_big = wb.create_sheet(label)
        ws_norm = wb[ws_norm_name]
        last_row = ws_norm.max_row

        norm_cols = []
        for c in range(2, ws_norm.max_column + 1):
            v = ws_norm.cell(2, c).value
            if v:
                norm_cols.append(v)

        ch = LineChart()
        ch.title = f"{label}（{BASE_DATE} = 1）"
        ch.height, ch.width = 25, 40
        ch.y_axis.number_format = "0.00"
        ch.y_axis.title = "归一化"
        ch.x_axis.number_format = "mm-dd"
        ch.legend.position = "b"
        added = 0
        for j, (u, y, b, r, _) in enumerate(tickers_list, start=2):
            if u not in norm_cols:
                continue
            ref = Reference(ws_norm, min_col=j, min_row=2, max_row=last_row)
            ch.add_data(ref, titles_from_data=True)
            added += 1
        ch.set_categories(Reference(ws_norm, min_col=1, min_row=3, max_row=last_row))
        for s in ch.series:
            s.smooth = False
            s.marker.symbol = "none"
            s.graphicalProperties.line.width = 15000
        ws_big.add_chart(ch, "A1")
        ws_big.cell(row=1, column=20,
                    value=f"共 {added} 只，基期 {BASE_DATE} = 1").font = font


# ══════════════════════════════════════════════════════════════════════
#  全量重建 — 当 --full 或 Excel 不存在时
# ══════════════════════════════════════════════════════════════════════
def full_rebuild(xlsx_path, demo=False):
    """全量重建所有 sheets（fallback 模式）。"""
    from openpyxl import Workbook
    from openpyxl.chart import LineChart, Reference
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    print("\n=== 全量重建模式 ===")

    def get_data(tickers, base, label):
        codes = [t[1] for t in tickers if t[1]]
        print(f"\n拉取 {label}（{len(codes)} 只）...")
        if demo:
            daily = _fake_daily(codes, base)
        else:
            daily = fetch_daily(codes, base)
        weekly = to_weekly(daily, base)
        y2u = {t[1]: t[0] for t in tickers if t[1]}
        weekly = weekly.rename(columns=y2u)
        return weekly

    sw_weekly = get_data(SPORTSWEAR, BASE_DATE, "运动服饰")
    lux_weekly = get_data(LUXURY, BASE_DATE, "奢侈品")

    etf_codes = [t[1] for t in ETFS]
    print(f"\n拉取 ETF...")
    if demo:
        etf_daily = _fake_daily(etf_codes, ETF_BASE_DATE)
    else:
        etf_daily = fetch_daily(etf_codes, ETF_BASE_DATE)
    etf_weekly = to_weekly(etf_daily, ETF_BASE_DATE)

    print("  全量重建模式暂不支持（请确保 Excel 文件已存在）。")
    print("  如需重建，请从备份恢复 Excel 或联系 David。")
    return

    # 宏观
    print("\n拉取宏观数据...")
    macro = fetch_macro() if not demo else _fake_macro()
    from openpyxl import load_workbook
    wb = load_workbook(xlsx_path)

    # 写周涨跌
    _build_weekly_changes_sheet(wb, sw_weekly, lux_weekly, macro)
    update_macro_sheets(wb, macro)
    wb.save(xlsx_path)
    print(f"\n全量重建完成：{xlsx_path}")


def _fake_daily(codes, base):
    idx = pd.bdate_range(pd.Timestamp(base) - pd.Timedelta(days=10), dt.date.today())
    rng = np.random.default_rng(42)
    data = {c: 50 * np.exp(np.cumsum(rng.normal(0, 0.02, len(idx)))) for c in codes}
    return pd.DataFrame(data, index=idx)


def _fake_macro():
    idx_w = pd.date_range(MACRO_BASE, dt.date.today(), freq="W-FRI")
    idx_m = pd.date_range(MACRO_BASE, dt.date.today(), freq="MS")
    rng = np.random.default_rng(99)
    return {
        "oil": pd.DataFrame({
            "Brent($/bbl)": 75 + np.cumsum(rng.normal(0, 1, len(idx_w))),
            "WTI($/bbl)": 72 + np.cumsum(rng.normal(0, 1, len(idx_w))),
        }, index=idx_w),
        "gas": pd.DataFrame({
            "全美汽油均价($/gal)": 3.2 + np.cumsum(rng.normal(0, 0.03, len(idx_w))),
        }, index=idx_w),
        "yield": pd.DataFrame({
            "2Y(%)": 4.0 + np.cumsum(rng.normal(0, 0.04, len(idx_w))),
            "10Y(%)": 4.5 + np.cumsum(rng.normal(0, 0.05, len(idx_w))),
            "30Y(%)": 4.8 + np.cumsum(rng.normal(0, 0.06, len(idx_w))),
        }, index=idx_w),
        "claims": pd.DataFrame({
            "初请失业金(千人)": 220 + rng.normal(0, 10, len(idx_w)),
        }, index=idx_w),
        "monthly": pd.DataFrame({n: rng.normal(2, 0.5, len(idx_m))
                                  for _, n, _, _ in FRED_MONTHLY}, index=idx_m),
    }


def _build_weekly_changes_sheet(wb, sw_weekly, lux_weekly, macro):
    """全量构建周涨跌 sheet。"""
    from openpyxl.styles import Font
    if "周涨跌" in wb.sheetnames:
        del wb["周涨跌"]
    ws = wb.create_sheet("周涨跌", 0)

    font = Font(name="Arial", size=10)
    bold = Font(name="Arial", size=10, bold=True)

    ws.cell(1, 1, value="周涨跌幅").font = bold
    ws.cell(2, 1, value="Ticker").font = bold
    ws.column_dimensions["A"].width = 14

    sw_codes = [t[0] for t in SPORTSWEAR]
    lux_codes = [t[0] for t in LUXURY]

    # Section labels
    ws.cell(3, 1, value="运动服饰").font = bold
    for i, code in enumerate(sw_codes):
        ws.cell(WC_ROW_SW_START + i, 1, value=code).font = bold

    ws.cell(38, 1, value="奢侈品").font = bold
    for i, code in enumerate(lux_codes):
        ws.cell(WC_ROW_LUX_START + i, 1, value=code).font = bold

    ws.cell(WC_ROW_MACRO_LABEL, 1, value="宏观指标").font = bold
    macro_row_labels = {
        WC_ROW_BRENT: "Brent($/bbl)", WC_ROW_WTI: "WTI($/bbl)",
        WC_ROW_GAS: "汽油($/gal)",
        WC_ROW_BRENT_CHG: "Brent周涨跌", WC_ROW_WTI_CHG: "WTI周涨跌",
        WC_ROW_GAS_CHG: "汽油周涨跌",
        WC_ROW_2Y: "2Y收益率(%)", WC_ROW_10Y: "10Y收益率(%)", WC_ROW_30Y: "30Y收益率(%)",
        WC_ROW_2Y_CHG: "2Y收益率周变动(bps)", WC_ROW_10Y_CHG: "10Y收益率周变动(bps)",
        WC_ROW_30Y_CHG: "30Y收益率周变动(bps)",
        WC_ROW_CLAIMS: "初请失业金(千人)",
    }
    for row, label in macro_row_labels.items():
        ws.cell(row, 1, value=label).font = bold

    # 写入所有周的数据
    sw_ret = sw_weekly.reindex(columns=sw_codes).pct_change()
    lux_ret = lux_weekly.reindex(columns=lux_codes).pct_change()

    all_dates = sw_ret.index[1:]

    oil_weekly = macro.get("oil", pd.DataFrame())
    gas_weekly = macro.get("gas", pd.DataFrame())
    yield_weekly = macro.get("yield", pd.DataFrame())
    claims = macro.get("claims", pd.DataFrame())

    oil_all = oil_weekly.copy()
    if not gas_weekly.empty:
        oil_all = oil_all.join(gas_weekly, how="outer").ffill()

    for col_offset, d in enumerate(all_dates):
        c = 2 + col_offset
        label = _week_label(d)
        ws.cell(2, c, value=label).font = bold
        ws.cell(WC_ROW_MACRO_HDR, c, value=label).font = bold
        ws.cell(WC_ROW_YIELD_HDR, c, value=label).font = bold

        for i, code in enumerate(sw_codes):
            v = sw_ret.loc[d, code] if code in sw_ret.columns else np.nan
            if not np.isnan(v):
                ws.cell(WC_ROW_SW_START + i, c, value=round(float(v), 6)).font = font

        for i, code in enumerate(lux_codes):
            v = lux_ret.loc[d, code] if code in lux_ret.columns else np.nan
            if not np.isnan(v):
                ws.cell(WC_ROW_LUX_START + i, c, value=round(float(v), 6)).font = font

        # Macro levels and changes (same logic as append)
        for row_n, col_name in [(WC_ROW_BRENT, "Brent($/bbl)"),
                                 (WC_ROW_WTI, "WTI($/bbl)")]:
            if not oil_all.empty and col_name in oil_all.columns and d in oil_all.index:
                v = oil_all.loc[d, col_name]
                if not np.isnan(v):
                    ws.cell(row_n, c, value=round(float(v), 2)).font = font

        if not gas_weekly.empty and "全美汽油均价($/gal)" in gas_weekly.columns and d in gas_weekly.index:
            v = gas_weekly.loc[d, "全美汽油均价($/gal)"]
            if not np.isnan(v):
                ws.cell(WC_ROW_GAS, c, value=round(float(v), 2)).font = font

        if not oil_all.empty:
            for row_n, col_name in [(WC_ROW_BRENT_CHG, "Brent($/bbl)"),
                                     (WC_ROW_WTI_CHG, "WTI($/bbl)")]:
                if col_name in oil_all.columns and d in oil_all.index:
                    idx = oil_all.index.get_loc(d)
                    if idx > 0:
                        prev_v = oil_all.iloc[idx - 1][col_name]
                        cur_v = oil_all.loc[d, col_name]
                        if not np.isnan(prev_v) and not np.isnan(cur_v) and prev_v != 0:
                            chg = (cur_v - prev_v) / prev_v * 100
                            ws.cell(row_n, c, value=round(float(chg), 2)).font = font

        if not gas_weekly.empty and "全美汽油均价($/gal)" in gas_weekly.columns and d in gas_weekly.index:
            idx = gas_weekly.index.get_loc(d)
            if idx > 0:
                prev_v = gas_weekly.iloc[idx - 1]["全美汽油均价($/gal)"]
                cur_v = gas_weekly.loc[d, "全美汽油均价($/gal)"]
                if not np.isnan(prev_v) and not np.isnan(cur_v) and prev_v != 0:
                    chg = (cur_v - prev_v) / prev_v * 100
                    ws.cell(WC_ROW_GAS_CHG, c, value=round(float(chg), 2)).font = font

        if not yield_weekly.empty:
            for row_n, col_name in [(WC_ROW_2Y, "2Y(%)"),
                                     (WC_ROW_10Y, "10Y(%)"),
                                     (WC_ROW_30Y, "30Y(%)")]:
                if col_name in yield_weekly.columns and d in yield_weekly.index:
                    v = yield_weekly.loc[d, col_name]
                    if not np.isnan(v):
                        ws.cell(row_n, c, value=round(float(v), 2)).font = font

            for row_n, col_name in [(WC_ROW_2Y_CHG, "2Y(%)"),
                                     (WC_ROW_10Y_CHG, "10Y(%)"),
                                     (WC_ROW_30Y_CHG, "30Y(%)")]:
                if col_name in yield_weekly.columns and d in yield_weekly.index:
                    idx = yield_weekly.index.get_loc(d)
                    if idx > 0:
                        prev_v = yield_weekly.iloc[idx - 1][col_name]
                        cur_v = yield_weekly.loc[d, col_name]
                        if not np.isnan(prev_v) and not np.isnan(cur_v):
                            bps = round((cur_v - prev_v) * 100)
                            ws.cell(row_n, c, value=int(bps)).font = font

        if not claims.empty and "初请失业金(千人)" in claims.columns:
            closest = claims.index[claims.index <= d]
            if len(closest) > 0:
                cd = closest[-1]
                if (d - cd).days <= 7:
                    v = claims.loc[cd, "初请失业金(千人)"]
                    if not np.isnan(v):
                        ws.cell(WC_ROW_CLAIMS, c, value=round(float(v), 0)).font = font


# ══════════════════════════════════════════════════════════════════════
#  主流程 — 增量更新
# ══════════════════════════════════════════════════════════════════════
def main():
    ap = argparse.ArgumentParser(description="增量更新 Apparel_updating file.xlsx")
    ap.add_argument("--demo", action="store_true", help="使用随机数据测试")
    ap.add_argument("--full", action="store_true", help="强制全量重建")
    ap.add_argument("--xlsx", default=XLSX_PATH)
    args = ap.parse_args()

    xlsx = Path(args.xlsx)
    if not xlsx.exists() or args.full:
        full_rebuild(args.xlsx, demo=args.demo)
        return

    from openpyxl import load_workbook

    print("=== 增量更新模式 ===")
    print(f"  文件：{args.xlsx}")

    wb = load_workbook(args.xlsx)
    last_date = read_last_date(wb)
    print(f"  现有最新日期：{last_date}")

    if last_date is None:
        print("  无法读取日期，切换全量重建")
        wb.close()
        full_rebuild(args.xlsx, demo=args.demo)
        return

    # 增量起始日期 = 最后日期的下一天
    start_date = last_date + dt.timedelta(days=1)
    today = dt.date.today()

    if start_date > today:
        print(f"  数据已是最新（{last_date}），无需更新")
        # 仍然更新宏观数据
        print("\n=== 更新宏观数据 ===")
        macro = fetch_macro() if not args.demo else _fake_macro()
        append_weekly_changes(wb, pd.DataFrame(), pd.DataFrame(), macro)
        update_macro_sheets(wb, macro)
        update_charts(wb)
        wb.save(args.xlsx)
        print(f"\n完成：{args.xlsx}")
        return

    print(f"  增量范围：{start_date} → {today}")

    # ── 1. 拉取增量日线数据 ──
    def get_incremental(tickers, base_date, label):
        codes = [t[1] for t in tickers if t[1]]
        print(f"\n拉取 {label}（{len(codes)} 只，从 {base_date} 起）...")
        if args.demo:
            daily = _fake_daily(codes, base_date)
        else:
            daily = fetch_daily(codes, base_date)
        weekly = to_weekly(daily, base_date)
        y2u = {t[1]: t[0] for t in tickers if t[1]}
        weekly = weekly.rename(columns=y2u)
        return weekly

    # 拉完整数据（需要基期起的全量来算归一化和周涨跌）
    sw_weekly_all = get_incremental(SPORTSWEAR, BASE_DATE, "运动服饰")
    lux_weekly_all = get_incremental(LUXURY, BASE_DATE, "奢侈品")

    etf_codes = [t[1] for t in ETFS]
    print(f"\n拉取 ETF...")
    if args.demo:
        etf_daily = _fake_daily(etf_codes, ETF_BASE_DATE)
    else:
        etf_daily = fetch_daily(etf_codes, ETF_BASE_DATE)
    etf_weekly_all = to_weekly(etf_daily, ETF_BASE_DATE)

    # ── 2. 筛选新增行 ──
    last_ts = pd.Timestamp(last_date)
    sw_new = sw_weekly_all[sw_weekly_all.index > last_ts]
    lux_new = lux_weekly_all[lux_weekly_all.index > last_ts]
    etf_new = etf_weekly_all[etf_weekly_all.index > last_ts]

    n_new = len(sw_new)
    print(f"\n新增 {n_new} 周数据")

    if n_new > 0:
        # ── 3. 追加 Prices ──
        print("\n=== 追加 Prices ===")
        sw_cols = [t[0] for t in SPORTSWEAR]
        lux_cols = [t[0] for t in LUXURY]
        append_prices_rows(wb, "运动服饰 Prices", sw_new.reindex(columns=sw_cols))
        append_prices_rows(wb, "奢侈品 Prices", lux_new.reindex(columns=lux_cols))
        print(f"  运动服饰 Prices: +{n_new} 行")
        print(f"  奢侈品 Prices: +{len(lux_new)} 行")

        # ── 4. 追加 Normalized ──
        print("\n=== 追加 Normalized ===")
        append_normalized_rows(wb, "运动服饰 Normalized", "运动服饰 Prices", n_new)
        append_normalized_rows(wb, "奢侈品 Normalized", "奢侈品 Prices", len(lux_new))

        # ── 5. 追加行业指数 ──
        print("\n=== 追加行业指数 ===")
        append_index_rows(wb, sw_new, lux_new, sw_weekly_all, lux_weekly_all)

        # ── 6. 追加 ETFs ──
        print("\n=== 追加 ETFs ===")
        append_etf_rows(wb, etf_new, etf_weekly_all)

    # ── 7. 宏观数据 ──
    print("\n=== 拉取宏观数据 ===")
    macro = fetch_macro() if not args.demo else _fake_macro()
    for k, df in macro.items():
        n = len(df) if isinstance(df, pd.DataFrame) else 0
        print(f"  {k}: {n} 行")

    # ── 8. 周涨跌 ──
    print("\n=== 更新周涨跌 ===")
    append_weekly_changes(wb, sw_weekly_all, lux_weekly_all, macro)

    # ── 9. 宏观详细 sheets ──
    print("\n=== 更新宏观 sheets ===")
    update_macro_sheets(wb, macro)

    # ── 10. Charts ──
    print("\n=== 更新 Charts ===")
    update_charts(wb)

    # ── 11. 保存 ──
    wb.save(args.xlsx)
    print(f"\n{'='*50}")
    print(f"增量更新完成：{args.xlsx}")
    print(f"  新增 {n_new} 周股价数据")

    # 汇总
    for label, weekly, tickers in [("运动服饰", sw_weekly_all, SPORTSWEAR),
                                    ("奢侈品", lux_weekly_all, LUXURY)]:
        if weekly.empty:
            continue
        norm = weekly / weekly.iloc[0]
        latest = norm.iloc[-1].dropna().sort_values()
        print(f"\n--- {label} 最新归一化 ---")
        print(f"  涨幅前5: {dict(latest.tail(5).round(3))}")
        print(f"  跌幅前5: {dict(latest.head(5).round(3))}")

    for key in ["oil", "gas", "yield"]:
        df = macro[key]
        if not df.empty:
            for c in df.columns:
                print(f"  {c}: {df[c].iloc[-1]:.2f}")


if __name__ == "__main__":
    main()
