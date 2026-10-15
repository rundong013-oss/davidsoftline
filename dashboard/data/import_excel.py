#!/usr/bin/env python3
"""
import_excel.py — 一次性将 3 个 Excel 文件导入 SQLite。

用法：
    cd dashboard
    python -m data.import_excel
"""
import datetime as dt
import json
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook

# 将 dashboard/ 加入 path，以便 import dashboard.config
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from config import (
    EXCEL_APPAREL, EXCEL_VALUATION, EXCEL_EARNINGS,
    SPORTSWEAR_TICKERS, LUXURY_TICKERS, ETF_TICKERS,
    SPORTSWEAR_NAMES, LUXURY_NAMES, BASE_DATE,
)
from data.database import (
    init_db, upsert_weekly_prices, upsert_etf_weekly,
    upsert_sector_index, upsert_macro_data, upsert_valuation,
    upsert_earnings, clear_earnings, set_meta, get_conn,
)

# ── update_monitor.py 的列表，用于列名映射 ─────────────────────
# (展示代码, Yahoo代码, Bloomberg, 区域, 备注)
SPORTSWEAR_MONITOR = [
    ("NKE.N","NKE","NKE US","US/EU","Nike"),
    ("ADDY","ADS.DE","ADS GY","US/EU","Adidas"),
    ("ZGN.N","ZGN","ZGN US","US/EU","Zegna"),
    ("WWW.N","WWW","WWW US","US/EU","Wolverine"),
    ("VSCO.N",None,None,"US/EU","疑似退市"),
    ("UAA.N","UAA","UAA US","US/EU","Under Armour"),
    ("SKX.N",None,None,"US/EU","私有化退市"),
    ("SHOO.O","SHOO","SHOO US","US/EU","Steven Madden"),
    ("RL.N","RL","RL US","US/EU","Ralph Lauren"),
    ("PUM.DF","PUM.DE","PUM GY","US/EU","Puma"),
    ("ONON.N","ONON","ONON US","US/EU","On Holding"),
    ("NWL.O","NWL","NWL US","US/EU","Newell Brands"),
    ("LULU.O","LULU","LULU US","US/EU","Lululemon"),
    ("LEVI.N","LEVI","LEVI US","US/EU","Levi Strauss"),
    ("KTB.N","KTB","KTB US","US/EU","Kontoor"),
    ("HBI.N",None,None,"US/EU","被收购退市"),
    ("GIL.N","GIL","GIL US","US/EU","Gildan"),
    ("GAP.N","GAP","GAP US","US/EU","Gap"),
    ("DECK.N","DECK","DECK US","US/EU","Deckers"),
    ("CROX.O","CROX","CROX US","US/EU","Crocs"),
    ("COLM.O","COLM","COLM US","US/EU","Columbia"),
    ("BZUN.O","BZUN","BZUN US","US/EU","宝尊电商"),
    ("AS.N","AS","AS US","US/EU","Amer Sports"),
    ("ANF.N","ANF","ANF US","US/EU","Abercrombie"),
    ("9983.T","9983.T","9983 JT","Japan","迅销"),
    ("8111.T","8111.T","8111 JT","Japan","Goldwin"),
    ("7936.T","7936.T","7936 JT","Japan","亚瑟士"),
    ("3998.HK","3998.HK","3998 HK","China/HK","波司登"),
    ("2331.HK","2331.HK","2331 HK","China/HK","李宁"),
    ("2020.HK","2020.HK","2020 HK","China/HK","安踏"),
    ("1836.HK","1836.HK","1836 HK","China/HK","九兴控股"),
    ("1368.HK","1368.HK","1368 HK","China/HK","特步"),
    ("1361.HK","1361.HK","1361 HK","China/HK","361度"),
    ("002832.SZ","002832.SZ","002832 CH","China/HK","比音勒芬"),
]

LUXURY_MONITOR = [
    ("VFC","VFC","VFC US","US","VF Corp"),
    ("CFR","CFR.SW","CFR SW","Europe","Richemont"),
    ("1913.HK","1913.HK","1913 HK","Asia","Prada"),
    ("MONC","MONC.MI","MONC IM","Europe","Moncler"),
    ("KER","KER.PA","KER FP","Europe","Kering"),
    ("SFER","SFER.MI","SFER IM","Europe","Ferragamo"),
    ("RACE","RACE","RACE US","US","Ferrari"),
    ("RMS","RMS.PA","RMS FP","Europe","Hermes"),
    ("MC","MC.PA","MC FP","Europe","LVMH"),
    ("BRBY","BRBY.L","BRBY LN","Europe","Burberry"),
    ("CPRI","CPRI","CPRI US","US","Capri Holdings"),
    ("PVH","PVH","PVH US","US","PVH Corp"),
    ("TPR","TPR","TPR US","US","Tapestry"),
    ("GOOS","GOOS.TO","GOOS CT","North America","Canada Goose"),
    ("ZGN","ZGN","ZGN US","US","Zegna"),
]

# valuation_update.py 的行常量
R_FY = 5; R_REV = 6; R_YOY = 7; R_GM = 8; R_OPM = 9
R_NI = 10; R_NPM = 11; R_EPS = 12; R_SHR = 13
R_DATE = 15; R_PRICE = 16; R_MCAP = 17; R_PE = 18; R_PS = 19; R_ANLST = 21
ANN_START = 3; Q_START = 10

VALUATION_TICKERS = [
    ("NKE.N","NKE"),("ADDY","ADS.DE"),("ZGN.N","ZGN"),("WWW.N","WWW"),
    ("UAA.N","UAA"),("SHOO.O","SHOO"),("RL.N","RL"),("PUM.DF","PUM.DE"),
    ("ONON.N","ONON"),("NWL.O","NWL"),("LULU.O","LULU"),("LEVI.N","LEVI"),
    ("KTB.N","KTB"),("GIL.N","GIL"),("GAP.N","GAP"),("DECK.N","DECK"),
    ("CROX.O","CROX"),("COLM.O","COLM"),("BZUN.O","BZUN"),("AS.N","AS"),
    ("ANF.N","ANF"),("VSCO.N","VSCO"),("9983.T","9983.T"),
    ("8111.T","8111.T"),("7936.T","7936.T"),
    ("3998.HK","3998.HK"),("2331.HK","2331.HK"),("2020.HK","2020.HK"),
    ("1836.HK","1836.HK"),("1368.HK","1368.HK"),("1361.HK","1361.HK"),
    ("002832.SZ","002832.SZ"),
    ("VFC","VFC"),("CFR","CFR.SW"),("1913.HK","1913.HK"),("MONC","MONC.MI"),
    ("KER","KER.PA"),("SFER","SFER.MI"),("RACE","RACE"),("RMS","RMS.PA"),
    ("MC","MC.PA"),("BRBY","BRBY.L"),("CPRI","CPRI"),("PVH","PVH"),
    ("TPR","TPR"),("GOOS","GOOS.TO"),
]

# 判断 ticker 属于哪个 sector
_sw_display = {t[0] for t in SPORTSWEAR_MONITOR}
_lux_display = {t[0] for t in LUXURY_MONITOR}


def _ticker_sector(display_code):
    if display_code in _sw_display:
        return "sportswear"
    if display_code in _lux_display:
        return "luxury"
    return "unknown"


# ══════════════════════════════════════════════════════════════════
#  1. 导入 Apparel_updating file.xlsx
# ══════════════════════════════════════════════════════════════════
def _read_prices_sheet(wb, sheet_name, ticker_list):
    """读取 Prices sheet，返回 DataFrame (date×tickers)。"""
    ws = wb[sheet_name]
    # header in row 2
    headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(2, c).value
        if v:
            headers[c] = str(v)

    rows = []
    for r in range(3, ws.max_row + 1):
        d = ws.cell(r, 1).value
        if d is None:
            continue
        if isinstance(d, dt.datetime):
            d = d.date()
        elif isinstance(d, str):
            try:
                d = pd.Timestamp(d).date()
            except Exception:
                continue
        row = {"date": d}
        for c, name in headers.items():
            if c == 1:
                continue
            v = ws.cell(r, c).value
            if v is not None:
                try:
                    row[name] = float(v)
                except (ValueError, TypeError):
                    pass
        rows.append(row)
    return pd.DataFrame(rows) if rows else pd.DataFrame()


def import_apparel(xlsx_path):
    """导入 Apparel_updating file.xlsx 的所有相关 sheets。"""
    print(f"\n=== 导入 {xlsx_path} ===")
    wb = load_workbook(str(xlsx_path), data_only=True, read_only=True)

    # ── 运动服饰 Prices ──────────────────────────────────────────
    if "运动服饰 Prices" in wb.sheetnames:
        print("  读取 运动服饰 Prices...")
        df = _read_prices_sheet(wb, "运动服饰 Prices",
                                [t[0] for t in SPORTSWEAR_MONITOR])
        if not df.empty:
            base_row = df.iloc[0]
            records = []
            for _, row in df.iterrows():
                date_str = str(row["date"])
                for t_display, _, _, _, _ in SPORTSWEAR_MONITOR:
                    if t_display in row and pd.notna(row.get(t_display)):
                        close = row[t_display]
                        base_val = base_row.get(t_display)
                        norm = close / base_val if (base_val and base_val > 0) else None
                        records.append({
                            "date": date_str, "ticker": t_display,
                            "sector": "sportswear", "close": close,
                            "normalized": norm, "weekly_return": None,
                        })
            if records:
                rdf = pd.DataFrame(records)
                # 计算 weekly_return
                for ticker in rdf["ticker"].unique():
                    mask = rdf["ticker"] == ticker
                    closes = rdf.loc[mask, "close"].values
                    returns = [None] + [
                        (closes[i] - closes[i-1]) / closes[i-1]
                        if closes[i-1] and closes[i-1] > 0 else None
                        for i in range(1, len(closes))
                    ]
                    rdf.loc[mask, "weekly_return"] = returns
                upsert_weekly_prices(rdf)
                print(f"    {len(rdf)} 条记录")

    # ── 奢侈品 Prices ────────────────────────────────────────────
    if "奢侈品 Prices" in wb.sheetnames:
        print("  读取 奢侈品 Prices...")
        df = _read_prices_sheet(wb, "奢侈品 Prices",
                                [t[0] for t in LUXURY_MONITOR])
        if not df.empty:
            base_row = df.iloc[0]
            records = []
            for _, row in df.iterrows():
                date_str = str(row["date"])
                for t_display, _, _, _, _ in LUXURY_MONITOR:
                    if t_display in row and pd.notna(row.get(t_display)):
                        close = row[t_display]
                        base_val = base_row.get(t_display)
                        norm = close / base_val if (base_val and base_val > 0) else None
                        records.append({
                            "date": date_str, "ticker": t_display,
                            "sector": "luxury", "close": close,
                            "normalized": norm, "weekly_return": None,
                        })
            if records:
                rdf = pd.DataFrame(records)
                for ticker in rdf["ticker"].unique():
                    mask = rdf["ticker"] == ticker
                    closes = rdf.loc[mask, "close"].values
                    returns = [None] + [
                        (closes[i] - closes[i-1]) / closes[i-1]
                        if closes[i-1] and closes[i-1] > 0 else None
                        for i in range(1, len(closes))
                    ]
                    rdf.loc[mask, "weekly_return"] = returns
                upsert_weekly_prices(rdf)
                print(f"    {len(rdf)} 条记录")

    # ── 行业指数 ──────────────────────────────────────────────────
    if "行业指数" in wb.sheetnames:
        print("  读取 行业指数...")
        ws = wb["行业指数"]
        records = []
        for r in range(3, ws.max_row + 1):
            d = ws.cell(r, 1).value
            if d is None:
                break
            if isinstance(d, dt.datetime):
                d = d.date()
            sw_val = ws.cell(r, 2).value
            lux_val = ws.cell(r, 3).value
            if sw_val is not None:
                records.append({"date": str(d), "sector": "sportswear",
                                "index_value": float(sw_val)})
            if lux_val is not None:
                records.append({"date": str(d), "sector": "luxury",
                                "index_value": float(lux_val)})
        if records:
            upsert_sector_index(pd.DataFrame(records))
            print(f"    {len(records)} 条记录")

    # ── 可选消费 ETFs ─────────────────────────────────────────────
    if "可选消费ETFs" in wb.sheetnames:
        print("  读取 可选消费ETFs...")
        ws = wb["可选消费ETFs"]
        # header in row 2: Date, XLY, RSPD, XRT, ..., (normalized cols)
        headers = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(2, c).value
            if v:
                headers[c] = str(v)
        etf_codes = ["XLY", "RSPD", "XRT"]
        records = []
        base_prices = {}
        first_data_row = True
        for r in range(3, ws.max_row + 1):
            d = ws.cell(r, 1).value
            if d is None:
                break
            if isinstance(d, dt.datetime):
                d = d.date()
            elif isinstance(d, str):
                try:
                    d = pd.Timestamp(d).date()
                except Exception:
                    break
            for c, name in headers.items():
                if name in etf_codes:
                    v = ws.cell(r, c).value
                    if v is not None:
                        try:
                            close = float(v)
                            if first_data_row:
                                base_prices[name] = close
                            norm = (close / base_prices[name]
                                    if name in base_prices and base_prices[name] > 0
                                    else None)
                            records.append({
                                "date": str(d), "ticker": name,
                                "close": close, "normalized": norm,
                            })
                        except (ValueError, TypeError):
                            pass
            first_data_row = False
        if records:
            upsert_etf_weekly(pd.DataFrame(records))
            print(f"    {len(records)} 条记录")

    # ── 宏观数据 sheets ───────────────────────────────────────────
    _import_macro_sheets(wb)

    wb.close()
    print("  Apparel 导入完成")


def _import_macro_sheets(wb):
    """导入油价汽油、美债收益率、宏观高频数据。"""
    # 油价汽油
    if "油价汽油" in wb.sheetnames:
        print("  读取 油价汽油...")
        ws = wb["油价汽油"]
        _import_ts_sheet(ws, "weekly")

    # 美债收益率
    if "美债收益率" in wb.sheetnames:
        print("  读取 美债收益率...")
        ws = wb["美债收益率"]
        _import_ts_sheet(ws, "weekly")

    # 宏观高频数据
    if "宏观高频数据" in wb.sheetnames:
        print("  读取 宏观高频数据...")
        ws = wb["宏观高频数据"]
        _import_macro_highfreq(ws)


def _import_ts_sheet(ws, frequency):
    """导入时序 sheet（油价、收益率）到 macro_data。"""
    # 找 header 行（Date 在 column 1）
    hdr_row = None
    for r in range(1, min(ws.max_row + 1, 10)):
        v = ws.cell(r, 1).value
        if v and str(v).strip().lower() == "date":
            hdr_row = r
            break
    if hdr_row is None:
        return

    headers = {}
    for c in range(2, ws.max_column + 1):
        v = ws.cell(hdr_row, c).value
        if v:
            headers[c] = str(v).strip()

    records = []
    for r in range(hdr_row + 1, ws.max_row + 1):
        d = ws.cell(r, 1).value
        if d is None:
            continue
        if isinstance(d, dt.datetime):
            d = d.date()
        elif isinstance(d, str):
            try:
                d = pd.Timestamp(d).date()
            except Exception:
                continue

        for c, series_name in headers.items():
            v = ws.cell(r, c).value
            if v is not None:
                try:
                    records.append({
                        "date": str(d), "series": series_name,
                        "value": float(v), "frequency": frequency,
                    })
                except (ValueError, TypeError):
                    pass
    if records:
        upsert_macro_data(pd.DataFrame(records))
        print(f"    {len(records)} 条记录")


def _import_macro_highfreq(ws):
    """导入宏观高频数据（周度 + 月度混合）。"""
    # 扫描所有带 Date header 的数据块
    r = 1
    while r <= ws.max_row:
        v = ws.cell(r, 1).value
        if v and str(v).strip().lower() == "date":
            # 这是一个数据块的 header
            headers = {}
            for c in range(2, ws.max_column + 1):
                h = ws.cell(r, c).value
                if h:
                    headers[c] = str(h).strip()
            if not headers:
                r += 1
                continue

            # 判断频率：如果系列名包含"初请"则周度，否则月度
            freq = "weekly"
            for name in headers.values():
                if any(kw in name for kw in ["CPI", "PPI", "PCE", "非农",
                                              "失业率", "零售", "信心", "信用卡"]):
                    freq = "monthly"
                    break

            records = []
            data_row = r + 1
            while data_row <= ws.max_row:
                d = ws.cell(data_row, 1).value
                if d is None:
                    break
                if isinstance(d, dt.datetime):
                    d = d.date()
                elif isinstance(d, str):
                    try:
                        d = pd.Timestamp(d).date()
                    except Exception:
                        break

                for c, series_name in headers.items():
                    val = ws.cell(data_row, c).value
                    if val is not None:
                        try:
                            records.append({
                                "date": str(d), "series": series_name,
                                "value": float(val), "frequency": freq,
                            })
                        except (ValueError, TypeError):
                            pass
                data_row += 1

            if records:
                upsert_macro_data(pd.DataFrame(records))
                print(f"    {len(records)} 条 ({freq})")
            r = data_row
        else:
            r += 1


# ══════════════════════════════════════════════════════════════════
#  2. 导入 Valuation screening.xlsx
# ══════════════════════════════════════════════════════════════════
def import_valuation(xlsx_path):
    """导入 Valuation screening.xlsx 的每个 ticker sheet。"""
    print(f"\n=== 导入 {xlsx_path} ===")
    wb = load_workbook(str(xlsx_path), data_only=True, read_only=True)

    val_display_set = {d for d, _ in VALUATION_TICKERS}
    count = 0

    for sheet_name in wb.sheetnames:
        if sheet_name == "Summary":
            continue
        if sheet_name not in val_display_set:
            continue

        ws = wb[sheet_name]
        sector = _ticker_sector(sheet_name)

        # 读取数据
        def cell(r, c):
            v = ws.cell(r, c).value
            if v is None:
                return None
            if isinstance(v, str):
                try:
                    return float(v)
                except ValueError:
                    return v
            return v

        # 年度数据
        annual = []
        for i in range(5):
            col = ANN_START + i
            fy_label = cell(R_FY, col)
            if fy_label is None:
                continue
            rev = cell(R_REV, col)
            yoy = cell(R_YOY, col)
            gm = cell(R_GM, col)
            opm = cell(R_OPM, col)
            ni = cell(R_NI, col)
            npm = cell(R_NPM, col)
            eps = cell(R_EPS, col)
            shr = cell(R_SHR, col)
            annual.append({
                "label": str(fy_label),
                "revenue": rev, "yoy": yoy, "gm": gm, "opm": opm,
                "net_income": ni, "npm": npm, "eps": eps, "shares": shr,
                "kind": "estimate" if str(fy_label).endswith("E") else "actual",
            })

        # 季度数据
        quarterly = []
        for i in range(6):
            col = Q_START + i
            q_label = cell(R_FY, col)
            if q_label is None:
                continue
            rev = cell(R_REV, col)
            gm = cell(R_GM, col)
            opm = cell(R_OPM, col)
            ni = cell(R_NI, col)
            npm = cell(R_NPM, col)
            eps = cell(R_EPS, col)
            shr = cell(R_SHR, col)
            quarterly.append({
                "label": str(q_label),
                "revenue": rev, "gm": gm, "opm": opm,
                "net_income": ni, "npm": npm, "eps": eps, "shares": shr,
                "kind": "estimate" if str(q_label).endswith("E") else "actual",
            })

        # 估值数据
        price = cell(R_PRICE, ANN_START)
        mcap = cell(R_MCAP, ANN_START)
        n_analysts = cell(R_ANLST, ANN_START + 3)  # FY+1E column

        # 提取 currency（从 row label）
        price_label = cell(R_PRICE, 2)
        currency = ""
        if price_label and isinstance(price_label, str) and "(" in price_label:
            currency = price_label.split("(")[1].split(")")[0]

        # Forward PE/PS
        fwd_pe = {}
        fwd_ps = {}
        for i, ann_item in enumerate(annual):
            col = ANN_START + i
            pe_val = cell(R_PE, col)
            ps_val = cell(R_PS, col)
            if pe_val is not None and isinstance(pe_val, (int, float)):
                fwd_pe[ann_item["label"]] = pe_val
            if ps_val is not None and isinstance(ps_val, (int, float)):
                fwd_ps[ann_item["label"]] = ps_val

        # 名称
        name_cell = cell(3, 2)
        name = str(name_cell).split(",")[1].strip() if (
            name_cell and isinstance(name_cell, str) and "," in name_cell
        ) else sheet_name

        upsert_valuation({
            "ticker": sheet_name,
            "name": name,
            "sector": sector,
            "currency": currency,
            "price": price,
            "market_cap": mcap,
            "shares_out": cell(R_SHR, ANN_START + len(annual) - 1) if annual else None,
            "annual_json": json.dumps(annual, ensure_ascii=False, default=str),
            "quarterly_json": json.dumps(quarterly, ensure_ascii=False, default=str),
            "fwd_pe_json": json.dumps(fwd_pe, ensure_ascii=False),
            "fwd_ps_json": json.dumps(fwd_ps, ensure_ascii=False),
            "n_analysts": int(n_analysts) if n_analysts else None,
        })
        count += 1

    wb.close()
    print(f"  导入 {count} 个 ticker")


# ══════════════════════════════════════════════════════════════════
#  3. 导入 Earnings calendar.xlsx
# ══════════════════════════════════════════════════════════════════
def import_earnings(xlsx_path):
    """导入 Earnings calendar.xlsx。"""
    print(f"\n=== 导入 {xlsx_path} ===")
    wb = load_workbook(str(xlsx_path), data_only=True, read_only=True)

    clear_earnings()
    count = 0
    now = dt.datetime.now(tz=dt.timezone.utc)

    for sheet_name, sector in [("运动服饰", "sportswear"), ("奢侈品", "luxury")]:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        for r in range(2, ws.max_row + 1):
            ticker = ws.cell(r, 1).value
            if ticker is None:
                continue
            name = ws.cell(r, 2).value or ""

            # 读取最多 2 个 earnings 日期
            for date_col, eps_col, qtr_col in [(3, 5, 6), (7, 9, 10)]:
                earnings_date = ws.cell(r, date_col).value
                if earnings_date is None:
                    continue
                if isinstance(earnings_date, str):
                    if earnings_date in ("非美股", ""):
                        continue
                    try:
                        earnings_date = pd.Timestamp(earnings_date)
                    except Exception:
                        continue
                if isinstance(earnings_date, dt.datetime):
                    edate = earnings_date.date()
                elif isinstance(earnings_date, dt.date):
                    edate = earnings_date
                else:
                    continue

                days = (dt.datetime.combine(edate, dt.time()) -
                        dt.datetime.combine(dt.date.today(), dt.time())).days
                eps_est = ws.cell(r, eps_col).value
                fiscal_qtr = ws.cell(r, qtr_col).value

                upsert_earnings({
                    "ticker": str(ticker),
                    "name": str(name),
                    "sector": sector,
                    "earnings_date": str(edate),
                    "days_from_today": days,
                    "eps_estimate": str(eps_est) if eps_est else None,
                    "fiscal_quarter": str(fiscal_qtr) if fiscal_qtr else None,
                })
                count += 1

    wb.close()
    print(f"  导入 {count} 条 earnings 记录")


# ══════════════════════════════════════════════════════════════════
#  主入口
# ══════════════════════════════════════════════════════════════════
def main():
    print("=" * 60)
    print("  Excel → SQLite 导入工具")
    print("=" * 60)

    init_db()

    if EXCEL_APPAREL.exists():
        import_apparel(EXCEL_APPAREL)
    else:
        print(f"\n  跳过：{EXCEL_APPAREL} 不存在")

    if EXCEL_VALUATION.exists():
        import_valuation(EXCEL_VALUATION)
    else:
        print(f"\n  跳过：{EXCEL_VALUATION} 不存在")

    if EXCEL_EARNINGS.exists():
        import_earnings(EXCEL_EARNINGS)
    else:
        print(f"\n  跳过：{EXCEL_EARNINGS} 不存在")

    set_meta("last_import", dt.datetime.now().isoformat())
    set_meta("import_source", "excel")

    # 统计
    with get_conn() as conn:
        for table in ["weekly_prices", "etf_weekly", "sector_index",
                       "macro_data", "valuation", "earnings_calendar"]:
            count = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
            print(f"  {table}: {count} 条")

    print("\n导入完成!")


if __name__ == "__main__":
    main()
