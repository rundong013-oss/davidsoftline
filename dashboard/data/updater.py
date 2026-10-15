"""
updater.py — API 增量更新模块。

复用 scripts/ 下的数据拉取逻辑，写入 SQLite。
"""
import datetime as dt
import json
import sys
from pathlib import Path

import numpy as np
import pandas as pd

# 将 scripts/ 加入 path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent / "scripts"))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from config import (
    BASE_DATE, ETF_BASE_DATE, MACRO_BASE, ETF_TICKERS,
    SPORTSWEAR_NAMES, LUXURY_NAMES,
)
from data.database import (
    init_db, upsert_weekly_prices, upsert_etf_weekly,
    upsert_sector_index, upsert_macro_data, upsert_valuation,
    upsert_earnings, clear_earnings, set_meta, get_conn,
)

# 从 update_monitor.py 导入
from update_monitor import (
    SPORTSWEAR as SW_MONITOR, LUXURY as LUX_MONITOR, ETFS,
    fetch_daily, to_weekly, fetch_macro,
)

# 从 valuation_update.py 导入
from valuation_update import (
    SPORTSWEAR as SW_VAL, LUXURY as LUX_VAL,
    fetch_one, build_annual, build_quarterly, KNOWN_FIN_CCY,
)


# ── Prices ───────────────────────────────────────────────────────
def update_prices():
    """拉取最新价格数据，写入 SQLite。"""
    print("=== 更新股价数据 ===")

    for label, tickers, sector, base_date in [
        ("运动服饰", SW_MONITOR, "sportswear", BASE_DATE),
        ("奢侈品", LUX_MONITOR, "luxury", BASE_DATE),
    ]:
        codes = [t[1] for t in tickers if t[1]]
        display_codes = [t[0] for t in tickers]
        y2u = {t[1]: t[0] for t in tickers if t[1]}

        print(f"  拉取 {label} ({len(codes)} 只)...")
        daily = fetch_daily(codes, base_date)
        weekly = to_weekly(daily, base_date)
        weekly = weekly.rename(columns=y2u)

        if weekly.empty:
            continue

        base_row = weekly.iloc[0]
        records = []
        for _, row in weekly.iterrows():
            date_str = str(row.name.date()) if hasattr(row.name, 'date') else str(row.name)[:10]
            for ticker in display_codes:
                if ticker in row and pd.notna(row[ticker]):
                    close = row[ticker]
                    base_val = base_row.get(ticker)
                    norm = close / base_val if (base_val and base_val > 0) else None
                    records.append({
                        "date": date_str, "ticker": ticker,
                        "sector": sector, "close": close,
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

    # ETFs
    print("  拉取 ETF...")
    etf_codes = [t[1] for t in ETFS]
    etf_daily = fetch_daily(etf_codes, ETF_BASE_DATE)
    etf_weekly = to_weekly(etf_daily, ETF_BASE_DATE)

    if not etf_weekly.empty:
        base_row = etf_weekly.iloc[0]
        records = []
        for _, row in etf_weekly.iterrows():
            date_str = str(row.name.date()) if hasattr(row.name, 'date') else str(row.name)[:10]
            for code in etf_codes:
                if code in row and pd.notna(row[code]):
                    close = row[code]
                    base_val = base_row.get(code)
                    norm = close / base_val if (base_val and base_val > 0) else None
                    records.append({
                        "date": date_str, "ticker": code,
                        "close": close, "normalized": norm,
                    })
        if records:
            upsert_etf_weekly(pd.DataFrame(records))
            print(f"    ETF: {len(records)} 条记录")

    # 行业指数
    print("  计算行业指数...")
    for label, tickers, sector in [
        ("运动服饰", SW_MONITOR, "sportswear"),
        ("奢侈品", LUX_MONITOR, "luxury"),
    ]:
        codes = [t[1] for t in tickers if t[1]]
        y2u = {t[1]: t[0] for t in tickers if t[1]}
        daily = fetch_daily(codes, BASE_DATE)
        wk = to_weekly(daily, BASE_DATE)
        wk = wk.rename(columns=y2u)
        if wk.empty:
            continue
        display_codes = [t[0] for t in tickers]
        wk_use = wk.reindex(columns=display_codes)
        norm = wk_use / wk_use.iloc[0]
        idx = norm.mean(axis=1, skipna=True)
        records = []
        for d, v in idx.items():
            if pd.notna(v):
                date_str = str(d.date()) if hasattr(d, 'date') else str(d)[:10]
                records.append({
                    "date": date_str, "sector": sector,
                    "index_value": float(v),
                })
        if records:
            upsert_sector_index(pd.DataFrame(records))
            print(f"    {sector} 指数: {len(records)} 条")

    set_meta("last_price_update", dt.datetime.now().isoformat())
    print("  股价更新完成")


# ── Macro ────────────────────────────────────────────────────────
def update_macro():
    """拉取最新宏观数据，写入 SQLite。"""
    print("\n=== 更新宏观数据 ===")
    macro = fetch_macro()

    records = []

    # 油价
    oil = macro.get("oil", pd.DataFrame())
    if not oil.empty:
        for d, row in oil.iterrows():
            date_str = str(d.date()) if hasattr(d, 'date') else str(d)[:10]
            for col in oil.columns:
                if pd.notna(row[col]):
                    records.append({
                        "date": date_str, "series": col,
                        "value": float(row[col]), "frequency": "weekly",
                    })

    # 汽油
    gas = macro.get("gas", pd.DataFrame())
    if not gas.empty:
        for d, row in gas.iterrows():
            date_str = str(d.date()) if hasattr(d, 'date') else str(d)[:10]
            for col in gas.columns:
                if pd.notna(row[col]):
                    records.append({
                        "date": date_str, "series": col,
                        "value": float(row[col]), "frequency": "weekly",
                    })

    # 收益率
    yld = macro.get("yield", pd.DataFrame())
    if not yld.empty:
        for d, row in yld.iterrows():
            date_str = str(d.date()) if hasattr(d, 'date') else str(d)[:10]
            for col in yld.columns:
                if pd.notna(row[col]):
                    records.append({
                        "date": date_str, "series": col,
                        "value": float(row[col]), "frequency": "weekly",
                    })

    # 初请失业金
    claims = macro.get("claims", pd.DataFrame())
    if not claims.empty:
        for d, row in claims.iterrows():
            date_str = str(d.date()) if hasattr(d, 'date') else str(d)[:10]
            for col in claims.columns:
                if pd.notna(row[col]):
                    records.append({
                        "date": date_str, "series": col,
                        "value": float(row[col]), "frequency": "weekly",
                    })

    # 月度数据
    monthly = macro.get("monthly", pd.DataFrame())
    if not monthly.empty:
        for d, row in monthly.iterrows():
            date_str = str(d.date()) if hasattr(d, 'date') else str(d)[:10]
            for col in monthly.columns:
                if pd.notna(row[col]):
                    records.append({
                        "date": date_str, "series": col,
                        "value": float(row[col]), "frequency": "monthly",
                    })

    if records:
        upsert_macro_data(pd.DataFrame(records))
        print(f"  {len(records)} 条宏观记录")

    set_meta("last_macro_update", dt.datetime.now().isoformat())
    print("  宏观更新完成")


# ── Valuation ────────────────────────────────────────────────────
def update_valuation():
    """拉取最新估值数据，写入 SQLite。"""
    print("\n=== 更新估值数据 ===")
    all_tickers = SW_VAL + LUX_VAL
    sw_set = {d for d, _ in SW_VAL}
    MM = 1e6
    ok, fail = 0, 0

    for disp, yahoo in all_tickers:
        sector = "sportswear" if disp in sw_set else "luxury"
        print(f"  {disp} ({yahoo})...", end=" ")
        d = fetch_one(yahoo)
        if not d:
            print("FAIL")
            fail += 1
            continue

        try:
            ann = build_annual(d)
            qtr = build_quarterly(d)

            # Forward PE/PS
            fwd_pe = {}
            fwd_ps = {}
            mcap = d.get("mcap", 0) or 0
            for r in ann:
                if r.get("net_income") and mcap:
                    pe = mcap / r["net_income"]
                    if pe > 0:
                        fwd_pe[r["label"]] = round(pe, 1)
                if r.get("revenue") and mcap:
                    ps = mcap / r["revenue"]
                    if ps > 0:
                        fwd_ps[r["label"]] = round(ps, 1)

            # n_analysts from estimate rows
            n_analysts = None
            for r in ann:
                if r["kind"] == "estimate" and r.get("n_analysts"):
                    n_analysts = r["n_analysts"]
                    break

            upsert_valuation({
                "ticker": disp,
                "name": d.get("name", disp),
                "sector": sector,
                "currency": d.get("ccy", ""),
                "price": d.get("price"),
                "market_cap": mcap / MM if mcap else None,
                "shares_out": d.get("shares", 0) / MM if d.get("shares") else None,
                "annual_json": json.dumps(ann, ensure_ascii=False, default=str),
                "quarterly_json": json.dumps(qtr, ensure_ascii=False, default=str),
                "fwd_pe_json": json.dumps(fwd_pe, ensure_ascii=False),
                "fwd_ps_json": json.dumps(fwd_ps, ensure_ascii=False),
                "n_analysts": n_analysts,
            })
            print(f"OK ({d.get('ccy','')} MCap={mcap/MM:,.0f}mm)")
            ok += 1
        except Exception as e:
            print(f"ERROR: {e}")
            fail += 1

    set_meta("last_valuation_update", dt.datetime.now().isoformat())
    print(f"  估值更新完成: {ok} 成功, {fail} 失败")


# ── Earnings ─────────────────────────────────────────────────────
def update_earnings():
    """拉取最新 earnings 日历，写入 SQLite。"""
    print("\n=== 更新 Earnings 日历 ===")
    try:
        from earnings_calendar import (
            SPORTSWEAR as EC_SW, LUXURY as EC_LUX,
            _get_session, scan_earnings, match_tickers,
        )
    except ImportError as e:
        print(f"  导入 earnings_calendar 失败: {e}")
        return

    session = _get_session()
    calendar = scan_earnings(session, months=6)

    clear_earnings()
    count = 0

    for ticker_list, sector in [(EC_SW, "sportswear"), (EC_LUX, "luxury")]:
        matched = match_tickers(calendar, ticker_list)
        for ticker, info in matched.items():
            for e in info.get("earnings", []):
                edate = e["date"]
                if hasattr(edate, 'date'):
                    edate = edate.date() if hasattr(edate, 'date') else edate
                elif hasattr(edate, 'strftime'):
                    pass
                days = (dt.datetime.combine(
                    edate if isinstance(edate, dt.date) else edate.date(),
                    dt.time()
                ) - dt.datetime.combine(dt.date.today(), dt.time())).days

                upsert_earnings({
                    "ticker": ticker,
                    "name": info.get("name", ""),
                    "sector": sector,
                    "earnings_date": str(edate)[:10] if hasattr(edate, 'strftime') else str(edate)[:10],
                    "days_from_today": days,
                    "eps_estimate": e.get("eps_estimate", ""),
                    "fiscal_quarter": e.get("fiscal_qtr", ""),
                })
                count += 1

    set_meta("last_earnings_update", dt.datetime.now().isoformat())
    print(f"  Earnings 更新完成: {count} 条")


# ── 全量更新 ─────────────────────────────────────────────────────
def update_all():
    """更新所有数据。"""
    init_db()
    update_prices()
    update_macro()
    update_valuation()
    update_earnings()
    set_meta("last_full_update", dt.datetime.now().isoformat())
    print("\n=== 全量更新完成 ===")


if __name__ == "__main__":
    update_all()
