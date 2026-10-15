#!/usr/bin/env python3
"""
fmp_snapshot.py — 从 FMP 拉取估值 snapshot 所需数据（先跑通数据，不涉及表格格式）

每个 ticker 输出：
  - 过去 3 个财年实际值：Revenue、Gross Profit、EBIT、Net Income（→ yoy、GM、OPM）
  - 未来 3 个财年一致预期：Revenue、EBIT、Net Income（均值 + 分析师数量）
  - 估值：股价、市值（直接取 FMP quote）、Forward PE、Forward PS

用法：
  export FMP_API_KEY=你的key
  python fmp_snapshot.py NKE DECK LULU
  python fmp_snapshot.py --file tickers.txt          # 每行一个 ticker，# 开头为注释
  python fmp_snapshot.py NKE --inspect               # 打印各接口原始字段，首次使用先跑这个核对字段名

每个 ticker 调用 4 次 API（利润表、一致预期、报价、公司概况），币种不一致时再加 1 次汇率。

输出（默认 ./fmp_out/）：
  snapshot_long.csv   每行 = 一个 ticker 的一个财年（actual / estimate）
  valuation.csv       每行 = 一个 ticker 的市值与 forward 估值
  warnings.csv        数据缺失、币种、覆盖分析师过少等问题
  raw/                各接口原始 JSON（便于排查）
"""
import argparse
import datetime as dt
import json
import os
import sys
import time
from pathlib import Path

import pandas as pd
import requests

BASE = "https://financialmodelingprep.com/stable"
N_HIST = 3          # 过去几个财年
N_FWD = 3           # 未来几个财年
MIN_ANALYSTS = 3    # 覆盖分析师少于此数则警告
MM = 1e6            # 输出单位：百万


class FMPError(Exception):
    pass


class FMP:
    def __init__(self, key: str, raw_dir: Path, sleep: float = 0.2):
        self.key, self.raw_dir, self.sleep = key, raw_dir, sleep
        self.calls = 0
        raw_dir.mkdir(parents=True, exist_ok=True)

    def get(self, endpoint: str, **params):
        params["apikey"] = self.key
        url = f"{BASE}/{endpoint}"
        for attempt in range(3):
            r = requests.get(url, params=params, timeout=30)
            self.calls += 1
            if r.status_code == 429:                      # 超速，退避重试
                time.sleep(5 * (attempt + 1))
                continue
            break
        time.sleep(self.sleep)
        if r.status_code in (401, 402, 403):
            raise FMPError(f"{endpoint}[{params.get('symbol')}] 无权限 HTTP {r.status_code}: {r.text[:200]}")
        if r.status_code != 200:
            raise FMPError(f"{endpoint}[{params.get('symbol')}] HTTP {r.status_code}: {r.text[:200]}")
        try:
            data = r.json()
        except ValueError:
            raise FMPError(f"{endpoint}[{params.get('symbol')}] 返回非 JSON: {r.text[:200]}")
        if isinstance(data, dict) and ("Error Message" in data or "error" in data):
            raise FMPError(f"{endpoint}[{params.get('symbol')}] {data}")
        sym = str(params.get("symbol", "")).replace("/", "_")
        (self.raw_dir / f"{sym}_{endpoint}.json").write_text(json.dumps(data, indent=1, ensure_ascii=False))
        return data


def pick(d: dict, *names):
    """按顺序取第一个存在的字段（兼容 stable 与旧 v3 的字段名）。"""
    for n in names:
        v = d.get(n)
        if v not in (None, ""):
            return v
    return None


def to_date(s):
    return dt.date.fromisoformat(str(s)[:10]) if s else None


def div(a, b):
    try:
        return a / b if a is not None and b not in (None, 0) else None
    except TypeError:
        return None


def fx_rate(fmp: FMP, frm: str, to: str):
    """返回 1 单位 frm 值多少 to。FMP 外汇代码形如 EURUSD。"""
    if not frm or not to or frm == to:
        return 1.0
    for sym, invert in ((f"{frm}{to}", False), (f"{to}{frm}", True)):
        try:
            q = fmp.get("quote", symbol=sym)
            if q and q[0].get("price"):
                p = float(q[0]["price"])
                return 1 / p if invert else p
        except FMPError:
            continue
    return None


def build(fmp: FMP, sym: str, actual_ebit_field: str):
    warns = []
    w = lambda msg: warns.append({"ticker": sym, "issue": msg})

    inc = fmp.get("income-statement", symbol=sym, period="annual", limit=N_HIST + 1)  # 多取 1 年算 yoy
    try:
        est = fmp.get("analyst-estimates", symbol=sym, period="annual", page=0, limit=10)
    except FMPError as e:
        est = []
        w(f"一致预期获取失败：{e}")
    quote = fmp.get("quote", symbol=sym)
    profile = fmp.get("profile", symbol=sym)

    if not inc:
        w("无利润表数据")
        return [], None, warns

    # ---------- 实际值 ----------
    inc = sorted(inc, key=lambda x: x.get("date", ""))
    rep_ccy = inc[-1].get("reportedCurrency")
    rows = []
    for x in inc:
        fy = pick(x, "fiscalYear", "calendarYear") or str(x.get("date", ""))[:4]
        ebit = pick(x, actual_ebit_field, "operatingIncome", "ebit")
        rows.append({
            "ticker": sym, "kind": "actual", "fy": int(fy), "period_end": x.get("date"),
            "currency": x.get("reportedCurrency"),
            "revenue": pick(x, "revenue"), "gross_profit": pick(x, "grossProfit"),
            "ebit": ebit, "net_income": pick(x, "netIncome"),
            "n_analysts_rev": None, "n_analysts_eps": None,
        })
    last_fy, last_end = rows[-1]["fy"], to_date(rows[-1]["period_end"])

    # ---------- 一致预期：取最近一个实际财年之后的 N_FWD 个 ----------
    fwd = [e for e in est if to_date(e.get("date")) and to_date(e["date"]) > last_end + dt.timedelta(days=45)]
    fwd = sorted(fwd, key=lambda e: e["date"])[:N_FWD]
    if len(fwd) < N_FWD:
        w(f"一致预期只有 {len(fwd)} 个未来财年")
    prev_end = last_end
    for i, e in enumerate(fwd, start=1):
        end = to_date(e["date"])
        gap = (end - prev_end).days
        if not 330 <= gap <= 400:
            w(f"预期期间 {e['date']} 与上一期相隔 {gap} 天，财年对齐可能有误")
        prev_end = end
        n_rev = pick(e, "numAnalystsRevenue", "numberAnalystEstimatedRevenue")
        n_eps = pick(e, "numAnalystsEps", "numberAnalystsEstimatedEps")
        if n_rev is not None and n_rev < MIN_ANALYSTS:
            w(f"FY{last_fy + i}E 收入预期仅 {n_rev} 位分析师")
        rows.append({
            "ticker": sym, "kind": "estimate", "fy": last_fy + i, "period_end": e["date"],
            "currency": rep_ccy,   # FMP 预期默认按报表币种；--inspect 可核对
            "revenue": pick(e, "revenueAvg", "estimatedRevenueAvg"),
            "gross_profit": None,  # FMP 一致预期不含毛利
            "ebit": pick(e, "ebitAvg", "estimatedEbitAvg"),
            "net_income": pick(e, "netIncomeAvg", "estimatedNetIncomeAvg"),
            "n_analysts_rev": n_rev, "n_analysts_eps": n_eps,
        })

    # ---------- 派生指标 ----------
    for i, r in enumerate(rows):
        prev = rows[i - 1]["revenue"] if i > 0 else None
        r["yoy"] = div(r["revenue"], prev) - 1 if div(r["revenue"], prev) is not None else None
        r["gm"] = div(r["gross_profit"], r["revenue"])
        r["opm"] = div(r["ebit"], r["revenue"])
        r["npm"] = div(r["net_income"], r["revenue"])
    n_act = sum(r["kind"] == "actual" for r in rows)
    rows = rows[n_act - N_HIST:] if n_act > N_HIST else rows   # 去掉只用于算 yoy 的最早一年

    # ---------- 估值 ----------
    q = quote[0] if quote else {}
    p = profile[0] if profile else {}
    mcap, price = pick(q, "marketCap") or pick(p, "marketCap", "mktCap"), pick(q, "price") or pick(p, "price")
    quote_ccy = p.get("currency")
    ts = q.get("timestamp")
    qdate = dt.datetime.fromtimestamp(ts, dt.timezone.utc).date().isoformat() if ts else dt.date.today().isoformat()
    if mcap is None:
        w("无市值数据")
    rate = fx_rate(fmp, quote_ccy, rep_ccy)
    if rate is None:
        w(f"无法获取 {quote_ccy}->{rep_ccy} 汇率，forward 估值留空")
    elif quote_ccy != rep_ccy:
        w(f"报价币种 {quote_ccy} ≠ 报表币种 {rep_ccy}，市值已按 {rate:.4f} 换算")
    mcap_rep = mcap * rate if (mcap and rate) else None

    val = {
        "ticker": sym, "name": p.get("companyName"), "date": qdate, "price": price,
        "quote_ccy": quote_ccy, "market_cap_mm": mcap / MM if mcap else None,
        "implied_shares_mm": div(mcap, price) / MM if div(mcap, price) else None,
        "reporting_ccy": rep_ccy, "fx_quote_to_reporting": rate,
        "market_cap_mm_reporting_ccy": mcap_rep / MM if mcap_rep else None,
    }
    for r in rows:
        if r["kind"] == "estimate":
            r["fwd_pe"] = div(mcap_rep, r["net_income"]) if (r["net_income"] or 0) > 0 else None
            r["fwd_ps"] = div(mcap_rep, r["revenue"])
            val[f"fwd_pe_FY{r['fy']}E"] = r["fwd_pe"]
            val[f"fwd_ps_FY{r['fy']}E"] = r["fwd_ps"]
        else:
            r["fwd_pe"] = r["fwd_ps"] = None
    for r in rows:   # 金额转百万
        for k in ("revenue", "gross_profit", "ebit", "net_income"):
            r[k] = r[k] / MM if r[k] is not None else None
    return rows, val, warns


def preview(df: pd.DataFrame, val: dict):
    sub = df.set_index("fy")
    lab = {f: f"{f}{'E' if k == 'estimate' else ''}" for f, k in zip(sub.index, sub["kind"])}
    t = pd.DataFrame({
        "Revenue (mm)": sub["revenue"].map(lambda v: f"{v:,.0f}" if pd.notna(v) else ""),
        "yoy": sub["yoy"].map(lambda v: f"{v:.1%}" if pd.notna(v) else ""),
        "GM": sub["gm"].map(lambda v: f"{v:.1%}" if pd.notna(v) else ""),
        "OPM": sub["opm"].map(lambda v: f"{v:.1%}" if pd.notna(v) else ""),
        "Net Profit (mm)": sub["net_income"].map(lambda v: f"{v:,.0f}" if pd.notna(v) else ""),
        "#analysts": sub["n_analysts_rev"].map(lambda v: f"{v:.0f}" if pd.notna(v) else ""),
        "Fwd PE": sub["fwd_pe"].map(lambda v: f"{v:.1f}x" if pd.notna(v) else ""),
        "Fwd PS": sub["fwd_ps"].map(lambda v: f"{v:.2f}x" if pd.notna(v) else ""),
    }).T.rename(columns=lab)
    print(f"\n=== {val['ticker']}  {val.get('name') or ''}  ({sub['currency'].iloc[-1]}) ===")
    print(t.to_string())
    mc = val.get("market_cap_mm")
    print(f"Date {val['date']} | Price {val['price']} {val['quote_ccy']} | "
          f"Market Cap {mc:,.0f}mm {val['quote_ccy']}" if mc else "Market Cap n/a")


def inspect(fmp: FMP, sym: str):
    for ep, params in [("income-statement", dict(period="annual", limit=1)),
                       ("analyst-estimates", dict(period="annual", page=0, limit=2)),
                       ("quote", {}), ("profile", {})]:
        try:
            d = fmp.get(ep, symbol=sym, **params)
            print(f"\n--- {ep} ---")
            print(json.dumps(d[:1] if isinstance(d, list) else d, indent=1, ensure_ascii=False)[:3000])
        except FMPError as e:
            print(f"\n--- {ep} --- 失败：{e}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("tickers", nargs="*")
    ap.add_argument("--file", help="ticker 列表文件，每行一个")
    ap.add_argument("--out", default="fmp_out")
    ap.add_argument("--inspect", action="store_true", help="只打印原始字段")
    ap.add_argument("--actual-ebit", default="operatingIncome", choices=["operatingIncome", "ebit"],
                    help="实际值 EBIT 口径（默认 operatingIncome）")
    a = ap.parse_args()

    key = os.environ.get("FMP_API_KEY")
    if not key:
        sys.exit("请先设置环境变量 FMP_API_KEY")
    tickers = list(a.tickers)
    if a.file:
        tickers += [l.split("#")[0].strip() for l in Path(a.file).read_text().splitlines()
                    if l.split("#")[0].strip()]
    if not tickers:
        sys.exit("请提供 ticker")

    out = Path(a.out)
    fmp = FMP(key, out / "raw")
    if a.inspect:
        for s in tickers:
            inspect(fmp, s)
        print(f"\nAPI 调用次数：{fmp.calls}")
        return

    all_rows, vals, warns = [], [], []
    for s in tickers:
        try:
            rows, val, w = build(fmp, s, a.actual_ebit)
        except FMPError as e:
            warns.append({"ticker": s, "issue": str(e)})
            print(f"[失败] {s}: {e}")
            continue
        all_rows += rows
        warns += w
        if val:
            vals.append(val)
            preview(pd.DataFrame(rows), val)

    out.mkdir(exist_ok=True)
    cols = ["ticker", "kind", "fy", "period_end", "currency", "revenue", "yoy", "gross_profit", "gm",
            "ebit", "opm", "net_income", "npm", "n_analysts_rev", "n_analysts_eps", "fwd_pe", "fwd_ps"]
    pd.DataFrame(all_rows).reindex(columns=cols).to_csv(out / "snapshot_long.csv", index=False)
    pd.DataFrame(vals).to_csv(out / "valuation.csv", index=False)
    pd.DataFrame(warns, columns=["ticker", "issue"]).to_csv(out / "warnings.csv", index=False)

    if warns:
        print("\n--- 警告 ---")
        for x in warns:
            print(f"{x['ticker']}: {x['issue']}")
    print(f"\n完成：{len(vals)}/{len(tickers)} 个 ticker，API 调用 {fmp.calls} 次，输出目录 {out.resolve()}")


if __name__ == "__main__":
    main()
