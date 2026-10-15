#!/usr/bin/env python3
"""
fmp_snapshot.py — 从 FMP 拉取估值 snapshot 所需数据
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

BASE = "https://fmp-v2.icu/stable"
N_HIST = 3
N_FWD = 3
MIN_ANALYSTS = 3
MM = 1e6


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
            if r.status_code == 429:
                time.sleep(5 * (attempt + 1))
                continue
            break
        time.sleep(self.sleep)
        if r.status_code in (401, 402, 403):
            raise FMPError(f"{endpoint}[{params.get('symbol')}] HTTP {r.status_code}: {r.text[:200]}")
        if r.status_code != 200:
            raise FMPError(f"{endpoint}[{params.get('symbol')}] HTTP {r.status_code}: {r.text[:200]}")
        try:
            data = r.json()
        except ValueError:
            raise FMPError(f"{endpoint}[{params.get('symbol')}] 返回非 JSON: {r.text[:200]}")
        if isinstance(data, dict) and ("Error Message" in data or "error" in data):
            raise FMPError(f"{endpoint}[{params.get('symbol')}] {data}")
        sym = str(params.get("symbol", "")).replace("/", "_")
        (self.raw_dir / f"{sym}_{endpoint}.json").write_text(
            json.dumps(data, indent=1, ensure_ascii=False))
        return data


def pick(d: dict, *names):
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

    inc = fmp.get("income-statement", symbol=sym, period="annual", limit=N_HIST + 1)
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
            "currency": rep_ccy,
            "revenue": pick(e, "revenueAvg", "estimatedRevenueAvg"),
            "gross_profit": None,
            "ebit": pick(e, "ebitAvg", "estimatedEbitAvg"),
            "net_income": pick(e, "netIncomeAvg", "estimatedNetIncomeAvg"),
            "n_analysts_rev": n_rev, "n_analysts_eps": n_eps,
        })

    for i, r in enumerate(rows):
        prev = rows[i - 1]["revenue"] if i > 0 else None
        r["yoy"] = div(r["revenue"], prev) - 1 if div(r["revenue"], prev) is not None else None
        r["gm"] = div(r["gross_profit"], r["revenue"])
        r["opm"] = div(r["ebit"], r["revenue"])
        r["npm"] = div(r["net_income"], r["revenue"])
    n_act = sum(r["kind"] == "actual" for r in rows)
    rows = rows[n_act - N_HIST:] if n_act > N_HIST else rows

    q = quote[0] if quote else {}
    p = profile[0] if profile else {}
    mcap = pick(q, "marketCap") or pick(p, "marketCap", "mktCap")
    price = pick(q, "price") or pick(p, "price")
    quote_ccy = p.get("currency")
    shares = pick(q, "sharesOutstanding") or pick(p, "sharesOutstanding")
    ts = q.get("timestamp")
    qdate = dt.datetime.fromtimestamp(ts, dt.timezone.utc).date().isoformat() if ts else dt.date.today().isoformat()
    if mcap is None:
        w("无市值数据")
    rate = fx_rate(fmp, quote_ccy, rep_ccy)
    if rate is None:
        w(f"无法获取 {quote_ccy}->{rep_ccy} 汇率")
    elif quote_ccy != rep_ccy:
        w(f"报价币种 {quote_ccy} ≠ 报表币种 {rep_ccy}，市值按 {rate:.4f} 换算")
    mcap_rep = mcap * rate if (mcap and rate) else None

    val = {
        "ticker": sym, "name": p.get("companyName"), "date": qdate, "price": price,
        "quote_ccy": quote_ccy, "market_cap": mcap,
        "shares": shares,
        "reporting_ccy": rep_ccy, "fx_quote_to_reporting": rate,
        "market_cap_reporting_ccy": mcap_rep,
    }
    for r in rows:
        if r["kind"] == "estimate":
            r["fwd_pe"] = div(mcap_rep, r["net_income"]) if (r["net_income"] or 0) > 0 else None
            r["fwd_ps"] = div(mcap_rep, r["revenue"])
            val[f"fwd_pe_FY{r['fy']}E"] = r["fwd_pe"]
            val[f"fwd_ps_FY{r['fy']}E"] = r["fwd_ps"]
        else:
            r["fwd_pe"] = r["fwd_ps"] = None
    for r in rows:
        for k in ("revenue", "gross_profit", "ebit", "net_income"):
            r[k] = r[k] / MM if r[k] is not None else None
    return rows, val, warns


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("tickers", nargs="*")
    ap.add_argument("--out", default="fmp_out")
    ap.add_argument("--actual-ebit", default="operatingIncome")
    a = ap.parse_args()

    key = os.environ.get("FMP_API_KEY", "8qnDx2XDkRQEa2MX")
    if not a.tickers:
        sys.exit("请提供 ticker")

    out = Path(a.out)
    fmp = FMP(key, out / "raw")

    all_rows, vals, warns = [], [], []
    for s in a.tickers:
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
            print(f"[OK] {val['ticker']} {val.get('name','')} | mcap={val.get('market_cap',0)/1e6:,.0f}mm")

    out.mkdir(exist_ok=True)
    cols = ["ticker", "kind", "fy", "period_end", "currency", "revenue", "yoy", "gross_profit", "gm",
            "ebit", "opm", "net_income", "npm", "n_analysts_rev", "n_analysts_eps", "fwd_pe", "fwd_ps"]
    pd.DataFrame(all_rows).reindex(columns=cols).to_csv(out / "snapshot_long.csv", index=False)
    pd.DataFrame(vals).to_csv(out / "valuation.csv", index=False)
    pd.DataFrame(warns, columns=["ticker", "issue"]).to_csv(out / "warnings.csv", index=False)

    if warns:
        print("\n--- 警告 ---")
        for x in warns:
            print(f"  {x['ticker']}: {x['issue']}")
    print(f"\n完成：{len(vals)}/{len(a.tickers)} ticker，API {fmp.calls} 次")


if __name__ == "__main__":
    main()
