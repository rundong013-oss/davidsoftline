#!/usr/bin/env python3
"""
match_setup.py — 用历史面板数据，给一家公司当前的 set-up 找最像的历史大机会起点。

定量匹配只负责"缩小范围"：给出指纹最接近的几个历史起点及其后来的路径。
最终是不是同一类机会，要再用 cases.json 里的人事/产品/品牌/渠道/Beta 定性对照确认。

用法（在 skill 根目录下运行）：
  # A. 公司在面板里：直接取某个月的指纹（回测或看历史）
  python scripts/match_setup.py --company PUM --month 2024-09

  # B. 公司不在面板里，或要用最新数据：手动输入指纹（缺的项不填即可）
  python scripts/match_setup.py --npm 0.018 --ps 0.35 --ps-pct 0.05 --rev-vs-peak -0.10 \
      --rev-g12 -0.04 --npm-chg12 -0.012 --rev 45 --pe 19

  常用参数：
  --type 困境反转|自身成长|行业Beta   只在该类型的历史起点里找
  --top 5                            返回几个类比
  --exclude-self                     回测时排除同一公司的窗口
  --include-rebound                  候选中包含 2020 年疫情后的纯反弹窗口（默认排除，因其由宏观驱动）
  --screen [--month YYYY-MM]         对面板里所有公司在某月做指纹扫描（默认最新月）
  --panel PATH                       使用追加了最新数据的面板（列同 data/monthly_panel.csv）

指纹定义（全部基于 Forward 一致预期，月度）：
  npm         Forward 净利率
  ps          Forward PS
  pe          Forward PE
  ps_pct      PS 在自身过去最多 120 个月中的分位（0 = 历史最低）
  rev_vs_peak Forward 收入相对过去最多 120 个月峰值的回撤
  rev_g12     Forward 收入 12 个月变化（预期的方向）
  npm_chg12   Forward 净利率 12 个月变化（百分点）
  rev         Forward 收入规模（亿美元；Adidas 为欧元口径）
"""
import argparse
import json
import math
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"

FEATURES = ["npm", "ps_log", "pe_log", "ps_pct", "rev_vs_peak", "rev_g12", "npm_chg12", "rev_log"]
# 权重：利润率水平、估值分位、收入回撤最能区分机会类型；规模次之
# 距离阈值：回测中历史案例起点的首个类比距离中位数约 0.5、90 分位约 0.7；
# 全部公司-月份的首个类比距离 90 分位约 0.8。超过阈值视为"没有足够相似的历史起点"。
WEAK_MATCH = 0.8

WEIGHTS = {"npm": 2.0, "ps_log": 1.5, "pe_log": 0.5, "ps_pct": 1.5, "rev_vs_peak": 1.0,
           "rev_g12": 1.0, "npm_chg12": 1.0, "rev_log": 0.5}


def month_index(m):
    return int(m[:4]) * 12 + int(m[5:7]) - 1


def load_panel(path):
    p = pd.read_csv(path)
    p = p.sort_values(["company", "month"]).reset_index(drop=True)
    out = []
    for co, g in p.groupby("company"):
        g = g.copy()
        g["mi"] = g.month.map(month_index)
        g = g.set_index("mi").sort_index()
        full = g.reindex(range(g.index.min(), g.index.max() + 1))
        full["company"] = co
        full["month"] = [f"{i // 12}-{i % 12 + 1:02d}" for i in full.index]
        ps, rev, npm = full.fwd_ps, full.fwd_rev, full.fwd_npm
        full["ps_pct"] = ps.rolling(120, min_periods=24).apply(
            lambda x: (np.sum(x[~np.isnan(x)] < x[-1]) / max(np.sum(~np.isnan(x)) - 1, 1)) if not np.isnan(x[-1]) else np.nan,
            raw=True)
        full["rev_vs_peak"] = rev / rev.rolling(120, min_periods=12).max() - 1
        full["rev_g12"] = rev / rev.shift(12) - 1
        full["npm_chg12"] = npm - npm.shift(12)
        full["mcap_fwd12"] = full.mkt_cap.shift(-12) / full.mkt_cap - 1
        full["mcap_fwd24"] = full.mkt_cap.shift(-24) / full.mkt_cap - 1
        full["npm_fwd12"] = npm.shift(-12)
        full["ps_fwd12"] = ps.shift(-12)
        out.append(full.reset_index(drop=True))
    df = pd.concat(out, ignore_index=True)
    return add_transforms(df)


def add_transforms(df):
    df = df.rename(columns={"fwd_npm": "npm", "fwd_ps": "ps", "fwd_pe": "pe", "fwd_rev": "rev"}) \
        if "fwd_npm" in df.columns else df
    df["ps_log"] = np.log(df.ps.where(df.ps > 0))
    df["pe_log"] = np.log(df.pe.where((df.pe > 0) & (df.pe < 300)))
    df["rev_log"] = np.log(df.rev.where(df.rev > 0))
    return df


def load_cases():
    return json.loads((DATA / "cases.json").read_text())


def q_start(window):
    """'22Q3–24Q4' -> '2022-07'"""
    w = window.replace(" ", "").replace("-", "–")
    try:
        a = w.split("–")[0]
        yy, q = a.split("Q")
        return f"20{yy}-{3 * int(q) - 2:02d}"
    except Exception:
        return None


def candidate_starts(panel, cases):
    """历史起点 = 案例起点 ∪ 未对应案例的高亮窗口起点。"""
    wins = pd.read_csv(DATA / "windows.csv")
    cands, used = [], set()
    for c in cases:
        s = q_start(c["window"])
        pw = c.get("panel_window")
        if pw:
            used.add((c["company"], pw["start"]))
            # 高亮窗口起点与案例起始季度相差不超过 6 个月时，以高亮窗口的起点月为准
            if s and abs(month_index(pw["start"]) - month_index(s)) <= 6:
                s = pw["start"]
        cands.append({"company": c["company"], "start": s, "label": f"{c['company']} {c['window']}",
                      "types": c.get("types") or [], "groups": c.get("groups") or [],
                      "return": c.get("return_multiple"), "case": c})
    for _, w in wins.iterrows():
        if (w.company, w.start) in used:
            continue
        rebound = "2020-03" <= w.start <= "2020-10"
        cands.append({"company": w.company, "start": w.start, "label": f"{w.company} {w.start}~{w.end}",
                      "types": ["行业Beta"] if rebound else ["未分类"],
                      "groups": ["2020 年疫情放水反弹（纯反弹）"] if rebound else [],
                      "return": None if pd.isna(w["return"]) else float(w["return"]), "case": None})
    return cands


def scales(panel):
    s = {}
    for f in FEATURES:
        v = panel[f].replace([np.inf, -np.inf], np.nan).dropna()
        s[f] = v.std() if len(v) > 10 and v.std() > 0 else 1.0
    return s


def distance(a, b, sc):
    num, den, used = 0.0, 0.0, []
    for f in FEATURES:
        x, y = a.get(f), b.get(f)
        if x is None or y is None or (isinstance(x, float) and math.isnan(x)) or (isinstance(y, float) and math.isnan(y)):
            continue
        w = WEIGHTS[f]
        num += w * ((x - y) / sc[f]) ** 2
        den += w
        used.append(f)
    if den < 4.0:          # 可比较的指纹太少（权重合计不足 4）
        return None, used
    return math.sqrt(num / den), used


def row_at(panel, co, month):
    r = panel[(panel.company == co) & (panel.month == month)]
    return None if r.empty else r.iloc[0]


def fmt(v, kind):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return "—"
    return {"pct": f"{v:.1%}", "x": f"{v:.2f}x", "pe": f"{v:.1f}x", "num": f"{v:,.0f}", "chg": f"{v:+.1%}",
            "pp": f"{v * 100:+.1f}pp"}[kind]


def describe(r):
    return {"npm": r.get("npm"), "ps": r.get("ps"), "pe": r.get("pe"), "ps_pct": r.get("ps_pct"),
            "rev_vs_peak": r.get("rev_vs_peak"), "rev_g12": r.get("rev_g12"),
            "npm_chg12": r.get("npm_chg12"), "rev": r.get("rev")}


def print_fp(title, d):
    print(f"{title:<28} 净利率 {fmt(d.get('npm'), 'pct'):>7} | PS {fmt(d.get('ps'), 'x'):>6} | PE {fmt(d.get('pe'), 'pe'):>6} | "
          f"PS分位 {fmt(d.get('ps_pct'), 'pct'):>6} | 收入距峰值 {fmt(d.get('rev_vs_peak'), 'chg'):>7} | "
          f"收入预期12M {fmt(d.get('rev_g12'), 'chg'):>7} | 净利率12M {fmt(d.get('npm_chg12'), 'pp'):>8} | 收入 {fmt(d.get('rev'), 'num'):>5}")


def match(panel, target, args, self_company=None):
    sc = scales(panel)
    cases = load_cases()
    res = []
    for c in candidate_starts(panel, cases):
        if not c["start"]:
            continue
        if args.type and args.type not in c["types"]:
            continue
        if not getattr(args, "include_rebound", False) and any("纯反弹" in g for g in c["groups"]):
            continue
        if args.exclude_self and self_company and c["company"] == self_company:
            continue
        r = row_at(panel, c["company"], c["start"])
        if r is None:
            continue
        feat = {f: (None if pd.isna(r[f]) else float(r[f])) for f in FEATURES}
        d, used = distance(target, feat, sc)
        if d is None:
            continue
        res.append((d, c, r, used))
    res.sort(key=lambda x: x[0])
    return res[: args.top]


def target_from_args(a):
    d = {"npm": a.npm, "ps": a.ps, "pe": a.pe, "ps_pct": a.ps_pct, "rev_vs_peak": a.rev_vs_peak,
         "rev_g12": a.rev_g12, "npm_chg12": a.npm_chg12, "rev": a.rev}
    t = dict(d)
    t["ps_log"] = math.log(a.ps) if a.ps and a.ps > 0 else None
    t["pe_log"] = math.log(a.pe) if a.pe and 0 < a.pe < 300 else None
    t["rev_log"] = math.log(a.rev) if a.rev and a.rev > 0 else None
    return t, d


def screen(panel, month):
    snap = panel[(panel.month == month) & (panel.npm.notna() | panel.ps.notna())]
    rows = []
    for _, r in snap.iterrows():
        flags = []
        if r.npm is not None and not pd.isna(r.npm) and r.npm <= 0.03:
            flags.append("净利率≤3%")
        if not pd.isna(r.ps_pct) and r.ps_pct <= 0.10:
            flags.append("PS≤10%分位")
        if not pd.isna(r.rev_vs_peak) and r.rev_vs_peak <= -0.08:
            flags.append("收入距峰值≤-8%")
        if not pd.isna(r.npm_chg12) and r.npm_chg12 > 0.005 and not pd.isna(r.npm) and r.npm <= 0.05:
            flags.append("低利润率且在修复")
        if not pd.isna(r.rev_g12) and r.rev_g12 >= 0.25 and not pd.isna(r.rev) and r.rev <= 30:
            flags.append("小体量高增长")
        rows.append((r.company, describe(r), flags))
    print(f"\n面板扫描：{month}（仅供发现候选，结论需结合定性判断）\n")
    for co, d, flags in sorted(rows, key=lambda x: -len(x[2])):
        print_fp(co, d)
        if flags:
            print(f"{'':<28} → {'；'.join(flags)}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--company"); ap.add_argument("--month")
    for f in ["npm", "ps", "pe", "ps-pct", "rev-vs-peak", "rev-g12", "npm-chg12", "rev"]:
        ap.add_argument(f"--{f}", type=float)
    ap.add_argument("--type"); ap.add_argument("--top", type=int, default=5)
    ap.add_argument("--exclude-self", action="store_true")
    ap.add_argument("--include-rebound", action="store_true", help="候选中包含 2020 年疫情纯反弹窗口（默认排除）")
    ap.add_argument("--screen", action="store_true")
    ap.add_argument("--panel", default=str(DATA / "monthly_panel.csv"))
    a = ap.parse_args()

    panel = load_panel(a.panel)
    if a.screen:
        cnt = panel[panel.npm.notna()].groupby("month").size()
        default = cnt[cnt >= cnt.max() * 0.5].index.max()
        screen(panel, a.month or default)
        return

    self_co = None
    if a.company:
        valid = panel[(panel.company == a.company) & panel.npm.notna()]
        month = a.month or (valid.month.max() if not valid.empty else panel[panel.company == a.company].month.max())
        r = row_at(panel, a.company, month)
        if r is None:
            raise SystemExit(f"面板中没有 {a.company} {month} 的数据。可用公司：{sorted(panel.company.unique())}")
        target = {f: (None if pd.isna(r[f]) else float(r[f])) for f in FEATURES}
        shown = describe(r)
        title = f"{a.company} {month}"
        self_co = a.company
    else:
        target, shown = target_from_args(a)
        title = "当前公司（手动输入）"

    print("\n=== 当前指纹 ===")
    print_fp(title, shown)
    results = match(panel, target, a, self_co)
    print(f"\n=== 最相近的历史起点（前 {a.top} 个；距离越小越像；d > {WEAK_MATCH} 视为不够相似）===")
    if not results or results[0][0] > WEAK_MATCH:
        best = f"{results[0][0]:.2f}" if results else "—"
        print(f"\n>>> 没有足够相似的历史起点（最小距离 {best}）。不要强行类比。")
    for d, c, r, used in results:
        print()
        flag = "" if d <= WEAK_MATCH else " 不够相似"
        print_fp(f"{c['label']}  [d={d:.2f}{flag}]", describe(r))
        tags = "、".join(c["types"]) + (f"｜{c['groups'][0].splitlines()[0][:40]}" if c["groups"] else "")
        ret = f"窗口涨幅 {c['return']:+.0%}" if c.get("return") is not None else ""
        print(f"{'':<28} 类型：{tags}  {ret}")
        print(f"{'':<28} 起点后 12M：市值 {fmt(r.get('mcap_fwd12'), 'chg')}，净利率 → {fmt(r.get('npm_fwd12'), 'pct')}，"
              f"PS → {fmt(r.get('ps_fwd12'), 'x')}；24M 市值 {fmt(r.get('mcap_fwd24'), 'chg')}")
        if c["case"]:
            b, af = c["case"].get("before", {}), c["case"].get("after", {})
            print(f"{'':<28} 兑现前→后：收入 {fmt(b.get('fwd_rev'), 'num')}→{fmt(af.get('fwd_rev'), 'num')}，"
                  f"净利率 {fmt(b.get('fwd_npm'), 'pct')}→{fmt(af.get('fwd_npm'), 'pct')}，"
                  f"PS {fmt(b.get('fwd_ps'), 'x')}→{fmt(af.get('fwd_ps'), 'x')}")
        print(f"{'':<28} 参与比较的指纹：{', '.join(used)}")


if __name__ == "__main__":
    main()
