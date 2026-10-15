#!/usr/bin/env python3
"""
render_obsidian.py — 把每家公司的扫描结果（JSON）写成 Obsidian 笔记，并重建 summary.md。

用法：
  python scripts/render_obsidian.py --vault <Obsidian vault 路径> [--dir 机会扫描] NKE.json ANF.json ...
  python scripts/render_obsidian.py --vault <路径> --summary-only        # 只根据现有笔记重建 summary

输出：
  <vault>/<dir>/<TICKER>.md   每家公司一份，分两部分：一、基本面；二、投资机会判断
  <vault>/<dir>/summary.md    所有公司的汇总表，按机会大小排序

JSON 结构见 references/output-template.md。重复扫描同一公司时，正文被覆盖，
"历史扫描记录"一节保留并追加本次记录（同一天重复扫描只保留最后一次）。
"""
import argparse
import json
import re
from pathlib import Path

SIZE_ORDER = {"大": 0, "中": 1, "小": 2, "—": 3}
PCT_HINT = re.compile(r"率|同比|增速|占比|OPM|margin|growth|yoy", re.I)
DIMENSIONS = ["管理层", "战略", "产品", "客群", "渠道"]


# ---------- 格式化 ----------
def fmt_val(name, v):
    if v is None or v == "":
        return "—"
    if isinstance(v, str):
        return v
    if isinstance(v, dict):                       # 业务拆分：{"rev": 1234, "yoy": 0.05}
        rev = fmt_val("收入", v.get("rev"))
        yoy = v.get("yoy")
        return rev if yoy is None else f"{rev}（{yoy:+.1%}）"
    if isinstance(v, (int, float)):
        if PCT_HINT.search(name) and abs(v) <= 5:
            return f"{v:+.1%}" if re.search(r"同比|增速|growth|yoy", name, re.I) else f"{v:.1%}"
        return f"{v:,.0f}" if abs(v) >= 100 else f"{v:,.1f}"
    return str(v)


def table(block, first_col="指标"):
    if not block or not block.get("rows"):
        return "_无数据_\n"
    periods = block["periods"]
    lines = [f"| {first_col} | " + " | ".join(periods) + " |",
             "|---|" + "---|" * len(periods)]
    for r in block["rows"]:
        name = r["name"]
        label = ("　　" * (r.get("level", 1) - 1)) + name
        vals = [fmt_val(name, v) for v in r["values"]] + ["—"] * (len(periods) - len(r["values"]))
        lines.append(f"| {label} | " + " | ".join(vals[:len(periods)]) + " |")
    out = "\n".join(lines) + "\n"
    if block.get("note"):
        out += f"\n_{block['note']}_\n"
    return out


def changes_table(ch):
    if not ch or not ch.get("rows"):
        return "_缺失_\n"
    periods = ch["periods"]
    lines = ["| 维度 | " + " | ".join(periods) + " |", "|---|" + "---|" * len(periods)]
    for dim in DIMENSIONS:
        vals = ch["rows"].get(dim, [])
        vals = [(v or "缺失").replace("\n", "<br>").replace("|", "／") for v in vals]
        vals += ["缺失"] * (len(periods) - len(vals))
        lines.append(f"| {dim} | " + " | ".join(vals[:len(periods)]) + " |")
    return "\n".join(lines) + "\n"


def yq(s):
    return json.dumps(s if s is not None else "", ensure_ascii=False)


# ---------- 公司笔记 ----------
def read_history(path):
    if not path.exists():
        return []
    text = path.read_text(encoding="utf-8")
    m = re.search(r"## 历史扫描记录\n(.*)$", text, re.S)
    return [l for l in m.group(1).splitlines() if l.startswith("- ")] if m else []


def render_company(d, folder):
    t, f, o = d["ticker"], d.get("fundamentals", {}), d.get("opportunity", {})
    path = folder / f"{t}.md"
    size = o.get("size") or "—"
    history = [h for h in read_history(path) if not h.startswith(f"- {d['scan_date']}")]
    history.append(f"- {d['scan_date']}｜{size}｜{o.get('type') or 'set-up 一般'}｜{o.get('setup') or ''}")

    unit = f"（{d.get('currency', '')}，{d.get('unit', '百万')}）"
    md = f"""---
ticker: {yq(t)}
name: {yq(d.get('name'))}
scan_date: {yq(d['scan_date'])}
size: {yq(size)}
type: {yq(o.get('type'))}
setup: {yq(o.get('setup'))}
analog: {yq(o.get('analog'))}
watch: {yq(o.get('watch'))}
tags: [机会扫描]
---

# {t} {d.get('name', '')}

扫描日期：{d['scan_date']}｜机会大小：**{size}**｜机会类型：{o.get('type') or 'set-up 一般'}

## 一、基本面

### 年度财务（过去 3 个财年）{unit}
{table(f.get('annual'))}
### 季度财务（过去 3 个季度）{unit}
{table(f.get('quarterly'))}
### 业务拆分（公司口径，过去 3 个季度）
{table(f.get('segments'), '业务')}
### 关键经营数据
{table(f.get('kpis'))}
### 小结
{f.get('summary') or '_缺失_'}

## 二、投资机会判断

{o.get('conclusion') or '_缺失_'}

**机会大小：{size}**。{o.get('size_reason') or ''}

### 最近 3 个季度基本面变化（依据业绩会 transcript）
{changes_table(o.get('changes'))}
### 数据来源与缺口
{d.get('sources') or '_未注明_'}

## 历史扫描记录
""" + "\n".join(history) + "\n"
    path.write_text(md, encoding="utf-8")
    return path


# ---------- summary ----------
def parse_frontmatter(text):
    m = re.match(r"---\n(.*?)\n---", text, re.S)
    fm = {}
    if not m:
        return fm
    for line in m.group(1).splitlines():
        if ":" in line:
            k, v = line.split(":", 1)
            v = v.strip()
            try:
                fm[k.strip()] = json.loads(v)
            except json.JSONDecodeError:
                fm[k.strip()] = v
    return fm


def render_summary(folder):
    rows = []
    for p in sorted(folder.glob("*.md")):
        if p.name == "summary.md":
            continue
        fm = parse_frontmatter(p.read_text(encoding="utf-8"))
        if "机会扫描" not in str(fm.get("tags", "")):
            continue
        rows.append(fm)
    rows.sort(key=lambda r: (SIZE_ORDER.get(r.get("size") or "—", 9), r.get("ticker", "")))
    counts = {k: sum(1 for r in rows if (r.get("size") or "—") == k) for k in SIZE_ORDER}
    cell = lambda s: (s or "").replace("|", "／").replace("\n", " ")
    lines = ["# 机会扫描 summary", "",
             f"共 {len(rows)} 家｜大 {counts['大']}｜中 {counts['中']}｜小 {counts['小']}｜set-up 一般 {counts['—']}", "",
             "| 公司 | 扫描日期 | 机会大小 | 机会类型 | set-up | 类比 | 关键观测 |",
             "|---|---|---|---|---|---|---|"]
    for r in rows:
        lines.append(f"| [[{r.get('ticker')}]] {cell(r.get('name'))} | {r.get('scan_date', '')} | {r.get('size') or '—'} | "
                     f"{cell(r.get('type')) or 'set-up 一般'} | {cell(r.get('setup'))} | {cell(r.get('analog')) or '—'} | "
                     f"{cell(r.get('watch')) or '—'} |")
    (folder / "summary.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    return folder / "summary.md"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="*")
    ap.add_argument("--vault", required=True)
    ap.add_argument("--dir", default="机会扫描")
    ap.add_argument("--summary-only", action="store_true")
    a = ap.parse_args()
    folder = Path(a.vault).expanduser() / a.dir
    folder.mkdir(parents=True, exist_ok=True)
    if not a.summary_only:
        for fp in a.files:
            d = json.loads(Path(fp).read_text(encoding="utf-8"))
            print("写入", render_company(d, folder))
    print("写入", render_summary(folder))


if __name__ == "__main__":
    main()
