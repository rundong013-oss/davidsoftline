"""
views/thesis_bank.py — Thesis Bank 页面：按公司浏览投资论文。
"""
import re
from pathlib import Path

import streamlit as st

from config import (
    LIBRARY_DIR,
    SPORTSWEAR_UNIVERSE,
    LUXURY_UNIVERSE,
)

# ── 常量 ────────────────────────────────────────────────────────
_STUB_THRESHOLD = 2048  # 文件 < 2KB 视为待补充


# ── 数据层 ──────────────────────────────────────────────────────
def _build_lib_dir_lookup():
    """构建 library 目录名 → {ticker, name, sector} 映射。

    universe.yaml 中 ticker 带交易所后缀（NKE.N），library 目录名通常去掉后缀
    （NKE），但 .HK / .SZ 保留。此函数推导出目录名并建立映射。
    """
    mapping = {}
    sectors = [
        ("运动服饰", SPORTSWEAR_UNIVERSE),
        ("奢侈品", LUXURY_UNIVERSE),
    ]
    for sector_name, entries in sectors:
        for entry in entries:
            ticker = entry["ticker"]
            name = entry["name"]
            # 去掉常见交易所后缀，但保留 .HK / .SZ
            lib_dir = ticker
            for suffix in (".N", ".O", ".PA", ".SW", ".MI", ".L", ".TO", ".DF", ".T"):
                if ticker.endswith(suffix):
                    lib_dir = ticker[: -len(suffix)]
                    break
            mapping[lib_dir] = {
                "ticker": ticker,
                "name": name,
                "sector": sector_name,
            }
    return mapping


_LIB_DIR_LOOKUP = _build_lib_dir_lookup()

# 元数据正则
_RE_OPP_SIZE = re.compile(r"机会大小：\*\*(.+?)\*\*")
_RE_OPP_TYPE = re.compile(r"机会类型：(.+?)$", re.MULTILINE)
_RE_DATE = re.compile(r"_thesis_(\d{4}-\d{2}-\d{2})\.md$")


def _date_sort_key(date_str: str) -> int:
    """将 YYYY-MM-DD 转为整数用于数值排序（可取负数实现倒序）。"""
    return int(date_str.replace("-", ""))


def _scan_theses():
    """扫描 library/ 下所有 thesis markdown，返回结构化列表。

    返回: [{lib_dir, date, path, opp_size, opp_type, is_stub, ticker, name, sector}, ...]
    """
    results = []
    for md_path in sorted(LIBRARY_DIR.glob("*/*_thesis_*.md")):
        # 解析目录名和日期
        lib_dir = md_path.parent.name
        m_date = _RE_DATE.search(md_path.name)
        if not m_date:
            continue
        date_str = m_date.group(1)

        # 查找对应公司信息
        info = _LIB_DIR_LOOKUP.get(lib_dir)
        if not info:
            # 未在 universe 中找到，跳过
            continue

        # 文件大小判断 stub
        is_stub = md_path.stat().st_size < _STUB_THRESHOLD

        # 解析元数据（前 5 行）
        head = md_path.read_text(encoding="utf-8")[:500]
        m_size = _RE_OPP_SIZE.search(head)
        m_type = _RE_OPP_TYPE.search(head)
        opp_size = m_size.group(1) if m_size else "—"
        opp_type = m_type.group(1).strip() if m_type else "—"

        results.append({
            "lib_dir": lib_dir,
            "date": date_str,
            "path": md_path,
            "opp_size": opp_size,
            "opp_type": opp_type,
            "is_stub": is_stub,
            "ticker": info["ticker"],
            "name": info["name"],
            "sector": info["sector"],
        })
    return results


# ── Sidebar 筛选 ────────────────────────────────────────────────
def render_sidebar():
    """在 sidebar 中渲染 Thesis Bank 筛选控件，返回选中的 lib_dir 或 None。"""
    theses = _scan_theses()

    # 按 lib_dir 分组，记录每家公司信息
    companies = {}
    for t in theses:
        key = t["lib_dir"]
        if key not in companies:
            companies[key] = {
                "ticker": t["ticker"],
                "name": t["name"],
                "sector": t["sector"],
                "versions": [],
            }
        companies[key]["versions"].append(t)
    # 每家公司的版本按日期倒序
    for comp in companies.values():
        comp["versions"].sort(key=lambda x: x["date"], reverse=True)

    # 行业筛选
    sector_options = ["全部", "运动服饰", "奢侈品"]
    sector = st.sidebar.selectbox("行业", sector_options, index=0)

    # 按行业过滤公司列表
    if sector == "全部":
        filtered = companies
    else:
        filtered = {k: v for k, v in companies.items() if v["sector"] == sector}

    # 公司选择（按最新更新日期倒序，同日期按 ticker 字母序）
    sorted_items = sorted(
        filtered.items(),
        key=lambda x: (-_date_sort_key(x[1]["versions"][0]["date"]), x[1]["ticker"]),
    )
    company_labels = ["总览"] + [
        f"{v['ticker']} — {v['name']}" for k, v in sorted_items
    ]
    company_keys = [None] + [k for k, v in sorted_items]
    # 默认选中 ONON
    default_idx = 0
    if "ONON" in company_keys:
        default_idx = company_keys.index("ONON")

    selected_idx = st.sidebar.selectbox(
        "公司",
        range(len(company_labels)),
        format_func=lambda i: company_labels[i],
        index=default_idx,
    )
    selected_lib_dir = company_keys[selected_idx]

    # 版本选择（多版本时）
    selected_date = None
    if selected_lib_dir and selected_lib_dir in filtered:
        versions = filtered[selected_lib_dir]["versions"]
        if len(versions) > 1:
            date_labels = [v["date"] for v in versions]
            selected_date = st.sidebar.selectbox("版本", date_labels, index=0)
        else:
            selected_date = versions[0]["date"]

    return theses, companies, filtered, selected_lib_dir, selected_date


# ── 主页面渲染 ──────────────────────────────────────────────────
def _render_overview(filtered):
    """渲染公司总览表。"""
    st.header("Thesis Bank")
    st.caption("投资论文总览 — 点击公司名称查看详细内容")

    # 构建表格数据
    rows = []
    for lib_dir, comp in sorted(
        filtered.items(),
        key=lambda x: (
            x[1]["versions"][0]["is_stub"],
            -_date_sort_key(x[1]["versions"][0]["date"]),
            x[1]["ticker"],
        ),
    ):
        latest = comp["versions"][0]
        rows.append({
            "Ticker": comp["ticker"],
            "公司": comp["name"],
            "行业": comp["sector"],
            "最新日期": latest["date"],
            "机会大小": latest["opp_size"],
            "版本数": len(comp["versions"]),
            "状态": "待补充" if latest["is_stub"] else "✓",
        })

    if not rows:
        st.info("没有找到 thesis 文件。")
        return

    st.dataframe(
        rows,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Ticker": st.column_config.TextColumn(width="small"),
            "公司": st.column_config.TextColumn(width="medium"),
            "行业": st.column_config.TextColumn(width="small"),
            "最新日期": st.column_config.TextColumn(width="small"),
            "机会大小": st.column_config.TextColumn(width="small"),
            "版本数": st.column_config.NumberColumn(width="small"),
            "状态": st.column_config.TextColumn(width="small"),
        },
    )

    # 统计
    total = len(rows)
    filled = sum(1 for r in rows if r["状态"] == "✓")
    st.caption(f"共 {total} 家公司，已完成 {filled} 篇，待补充 {total - filled} 篇")


def _render_thesis(companies, lib_dir, selected_date):
    """渲染单篇 thesis 内容。"""
    comp = companies[lib_dir]
    versions = comp["versions"]

    # 找到选中版本
    thesis = None
    for v in versions:
        if v["date"] == selected_date:
            thesis = v
            break
    if thesis is None:
        thesis = versions[0]

    # 公司信息卡片
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Ticker", comp["ticker"])
    col2.metric("公司", comp["name"])
    col3.metric("机会大小", thesis["opp_size"])
    col4.metric("日期", thesis["date"])

    if thesis["opp_type"] != "—":
        st.caption(f"机会类型：{thesis['opp_type']}")

    if thesis["is_stub"]:
        st.warning("此 thesis 为待补充状态（模板），尚无完整内容。")

    st.divider()

    # 渲染 markdown 内容
    content = thesis["path"].read_text(encoding="utf-8")
    # 转义 $ 符号，防止 Streamlit 将 $2B 等金额误解析为 LaTeX 公式
    content = content.replace("$", r"\$")
    st.markdown(content, unsafe_allow_html=False)


# ── 入口 ────────────────────────────────────────────────────────
def render():
    """Thesis Bank 页面入口。"""
    theses, companies, filtered, selected_lib_dir, selected_date = render_sidebar()

    if selected_lib_dir is None:
        _render_overview(filtered)
    else:
        _render_thesis(companies, selected_lib_dir, selected_date)
