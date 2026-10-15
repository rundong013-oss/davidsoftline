"""
valuation.py — 估值筛选（Summary + 个股详情）。
"""
import json

import pandas as pd
import streamlit as st

from data.database import get_valuation


def _safe_json(s):
    if not s:
        return []
    try:
        return json.loads(s)
    except (json.JSONDecodeError, TypeError):
        return []


def _safe_json_dict(s):
    if not s:
        return {}
    try:
        return json.loads(s)
    except (json.JSONDecodeError, TypeError):
        return {}


def render():
    st.header("估值筛选")

    df = get_valuation()
    if df.empty:
        st.warning("无估值数据。请先导入 Excel 或运行 API 更新。")
        return

    view = st.radio(
        "视图", ["Summary 一览", "个股详情"],
        horizontal=True, label_visibility="collapsed",
    )

    if view == "Summary 一览":
        _render_summary(df)
    else:
        _render_detail(df)


def _render_summary(df):
    """全部公司 Summary 表。"""
    rows = []
    for _, r in df.iterrows():
        fwd_pe = _safe_json_dict(r.get("fwd_pe_json"))
        fwd_ps = _safe_json_dict(r.get("fwd_ps_json"))

        # 获取 FY+1E 和 FY+2E 的 PE/PS
        pe_keys = sorted([k for k in fwd_pe if k.endswith("E")])
        ps_keys = sorted([k for k in fwd_ps if k.endswith("E")])

        row = {
            "Ticker": r["ticker"],
            "Name": r.get("name", ""),
            "Sector": "运动服饰" if r.get("sector") == "sportswear" else "奢侈品",
            "CCY": r.get("currency", ""),
            "Price": r.get("price"),
            "MCap (mm)": r.get("market_cap"),
            "FY+1E PE": fwd_pe.get(pe_keys[0]) if pe_keys else None,
            "FY+2E PE": fwd_pe.get(pe_keys[1]) if len(pe_keys) > 1 else None,
            "FY+1E PS": fwd_ps.get(ps_keys[0]) if ps_keys else None,
            "FY+2E PS": fwd_ps.get(ps_keys[1]) if len(ps_keys) > 1 else None,
            "Analysts": r.get("n_analysts"),
        }
        rows.append(row)

    summary = pd.DataFrame(rows)

    # 筛选
    sector_filter = st.selectbox(
        "筛选行业", ["全部", "运动服饰", "奢侈品"]
    )
    if sector_filter != "全部":
        summary = summary[summary["Sector"] == sector_filter]

    # 排序选项
    sort_col = st.selectbox(
        "排序", ["Ticker", "MCap (mm)", "FY+1E PE", "FY+2E PE",
                 "FY+1E PS", "FY+2E PS"],
        index=1,
    )
    ascending = sort_col in ("Ticker", "FY+1E PE", "FY+2E PE", "FY+1E PS", "FY+2E PS")
    summary = summary.sort_values(sort_col, ascending=ascending, na_position="last")

    # 格式化
    format_dict = {
        "Price": "{:.2f}",
        "MCap (mm)": "{:,.0f}",
        "FY+1E PE": "{:.1f}x",
        "FY+2E PE": "{:.1f}x",
        "FY+1E PS": "{:.1f}x",
        "FY+2E PS": "{:.1f}x",
    }

    st.dataframe(
        summary.style.format(format_dict, na_rep="—"),
        use_container_width=True,
        height=min(800, len(summary) * 35 + 50),
    )

    st.caption(f"共 {len(summary)} 家公司")


def _render_detail(df):
    """个股详情视图。"""
    tickers = df["ticker"].tolist()
    selected = st.selectbox("选择公司", tickers, index=0)

    row = df[df["ticker"] == selected].iloc[0]
    annual = _safe_json(row.get("annual_json"))
    quarterly = _safe_json(row.get("quarterly_json"))
    fwd_pe = _safe_json_dict(row.get("fwd_pe_json"))
    fwd_ps = _safe_json_dict(row.get("fwd_ps_json"))

    # 标题
    name = row.get("name", selected)
    sector_label = "运动服饰" if row.get("sector") == "sportswear" else "奢侈品"
    st.subheader(f"{selected} — {name} ({sector_label})")

    # 估值信息卡片
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        price = row.get("price")
        ccy = row.get("currency", "")
        st.metric("Price", f"{ccy} {price:.2f}" if price else "—")
    with col2:
        mcap = row.get("market_cap")
        st.metric("MCap (mm)", f"{mcap:,.0f}" if mcap else "—")
    with col3:
        # FY+1E PE
        pe_keys = sorted([k for k in fwd_pe if k.endswith("E")])
        pe_val = fwd_pe.get(pe_keys[0]) if pe_keys else None
        st.metric(f"FY+1E PE", f"{pe_val:.1f}x" if pe_val else "—")
    with col4:
        ps_keys = sorted([k for k in fwd_ps if k.endswith("E")])
        ps_val = fwd_ps.get(ps_keys[0]) if ps_keys else None
        st.metric(f"FY+1E PS", f"{ps_val:.1f}x" if ps_val else "—")

    st.divider()

    # ── 年度数据 ──────────────────────────────────────────────────
    if annual:
        st.markdown("**年度数据**")
        ann_df = _build_period_table(annual)
        st.dataframe(ann_df, use_container_width=True)

    # ── 季度数据 ──────────────────────────────────────────────────
    if quarterly:
        st.markdown("**季度数据**")
        qtr_df = _build_period_table(quarterly)
        st.dataframe(qtr_df, use_container_width=True)

    # ── Forward PE/PS ─────────────────────────────────────────────
    if fwd_pe or fwd_ps:
        st.markdown("**Forward Multiples**")
        mult_data = {}
        if fwd_pe:
            mult_data["Forward PE"] = fwd_pe
        if fwd_ps:
            mult_data["Forward PS"] = fwd_ps
        mult_df = pd.DataFrame(mult_data)
        st.dataframe(mult_df.style.format("{:.1f}x", na_rep="—"), use_container_width=True)

    # Analysts
    n_analysts = row.get("n_analysts")
    if n_analysts:
        st.caption(f"Analyst coverage: {n_analysts} analysts (revenue estimates)")


def _build_period_table(periods):
    """将 annual/quarterly 的 JSON 列表转为展示用 DataFrame。"""
    if not periods:
        return pd.DataFrame()

    labels = [p.get("label", "") for p in periods]
    metrics = {
        "Revenue (mm)": [],
        "YoY %": [],
        "GM %": [],
        "OPM %": [],
        "Net Income (mm)": [],
        "NPM %": [],
        "EPS": [],
        "Shares (mm)": [],
    }
    for p in periods:
        rev = p.get("revenue")
        metrics["Revenue (mm)"].append(
            f"{rev:,.0f}" if rev and isinstance(rev, (int, float)) else "—"
        )
        yoy = p.get("yoy")
        metrics["YoY %"].append(f"{yoy:.1%}" if yoy and isinstance(yoy, (int, float)) else "—")
        gm = p.get("gm")
        metrics["GM %"].append(f"{gm:.1%}" if gm and isinstance(gm, (int, float)) else "—")
        opm = p.get("opm")
        metrics["OPM %"].append(f"{opm:.1%}" if opm and isinstance(opm, (int, float)) else "—")
        ni = p.get("net_income")
        metrics["Net Income (mm)"].append(
            f"{ni:,.0f}" if ni and isinstance(ni, (int, float)) else "—"
        )
        npm = p.get("npm")
        metrics["NPM %"].append(f"{npm:.1%}" if npm and isinstance(npm, (int, float)) else "—")
        eps = p.get("eps")
        metrics["EPS"].append(f"{eps:.2f}" if eps and isinstance(eps, (int, float)) else "—")
        shr = p.get("shares")
        metrics["Shares (mm)"].append(
            f"{shr:,.1f}" if shr and isinstance(shr, (int, float)) else "—"
        )

    df = pd.DataFrame(metrics, index=labels).T
    return df
