"""
earnings_calendar.py — Earnings 日历页面。
"""
import datetime as dt

import pandas as pd
import streamlit as st

from data.database import get_earnings


def get_this_week_earnings(df):
    """返回本周（下7天内）出 earnings 的公司。"""
    if df.empty:
        return df
    today = dt.date.today()
    week_end = today + dt.timedelta(days=7)
    df_copy = df.copy()
    df_copy["_edate"] = pd.to_datetime(df_copy["earnings_date"]).dt.date
    mask = (df_copy["_edate"] >= today) & (df_copy["_edate"] <= week_end)
    return df_copy[mask].drop(columns=["_edate"])


def render():
    st.header("Earnings 日历")

    df = get_earnings()
    if df.empty:
        st.warning("无 Earnings 数据。请先导入 Excel 或运行 API 更新。")
        return

    # 更新 days_from_today（相对于今天）
    today = dt.date.today()
    df["earnings_date_dt"] = pd.to_datetime(df["earnings_date"]).dt.date
    df["days_from_today"] = df["earnings_date_dt"].apply(
        lambda d: (d - today).days if pd.notna(d) else None
    )

    # ── 本周 Earnings 提示 ────────────────────────────────────────
    this_week = get_this_week_earnings(df)
    if not this_week.empty:
        st.markdown("### 本周 Earnings")
        cols = st.columns(min(4, len(this_week)))
        for i, (_, row) in enumerate(this_week.iterrows()):
            with cols[i % len(cols)]:
                sector_emoji = ""
                sector_label = ""
                if row.get("sector") == "sportswear":
                    sector_label = "运动服饰"
                elif row.get("sector") == "luxury":
                    sector_label = "奢侈品"
                days = row.get("days_from_today", "")
                date_str = str(row["earnings_date"])[:10]
                eps = row.get("eps_estimate", "")

                st.markdown(f"""
                **{row['ticker']}**
                {row.get('name', '')}
                {date_str} ({days}天后)
                {f'EPS Est: {eps}' if eps else ''}
                *{sector_label}*
                """)
        st.divider()

    # ── 筛选 ──────────────────────────────────────────────────────
    col1, col2 = st.columns([1, 3])
    with col1:
        filter_opt = st.radio(
            "筛选",
            ["全部", "运动服饰", "奢侈品", "30天内"],
            horizontal=False,
        )

    filtered = df.copy()
    if filter_opt == "运动服饰":
        filtered = filtered[filtered["sector"] == "sportswear"]
    elif filter_opt == "奢侈品":
        filtered = filtered[filtered["sector"] == "luxury"]
    elif filter_opt == "30天内":
        filtered = filtered[
            (filtered["days_from_today"] >= 0) &
            (filtered["days_from_today"] <= 30)
        ]

    # 排序
    filtered = filtered.sort_values("earnings_date")

    # ── 表格 ──────────────────────────────────────────────────────
    with col2:
        display_df = filtered[[
            "ticker", "name", "sector", "earnings_date",
            "days_from_today", "eps_estimate", "fiscal_quarter",
        ]].copy()
        display_df.columns = [
            "Ticker", "Company", "Sector", "Earnings Date",
            "Days", "EPS Est", "Quarter",
        ]
        # Sector 中文化
        display_df["Sector"] = display_df["Sector"].map(
            {"sportswear": "运动服饰", "luxury": "奢侈品"}
        ).fillna(display_df["Sector"])

        def _highlight_soon(row):
            days = row.get("Days")
            if days is not None and 0 <= days <= 30:
                return ["background-color: #fff8e1"] * len(row)
            return [""] * len(row)

        styled = display_df.style.apply(_highlight_soon, axis=1)
        st.dataframe(
            styled,
            use_container_width=True,
            height=min(800, len(display_df) * 35 + 50),
        )

    st.caption(f"共 {len(filtered)} 条记录 | 黄色高亮 = 30天内")


def render_sidebar_widget():
    """侧边栏「本周 Earnings」小组件。"""
    df = get_earnings()
    if df.empty:
        return

    today = dt.date.today()
    df["earnings_date_dt"] = pd.to_datetime(df["earnings_date"]).dt.date
    df["days_from_today"] = df["earnings_date_dt"].apply(
        lambda d: (d - today).days if pd.notna(d) else None
    )

    this_week = get_this_week_earnings(df)
    if this_week.empty:
        st.sidebar.markdown("**本周 Earnings**: 无")
        return

    st.sidebar.markdown("**本周 Earnings**")
    for _, row in this_week.iterrows():
        date_str = str(row["earnings_date"])[:10]
        st.sidebar.markdown(
            f"- **{row['ticker']}** {row.get('name','')} ({date_str})"
        )
