"""
macro.py — 宏观数据页面。
"""
import pandas as pd
import streamlit as st

from data.database import get_macro_data
from views.components import line_chart, styled_dataframe


def render():
    st.header("宏观数据")

    tab_oil, tab_yield, tab_weekly, tab_monthly = st.tabs([
        "油价", "美债收益率", "周度（初请失业金）", "月度指标"
    ])

    # ── 油价 ──────────────────────────────────────────────────────
    with tab_oil:
        oil_series = ["Brent($/bbl)", "WTI($/bbl)", "全美汽油均价($/gal)"]
        df = get_macro_data(series_list=oil_series)
        if not df.empty:
            df["date"] = pd.to_datetime(df["date"])
            pivot = df.pivot_table(index="date", columns="series", values="value")
            # 保持列顺序
            cols = [c for c in oil_series if c in pivot.columns]
            pivot = pivot[cols]

            fig = line_chart(
                pivot, title="油价 & 汽油价格", y_label="$/bbl 或 $/gal",
                height=450,
            )
            st.plotly_chart(fig, use_container_width=True)

            # 最新值
            if not pivot.empty:
                latest = pivot.iloc[-1].dropna()
                cols_ui = st.columns(len(latest))
                for i, (series, val) in enumerate(latest.items()):
                    with cols_ui[i]:
                        prev = pivot.iloc[-2][series] if (
                            len(pivot) > 1 and series in pivot.columns
                        ) else val
                        chg = (val - prev) / prev * 100 if prev else 0
                        st.metric(series, f"{val:.2f}", delta=f"{chg:+.1f}%")
        else:
            st.info("无油价数据")

    # ── 美债收益率 ────────────────────────────────────────────────
    with tab_yield:
        yield_series = ["2Y(%)", "10Y(%)", "30Y(%)"]
        df = get_macro_data(series_list=yield_series)
        if not df.empty:
            df["date"] = pd.to_datetime(df["date"])
            pivot = df.pivot_table(index="date", columns="series", values="value")
            cols = [c for c in yield_series if c in pivot.columns]
            pivot = pivot[cols]

            fig = line_chart(
                pivot, title="US Treasury Yields", y_label="%",
                height=450,
            )
            st.plotly_chart(fig, use_container_width=True)

            if not pivot.empty:
                latest = pivot.iloc[-1].dropna()
                cols_ui = st.columns(len(latest))
                for i, (series, val) in enumerate(latest.items()):
                    with cols_ui[i]:
                        prev = pivot.iloc[-2][series] if (
                            len(pivot) > 1 and series in pivot.columns
                        ) else val
                        bps = round((val - prev) * 100) if prev else 0
                        st.metric(series, f"{val:.2f}%",
                                  delta=f"{bps:+d} bps")
        else:
            st.info("无收益率数据")

    # ── 周度数据 ──────────────────────────────────────────────────
    with tab_weekly:
        claims_series = ["初请失业金(千人)"]
        df = get_macro_data(series_list=claims_series)
        if not df.empty:
            df["date"] = pd.to_datetime(df["date"])
            pivot = df.pivot_table(index="date", columns="series", values="value")

            fig = line_chart(
                pivot, title="初请失业金", y_label="千人",
                height=400,
            )
            st.plotly_chart(fig, use_container_width=True)

            if not pivot.empty and "初请失业金(千人)" in pivot.columns:
                latest = pivot["初请失业金(千人)"].dropna().iloc[-1]
                prev = pivot["初请失业金(千人)"].dropna().iloc[-2] if len(pivot.dropna()) > 1 else latest
                st.metric("最新初请失业金",
                          f"{latest:,.0f} 千人",
                          delta=f"{latest - prev:+,.0f}")
        else:
            st.info("无初请失业金数据")

    # ── 月度指标 ──────────────────────────────────────────────────
    with tab_monthly:
        monthly_series = [
            "CPI YoY%", "核心CPI YoY%", "PPI YoY%",
            "PCE YoY%", "核心PCE YoY%",
            "非农delta(千人)", "失业率(%)", "零售销售MoM%",
            "消费者信心", "信用卡余额YoY%",
        ]
        df = get_macro_data(series_list=monthly_series, frequency="monthly")
        if not df.empty:
            df["date"] = pd.to_datetime(df["date"])
            pivot = df.pivot_table(index="date", columns="series", values="value")

            # 只保留已有的列，按定义顺序排列
            cols = [c for c in monthly_series if c in pivot.columns]
            pivot = pivot[cols]

            # 显示最新值摘要
            if not pivot.empty:
                st.subheader("最新月度指标")
                latest_row = pivot.dropna(how="all").iloc[-1]
                prev_row = pivot.dropna(how="all").iloc[-2] if len(pivot.dropna(how="all")) > 1 else latest_row

                # 分两行展示
                n_cols = min(5, len(cols))
                for chunk_start in range(0, len(cols), n_cols):
                    chunk = cols[chunk_start:chunk_start + n_cols]
                    ui_cols = st.columns(len(chunk))
                    for i, series in enumerate(chunk):
                        with ui_cols[i]:
                            val = latest_row.get(series)
                            prev_val = prev_row.get(series)
                            if pd.notna(val):
                                delta = None
                                if pd.notna(prev_val):
                                    delta = f"{val - prev_val:+.1f}"
                                st.metric(series, f"{val:.1f}", delta=delta)

                st.divider()

            # 完整表格
            st.subheader("历史月度数据")
            display_df = pivot.copy()
            display_df.index = display_df.index.strftime("%Y-%m")
            # 最新的在上面
            display_df = display_df.iloc[::-1]
            styled_dataframe(display_df, height=500)
        else:
            st.info("无月度宏观数据")

        st.caption(
            "说明：CPI/PPI/PCE/信用卡余额为同比%；零售销售为环比%；"
            "非农为环比变化千人。来源 FRED。"
        )
