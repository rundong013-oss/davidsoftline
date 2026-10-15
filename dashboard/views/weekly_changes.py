"""
weekly_changes.py — 周涨跌幅个股 heatmap。
"""
import pandas as pd
import streamlit as st

from data.database import get_weekly_prices
from views.components import heatmap_chart


def render():
    st.header("周涨跌幅")

    # 周数滑块
    n_weeks = st.slider("显示周数", min_value=4, max_value=52, value=12, step=4)

    for sector, label in [("sportswear", "运动服饰"), ("luxury", "奢侈品")]:
        df = get_weekly_prices(sector=sector)
        if df.empty:
            st.warning(f"{label}: 无数据")
            continue

        df["date"] = pd.to_datetime(df["date"])

        # pivot: tickers × dates → weekly_return
        pivot = df.pivot_table(
            index="ticker", columns="date", values="weekly_return"
        )

        if pivot.empty:
            st.warning(f"{label}: 无涨跌数据")
            continue

        # 只取最近 n_weeks 列
        pivot = pivot.iloc[:, -n_weeks:] if pivot.shape[1] > n_weeks else pivot

        # 排序：按最新一周涨跌排序
        if pivot.shape[1] > 0:
            last_col = pivot.columns[-1]
            pivot = pivot.sort_values(last_col, ascending=False, na_position="last")

        fig = heatmap_chart(
            pivot,
            title=f"{label} 周涨跌幅",
            height=max(300, len(pivot) * 22 + 100),
        )
        st.plotly_chart(fig, use_container_width=True)

        # 最新一周概览
        if pivot.shape[1] > 0:
            latest = pivot.iloc[:, -1].dropna().sort_values(ascending=False)
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"**{label} 本周涨幅前5**")
                top5 = latest.head(5)
                for ticker, ret in top5.items():
                    st.write(f"  {ticker}: {ret:+.1%}")
            with col2:
                st.markdown(f"**{label} 本周跌幅前5**")
                bot5 = latest.tail(5)
                for ticker, ret in bot5.items():
                    st.write(f"  {ticker}: {ret:+.1%}")
