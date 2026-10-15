"""
index_etf.py — 行业等权指数 + ETF 归一化走势。
"""
import pandas as pd
import streamlit as st

from data.database import get_sector_index, get_etf_weekly
from views.components import line_chart


def render():
    st.header("指数和ETF")

    # ── 可选消费 ETF ──────────────────────────────────────────────
    st.subheader("可选消费 ETF")
    etf_df = get_etf_weekly()
    if not etf_df.empty:
        etf_df["date"] = pd.to_datetime(etf_df["date"])

        # 价格走势
        price_pivot = etf_df.pivot_table(
            index="date", columns="ticker", values="close"
        )
        if not price_pivot.empty:
            # 归一化
            norm_pivot = etf_df.pivot_table(
                index="date", columns="ticker", values="normalized"
            )
            if not norm_pivot.empty:
                fig = line_chart(
                    norm_pivot,
                    title="ETF 归一化走势",
                    y_label="归一化",
                    height=400,
                    normalized=True,
                )
                st.plotly_chart(fig, use_container_width=True)

            # 最新价格和涨跌
            latest = price_pivot.iloc[-1]
            cols = st.columns(len(latest))
            for i, (ticker, price) in enumerate(latest.items()):
                with cols[i]:
                    prev = price_pivot.iloc[-2][ticker] if len(price_pivot) > 1 else price
                    delta = f"{(price / prev - 1):+.2%}" if prev else None
                    st.metric(ticker, f"${price:.2f}", delta=delta)

            # ETF 说明
            etf_desc = {
                "XLY": "可选消费 ETF（市值加权，被亚马逊/特斯拉主导）",
                "RSPD": "标普500可选消费等权 ETF（去掉巨头集中度）",
                "XRT": "标普零售精选 ETF（等权，零售属性强）",
            }
            with st.expander("ETF 说明"):
                for ticker, desc in etf_desc.items():
                    st.write(f"**{ticker}**: {desc}")
    else:
        st.info("无 ETF 数据")

    st.divider()

    # ── 行业等权指数 ──────────────────────────────────────────────
    st.subheader("行业等权指数")
    idx_df = get_sector_index()
    if not idx_df.empty:
        idx_df["date"] = pd.to_datetime(idx_df["date"])
        pivot = idx_df.pivot_table(
            index="date", columns="sector", values="index_value"
        )
        # 重命名列
        rename_map = {"sportswear": "运动服饰", "luxury": "奢侈品"}
        pivot = pivot.rename(columns=rename_map)

        fig = line_chart(
            pivot,
            title="行业等权指数（基期 = 1）",
            y_label="归一化",
            height=450,
            normalized=True,
        )
        st.plotly_chart(fig, use_container_width=True)

        # 最新值
        if not pivot.empty:
            latest = pivot.iloc[-1]
            cols = st.columns(len(latest))
            for i, (col_name, val) in enumerate(latest.items()):
                with cols[i]:
                    prev = pivot.iloc[-2][col_name] if len(pivot) > 1 else val
                    delta = f"{(val / prev - 1):+.2%}" if prev else None
                    st.metric(col_name, f"{val:.4f}", delta=delta)
    else:
        st.info("无行业指数数据")
