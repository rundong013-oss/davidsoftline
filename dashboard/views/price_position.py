"""
price_position.py — 股价位置：当前价在历史区间的分位数。
"""
import streamlit as st
import pandas as pd


def _compute_percentiles(df):
    """从 weekly_prices DataFrame 计算每个 ticker 的分位数。"""
    df = df.dropna(subset=["close"])
    if df.empty:
        return pd.DataFrame()

    # 每个 ticker 的统计量
    agg = df.groupby("ticker").agg(
        sector=("sector", "first"),
        weeks=("close", "count"),
        low=("close", "min"),
        high=("close", "max"),
    ).reset_index()

    # 最新一周收盘价：取每个 ticker 最大 date 对应的 close
    latest = df.sort_values("date").groupby("ticker").tail(1)[["ticker", "close"]]
    latest = latest.rename(columns={"close": "current_price"})

    result = agg.merge(latest, on="ticker", how="left")
    span = result["high"] - result["low"]
    result["percentile"] = ((result["current_price"] - result["low"]) / span).clip(0, 1)
    # span == 0 时设为 0.5
    result.loc[span == 0, "percentile"] = 0.5
    return result


def _build_table(df, tickers, names):
    """筛选指定 tickers，拼接公司名，按分位数升序排列。"""
    sub = df[df["ticker"].isin(tickers)].copy()
    if sub.empty:
        return sub
    sub["公司"] = sub["ticker"].map(names).fillna("")
    sub = sub.rename(columns={
        "ticker": "Ticker",
        "current_price": "当前价",
        "low": "历史最低",
        "high": "历史最高",
        "percentile": "分位数",
        "weeks": "数据周数",
    })
    sub = sub[["Ticker", "公司", "当前价", "历史最低", "历史最高", "分位数", "数据周数"]]
    sub = sub.sort_values("分位数", ascending=True).reset_index(drop=True)
    return sub


def render():
    from config import ALL_NAMES, SPORTSWEAR_TICKERS, LUXURY_TICKERS
    from data.database import get_weekly_prices

    st.header("股价位置")
    st.caption("当前价在历史周线数据中的分位数（0% = 历史最低，100% = 历史最高）")

    raw = get_weekly_prices()
    df = _compute_percentiles(raw)
    if df.empty:
        st.warning("无周线数据。请先导入 Excel 或运行 API 更新。")
        return

    col_config = {
        "当前价": st.column_config.NumberColumn(format="%.2f"),
        "历史最低": st.column_config.NumberColumn(format="%.2f"),
        "历史最高": st.column_config.NumberColumn(format="%.2f"),
        "分位数": st.column_config.ProgressColumn(
            min_value=0, max_value=1, format="%.0%%",
        ),
        "数据周数": st.column_config.NumberColumn(format="%d"),
    }

    # ── 运动服饰 ──────────────────────────────────────────────────
    st.subheader("运动服饰")
    sport = _build_table(df, SPORTSWEAR_TICKERS, ALL_NAMES)
    if sport.empty:
        st.info("暂无运动服饰数据")
    else:
        st.dataframe(
            sport,
            column_config=col_config,
            use_container_width=True,
            hide_index=True,
            height=min(800, len(sport) * 35 + 50),
        )

    # ── 奢侈品 ────────────────────────────────────────────────────
    st.subheader("奢侈品")
    luxury = _build_table(df, LUXURY_TICKERS, ALL_NAMES)
    if luxury.empty:
        st.info("暂无奢侈品数据")
    else:
        st.dataframe(
            luxury,
            column_config=col_config,
            use_container_width=True,
            hide_index=True,
            height=min(800, len(luxury) * 35 + 50),
        )
