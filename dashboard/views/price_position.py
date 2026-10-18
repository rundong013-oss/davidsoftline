"""
price_position.py — 股价位置：当前价在历史价格分布中的分位数。
"""
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from config import ALL_NAMES, SPORTSWEAR_TICKERS, LUXURY_TICKERS
from data.database import get_weekly_prices


def _compute_percentiles():
    """用 get_weekly_prices 计算每个 ticker 的历史分位数。"""
    df = get_weekly_prices()
    if df.empty:
        return pd.DataFrame()

    df = df.dropna(subset=["close"]).copy()
    if df.empty:
        return pd.DataFrame()

    rows = []
    for ticker, grp in df.groupby("ticker"):
        closes = grp["close"]
        low = float(closes.min())
        high = float(closes.max())
        current = float(grp.sort_values("date")["close"].iloc[-1])
        # 区间分位数：(current - low) / (high - low)
        if high > low:
            range_pct = (current - low) / (high - low)
            range_pct = max(0.0, min(1.0, range_pct))
        else:
            range_pct = 0.5
        # 排名分位数：count(closes <= current) / n
        n = len(closes)
        if n > 1:
            rank_pct = float((closes <= current).sum()) / n
        else:
            rank_pct = 0.5
        rows.append({
            "ticker": ticker,
            "sector": grp["sector"].iloc[0],
            "weeks": len(grp),
            "low": low,
            "high": high,
            "current_price": current,
            "range_percentile": round(range_pct * 100),
            "rank_percentile": round(rank_pct * 100),
        })
    return pd.DataFrame(rows)


def _build_table(df, tickers, names):
    """筛选指定 tickers，拼接公司名，按排名分位数升序排列。"""
    sub = df[df["ticker"].isin(tickers)].copy()
    if sub.empty:
        return sub
    sub["公司"] = sub["ticker"].map(names).fillna("")
    sub = sub.rename(columns={
        "ticker": "Ticker",
        "current_price": "当前价",
        "low": "历史最低",
        "high": "历史最高",
        "range_percentile": "区间分位数",
        "rank_percentile": "排名分位数",
        "weeks": "数据周数",
    })
    sub = sub[["Ticker", "公司", "当前价", "历史最低", "历史最高", "区间分位数", "排名分位数", "数据周数"]]
    sub = sub.sort_values("排名分位数", ascending=True).reset_index(drop=True)
    return sub


def _make_chart(table: pd.DataFrame, title: str):
    """横向条形图：排名分位数 vs 区间分位数，按排名分位数排序。"""
    if table.empty:
        return None
    # 按排名分位数降序排列（图表从下往上读，降序让最低的在最上面）
    t = table.sort_values("排名分位数", ascending=True)
    labels = t["Ticker"] + " " + t["公司"]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=labels, x=t["排名分位数"], name="排名分位数",
        orientation="h", marker_color="#5470c6",
        text=t["排名分位数"].apply(lambda v: f"{v}%"), textposition="outside",
    ))
    fig.add_trace(go.Bar(
        y=labels, x=t["区间分位数"], name="区间分位数",
        orientation="h", marker_color="#91cc75", opacity=0.7,
        text=t["区间分位数"].apply(lambda v: f"{v}%"), textposition="outside",
    ))
    fig.update_layout(
        title=title,
        xaxis=dict(title="分位数 (%)", range=[0, 110]),
        barmode="group",
        height=max(300, len(t) * 40 + 100),
        margin=dict(l=10, r=10, t=40, b=40),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    return fig


def render():
    st.header("股价位置")
    st.caption(
        "排名分位数 = 历史收盘价 ≤ 当前价的比例；"
        "区间分位数 = (当前价 − 最低) / (最高 − 最低)"
    )

    df = _compute_percentiles()
    if df.empty:
        st.warning("无周线数据。请先导入 Excel 或运行 API 更新。")
        return

    col_config = {
        "当前价": st.column_config.NumberColumn(format="%.2f"),
        "历史最低": st.column_config.NumberColumn(format="%.2f"),
        "历史最高": st.column_config.NumberColumn(format="%.2f"),
        "排名分位数": st.column_config.ProgressColumn(
            min_value=0, max_value=100, format="%d%%",
        ),
        "区间分位数": st.column_config.ProgressColumn(
            min_value=0, max_value=100, format="%d%%",
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
        fig = _make_chart(sport, "运动服饰 — 股价分位数")
        if fig:
            st.plotly_chart(fig, use_container_width=True)

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
        fig = _make_chart(luxury, "奢侈品 — 股价分位数")
        if fig:
            st.plotly_chart(fig, use_container_width=True)
