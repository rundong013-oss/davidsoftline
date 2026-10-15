"""
components.py — 共享可视化组件。
"""
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


def heatmap_chart(pivot_df, title="", height=600, color_scale=None, zmid=0,
                  fmt=".1%", text_auto=True):
    """绘制周涨跌幅 heatmap。

    pivot_df: index=tickers, columns=week_dates, values=returns
    """
    if color_scale is None:
        color_scale = [[0, "#c0392b"], [0.5, "#ffffff"], [1, "#27ae60"]]

    fig = go.Figure(data=go.Heatmap(
        z=pivot_df.values,
        x=[str(c)[:10] for c in pivot_df.columns],
        y=pivot_df.index.tolist(),
        colorscale=color_scale,
        zmid=zmid,
        text=[[f"{v:.1%}" if v == v else "" for v in row]
              for row in pivot_df.values],
        texttemplate="%{text}",
        textfont={"size": 10},
        hovertemplate="Ticker: %{y}<br>Week: %{x}<br>Return: %{z:.2%}<extra></extra>",
        colorbar=dict(tickformat=".0%", title="Return"),
    ))
    fig.update_layout(
        title=title,
        height=height,
        yaxis=dict(autorange="reversed"),
        xaxis=dict(side="top"),
        margin=dict(l=100, r=20, t=60, b=20),
    )
    return fig


def line_chart(df, title="", y_label="", height=400, normalized=False):
    """绘制多序列折线图。

    df: index=dates, columns=series names
    """
    fig = px.line(df, title=title, height=height)
    fig.update_layout(
        yaxis_title=y_label,
        xaxis_title="",
        legend=dict(orientation="h", yanchor="bottom", y=-0.3, x=0),
        margin=dict(l=60, r=20, t=50, b=20),
        hovermode="x unified",
    )
    if normalized:
        fig.update_layout(yaxis_tickformat=".3f")
    return fig


def metric_card(label, value, delta=None, delta_color="normal"):
    """Streamlit metric card wrapper."""
    st.metric(label=label, value=value, delta=delta, delta_color=delta_color)


def styled_dataframe(df, height=400):
    """带格式的 dataframe。"""
    st.dataframe(df, height=height, use_container_width=True)
