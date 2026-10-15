"""
app.py — Streamlit Dashboard 入口。

用法:
    cd 运动服饰研究系统/dashboard
    streamlit run app.py
"""
import sys
from pathlib import Path

# 确保 dashboard/ 和 scripts/ 在 path 中
_dashboard_dir = Path(__file__).resolve().parent
_project_root = _dashboard_dir.parent
sys.path.insert(0, str(_dashboard_dir))
sys.path.insert(0, str(_project_root / "scripts"))

import streamlit as st

from data.database import init_db, get_meta
from views import weekly_changes, index_etf, macro, valuation, earnings_calendar
from views import thesis_bank

# ── 页面配置 ──────────────────────────────────────────────────────
st.set_page_config(
    page_title="运动服饰研究",
    page_icon="",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── 初始化数据库 ──────────────────────────────────────────────────
init_db()

# ── Sidebar ──────────────────────────────────────────────────────
st.sidebar.title("研究工作台")
st.sidebar.caption("运动服饰 & 奢侈品行业研究")

# 顶层模块切换
module = st.sidebar.radio("模块", ["Monitor", "Thesis Bank"], index=0, horizontal=True)

st.sidebar.divider()

if module == "Monitor":
    # ── Monitor 子导航 ────────────────────────────────────────────
    page = st.sidebar.radio(
        "导航",
        [
            "周涨跌幅",
            "估值筛选",
            "指数和ETF",
            "宏观数据",
            "Earnings 日历",
        ],
        index=0,
    )

    st.sidebar.divider()

    # ── 数据更新 ──────────────────────────────────────────────────
    if st.sidebar.button("Refresh Data", use_container_width=True):
        with st.spinner("正在更新数据..."):
            try:
                from data.updater import update_all
                update_all()
                st.sidebar.success("数据更新完成!")
                st.rerun()
            except Exception as e:
                st.sidebar.error(f"更新失败: {e}")

    # 上次更新时间
    last_update = get_meta("last_full_update") or get_meta("last_import")
    if last_update:
        st.sidebar.caption(f"上次更新: {last_update[:16]}")
    else:
        st.sidebar.caption("尚未导入数据")

        # 提供 Excel 导入按钮
        if st.sidebar.button("Import from Excel", use_container_width=True):
            with st.spinner("正在导入 Excel 数据..."):
                try:
                    from data.import_excel import main as import_main
                    import_main()
                    st.sidebar.success("Excel 导入完成!")
                    st.rerun()
                except Exception as e:
                    st.sidebar.error(f"导入失败: {e}")

    st.sidebar.divider()

    # ── 本周 Earnings 小组件 ─────────────────────────────────────
    earnings_calendar.render_sidebar_widget()

    # ── 页面路由 ──────────────────────────────────────────────────
    if page == "周涨跌幅":
        weekly_changes.render()
    elif page == "指数和ETF":
        index_etf.render()
    elif page == "宏观数据":
        macro.render()
    elif page == "估值筛选":
        valuation.render()
    elif page == "Earnings 日历":
        earnings_calendar.render()

elif module == "Thesis Bank":
    # ── Thesis Bank ───────────────────────────────────────────────
    thesis_bank.render()
