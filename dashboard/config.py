"""
config.py — 读取 universe.yaml，暴露 ticker 列表和常量。
"""
from pathlib import Path

import yaml

# ── 路径 ──────────────────────────────────────────────────────────
PROJECT_ROOT = Path(__file__).resolve().parent.parent  # 运动服饰研究系统/
DASHBOARD_DIR = Path(__file__).resolve().parent
DB_PATH = DASHBOARD_DIR / "data" / "monitor.db"
UNIVERSE_PATH = PROJECT_ROOT / "config" / "universe.yaml"
SCRIPTS_DIR = PROJECT_ROOT / "scripts"
LIBRARY_DIR = PROJECT_ROOT / "library"

# Excel 文件路径
EXCEL_APPAREL = PROJECT_ROOT / "output" / "monitor" / "Apparel_updating file.xlsx"
EXCEL_VALUATION = PROJECT_ROOT / "output" / "monitor" / "Valuation screening.xlsx"
EXCEL_EARNINGS = PROJECT_ROOT / "output" / "monitor" / "Earnings calendar.xlsx"

# ── 基期常量（与 update_monitor.py 一致）────────────────────────
BASE_DATE = "2025-12-01"
ETF_BASE_DATE = "2025-01-01"
MACRO_BASE = "2025-01-01"

# ── 读 universe.yaml ─────────────────────────────────────────────
def _load_universe():
    with open(UNIVERSE_PATH, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

_universe = _load_universe()

SPORTSWEAR_UNIVERSE = _universe.get("sportswear", [])
LUXURY_UNIVERSE = _universe.get("luxury", [])
BENCHMARK = _universe.get("benchmark", "SPY")

# 构建 ticker→name 映射
def _build_ticker_map(entries):
    """返回 {ticker: name, ...}"""
    return {e["ticker"]: e["name"] for e in entries}

SPORTSWEAR_NAMES = _build_ticker_map(SPORTSWEAR_UNIVERSE)
LUXURY_NAMES = _build_ticker_map(LUXURY_UNIVERSE)
ALL_NAMES = {**SPORTSWEAR_NAMES, **LUXURY_NAMES}

# ticker 列表
SPORTSWEAR_TICKERS = list(SPORTSWEAR_NAMES.keys())
LUXURY_TICKERS = list(LUXURY_NAMES.keys())
ALL_TICKERS = SPORTSWEAR_TICKERS + LUXURY_TICKERS

# fmp_ticker 映射
def _build_fmp_map(entries):
    return {e["ticker"]: e.get("fmp_ticker", e["ticker"]) for e in entries}

SPORTSWEAR_FMP = _build_fmp_map(SPORTSWEAR_UNIVERSE)
LUXURY_FMP = _build_fmp_map(LUXURY_UNIVERSE)
ALL_FMP = {**SPORTSWEAR_FMP, **LUXURY_FMP}

# ── ETF 列表（与 update_monitor.py 一致）─────────────────────────
ETFS = [
    ("XLY",  "可选消费 ETF 市值加权"),
    ("RSPD", "标普500可选消费等权 ETF"),
    ("XRT",  "标普零售精选 ETF 等权"),
]
ETF_TICKERS = [t[0] for t in ETFS]

# ── 行业分类 ─────────────────────────────────────────────────────
SECTORS = {
    "sportswear": "运动服饰",
    "luxury": "奢侈品",
}
