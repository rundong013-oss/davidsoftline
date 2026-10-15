#!/usr/bin/env python3
"""
fetch_prices.py — 拉取覆盖公司的最新股价数据（FMP 优先，yfinance fallback）。

读取 config/universe.yaml，输出 JSON 格式的涨跌幅、量比、相对基准偏离等指标。
"""

import json
import sys
from pathlib import Path
from typing import Optional

import yaml

# FMP client
sys.path.insert(0, str(Path(__file__).resolve().parent))
from fmp_client import FMPClient, FMPError


def load_universe(config_path: Path) -> dict:
    """读取 universe.yaml 配置。"""
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def fetch_stock_data_fmp(client: FMPClient, ticker: str) -> Optional[dict]:
    """通过 FMP quote() 拉取股票数据。"""
    try:
        q = client.quote(ticker)
        if not q or not q.get("price"):
            return None

        price = float(q["price"])
        prev_close = float(q.get("previousClose", 0))
        # API 有时返回 0，自行计算
        change_pct = float(q.get("changesPercentage", 0))
        if change_pct == 0 and prev_close > 0:
            change_pct = round((price - prev_close) / prev_close * 100, 2)
        volume = int(q.get("volume", 0))
        avg_volume = int(q.get("avgVolume", 0)) if q.get("avgVolume") else volume
        volume_ratio = round(volume / avg_volume, 2) if avg_volume > 0 else 0

        return {
            "ticker": ticker,
            "close": round(price, 2),
            "prev_close": round(prev_close, 2),
            "change_pct": round(change_pct, 2),
            "volume": volume,
            "avg_volume": avg_volume,
            "volume_ratio": volume_ratio,
            "date": q.get("timestamp", ""),
        }
    except FMPError:
        return None


def fetch_stock_data_yf(ticker: str, period: str = "5d") -> Optional[dict]:
    """yfinance fallback。"""
    try:
        import yfinance as yf
        stock = yf.Ticker(ticker)
        hist = stock.history(period=period)

        if hist.empty or len(hist) < 2:
            return None

        latest = hist.iloc[-1]
        prev = hist.iloc[-2]

        close = float(latest["Close"])
        prev_close = float(prev["Close"])
        change_pct = (close - prev_close) / prev_close * 100

        volume = int(latest["Volume"])
        avg_volume = int(hist["Volume"].iloc[:-1].mean()) if len(hist) > 1 else volume
        volume_ratio = volume / avg_volume if avg_volume > 0 else 0

        return {
            "ticker": ticker,
            "close": round(close, 2),
            "prev_close": round(prev_close, 2),
            "change_pct": round(change_pct, 2),
            "volume": volume,
            "avg_volume": avg_volume,
            "volume_ratio": round(volume_ratio, 2),
            "date": str(hist.index[-1].date()),
        }
    except Exception as e:
        print(f"[WARN] yfinance fallback failed for {ticker}: {e}", file=sys.stderr)
        return None


def fetch_stock_data(client: FMPClient, ticker: str) -> Optional[dict]:
    """FMP 优先，失败回退 yfinance。"""
    data = fetch_stock_data_fmp(client, ticker)
    if data:
        return data
    print(f"[INFO] FMP failed for {ticker}, trying yfinance...", file=sys.stderr)
    return fetch_stock_data_yf(ticker)


def compute_relative_performance(stock_data: dict, benchmark_data: dict) -> Optional[float]:
    """计算相对基准的超额收益。"""
    if stock_data and benchmark_data:
        return round(stock_data["change_pct"] - benchmark_data["change_pct"], 2)
    return None


def main():
    workspace = Path(__file__).resolve().parent.parent
    config_path = workspace / "config" / "universe.yaml"

    if not config_path.exists():
        print(f"[ERROR] Config not found: {config_path}", file=sys.stderr)
        sys.exit(1)

    universe = load_universe(config_path)
    benchmark_ticker = universe.get("benchmark", "SPY")

    # 合并 sportswear + luxury 两个列表
    companies = universe.get("sportswear", []) + universe.get("luxury", [])
    # 兼容旧格式
    if not companies:
        companies = universe.get("companies", [])

    if not companies:
        print("[WARN] No companies in universe.yaml", file=sys.stderr)

    client = FMPClient()

    # 拉取基准指数数据
    benchmark_data = fetch_stock_data(client, benchmark_ticker)

    results = []
    for company in companies:
        # 优先使用 fmp_ticker 作为 API 查询 ticker
        fmp_ticker = company.get("fmp_ticker", company["ticker"])
        display_ticker = company["ticker"]
        name = company.get("name", display_ticker)

        stock_data = fetch_stock_data(client, fmp_ticker)
        if stock_data is None:
            continue

        stock_data["ticker"] = display_ticker
        stock_data["name"] = name
        stock_data["market"] = company.get("market", "US")
        stock_data["vs_benchmark"] = compute_relative_performance(stock_data, benchmark_data)

        results.append(stock_data)

    output = {
        "date": benchmark_data["date"] if benchmark_data else "N/A",
        "benchmark": {
            "ticker": benchmark_ticker,
            **(benchmark_data or {}),
        },
        "stocks": results,
    }

    print(json.dumps(output, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
