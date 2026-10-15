#!/usr/bin/env python3
"""
fmp_client.py — 共享 FMP API 客户端，供所有脚本复用。

Base URL: https://fmp-v2.icu/stable
API Key:  硬编码默认值，可通过环境变量 FMP_API_KEY 覆盖。
"""
import os
import time
from typing import Optional

import pandas as pd
import requests

_DEFAULT_BASE = "https://fmp-v2.icu/stable"
_DEFAULT_KEY = "8qnDx2XDkRQEa2MX"


class FMPError(Exception):
    pass


class FMPClient:
    """通用 FMP API 客户端，内置 rate limit + 429 重试。"""

    def __init__(self, base_url: str = None, api_key: str = None,
                 sleep: float = 0.2, timeout: int = 30):
        self.base_url = base_url or os.environ.get("FMP_BASE_URL", _DEFAULT_BASE)
        self.api_key = api_key or os.environ.get("FMP_API_KEY", _DEFAULT_KEY)
        self.sleep = sleep
        self.timeout = timeout
        self.calls = 0

    # ── 通用请求 ──────────────────────────────────────────────────
    def get(self, endpoint: str, **params):
        """发起 GET 请求，返回 JSON（list 或 dict）。"""
        params["apikey"] = self.api_key
        url = f"{self.base_url}/{endpoint}"
        last_exc = None
        for attempt in range(4):
            try:
                r = requests.get(url, params=params, timeout=self.timeout)
            except requests.RequestException as e:
                last_exc = e
                time.sleep(2 * (attempt + 1))
                continue
            self.calls += 1
            if r.status_code == 429:
                time.sleep(5 * (attempt + 1))
                continue
            if r.status_code in (401, 402, 403):
                raise FMPError(f"{endpoint} HTTP {r.status_code}: {r.text[:200]}")
            if r.status_code != 200:
                raise FMPError(f"{endpoint} HTTP {r.status_code}: {r.text[:200]}")
            try:
                data = r.json()
            except ValueError:
                raise FMPError(f"{endpoint} non-JSON response: {r.text[:200]}")
            if isinstance(data, dict) and ("Error Message" in data or "error" in data):
                raise FMPError(f"{endpoint} API error: {data}")
            time.sleep(self.sleep)
            return data
        if last_exc:
            raise FMPError(f"{endpoint} request failed after retries: {last_exc}")
        raise FMPError(f"{endpoint} exhausted retries (429)")

    # ── 历史日线 ──────────────────────────────────────────────────
    def historical_prices(self, symbol: str, start: str,
                          end: str = None) -> pd.DataFrame:
        """返回 DataFrame，columns=[date, open, high, low, close, volume]。"""
        params = {"symbol": symbol, "from": start}
        if end:
            params["to"] = end
        data = self.get("historical-price-eod/light", **params)
        if not data:
            return pd.DataFrame()
        df = pd.DataFrame(data)
        if "date" not in df.columns:
            return pd.DataFrame()
        df["date"] = pd.to_datetime(df["date"])
        df = df.sort_values("date").set_index("date")
        return df

    # ── 实时报价 ──────────────────────────────────────────────────
    def quote(self, symbol: str) -> dict:
        """返回单只股票的实时报价 dict。"""
        data = self.get("quote", symbol=symbol)
        if isinstance(data, list) and data:
            return data[0]
        if isinstance(data, dict):
            return data
        return {}

    # ── 财务报表 ──────────────────────────────────────────────────
    def income_statement(self, symbol: str, period: str = "annual",
                         limit: int = 5) -> list:
        return self.get("income-statement", symbol=symbol,
                        period=period, limit=limit)

    def balance_sheet(self, symbol: str, period: str = "annual",
                      limit: int = 5) -> list:
        return self.get("balance-sheet-statement", symbol=symbol,
                        period=period, limit=limit)

    # ── 公司概况 ──────────────────────────────────────────────────
    def profile(self, symbol: str) -> dict:
        data = self.get("profile", symbol=symbol)
        if isinstance(data, list) and data:
            return data[0]
        if isinstance(data, dict):
            return data
        return {}

    # ── 一致预期 ──────────────────────────────────────────────────
    def analyst_estimates(self, symbol: str, period: str = "annual",
                          limit: int = 10) -> list:
        return self.get("analyst-estimates", symbol=symbol,
                        period=period, limit=limit)

    # ── FX 汇率 ──────────────────────────────────────────────────
    def fx_rate(self, from_ccy: str, to_ccy: str) -> Optional[float]:
        """返回 1 from_ccy = ? to_ccy。"""
        if not from_ccy or not to_ccy or from_ccy == to_ccy:
            return 1.0
        for sym, invert in ((f"{from_ccy}{to_ccy}", False),
                            (f"{to_ccy}{from_ccy}", True)):
            try:
                q = self.get("quote", symbol=sym)
                if isinstance(q, list) and q:
                    q = q[0]
                p = q.get("price") if isinstance(q, dict) else None
                if p and float(p) > 0:
                    return 1.0 / float(p) if invert else float(p)
            except FMPError:
                continue
        return None
