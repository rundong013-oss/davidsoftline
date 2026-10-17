"""
database.py — SQLite 读写封装（WAL 模式，单文件）。
"""
import json
import sqlite3
from contextlib import contextmanager
from datetime import datetime

import pandas as pd

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from config import DB_PATH


# ── 连接管理 ─────────────────────────────────────────────────────
@contextmanager
def get_conn():
    conn = sqlite3.connect(str(DB_PATH), timeout=15)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db():
    """建表（幂等）。"""
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    with get_conn() as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS weekly_prices (
            date       TEXT NOT NULL,
            ticker     TEXT NOT NULL,
            sector     TEXT,
            close      REAL,
            normalized REAL,
            weekly_return REAL,
            PRIMARY KEY (date, ticker)
        );

        CREATE TABLE IF NOT EXISTS etf_weekly (
            date       TEXT NOT NULL,
            ticker     TEXT NOT NULL,
            close      REAL,
            normalized REAL,
            PRIMARY KEY (date, ticker)
        );

        CREATE TABLE IF NOT EXISTS sector_index (
            date        TEXT NOT NULL,
            sector      TEXT NOT NULL,
            index_value REAL,
            PRIMARY KEY (date, sector)
        );

        CREATE TABLE IF NOT EXISTS macro_data (
            date      TEXT NOT NULL,
            series    TEXT NOT NULL,
            value     REAL,
            frequency TEXT,
            PRIMARY KEY (date, series)
        );

        CREATE TABLE IF NOT EXISTS valuation (
            ticker         TEXT PRIMARY KEY,
            name           TEXT,
            sector         TEXT,
            currency       TEXT,
            price          REAL,
            market_cap     REAL,
            shares_out     REAL,
            updated_at     TEXT,
            annual_json    TEXT,
            quarterly_json TEXT,
            fwd_pe_json    TEXT,
            fwd_ps_json    TEXT,
            n_analysts     INTEGER
        );

        CREATE TABLE IF NOT EXISTS earnings_calendar (
            ticker         TEXT NOT NULL,
            name           TEXT,
            sector         TEXT,
            earnings_date  TEXT,
            days_from_today INTEGER,
            eps_estimate   TEXT,
            fiscal_quarter TEXT,
            updated_at     TEXT,
            PRIMARY KEY (ticker, earnings_date)
        );

        CREATE TABLE IF NOT EXISTS meta (
            key   TEXT PRIMARY KEY,
            value TEXT
        );
        """)


# ── Meta ─────────────────────────────────────────────────────────
def get_meta(key, default=None):
    with get_conn() as conn:
        row = conn.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
    return row[0] if row else default


def set_meta(key, value):
    with get_conn() as conn:
        conn.execute(
            "INSERT OR REPLACE INTO meta (key, value) VALUES (?, ?)",
            (key, str(value))
        )


# ── Weekly Prices ────────────────────────────────────────────────
def upsert_weekly_prices(df):
    """df: columns=[date, ticker, sector, close, normalized, weekly_return]"""
    if df.empty:
        return
    with get_conn() as conn:
        for _, row in df.iterrows():
            conn.execute(
                """INSERT OR REPLACE INTO weekly_prices
                   (date, ticker, sector, close, normalized, weekly_return)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                (str(row["date"])[:10], row["ticker"], row.get("sector"),
                 row.get("close"), row.get("normalized"), row.get("weekly_return"))
            )


def get_weekly_prices(sector=None, start_date=None, end_date=None):
    """返回 DataFrame。"""
    sql = "SELECT * FROM weekly_prices WHERE 1=1"
    params = []
    if sector:
        sql += " AND sector=?"
        params.append(sector)
    if start_date:
        sql += " AND date>=?"
        params.append(str(start_date)[:10])
    if end_date:
        sql += " AND date<=?"
        params.append(str(end_date)[:10])
    sql += " ORDER BY date, ticker"
    with get_conn() as conn:
        return pd.read_sql_query(sql, conn, params=params)


# ── ETF Weekly ───────────────────────────────────────────────────
def upsert_etf_weekly(df):
    """df: columns=[date, ticker, close, normalized]"""
    if df.empty:
        return
    with get_conn() as conn:
        for _, row in df.iterrows():
            conn.execute(
                """INSERT OR REPLACE INTO etf_weekly
                   (date, ticker, close, normalized)
                   VALUES (?, ?, ?, ?)""",
                (str(row["date"])[:10], row["ticker"],
                 row.get("close"), row.get("normalized"))
            )


def get_etf_weekly(start_date=None, end_date=None):
    sql = "SELECT * FROM etf_weekly WHERE 1=1"
    params = []
    if start_date:
        sql += " AND date>=?"
        params.append(str(start_date)[:10])
    if end_date:
        sql += " AND date<=?"
        params.append(str(end_date)[:10])
    sql += " ORDER BY date, ticker"
    with get_conn() as conn:
        return pd.read_sql_query(sql, conn, params=params)


# ── Sector Index ─────────────────────────────────────────────────
def upsert_sector_index(df):
    """df: columns=[date, sector, index_value]"""
    if df.empty:
        return
    with get_conn() as conn:
        for _, row in df.iterrows():
            conn.execute(
                """INSERT OR REPLACE INTO sector_index
                   (date, sector, index_value)
                   VALUES (?, ?, ?)""",
                (str(row["date"])[:10], row["sector"], row.get("index_value"))
            )


def get_sector_index(sector=None, start_date=None):
    sql = "SELECT * FROM sector_index WHERE 1=1"
    params = []
    if sector:
        sql += " AND sector=?"
        params.append(sector)
    if start_date:
        sql += " AND date>=?"
        params.append(str(start_date)[:10])
    sql += " ORDER BY date, sector"
    with get_conn() as conn:
        return pd.read_sql_query(sql, conn, params=params)


# ── Macro Data ───────────────────────────────────────────────────
def upsert_macro_data(df):
    """df: columns=[date, series, value, frequency]"""
    if df.empty:
        return
    with get_conn() as conn:
        for _, row in df.iterrows():
            conn.execute(
                """INSERT OR REPLACE INTO macro_data
                   (date, series, value, frequency)
                   VALUES (?, ?, ?, ?)""",
                (str(row["date"])[:10], row["series"],
                 row.get("value"), row.get("frequency"))
            )


def get_macro_data(series_list=None, frequency=None, start_date=None):
    sql = "SELECT * FROM macro_data WHERE 1=1"
    params = []
    if series_list:
        placeholders = ",".join("?" * len(series_list))
        sql += f" AND series IN ({placeholders})"
        params.extend(series_list)
    if frequency:
        sql += " AND frequency=?"
        params.append(frequency)
    if start_date:
        sql += " AND date>=?"
        params.append(str(start_date)[:10])
    sql += " ORDER BY date, series"
    with get_conn() as conn:
        return pd.read_sql_query(sql, conn, params=params)


# ── Valuation ────────────────────────────────────────────────────
def upsert_valuation(data: dict):
    """data: dict with keys matching valuation table columns."""
    with get_conn() as conn:
        conn.execute(
            """INSERT OR REPLACE INTO valuation
               (ticker, name, sector, currency, price, market_cap, shares_out,
                updated_at, annual_json, quarterly_json, fwd_pe_json, fwd_ps_json,
                n_analysts)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (data["ticker"], data.get("name"), data.get("sector"),
             data.get("currency"), data.get("price"), data.get("market_cap"),
             data.get("shares_out"), data.get("updated_at", datetime.now().isoformat()),
             data.get("annual_json"), data.get("quarterly_json"),
             data.get("fwd_pe_json"), data.get("fwd_ps_json"),
             data.get("n_analysts"))
        )


def get_valuation(ticker=None, sector=None):
    sql = "SELECT * FROM valuation WHERE 1=1"
    params = []
    if ticker:
        sql += " AND ticker=?"
        params.append(ticker)
    if sector:
        sql += " AND sector=?"
        params.append(sector)
    sql += " ORDER BY ticker"
    with get_conn() as conn:
        return pd.read_sql_query(sql, conn, params=params)


# ── Earnings Calendar ────────────────────────────────────────────
def upsert_earnings(data: dict):
    """data: dict with keys matching earnings_calendar columns."""
    with get_conn() as conn:
        conn.execute(
            """INSERT OR REPLACE INTO earnings_calendar
               (ticker, name, sector, earnings_date, days_from_today,
                eps_estimate, fiscal_quarter, updated_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (data["ticker"], data.get("name"), data.get("sector"),
             data.get("earnings_date"), data.get("days_from_today"),
             data.get("eps_estimate"), data.get("fiscal_quarter"),
             data.get("updated_at", datetime.now().isoformat()))
        )


def clear_earnings():
    with get_conn() as conn:
        conn.execute("DELETE FROM earnings_calendar")


def get_earnings(sector=None):
    sql = "SELECT * FROM earnings_calendar WHERE 1=1"
    params = []
    if sector:
        sql += " AND sector=?"
        params.append(sector)
    sql += " ORDER BY earnings_date"
    with get_conn() as conn:
        return pd.read_sql_query(sql, conn, params=params)


# ── Price Percentiles ────────────────────────────────────────────
def get_price_percentiles():
    """计算每个 ticker 当前价在历史 weekly_prices 中的分位数。"""
    sql = """
    SELECT
        ticker,
        sector,
        COUNT(*)          AS weeks,
        MIN(close)        AS low,
        MAX(close)        AS high,
        -- 最新一周的收盘价
        (SELECT wp2.close
         FROM weekly_prices wp2
         WHERE wp2.ticker = wp.ticker
         ORDER BY wp2.date DESC LIMIT 1) AS current_price
    FROM weekly_prices wp
    WHERE close IS NOT NULL
    GROUP BY ticker
    """
    with get_conn() as conn:
        df = pd.read_sql_query(sql, conn)
    if df.empty:
        return df
    # percentile = (current - low) / (high - low)
    df["percentile"] = (df["current_price"] - df["low"]) / (df["high"] - df["low"])
    df["percentile"] = df["percentile"].clip(0, 1)
    return df
