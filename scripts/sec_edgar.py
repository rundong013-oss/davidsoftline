#!/usr/bin/env python3
"""
sec_edgar.py — SEC EDGAR 财报原文下载器

从 SEC EDGAR 下载 10-K（年报）和 10-Q（季报）原文，转换为纯文本存储。
仅支持美股 SEC 上市公司（NKE, LULU, ONON, DECK 等）。

Usage:
    python sec_edgar.py NKE                        # 下载最新 1 份 10-K
    python sec_edgar.py NKE --form 10-Q            # 下载最新 1 份 10-Q
    python sec_edgar.py NKE --form 10-K --limit 3  # 下载最近 3 份 10-K
    python sec_edgar.py NKE LULU --form 10-Q       # 多 ticker
    python sec_edgar.py NKE --list                 # 仅列出可用 filing，不下载
    python sec_edgar.py NKE --form 10-K --raw      # 同时保留 HTML 原文

输出目录: library/{TICKER}/filings/
文件命名: {TICKER}_{form}_{period_end}.txt

可作为模块导入:
    from sec_edgar import fetch_filing_text
    text = fetch_filing_text("NKE", "10-K")
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import time
from html import unescape
from pathlib import Path
from typing import Dict, List, Optional, Union
from urllib.request import Request, urlopen
from urllib.error import HTTPError, URLError

# ── 路径 ─────────────────────────────────────────────────────────
WORKSPACE = Path(__file__).resolve().parent.parent
LIBRARY = WORKSPACE / "library"

# ── SEC EDGAR 配置 ───────────────────────────────────────────────
USER_AGENT = "SportswearResearch david@research.com"
BASE_DATA = "https://data.sec.gov"
BASE_ARCHIVES = "https://www.sec.gov/Archives/edgar/data"
TICKER_MAP_URL = "https://www.sec.gov/files/company_tickers.json"
REQ_INTERVAL = 0.12  # 秒，确保 < 10 req/s

_last_req_time = 0.0


def _throttled_get(url: str, *, decode: bool = True) -> str | bytes:
    """发送 GET 请求，自动限速 + User-Agent。"""
    global _last_req_time
    elapsed = time.time() - _last_req_time
    if elapsed < REQ_INTERVAL:
        time.sleep(REQ_INTERVAL - elapsed)

    req = Request(url, headers={"User-Agent": USER_AGENT})
    for attempt in range(3):
        try:
            with urlopen(req, timeout=30) as resp:
                data = resp.read()
                _last_req_time = time.time()
                return data.decode("utf-8") if decode else data
        except HTTPError as e:
            if e.code == 429:
                wait = 2 * (attempt + 1)
                print(f"  [WARN] 429 Too Many Requests, 等待 {wait}s...",
                      file=sys.stderr)
                time.sleep(wait)
                continue
            raise
        except URLError:
            if attempt < 2:
                time.sleep(1)
                continue
            raise
    return ""


# ── CIK 查询 ────────────────────────────────────────────────────
_cik_cache: dict[str, str] = {}


def get_cik(ticker: str) -> str | None:
    """通过 ticker 查询 SEC CIK 编号（10 位零填充）。"""
    ticker_upper = ticker.upper().replace(".N", "").replace(".O", "")
    if ticker_upper in _cik_cache:
        return _cik_cache[ticker_upper]

    try:
        raw = _throttled_get(TICKER_MAP_URL)
        data = json.loads(raw)
        for entry in data.values():
            t = entry["ticker"].upper()
            cik = str(entry["cik_str"]).zfill(10)
            _cik_cache[t] = cik
            if t == ticker_upper:
                return cik
    except Exception as e:
        print(f"  [ERROR] CIK 查询失败: {e}", file=sys.stderr)

    return None


# ── Filing 索引 ──────────────────────────────────────────────────
def list_filings(ticker: str, form_type: str = "10-K",
                 limit: int = 10) -> list[dict]:
    """
    获取指定 ticker 的 filing 列表。

    返回 list[dict]，每项包含:
        accessionNumber, filingDate, primaryDocument, reportDate, form
    """
    cik = get_cik(ticker)
    if not cik:
        print(f"  [ERROR] 未找到 {ticker} 的 CIK", file=sys.stderr)
        return []

    url = f"{BASE_DATA}/submissions/CIK{cik}.json"
    try:
        raw = _throttled_get(url)
    except Exception as e:
        print(f"  [ERROR] 获取 {ticker} filing 索引失败: {e}", file=sys.stderr)
        return []

    data = json.loads(raw)
    recent = data.get("filings", {}).get("recent", {})
    forms = recent.get("form", [])
    results = []

    for i, form in enumerate(forms):
        if form == form_type and len(results) < limit:
            results.append({
                "cik": cik,
                "accessionNumber": recent["accessionNumber"][i],
                "filingDate": recent["filingDate"][i],
                "primaryDocument": recent["primaryDocument"][i],
                "reportDate": recent.get("reportDate", [""] * len(forms))[i],
                "form": form,
            })

    return results


# ── HTML → 纯文本 ────────────────────────────────────────────────
def _html_to_text(html: str) -> str:
    """轻量 HTML 转纯文本，不依赖 BeautifulSoup。"""
    # 移除 Inline XBRL 隐藏块（含 XBRL context/metadata）
    text = re.sub(r"<ix:hidden[^>]*>.*?</ix:hidden>", "", html,
                  flags=re.DOTALL | re.IGNORECASE)
    # 移除 script/style 块
    text = re.sub(r"<(script|style|head)[^>]*>.*?</\1>", "", text,
                  flags=re.DOTALL | re.IGNORECASE)
    # <br> / <p> / <div> / <tr> → 换行
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</(p|div|tr|li|h[1-6])>", "\n", text,
                  flags=re.IGNORECASE)
    text = re.sub(r"<(p|div|tr|li|h[1-6])[^>]*>", "\n", text,
                  flags=re.IGNORECASE)
    # <td> → tab 分隔
    text = re.sub(r"<td[^>]*>", "\t", text, flags=re.IGNORECASE)
    # 移除所有剩余标签（含 ix: XBRL 内联标签）
    text = re.sub(r"<[^>]+>", "", text)
    # 解码 HTML 实体
    text = unescape(text)
    # 过滤 XBRL 残留行（特征：无空格的超长字符串，含 CIK/us-gaap 等关键词）
    lines = []
    blank_count = 0
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            blank_count += 1
            if blank_count <= 2:
                lines.append("")
            continue
        # 跳过 XBRL 元数据残留（超长无空格行 + 含 XBRL 关键词）
        if (len(stripped) > 500 and " " not in stripped[:200]
                and ("us-gaap" in stripped or "xbrli:" in stripped
                     or "iso4217:" in stripped)):
            continue
        blank_count = 0
        lines.append(stripped)
    return "\n".join(lines).strip()


# ── 下载单份 filing ──────────────────────────────────────────────
def download_filing(ticker: str, filing: dict, *,
                    save_raw: bool = False) -> Path | None:
    """
    下载一份 filing 并保存为纯文本。

    Args:
        ticker: 股票代码
        filing: list_filings() 返回的 dict
        save_raw: 是否同时保存 HTML 原文

    Returns:
        保存的 txt 文件路径，失败返回 None
    """
    cik = filing["cik"]
    accession_clean = filing["accessionNumber"].replace("-", "")
    primary_doc = filing["primaryDocument"]
    doc_url = f"{BASE_ARCHIVES}/{int(cik)}/{accession_clean}/{primary_doc}"

    try:
        html = _throttled_get(doc_url)
    except Exception as e:
        print(f"  [ERROR] 下载失败 {doc_url}: {e}", file=sys.stderr)
        return None

    # 确定输出目录
    out_dir = LIBRARY / ticker.upper() / "filings"
    out_dir.mkdir(parents=True, exist_ok=True)

    # 文件名: NKE_10-K_2024-05-31.txt (用 reportDate，fallback filingDate)
    date_str = filing["reportDate"] or filing["filingDate"]
    form_safe = filing["form"].replace("/", "-")
    base_name = f"{ticker.upper()}_{form_safe}_{date_str}"

    # 保存纯文本
    plain_text = _html_to_text(html)
    txt_path = out_dir / f"{base_name}.txt"
    txt_path.write_text(plain_text, encoding="utf-8")

    # 可选保存 HTML 原文
    if save_raw:
        htm_path = out_dir / f"{base_name}.htm"
        htm_path.write_text(html, encoding="utf-8")

    return txt_path


# ── 便捷 API（供其他脚本调用） ───────────────────────────────────
def fetch_filing_text(ticker: str, form_type: str = "10-K",
                      limit: int = 1) -> list[dict]:
    """
    一站式获取 filing 纯文本。

    Returns:
        list[dict]，每项包含:
            path: 文件路径 (Path)
            filingDate: 提交日期
            reportDate: 报告期末日期
            form: 表单类型
            text: 纯文本内容
    """
    filings = list_filings(ticker, form_type, limit=limit)
    results = []
    for f in filings:
        path = download_filing(ticker, f)
        if path and path.exists():
            results.append({
                "path": path,
                "filingDate": f["filingDate"],
                "reportDate": f["reportDate"],
                "form": f["form"],
                "text": path.read_text(encoding="utf-8"),
            })
    return results


# ── CLI ──────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="SEC EDGAR 财报原文下载器 (10-K / 10-Q)")
    parser.add_argument("tickers", nargs="+",
                        help="股票代码（如 NKE LULU）")
    parser.add_argument("--form", default="10-K",
                        choices=["10-K", "10-Q", "8-K", "20-F"],
                        help="SEC 表单类型（默认 10-K）")
    parser.add_argument("--limit", type=int, default=1,
                        help="每个 ticker 下载份数（默认 1）")
    parser.add_argument("--list", action="store_true", dest="list_only",
                        help="仅列出可用 filing，不下载")
    parser.add_argument("--raw", action="store_true",
                        help="同时保留 HTML 原文")
    args = parser.parse_args()

    for ticker in args.tickers:
        ticker = ticker.upper()
        print(f"\n{'='*60}")
        print(f"  {ticker} | {args.form} | limit={args.limit}")
        print(f"{'='*60}")

        cik = get_cik(ticker)
        if not cik:
            print(f"  [SKIP] 未找到 {ticker} 的 CIK，"
                  "可能不是美股 SEC 上市公司")
            continue
        print(f"  CIK: {cik}")

        filings = list_filings(ticker, args.form, limit=args.limit)
        if not filings:
            print(f"  [SKIP] 未找到 {args.form} filing")
            continue

        print(f"  找到 {len(filings)} 份 {args.form}:")
        for i, f in enumerate(filings):
            tag = f"  [{i+1}] {f['filingDate']}"
            if f["reportDate"]:
                tag += f" (报告期: {f['reportDate']})"
            tag += f"  {f['primaryDocument']}"
            print(tag)

        if args.list_only:
            continue

        # 下载
        print()
        for f in filings:
            path = download_filing(ticker, f, save_raw=args.raw)
            if path:
                size_kb = path.stat().st_size / 1024
                print(f"  ✓ {path.name}  ({size_kb:,.0f} KB)")
            else:
                print(f"  ✗ {f['filingDate']} 下载失败")

    print()


if __name__ == "__main__":
    main()
