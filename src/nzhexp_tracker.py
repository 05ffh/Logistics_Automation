"""宁致物流轨迹查询 - nzhexp.nextsls.com。

通过 CDP 操控 Edge，批量查询物流单号，从抽屉面板中提取最新路由信息。
网站使用 Ant Design Vue，以下选择器已在 2026-07-10 实际探查确认。
"""

from __future__ import annotations

import re
import time
from dataclasses import dataclass

try:
    from .cdp_client import CdpClient
except ImportError:
    from cdp_client import CdpClient

NZHEXP_DOMAIN = "nzhexp.nextsls.com"
NZHEXP_SHIPMENT_URL = "http://nzhexp.nextsls.com/tms/wos/shipment"

# -- 搜索页选择器 --
SEL_SEARCH_INPUT = "input.ant-input"
SEL_SEARCH_BTN = "button.pod-btn-success.pod-btn-block"
SEL_TABLE_ROWS = ".ant-table-tbody tr.ant-table-row"
SEL_TRACKING_CELL = "td:nth-child(2)"
SEL_SLIDEPANEL_BTN = "span.btn-slidepanel-open"

# -- 抽屉面板选择器 --
SEL_DRAWER_BODY = ".ant-drawer-body"
SEL_DRAWER_CLOSE = ".ant-drawer-close"

# 每批查询单号数
BATCH_SIZE = 1

@dataclass
class TrackingResult:
    tracking_no: str
    routing_info: str | None


def check_logged_in(cdp: CdpClient) -> bool:
    """检查 nzhexp 是否已登录。"""
    url = cdp.evaluate("window.location.href")
    current = _extract_value(url, "")
    return "/login" not in current


def query_tracking(cdp: CdpClient, tracking_nos: list[str]) -> list[TrackingResult]:
    """批量查询物流单号（每批 BATCH_SIZE 个）。"""
    results: dict[str, str | None] = {}
    total = len(tracking_nos)

    for batch_start in range(0, total, BATCH_SIZE):
        batch = tracking_nos[batch_start : batch_start + BATCH_SIZE]
        batch_results = _query_batch(cdp, batch)

        for r in batch_results:
            results[r.tracking_no] = r.routing_info

        done = min(batch_start + BATCH_SIZE, total)
        for tn in batch:
            status = "OK" if results.get(tn) else "MISS"
            print(f"  [{done}/{total}] {tn} {status}")

    return [TrackingResult(tn, results.get(tn)) for tn in tracking_nos]


def _query_batch(cdp: CdpClient, tracking_nos: list[str]) -> list[TrackingResult]:
    """查询一批单号（逐个搜，BATCH_SIZE=1 确保稳定）。"""
    keyword = tracking_nos[0]

    # 0. 完整清空并导航到干净页面
    cdp.evaluate(
        f"sessionStorage.clear();localStorage.removeItem('keywords');"
        f"window.location.href='about:blank';"
    )
    time.sleep(0.5)
    cdp.evaluate(
        f"window.location.href='{NZHEXP_SHIPMENT_URL}?page=1&pageSize=30&_t='+Date.now();"
    )
    time.sleep(3)

    # 1. 填入单号
    cdp.evaluate(
        f"(function(){{"
        f"var el=document.querySelector('{SEL_SEARCH_INPUT}');"
        f"if(el){{el.focus();el.value='';}}"
        f"}})()"
    )
    time.sleep(0.2)
    cdp._send({"method": "Input.insertText", "params": {"text": keyword}})
    time.sleep(0.3)

    # 2. 搜索
    cdp.click(SEL_SEARCH_BTN)
    time.sleep(3)

    # 3. 从结果行逐个提取
    results: list[TrackingResult] = []
    found_tns: set[str] = set()

    rows_js = f"document.querySelectorAll('{SEL_TABLE_ROWS}').length"
    row_count = _extract_value(cdp.evaluate(rows_js), 0)

    for _ in range(row_count):
        # 找第一个未处理的匹配行
        for tn in tracking_nos:
            if tn in found_tns:
                continue
            clicked = cdp.evaluate(
                f"(function(){{"
                f"var rows=document.querySelectorAll('{SEL_TABLE_ROWS}');"
                f"for(var i=0;i<rows.length;i++){{"
                f"var cell=rows[i].querySelector('{SEL_TRACKING_CELL}');"
                f"var btn=cell?cell.querySelector('{SEL_SLIDEPANEL_BTN}'):null;"
                f"if(btn&&btn.textContent.trim()=='{tn}'){{btn.click();return 1;}}"
                f"}}"
                f"return 0;"
                f"}})()"
            )
            if _extract_value(clicked, 0):
                time.sleep(1.5)
                routing = _extract_routing_from_drawer(cdp)
                results.append(TrackingResult(tn, routing))
                found_tns.add(tn)
                cdp.click(SEL_DRAWER_CLOSE)
                time.sleep(1)
                break

    # 未找到结果的单号
    for tn in tracking_nos:
        if tn not in found_tns:
            results.append(TrackingResult(tn, None))

    return results


def _extract_routing_from_drawer(cdp: CdpClient) -> str | None:
    drawer_check = cdp.evaluate(
        f"(function(){{var d=document.querySelector('{SEL_DRAWER_BODY}');"
        f"return d?d.textContent.trim():null;}})()"
    )
    text = _extract_value(drawer_check)
    if not text:
        return None
    return _parse_routing_info(text)


def _parse_routing_info(drawer_text: str) -> str | None:
    idx = drawer_text.find("路由信息")
    if idx == -1:
        return None
    section = drawer_text[idx:]
    pattern = re.compile(
        r"(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})\s+(.*?)(?=\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}|$)",
        re.DOTALL,
    )
    matches = pattern.findall(section)
    if not matches:
        return None
    timestamp, description = matches[0]
    description = description.strip()
    if len(matches) > 1 and matches[1][0] == timestamp:
        description += " " + matches[1][1].strip()
    return f"{timestamp}\n{description}"


# -- Tab 管理 --
def find_nzhexp_tab(cdp: CdpClient) -> str | None:
    tabs = cdp.list_tabs()
    for t in tabs:
        if t.get("type") == "page" and NZHEXP_DOMAIN in (t.get("url") or ""):
            return t.get("webSocketDebuggerUrl", "")
    return None


def ensure_nzhexp_tab(cdp: CdpClient) -> str:
    ws_url = find_nzhexp_tab(cdp)
    if ws_url:
        return ws_url
    tabs = cdp.list_tabs()
    page_tabs = [t for t in tabs if t.get("type") == "page"]
    if not page_tabs:
        raise RuntimeError("No browser tabs found. Is Edge running?")
    cdp.connect_tab(page_tabs[0]["webSocketDebuggerUrl"])
    cdp.evaluate(f"window.open('{NZHEXP_SHIPMENT_URL}', '_blank')")
    cdp.close()
    time.sleep(2)
    ws_url = find_nzhexp_tab(cdp)
    if ws_url:
        return ws_url
    raise RuntimeError("Cannot open nzhexp tab. Is Edge running with the logistics profile?")


def _extract_value(cdp_result: dict, default=None):
    return cdp_result.get("result", {}).get("result", {}).get("value", default)
