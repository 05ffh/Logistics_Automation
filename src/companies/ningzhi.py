"""宁致物流适配器 - nzhexp.nextsls.com。"""

from __future__ import annotations

import re
import time

from .base import CompanyAdapter, TrackingResult

try:
    from ..validation import is_valid_routing
except ImportError:
    from validation import is_valid_routing

NZHEXP_SHIPMENT_URL = "http://nzhexp.nextsls.com/tms/wos/shipment"
NZHEXP_DOMAIN = "nzhexp.nextsls.com"

# DOM 选择器
SEL_SEARCH_INPUT = "input.ant-input"
SEL_SEARCH_BTN = "button.pod-btn-success.pod-btn-block"
SEL_TABLE_ROWS = ".ant-table-tbody tr.ant-table-row"
SEL_TRACKING_CELL = "td:nth-child(2)"
SEL_SLIDEPANEL_BTN = "span.btn-slidepanel-open"
SEL_DRAWER_BODY = ".ant-drawer-body"
SEL_DRAWER_CLOSE = ".ant-drawer-close"


class NingZhiAdapter(CompanyAdapter):
    name = "宁致"
    prefix = "NZ"
    batch_size = 1
    canary_number = "NZ2605063839"  # 自检用已知单号（失效时更新）

    def check_ready(self, cdp) -> bool:
        # 导航到运单页，检查是否被重定向到登录
        cdp.evaluate(f"window.location.href='{NZHEXP_SHIPMENT_URL}?page=1&pageSize=30';")
        time.sleep(3)
        url = cdp.evaluate("window.location.href")
        current = _extract_value(url, "")
        return "/login" not in current

    def query(self, cdp, tracking_nos: list[str]) -> list[TrackingResult]:
        results: dict[str, str | None] = {}
        total = len(tracking_nos)

        for i, tn in enumerate(tracking_nos):
            routing = self._query_one(cdp, tn)
            results[tn] = routing
            status = "OK" if routing else "MISS"
            print(f"  [{self.name}] {i+1}/{total} {tn} {status}")

        return [TrackingResult(tn, results.get(tn)) for tn in tracking_nos]

    def ensure_tab(self, cdp) -> str:
        tabs = cdp.list_tabs()
        for t in tabs:
            if t.get("type") == "page" and NZHEXP_DOMAIN in (t.get("url") or ""):
                return t.get("webSocketDebuggerUrl", "")
        # 打开新标签页
        page_tabs = [t for t in tabs if t.get("type") == "page"]
        if not page_tabs:
            raise RuntimeError("No browser tabs. Is Edge running?")
        cdp.connect_tab(page_tabs[0]["webSocketDebuggerUrl"])
        cdp.evaluate(f"window.open('{NZHEXP_SHIPMENT_URL}', '_blank')")
        cdp.close()
        time.sleep(2)
        for t in cdp.list_tabs():
            if t.get("type") == "page" and NZHEXP_DOMAIN in (t.get("url") or ""):
                return t.get("webSocketDebuggerUrl", "")
        raise RuntimeError("Cannot open nzhexp tab.")

    def _query_one(self, cdp, tracking_no: str) -> str | None:
        # 查空重试一次（缓解 SPA 加载慢/偶发时序竞态）
        routing = self._query_once(cdp, tracking_no)
        if routing is None:
            routing = self._query_once(cdp, tracking_no)
        return routing

    def _query_once(self, cdp, tracking_no: str) -> str | None:
        # 0. about:blank 重置
        cdp.evaluate("sessionStorage.clear();localStorage.removeItem('keywords');window.location.href='about:blank';")
        time.sleep(0.5)
        cdp.evaluate(
            f"window.location.href='{NZHEXP_SHIPMENT_URL}?page=1&pageSize=30&_t='+Date.now();"
        )
        time.sleep(3)

        # 1. 聚焦 + 输入
        cdp.evaluate(
            f"(function(){{var el=document.querySelector('{SEL_SEARCH_INPUT}');"
            f"if(el){{el.focus();el.value='';}}}})()"
        )
        time.sleep(0.2)
        cdp._send({"method": "Input.insertText", "params": {"text": tracking_no}})
        time.sleep(0.3)

        # 2. 搜索
        cdp.click(SEL_SEARCH_BTN)
        time.sleep(3)

        # 3. 匹配并打开抽屉
        clicked = cdp.evaluate(
            f"(function(){{"
            f"var rows=document.querySelectorAll('{SEL_TABLE_ROWS}');"
            f"for(var i=0;i<rows.length;i++){{"
            f"var cell=rows[i].querySelector('{SEL_TRACKING_CELL}');"
            f"var btn=cell?cell.querySelector('{SEL_SLIDEPANEL_BTN}'):null;"
            f"if(btn&&btn.textContent.trim()=='{tracking_no}'){{btn.click();return 1;}}"
            f"}}"
            f"return 0;"
            f"}})()"
        )
        if not _extract_value(clicked, 0):
            return None

        time.sleep(1.5)

        # 4. 提取路由
        drawer = cdp.evaluate(
            f"(function(){{var d=document.querySelector('{SEL_DRAWER_BODY}');"
            f"return d?d.textContent.trim():null;}})()"
        )
        text = _extract_value(drawer)
        routing = _parse_nz_routing(text) if text else None

        # 5. 关闭抽屉
        cdp.click(SEL_DRAWER_CLOSE)
        time.sleep(1)

        # 6. 校验：不合格当作未查到
        return routing if is_valid_routing(routing) else None


def _parse_nz_routing(drawer_text: str) -> str | None:
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


def _extract_value(cdp_result: dict, default=None):
    return cdp_result.get("result", {}).get("result", {}).get("value", default)
