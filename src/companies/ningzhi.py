"""宁致物流适配器 - nzhexp.nextsls.com。

需要登录。通过浏览器内 fetch() 调用内部 JSON API，不抓 DOM。
借鉴象往项目的 CDP fetch 策略：fetch 在浏览器内执行，携带完整 Cookie/会话。

API: /tracking/app?inajax=1&tracking_number={tn}
返回: {data: {shipment: {traces: [{time, info}, ...]}}}

check_ready / ensure_tab 保留 DOM 方式（验证登录态、打开标签页）。
"""

from __future__ import annotations

import random
import time

from .base import CompanyAdapter, TrackingResult

try:
    from ..validation import is_valid_routing
    from ..cdp_util import val
except ImportError:
    from validation import is_valid_routing
    from cdp_util import val

NZHEXP_DOMAIN = "nzhexp.nextsls.com"
NZHEXP_SHIPMENT_URL = "http://nzhexp.nextsls.com/tms/wos/shipment"
API_URL = "http://nzhexp.nextsls.com/tracking/app?inajax=1&tracking_number="

# 象往式礼貌间隔
BATCH_INTERVAL_MIN = 2.0
BATCH_INTERVAL_MAX = 5.0


class NingZhiAdapter(CompanyAdapter):
    name = "宁致"
    prefix = "NZ"
    batch_size = 1
    canary_number = "NZ2605063839"

    def check_ready(self, cdp) -> bool:
        """通过导航到运单页检查是否被重定向到登录。"""
        cdp.evaluate(f"window.location.href='{NZHEXP_SHIPMENT_URL}?page=1&pageSize=30';")
        time.sleep(3)
        url = val(cdp.evaluate("window.location.href"), "")
        return "/login" not in url

    def query(self, cdp, tracking_nos: list[str]) -> list[TrackingResult]:
        results: dict[str, str | None] = {}
        total = len(tracking_nos)

        # 确保标签页在宁致域上（fetch 需要同源 Cookie）
        url = val(cdp.evaluate("window.location.href"), "")
        if NZHEXP_DOMAIN not in url:
            cdp.evaluate(f"window.location.href='{NZHEXP_SHIPMENT_URL}?page=1&pageSize=30';")
            time.sleep(3)

        for i, tn in enumerate(tracking_nos):
            if i > 0:
                interval = random.uniform(BATCH_INTERVAL_MIN, BATCH_INTERVAL_MAX)
                time.sleep(interval)

            routing = self._fetch_one(cdp, tn)
            results[tn] = routing
            status = "OK" if routing else "MISS"
            print(f"  [{self.name}] {i+1}/{total} {tn} {status}")

        ok = sum(1 for tn in tracking_nos if results.get(tn))
        print(f"  [{self.name}] 合计 {ok}/{total} OK")
        return [TrackingResult(tn, results.get(tn)) for tn in tracking_nos]

    def ensure_tab(self, cdp) -> str:
        tabs = cdp.list_tabs()
        for t in tabs:
            if t.get("type") == "page" and NZHEXP_DOMAIN in (t.get("url") or ""):
                return t.get("webSocketDebuggerUrl", "")
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

    # ── fetch API 查询 ────────────────────────────────────────

    def _fetch_one(self, cdp, tracking_no: str) -> str | None:
        routing = self._fetch_once(cdp, tracking_no)
        if routing is None:
            time.sleep(2)
            routing = self._fetch_once(cdp, tracking_no)
        return routing

    def _fetch_once(self, cdp, tracking_no: str) -> str | None:
        resp = cdp.fetch_api(API_URL + tracking_no, timeout=10)
        if not resp.get("ok"):
            return None
        try:
            data = resp.get("data", {})
            shipment = data.get("data", {}).get("shipment", {})
            traces = shipment.get("traces", [])
            if not traces:
                return None
            latest = traces[0]
            ts = latest.get("time", "")
            info = latest.get("info", "")
            if not ts or not info:
                return None
            result = f"{ts}\n{info}"
            return result if is_valid_routing(result) else None
        except (KeyError, IndexError, TypeError, AttributeError):
            return None


