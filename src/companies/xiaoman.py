"""小满物流适配器 - xmsdwl.nextsls.com。

无需登录。通过浏览器内 fetch() 调用内部 JSON API，不抓 DOM。
借鉴象往项目的 CDP fetch 策略：fetch 在浏览器内执行，携带完整 Cookie/会话，
从服务器角度看与页面自身的 AJAX 请求无法区分。

API: /tracking/app?inajax=1&tracking_number={tn}
返回: {data: {shipment: {traces: [{time, info}, ...]}}}
"""

from __future__ import annotations

import json
import random
import time

from .base import CompanyAdapter, TrackingResult

try:
    from ..validation import is_valid_routing
    from ..cdp_util import val
except ImportError:
    from validation import is_valid_routing
    from cdp_util import val

XM_DOMAIN = "xmsdwl.nextsls.com"
TRACKING_URL = "https://xmsdwl.nextsls.com/tracking/app#/tracking"
API_URL = "https://xmsdwl.nextsls.com/tracking/app?inajax=1&tracking_number="
MAX_BATCH = 5
# 象往式礼貌间隔：批间随机 2-5s，避免连续高速请求触发限流
BATCH_INTERVAL_MIN = 2.0
BATCH_INTERVAL_MAX = 5.0


class XiaoManAdapter(CompanyAdapter):
    name = "小满"
    prefix = "XM"
    batch_size = MAX_BATCH
    canary_number = "XM26070315932"

    def check_ready(self, cdp) -> bool:
        cdp.evaluate(f"window.location.href='{TRACKING_URL}';")
        time.sleep(3)
        url = val(cdp.evaluate("window.location.href"), "")
        return XM_DOMAIN in url

    def query(self, cdp, tracking_nos: list[str]) -> list[TrackingResult]:
        results: dict[str, str | None] = {}
        total = len(tracking_nos)

        # 确保标签页在小满域上（fetch 需要同源 Cookie）
        url = val(cdp.evaluate("window.location.href"), "")
        if XM_DOMAIN not in url:
            cdp.evaluate(f"window.location.href='{TRACKING_URL}';")
            time.sleep(3)

        for batch_idx, start in enumerate(range(0, total, MAX_BATCH)):
            if batch_idx > 0:
                interval = random.uniform(BATCH_INTERVAL_MIN, BATCH_INTERVAL_MAX)
                time.sleep(interval)

            batch = tracking_nos[start:start + MAX_BATCH]
            for tn in batch:
                routing = self._fetch_one(cdp, tn)
                results[tn] = routing
                status = "OK" if routing else "MISS"
                print(f"  [{self.name}] {start+1}-{start+len(batch)}/{total} {tn} {status}")

        ok = sum(1 for tn in tracking_nos if results.get(tn))
        print(f"  [{self.name}] 合计 {ok}/{total} OK")
        return [TrackingResult(tn, results.get(tn)) for tn in tracking_nos]

    def ensure_tab(self, cdp) -> str:
        tabs = cdp.list_tabs()
        for t in tabs:
            if t.get("type") == "page" and XM_DOMAIN in (t.get("url") or ""):
                return t.get("webSocketDebuggerUrl", "")
        page_tabs = [t for t in tabs if t.get("type") == "page"]
        if not page_tabs:
            raise RuntimeError("No browser tabs. Is Edge running?")
        cdp.connect_tab(page_tabs[0]["webSocketDebuggerUrl"])
        cdp.evaluate(f"window.open('{TRACKING_URL}', '_blank')")
        cdp.close()
        time.sleep(2)
        for t in cdp.list_tabs():
            if t.get("type") == "page" and XM_DOMAIN in (t.get("url") or ""):
                return t.get("webSocketDebuggerUrl", "")
        raise RuntimeError("Cannot open xmsdwl tab.")

    # ── fetch API 查询 ────────────────────────────────────────

    def _fetch_one(self, cdp, tracking_no: str) -> str | None:
        """通过浏览器内 fetch() 调用内部 API。

        重试一次（象往式固定间隔），缓解偶发网络抖动。
        """
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


