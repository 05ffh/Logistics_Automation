"""数据驱动的 Fetch API 物流适配器基类。

宁致(NZ) 和小满(XM) 都通过浏览器内 fetch() 调用内部 JSON API，逻辑高度一致。
此基类抽取公共逻辑，子类只需声明 name/prefix/domain/url 等配置。

API: /tracking/app?inajax=1&tracking_number={tn}
返回: {data: {shipment: {traces: [{time, info}, ...]}}}
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


class FetchApiAdapter(CompanyAdapter):
    """通过 CDP fetch API 查询物流轨迹的通用适配器。

    子类只需定义以下类属性:
        name, prefix, canary_number
        domain, shipment_url, api_url
        batch_size (default 1)
        batch_interval_min / batch_interval_max
    并可按需重写 check_ready()。
    """

    # ── 子类必须定义 ──
    domain: str = ""
    shipment_url: str = ""
    api_url: str = ""

    # ── 可选配置 ──
    batch_interval_min: float = 2.0
    batch_interval_max: float = 5.0

    # ── check_ready ──────────────────────────────────────────────

    def check_ready(self, cdp) -> bool:
        """导航到页面并检查域名是否可达。"""
        cdp.evaluate(f"window.location.href='{self.shipment_url}';")
        time.sleep(3)
        url = val(cdp.evaluate("window.location.href"), "")
        return self.domain in url

    # ── query ────────────────────────────────────────────────────

    def query(self, cdp, tracking_nos: list[str]) -> list[TrackingResult]:
        results: dict[str, str | None] = {}
        errors: dict[str, str | None] = {}
        total = len(tracking_nos)

        # 刷新页面确保 JS 上下文新鲜
        cdp.evaluate("location.reload()")
        time.sleep(4)

        # 确保标签页在正确域上（fetch 需要同源 Cookie）
        url = val(cdp.evaluate("window.location.href"), "")
        if self.domain not in url:
            cdp.evaluate(f"window.location.href='{self.shipment_url}';")
            time.sleep(3)

        for batch_idx, start in enumerate(range(0, total, self.batch_size)):
            if batch_idx > 0:
                interval = random.uniform(self.batch_interval_min,
                                          self.batch_interval_max)
                time.sleep(interval)

            batch = tracking_nos[start:start + self.batch_size]
            for tn in batch:
                routing, error = self._fetch_one(cdp, tn)
                results[tn] = routing
                errors[tn] = error
                if error:
                    print(f"  [{self.name}] {start+1}-{start+len(batch)}/{total} {tn} ERROR: {error}")
                else:
                    status = "OK" if routing else "MISS"
                    print(f"  [{self.name}] {start+1}-{start+len(batch)}/{total} {tn} {status}")

        ok = sum(1 for tn in tracking_nos if results.get(tn))
        print(f"  [{self.name}] 合计 {ok}/{total} OK")
        return [TrackingResult(tn, results.get(tn), error=errors.get(tn))
                for tn in tracking_nos]

    # ── ensure_tab ───────────────────────────────────────────────

    def ensure_tab(self, cdp) -> str:
        tabs = cdp.list_tabs()
        for t in tabs:
            if t.get("type") == "page" and self.domain in (t.get("url") or ""):
                return t.get("webSocketDebuggerUrl", "")
        page_tabs = [t for t in tabs if t.get("type") == "page"]
        if not page_tabs:
            raise RuntimeError("No browser tabs. Is Chrome running?")
        cdp.connect_tab(page_tabs[0]["webSocketDebuggerUrl"])
        cdp.evaluate(f"window.open('{self.shipment_url}', '_blank')")
        cdp.close()
        time.sleep(2)
        for t in cdp.list_tabs():
            if t.get("type") == "page" and self.domain in (t.get("url") or ""):
                return t.get("webSocketDebuggerUrl", "")
        raise RuntimeError(f"Cannot open {self.name} tab.")

    # ── fetch ────────────────────────────────────────────────────

    def _fetch_one(self, cdp, tracking_no: str) -> tuple[str | None, str | None]:
        """查询一个单号，重试一次缓解偶发网络抖动。"""
        routing, error = self._fetch_once(cdp, tracking_no)
        if routing is None and error is None:
            time.sleep(2)
            routing, error = self._fetch_once(cdp, tracking_no)
        return routing, error

    def _fetch_once(self, cdp, tracking_no: str) -> tuple[str | None, str | None]:
        resp = cdp.fetch_api(self.api_url + tracking_no, timeout=10)
        if not resp.get("ok"):
            return None, resp.get("error", "fetch failed")
        try:
            data = resp.get("data", {})
            shipment = data.get("data", {}).get("shipment", {})
            traces = shipment.get("traces", [])
            if not traces:
                return None, None
            latest = traces[0]
            ts = latest.get("time", "")
            info = latest.get("info", "")
            if not ts or not info:
                return None, None
            result = f"{ts}\n{info}"
            return (result, None) if is_valid_routing(result) else (None, None)
        except (KeyError, IndexError, TypeError, AttributeError) as e:
            return None, f"parse error: {e}"
