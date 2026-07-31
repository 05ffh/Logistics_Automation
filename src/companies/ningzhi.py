"""宁致物流适配器 - nzhexp.nextsls.com。

需要登录。通过浏览器内 fetch() 调用内部 JSON API。
继承 FetchApiAdapter，仅重写 check_ready 检测登录态。
"""

from __future__ import annotations

import time

from .fetch_adapter import FetchApiAdapter

try:
    from ..cdp_util import val
except ImportError:
    from cdp_util import val


class NingZhiAdapter(FetchApiAdapter):
    name = "宁致"
    prefix = "NZ"
    batch_size = 1
    canary_number = "NZ2605063839"
    domain = "nzhexp.nextsls.com"
    shipment_url = "https://nzhexp.nextsls.com/tms/wos/shipment"
    api_url = "https://nzhexp.nextsls.com/tracking/app?inajax=1&tracking_number="

    def check_ready(self, cdp) -> bool:
        """导航到运单页检查是否被重定向到登录。"""
        cdp.evaluate(
            f"window.location.href='{self.shipment_url}?page=1&pageSize=30';"
        )
        time.sleep(3)
        url = val(cdp.evaluate("window.location.href"), "")
        return "/login" not in url
