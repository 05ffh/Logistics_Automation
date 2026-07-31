"""小满物流适配器 - xmsdwl.nextsls.com。

无需登录。通过浏览器内 fetch() 调用内部 JSON API。
继承 FetchApiAdapter，完全使用基类默认实现。
"""

from __future__ import annotations

from .fetch_adapter import FetchApiAdapter


class XiaoManAdapter(FetchApiAdapter):
    name = "小满"
    prefix = "XM"
    batch_size = 5
    canary_number = "XM26070315932"
    domain = "xmsdwl.nextsls.com"
    shipment_url = "https://xmsdwl.nextsls.com/tracking/app#/tracking"
    api_url = "https://xmsdwl.nextsls.com/tracking/app?inajax=1&tracking_number="
