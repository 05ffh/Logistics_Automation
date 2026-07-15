"""CDP 相关工具函数，各适配器共用。"""


def val(cdp_result: dict, default=None):
    """从 CDP Runtime.evaluate 响应中提取 value。"""
    return cdp_result.get("result", {}).get("result", {}).get("value", default)
