"""路由数据合理性校验。

写回 Excel 前用它拦掉"看起来对、其实是页面结构变化抓错"的垃圾数据：
页面改版后正则可能抓到无意义文本，宁可当作 MISS 也不要把脏数据写进业务表。
"""

from __future__ import annotations

import re
from datetime import datetime

# 轨迹里应含形如 YYYY-MM-DD 的时间戳
_DATE = re.compile(r"(\d{4})-(\d{2})-(\d{2})")


def is_valid_routing(routing: str | None) -> bool:
    """判断一条路由文本是否合理可写。

    要求：非空、含合理年份范围内的日期、除时间戳外还有描述文字。
    """
    if not routing or not routing.strip():
        return False

    m = _DATE.search(routing)
    if not m:
        return False

    year, month, day = int(m.group(1)), int(m.group(2)), int(m.group(3))
    now_year = datetime.now().year
    if not (2020 <= year <= now_year + 1):
        return False
    if not (1 <= month <= 12 and 1 <= day <= 31):
        return False

    # 至少要有 时间戳 + 一行描述（否则只是抓到个孤零零的日期）
    lines = [ln.strip() for ln in routing.split("\n") if ln.strip()]
    if len(lines) < 2:
        return False

    return True
