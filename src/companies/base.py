"""物流公司适配器基类。"""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass


@dataclass
class TrackingResult:
    tracking_no: str
    routing_info: str | None


class CompanyAdapter(ABC):
    """物流公司查询适配器基类。

    每个发货公司继承此类，实现对应的查询逻辑。
    """

    # 子类必须定义
    name: str           # 公司名称（与 Excel K/J 列匹配）
    prefix: str         # 物流单号前缀
    batch_size: int = 1

    @abstractmethod
    def query(self, cdp, tracking_nos: list[str]) -> list[TrackingResult]:
        """批量查询物流单号，返回每条的最新路由信息。"""
        ...

    def check_ready(self, cdp) -> bool:
        """检查前置条件（登录、页面可达等），返回是否就绪。"""
        return True
