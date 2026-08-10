"""标签页刷新保活守护 — 长时间运行时定期刷新物流网站标签页，防止 session 过期。

用法:
    from .keepalive import TabKeepAlive
    keeper = TabKeepAlive.start(host, port, domains=["nzhexp", "xmsdwl", "17track"])
    # ... long query run ...
    keeper.stop()

Windows 常驻模式 (独立进程，配合 .bat 使用):
    python -m src.keepalive --daemon
"""

from __future__ import annotations

import json
import logging
import threading
import time
from urllib.error import URLError
from urllib.request import urlopen

logger = logging.getLogger(__name__)

# 需要保活的物流域名关键词
DEFAULT_DOMAINS = ["nzhexp", "xmsdwl", "17track", "sfgjdl", "360vipwuliu",
                   "yplogistics", "smtgyl"]

# 默认刷新间隔 (秒)
DEFAULT_INTERVAL = 480  # 8 分钟


class TabKeepAlive:
    """后台守护线程，周期性刷新物流标签页防止 session 超时。

    每次刷新创建短暂 CDP 连接，刷新后立即断开，不占用主查询的连接。
    """

    def __init__(self, host: str = "localhost", port: int = 9222,
                 domains: list[str] | None = None,
                 interval: float = DEFAULT_INTERVAL):
        self._host = host
        self._port = port
        self._domains = domains or DEFAULT_DOMAINS
        self._interval = interval
        self._thread: threading.Thread | None = None
        self._stop = threading.Event()

    @classmethod
    def start(cls, host: str = "localhost", port: int = 9222,
              domains: list[str] | None = None,
              interval: float = DEFAULT_INTERVAL) -> "TabKeepAlive":
        """工厂方法：创建并启动守护。"""
        keeper = cls(host, port, domains, interval)
        keeper._start_thread()
        return keeper

    def _start_thread(self):
        self._thread = threading.Thread(target=self._run, daemon=True,
                                        name="tab-keepalive")
        self._thread.start()
        logger.info("TabKeepAlive started (interval=%ds, domains=%s)",
                     self._interval, self._domains)

    def stop(self):
        """停止守护线程。"""
        self._stop.set()
        logger.info("TabKeepAlive stopped")

    def _run(self):
        while not self._stop.wait(self._interval):
            try:
                self._touch_tabs()
            except Exception:
                pass  # 保活是最佳努力，失败不崩

    def _touch_tabs(self):
        """连接每个物流标签页并做软刷新，保持会话活跃。"""
        from .cdp_client import CdpClient

        # 1. 通过 HTTP 列出所有标签页
        try:
            raw = urlopen(f"http://{self._host}:{self._port}/json", timeout=5)
            tabs = json.loads(raw.read())
        except (URLError, OSError, json.JSONDecodeError):
            return

        # 2. 找出物流标签页
        logistics_tabs = []
        for t in tabs:
            url = t.get("url", "")
            if t.get("type") == "page" and any(d in url for d in self._domains):
                logistics_tabs.append(t)

        if not logistics_tabs:
            return

        # 3. 逐个触摸（短暂连接 → 刷新 → 断开）
        touched = 0
        for t in logistics_tabs:
            ws_url = t.get("webSocketDebuggerUrl", "")
            if not ws_url:
                continue
            try:
                cdp = CdpClient(self._host, self._port, timeout=8)
                cdp.connect_tab(ws_url)
                cdp.evaluate("location.reload()")
                cdp.close()
                touched += 1
                time.sleep(2)  # 页面加载间隙
            except Exception:
                continue

        if touched:
            logger.debug("TabKeepAlive: refreshed %d logistics tabs", touched)

    # ── HTTP 直查 (独立进程模式) ──

    @staticmethod
    def status(host: str = "localhost", port: int = 9222) -> dict:
        """查询 Chrome CDP 状态，返回 {running: bool, logistics_tabs: int}。"""
        try:
            raw = urlopen(f"http://{host}:{port}/json", timeout=3)
            tabs = json.loads(raw.read())
            logistics = sum(
                1 for t in tabs
                if t.get("type") == "page" and any(
                    d in (t.get("url") or "") for d in DEFAULT_DOMAINS
                )
            )
            return {"running": True, "total_tabs": len(tabs),
                    "logistics_tabs": logistics}
        except Exception as e:
            return {"running": False, "error": str(e)}


# ── CLI ─────────────────────────────────────────────────────────────

def main():
    import argparse
    parser = argparse.ArgumentParser(description="标签页保活守护")
    parser.add_argument("--daemon", action="store_true",
                        help="独立进程模式，持续运行直到 Ctrl+C")
    parser.add_argument("--interval", type=int, default=DEFAULT_INTERVAL,
                        help=f"刷新间隔秒数 (默认 {DEFAULT_INTERVAL})")
    parser.add_argument("--status", action="store_true",
                        help="输出 CDP 状态 JSON 后退出")
    parser.add_argument("--host", default="localhost")
    parser.add_argument("--port", type=int, default=9222)
    args = parser.parse_args()

    if args.status:
        print(json.dumps(TabKeepAlive.status(args.host, args.port),
                         ensure_ascii=False))
        return

    if args.daemon:
        print(f"TabKeepAlive daemon starting (interval={args.interval}s)...")
        keeper = TabKeepAlive.start(args.host, args.port, interval=args.interval)
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            keeper.stop()
            print("Daemon stopped.")
    else:
        print("Usage: python -m src.keepalive --daemon [--interval 480]")
        print("       python -m src.keepalive --status")


if __name__ == "__main__":
    main()
