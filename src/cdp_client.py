"""CDP 客户端 - 通过 Chrome DevTools Protocol 操控 Edge 浏览器。"""

from __future__ import annotations

import base64
import hashlib
import json
import os
import socket
import struct
import time
from urllib.error import URLError
from urllib.parse import urlparse
from urllib.request import urlopen


class CdpClient:
    """Chrome DevTools Protocol 客户端。"""

    def __init__(self, host: str = "localhost", port: int = 9222, timeout: float = 15.0):
        self.host = host
        self.port = port
        self.timeout = timeout
        self._sock: socket.socket | None = None
        self._msg_id = 0

    # ── connection ────────────────────────────────────────────────

    def list_tabs(self) -> list[dict]:
        """列出所有打开的标签页。"""
        raw = urlopen(f"http://{self.host}:{self.port}/json", timeout=self.timeout)
        return json.load(raw)

    def connect_tab(self, ws_url: str) -> None:
        """通过 WebSocket URL 连接到一个标签页。"""
        self.close()
        self._msg_id = 0  # 重置消息计数器（避免跨适配器污染）
        parsed = urlparse(ws_url)
        host = parsed.hostname or self.host
        port = parsed.port or self.port
        path = parsed.path or "/"
        if parsed.query:
            path = f"{path}?{parsed.query}"

        key = base64.b64encode(os.urandom(16)).decode("ascii")
        req = (
            f"GET {path} HTTP/1.1\r\n"
            f"Host: {host}:{port}\r\n"
            "Upgrade: websocket\r\n"
            "Connection: Upgrade\r\n"
            f"Sec-WebSocket-Key: {key}\r\n"
            "Sec-WebSocket-Version: 13\r\n"
            "\r\n"
        ).encode("ascii")

        expected_accept = base64.b64encode(
            hashlib.sha1((key + "258EAFA5-E914-47DA-95CA-C5AB0DC85B11").encode()).digest()
        ).decode("ascii")

        sock = socket.create_connection((host, port), timeout=self.timeout)
        sock.settimeout(self.timeout)
        sock.sendall(req)

        response = b""
        while b"\r\n\r\n" not in response:
            response += sock.recv(4096)
            if not response:
                raise ConnectionError("empty WebSocket handshake response")

        headers = response.decode("iso-8859-1", errors="replace")
        if " 101 " not in headers or expected_accept not in headers:
            raise ConnectionError(f"WebSocket handshake failed: {headers[:200]}")

        self._sock = sock
        self._msg_id = 0

    def close(self) -> None:
        """关闭连接。"""
        if self._sock:
            try:
                self._sock.close()
            except OSError:
                pass
            self._sock = None

    # ── CDP commands ──────────────────────────────────────────────

    def evaluate(self, expression: str, await_promise: bool = False) -> dict:
        """执行 JavaScript 表达式并返回结果。"""
        return self._send({
            "method": "Runtime.evaluate",
            "params": {
                "expression": expression,
                "returnByValue": True,
                "awaitPromise": await_promise,
            },
        })

    def fetch_api(self, api_url: str, timeout: float = 15.0) -> dict:
        """在浏览器内通过 fetch() 调用 API，自动携带 Cookie/会话。

        类似象往项目的 CdpClient.fetch_api()，利用浏览器标签页已有的
        登录态发起请求，从服务器角度看与页面自身的 AJAX 请求无法区分。
        """
        js = (
            "(async()=>{"
            "try{"
            f"var r=await fetch('{api_url}',{{headers:{{Accept:'application/json'}}}});"
            "var d=await r.json();"
            "return JSON.stringify({ok:r.ok,status:r.status,data:d});"
            "}catch(e){"
            "return JSON.stringify({ok:false,error:e.message});"
            "}"
            "})()"
        )
        deadline = time.time() + timeout
        last_err = None
        while time.time() < deadline:
            result = self.evaluate(js, await_promise=True)
            value = result.get("result", {}).get("result", {}).get("value")
            if value is None:
                last_err = result.get("result", {}).get("result", {}).get("description",
                           result.get("error", "no value"))
                time.sleep(0.5)
                continue
            try:
                return json.loads(value)
            except (json.JSONDecodeError, TypeError):
                return {"ok": False, "error": str(value)[:200]}
        return {"ok": False, "error": f"timeout: {last_err}"}

    def capture_network(self, url_substring: str, timeout: float = 10.0) -> list[dict]:
        """启用 Network 域，导航后捕获匹配 URL 的请求/响应。

        返回 [{url, method, status, responseBody}] 列表。
        """
        self._send({"method": "Network.enable", "params": {"maxTotalBufferSize": 10000000}})
        # 等待一段时间收集请求
        time.sleep(timeout)
        results: list[dict] = []
        # 读取积压的 WebSocket 消息
        old_timeout = self._sock.gettimeout() if self._sock else None
        try:
            if self._sock:
                self._sock.settimeout(0.3)
            for _ in range(200):
                try:
                    msg = _read_ws_message(self._sock)
                except (ConnectionError, BlockingIOError, TimeoutError):
                    break
                method = msg.get("method", "")
                if method == "Network.responseReceived":
                    params = msg.get("params", {})
                    resp = params.get("response", {})
                    url = resp.get("url", "")
                    if url_substring in url:
                        req_id = params.get("requestId", "")
                        results.append({
                            "url": url,
                            "method": resp.get("method", ""),
                            "status": resp.get("status", 0),
                            "mimeType": resp.get("mimeType", ""),
                            "requestId": req_id,
                        })
        finally:
            if self._sock and old_timeout is not None:
                self._sock.settimeout(old_timeout)
        return results

    def click(self, selector: str) -> dict:
        """点击匹配选择器的第一个元素。"""
        return self.evaluate(
            f"const el=document.querySelector('{selector}');"
            f"if(el){{el.click();'clicked'}}else{{'not found: {selector}'}}"
        )

    def type_text(self, selector: str, text: str) -> dict:
        """向输入框填入文本（兼容 React/Vue 的 v-model）。"""
        escaped = text.replace("\\", "\\\\").replace("'", "\\'").replace("\n", "\\n")
        return self.evaluate(
            f"(function(){{"
            f"var el=document.querySelector('{selector}');"
            f"if(!el)return 'not found: {selector}';"
            f"var native=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value');"
            f"if(native&&native.set){{native.set.call(el,'{escaped}');}}else{{el.value='{escaped}';}}"
            f"el.dispatchEvent(new Event('input',{{bubbles:true}}));"
            f"el.dispatchEvent(new Event('change',{{bubbles:true}}));"
            f"return 'typed';"
            f"}})()"
        )

    def wait_for_selector(self, selector: str, timeout_ms: int = 5000, interval_ms: int = 200) -> bool:
        """等待选择器匹配到元素。"""
        deadline = time.time() + timeout_ms / 1000
        while time.time() < deadline:
            result = self.evaluate(
                f"document.querySelector('{selector}') !== null"
            )
            if result.get("result", {}).get("result", {}).get("value"):
                return True
            time.sleep(interval_ms / 1000)
        return False

    def get_text(self, selector: str) -> str:
        """获取元素的文本内容。"""
        result = self.evaluate(
            f"const el=document.querySelector('{selector}');"
            f"el?el.textContent.trim():''"
        )
        return result.get("result", {}).get("result", {}).get("value", "")

    # ── WebSocket protocol ────────────────────────────────────────

    def _send(self, message: dict) -> dict:
        if not self._sock:
            raise RuntimeError("not connected")
        self._msg_id += 1
        message["id"] = self._msg_id
        payload = json.dumps(message, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
        self._sock.sendall(_make_ws_frame(payload))

        # 收集响应直到收到匹配的 id
        while True:
            response = _read_ws_message(self._sock)
            if response.get("id") == self._msg_id:
                return response
            # ignore other messages (e.g. Page.frameStoppedLoading)


# ── WebSocket framing helpers ─────────────────────────────────────

def _make_ws_frame(payload: bytes) -> bytes:
    header = bytearray([0x81])  # text frame, FIN
    length = len(payload)
    if length < 126:
        header.append(0x80 | length)
    elif length <= 0xFFFF:
        header.append(0x80 | 126)
        header.extend(struct.pack(">H", length))
    else:
        header.append(0x80 | 127)
        header.extend(struct.pack(">Q", length))
    mask = os.urandom(4)
    masked = bytes(b ^ mask[i % 4] for i, b in enumerate(payload))
    return bytes(header) + mask + masked


def _read_ws_message(sock: socket.socket) -> dict:
    first, second = _recv_exact(sock, 2)
    length = second & 0x7F
    if length == 126:
        length = struct.unpack(">H", _recv_exact(sock, 2))[0]
    elif length == 127:
        length = struct.unpack(">Q", _recv_exact(sock, 8))[0]
    if second & 0x80:
        _recv_exact(sock, 4)  # skip mask bytes (only present on client→server frames)
    payload = _recv_exact(sock, length) if length else b""
    try:
        return json.loads(payload.decode("utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError):
        return {}


def _recv_exact(sock: socket.socket, size: int) -> bytes:
    chunks = []
    remaining = size
    while remaining:
        chunk = sock.recv(remaining)
        if not chunk:
            raise ConnectionError("websocket closed")
        chunks.append(chunk)
        remaining -= len(chunk)
    return b"".join(chunks)
