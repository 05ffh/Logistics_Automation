"""云驼物流适配器 - 17track.net。

逐单查询策略：批量模式对"愿景征途"自动识别几乎不命中（"冷"状态），
批量 → 回退的双阶段反而不如逐单高效。改为逐单查询，每条含选运输商兜底。
"""

from __future__ import annotations

import json
import random
import re
import time

from .base import CompanyAdapter, TrackingResult

try:
    from ..validation import is_valid_routing
    from ..cdp_util import val as _val
except ImportError:
    from validation import is_valid_routing
    from cdp_util import val as _val

MAIN_URL = "https://www.17track.net/zh-cn"
RESULT_URL = "https://t.17track.net/zh-cn#nums="
CARRIER_NAME = "愿景征途"
# 象往式礼貌间隔：逐单之间随机 3-6s，避免连续请求触发安全验证
QUERY_INTERVAL_MIN = 3.0
QUERY_INTERVAL_MAX = 6.0


class YunTuoAdapter(CompanyAdapter):
    name = "云驼"
    prefix = "999"
    batch_size = 1
    canary_number = "999260530000730"  # 自检用已知单号（失效时更新）

    def query(self, cdp, tracking_nos: list[str]) -> list[TrackingResult]:
        results: dict[str, str | None] = {}
        total = len(tracking_nos)

        for i, tn in enumerate(tracking_nos):
            if i > 0:
                interval = random.uniform(QUERY_INTERVAL_MIN, QUERY_INTERVAL_MAX)
                time.sleep(interval)

            routing = self._query_one(cdp, tn)
            results[tn] = routing
            status = "OK" if routing else "MISS"
            print(f"  [{self.name}] {i+1}/{total} {tn} {status}")

        ok = sum(1 for tn in tracking_nos if results.get(tn))
        print(f"  [{self.name}] 合计 {ok}/{total} OK")
        return [TrackingResult(tn, results.get(tn)) for tn in tracking_nos]

    # ── 批量查询 ──────────────────────────────────────────────

    def _query_batch(self, cdp, nums: list[str]) -> dict[str, str]:
        """批量提交并提取，返回 {单号: 时间戳+描述}（仅含自动识别成功的）。"""
        # 确保页面 UI 就绪（textarea + 按钮已渲染），避免 React 未挂载导致填入静默失败
        if not self._page_stable(cdp, timeout=8):
            return {}
        if not self._fill_batch(cdp, nums):
            return {}
        time.sleep(0.6)
        if not self._click_search(cdp):
            return {}
        self._wait_batch(cdp, len(nums), timeout=max(15, len(nums)))
        return self._extract_batch(cdp)

    def _fill_batch(self, cdp, nums: list[str]) -> bool:
        """把多个单号按行填入 textarea。"""
        arr = ",".join("'" + n + "'" for n in nums)
        r = cdp.evaluate(
            "(function(){"
            "var tas=document.querySelectorAll('textarea');var ta=null;"
            "for(var i=0;i<tas.length;i++){var ph=tas[i].placeholder||'';"
            "if(ph.indexOf('每行输入')>=0||tas[i].id==='auto-size-textarea'){ta=tas[i];break;}}"
            "if(!ta&&tas.length)ta=tas[0];if(!ta)return 'no';"
            "ta.focus();"
            "var d=Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype,'value');"
            f"d.set.call(ta,[{arr}].join(String.fromCharCode(10)));"
            "ta.dispatchEvent(new Event('input',{bubbles:true}));"
            "return 'ok';})()"
        )
        return _val(r) == "ok"

    def _wait_batch(self, cdp, n: int, timeout: int = 20) -> None:
        """轮询等待 n 个单号的卡片全部渲染。"""
        deadline = time.time() + timeout
        while time.time() < deadline:
            r = cdp.evaluate(
                "(function(){"
                "var cards=document.querySelectorAll('[data-state]');var s={},c=0;"
                "for(var i=0;i<cards.length;i++){"
                "var m=(cards[i].innerText||'').match(/999\\d{10,}/);"
                "if(m&&!s[m[0]]){s[m[0]]=1;c++;}}"
                "return c;})()"
            )
            if _val(r, 0) >= n:
                time.sleep(1)  # 结算
                return
            time.sleep(0.5)

    def _extract_batch(self, cdp) -> dict[str, str]:
        """遍历所有结果卡片，提取 {单号: 最新时间戳+描述}。"""
        r = cdp.evaluate(
            "(function(){"
            "var cards=document.querySelectorAll('[data-state]');"
            "var seen={},res={};"
            "for(var i=0;i<cards.length;i++){"
            "var t=(cards[i].innerText||'').trim();"
            "var nm=t.match(/999\\d{10,}/);if(!nm)continue;"
            "var num=nm[0];if(seen[num])continue;"
            "var lines=t.split(String.fromCharCode(10)).map(function(s){return s.trim();})"
            ".filter(function(s){return s;});"
            "var ts=null,desc=null;"
            "for(var j=0;j<lines.length;j++){"
            "if(/^\\d{4}-\\d{2}-\\d{2}\\s+\\d{2}:\\d{2}$/.test(lines[j])){"
            "ts=lines[j];desc=lines[j+1]||'';break;}}"
            "if(ts){seen[num]=1;res[num]=ts+String.fromCharCode(10)+desc;}}"
            "return JSON.stringify(res);})()"
        )
        try:
            raw = json.loads(_val(r, "{}"))
        except (ValueError, TypeError):
            return {}
        # 校验：拦掉页面结构变化抓到的垃圾
        return {tn: v for tn, v in raw.items() if is_valid_routing(v)}

    def ensure_tab(self, cdp) -> str:
        tabs = cdp.list_tabs()
        for t in tabs:
            if t.get("type") == "page" and "17track" in (t.get("url") or ""):
                return t.get("webSocketDebuggerUrl", "")
        page_tabs = [t for t in tabs if t.get("type") == "page"]
        if not page_tabs:
            raise RuntimeError("No browser tabs.")
        cdp.connect_tab(page_tabs[0]["webSocketDebuggerUrl"])
        cdp.evaluate(f"window.open('{MAIN_URL}', '_blank')")
        cdp.close()
        time.sleep(2)
        for t in cdp.list_tabs():
            if t.get("type") == "page" and "17track" in (t.get("url") or ""):
                return t.get("webSocketDebuggerUrl", "")
        raise RuntimeError("Cannot open 17track tab.")

    def _query_one(self, cdp, tracking_no: str) -> str | None:
        # 0. 确保在结果页壳（提供可复用的 textarea + 查询按钮）
        url = _val(cdp.evaluate("location.href"), "")
        if "t.17track.net" not in url:
            cdp.evaluate(f"location.href='{RESULT_URL}{tracking_no}';")
            time.sleep(6)

        # 1. 填入单号
        if not self._fill_number(cdp, tracking_no):
            return None
        time.sleep(0.6)

        # 2. 点查询
        if not self._click_search(cdp):
            return None

        # 3. 轮询等待新结果加载（时间线出现 或 弹出运输商候选），避免时序竞态
        state = self._wait_result(cdp, timeout=12)
        # 页面可能处于 mid-load 状态（loading 提示出现但内容未就绪），等它安定
        if not self._page_stable(cdp, tracking_no, timeout=5):
            return None

        # 4. 自动识别成功 → 直接提取
        routing = self._extract_routing(cdp, tracking_no)
        if routing:
            return routing

        # 5. 需手动选运输商 → 点"愿景征途"
        if state == "carrier" or self._select_carrier(cdp):
            self._wait_result(cdp, timeout=8, want="timeline")
            time.sleep(1)
            return self._extract_routing(cdp, tracking_no)

        return None

    # ── 步骤实现 ──────────────────────────────────────────────

    def _fill_number(self, cdp, tracking_no: str) -> bool:
        """填入单号到结果页 textarea（React 受控组件用原生 setter）。"""
        r = cdp.evaluate(
            "(function(){"
            "var tas=document.querySelectorAll('textarea');"
            "var ta=null;"
            "for(var i=0;i<tas.length;i++){"
            "var ph=tas[i].placeholder||'';"
            "if(ph.indexOf('每行输入')>=0||tas[i].id==='auto-size-textarea'){ta=tas[i];break;}}"
            "if(!ta&&tas.length)ta=tas[0];"
            "if(!ta)return 'no';"
            "ta.focus();"
            "var d=Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype,'value');"
            f"d.set.call(ta,'{tracking_no}');"
            "ta.dispatchEvent(new Event('input',{bubbles:true}));"
            "return 'ok';"
            "})()"
        )
        return _val(r) == "ok"

    def _click_search(self, cdp) -> bool:
        """点击"查询(N)"按钮（结果页），回退到主页搜索区域。"""
        r = cdp.evaluate(
            "(function(){"
            "var btns=document.querySelectorAll('button');"
            "for(var i=0;i<btns.length;i++){"
            "var t=(btns[i].textContent||'').trim();"
            "if(t.indexOf('查询(')===0){btns[i].click();return 'result';}}"
            "var a=document.querySelector('[class*=batch_track_search-area]');"
            "if(a){a.click();return 'main';}"
            "return 'no';"
            "})()"
        )
        return _val(r) in ("result", "main")

    def _wait_result(self, cdp, timeout: int = 12, want: str = "any") -> str:
        """轮询等待结果加载。

        返回 'timeline'（时间线已出现）/ 'carrier'（弹出运输商候选）/ 'timeout'。
        want='timeline' 时只等时间线（用于选完运输商后）。
        """
        deadline = time.time() + timeout
        while time.time() < deadline:
            r = cdp.evaluate(
                "(function(){"
                "var b=document.body.innerText||'';"
                "if(b.indexOf('同步时间')>=0)return 'timeline';"
                "if(b.indexOf('选择运输商')>=0||b.indexOf('可能的运输商')>=0)return 'carrier';"
                "return 'wait';"
                "})()"
            )
            state = _val(r, "wait")
            if state == "timeline":
                return "timeline"
            if state == "carrier" and want != "timeline":
                return "carrier"
            time.sleep(0.5)
        return "timeout"

    def _page_stable(self, cdp, tracking_no: str = "", timeout: int = 5) -> bool:
        """确认页面真正就绪：body 不含 loading 提示。

        17track SPA 在 loading 阶段可能触发 _wait_result 的"同步时间"匹配
        但实际内容尚未渲染。此方法轮询确认页面已脱离 loading 状态。
        若 tracking_no 非空，额外要求该单号已出现在 body 中。
        """
        target_check = f"var hasTarget=b.indexOf('{tracking_no}')>=0;" if tracking_no else "var hasTarget=true;"
        deadline = time.time() + timeout
        while time.time() < deadline:
            r = cdp.evaluate(
                "(function(){"
                "var b=document.body.innerText||'';"
                "var loading=b.indexOf('check the logistics trajectory')>=0;"
                + target_check +
                "if(!loading&&hasTarget)return 'ok';"
                "if(loading)return 'loading';"
                "return 'wait';"
                "})()"
            )
            state = _val(r, "wait")
            if state == "ok":
                return True
            time.sleep(0.4)
        return False

    def _select_carrier(self, cdp) -> bool:
        """在运输商候选中点击含"愿景征途"的 LI 行。"""
        r = cdp.evaluate(
            "(function(){"
            "var spans=document.querySelectorAll('span');"
            "var span=null;"
            "for(var i=0;i<spans.length;i++){"
            f"if((spans[i].textContent||'').trim()==='{CARRIER_NAME}'){{span=spans[i];break;}}}}"
            "if(!span)return 'no-carrier';"
            "var e=span;"
            "while(e&&e.tagName!=='LI')e=e.parentElement;"
            "if(!e){e=span;e.click();return 'clicked-span';}"
            "e.click();return 'clicked-li';"
            "})()"
        )
        return _val(r) in ("clicked-li", "clicked-span")

    def _extract_routing(self, cdp, tracking_no: str) -> str | None:
        """从"同步时间:"之后的时间线提取最新一条（时间戳 + 描述）。"""
        r = cdp.evaluate(
            "(function(){return (document.body.innerText||'').substring(0,6000);})()"
        )
        body = _val(r, "")
        if not body:
            return None
        return _parse_yt_routing(body)


# ── 解析 ──────────────────────────────────────────────────────

# 时间线时间戳格式 YYYY-MM-DD HH:mm（不含秒；同步时间含秒会被排除）
_TS = re.compile(
    r"(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2})(?!:)\s+(.*?)"
    r"(?=\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}(?!:)|$)",
    re.DOTALL,
)


def _parse_yt_routing(body: str) -> str | None:
    # 锚定到"同步时间: ... (GMT...)"之后，跳过顶部摘要区
    anchor = body.find("同步时间")
    if anchor >= 0:
        gmt = body.find(")", anchor)
        section = body[gmt + 1:] if gmt > anchor else body[anchor:]
    else:
        section = body

    matches = _TS.findall(section)
    if not matches:
        return None
    timestamp, desc = matches[0]
    desc = _clean(desc)
    if not desc:
        return None
    result = f"{timestamp}\n{desc}"
    return result if is_valid_routing(result) else None


def _clean(text: str) -> str:
    lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
    # 去掉界面噪声行
    noise = {"我没收到货", "签收时间:", "FAQ>", "轨迹信息", "复制详细", "复制链接", "更多信息"}
    lines = [ln for ln in lines if ln not in noise]
    return " ".join(lines)


