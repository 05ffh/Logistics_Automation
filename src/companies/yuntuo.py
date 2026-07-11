"""云驼物流适配器 - 17track.net。

17track 结果页 textarea 原生支持批量（每行一个单号，最多 40 个），所以策略为:
    1. 批量填入 ≤40 个单号 → 点"查询(N)"按钮
    2. 等所有卡片渲染，遍历 [data-state] 卡片一次性提取 {单号: 最新时间戳+描述}
    3. 对未命中的单号（需手动选运输商 / 加载慢）回退到单条 _query_one（含选"愿景征途"）
"""

from __future__ import annotations

import json
import re
import time

from .base import CompanyAdapter, TrackingResult

MAIN_URL = "https://www.17track.net/zh-cn"
RESULT_URL = "https://t.17track.net/zh-cn#nums="
CARRIER_NAME = "愿景征途"
MAX_BATCH = 40  # 17track 单次提交上限


class YunTuoAdapter(CompanyAdapter):
    name = "云驼"
    prefix = "999"
    batch_size = MAX_BATCH

    def query(self, cdp, tracking_nos: list[str]) -> list[TrackingResult]:
        results: dict[str, str | None] = {}
        total = len(tracking_nos)

        # 确保结果页壳（提供可复用的 textarea + 查询按钮）
        url = _val(cdp.evaluate("location.href"), "")
        if "t.17track.net" not in url and tracking_nos:
            cdp.evaluate(f"location.href='{RESULT_URL}{tracking_nos[0]}';")
            time.sleep(6)

        # 1. 分批批量查询
        for start in range(0, total, MAX_BATCH):
            batch = tracking_nos[start:start + MAX_BATCH]
            found = self._query_batch(cdp, batch)
            for tn in batch:
                if found.get(tn):
                    results[tn] = found[tn]
            hit = sum(1 for tn in batch if results.get(tn))
            print(f"  [{self.name}] 批量 {start+1}-{start+len(batch)}/{total}: "
                  f"命中 {hit}/{len(batch)}")

        # 2. 回退：未命中的单号逐个查（处理选运输商 / 慢加载）
        misses = [tn for tn in tracking_nos if not results.get(tn)]
        if misses:
            print(f"  [{self.name}] 回退单条查询 {len(misses)} 个未命中...")
            for i, tn in enumerate(misses):
                results[tn] = self._query_one(cdp, tn)
                st = "OK" if results.get(tn) else "MISS"
                print(f"  [{self.name}] 回退 {i+1}/{len(misses)} {tn} {st}")

        ok = sum(1 for tn in tracking_nos if results.get(tn))
        print(f"  [{self.name}] 合计 {ok}/{total} OK")
        return [TrackingResult(tn, results.get(tn)) for tn in tracking_nos]

    # ── 批量查询 ──────────────────────────────────────────────

    def _query_batch(self, cdp, nums: list[str]) -> dict[str, str]:
        """批量提交并提取，返回 {单号: 时间戳+描述}（仅含自动识别成功的）。"""
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
            return json.loads(_val(r, "{}"))
        except (ValueError, TypeError):
            return {}

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

        # 4. 自动识别成功 → 直接提取
        routing = self._extract_routing(cdp, tracking_no)
        if routing:
            return routing

        # 5. 需手动选运输商 → 点"愿景征途"
        if state == "carrier" or self._select_carrier(cdp):
            self._wait_result(cdp, timeout=8, want="timeline")
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
    return f"{timestamp}\n{desc}"


def _clean(text: str) -> str:
    lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
    # 去掉界面噪声行
    noise = {"我没收到货", "签收时间:", "FAQ>", "轨迹信息", "复制详细", "复制链接", "更多信息"}
    lines = [ln for ln in lines if ln not in noise]
    return " ".join(lines)


def _val(cdp_result: dict, default=None):
    return cdp_result.get("result", {}).get("result", {}).get("value", default)
