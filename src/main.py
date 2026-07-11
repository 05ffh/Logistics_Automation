"""物流轨迹查询 - 主流程（多公司支持）。

用法:
    python -m src.main <excel_path>              # 处理所有 sheet + 所有公司
    python -m src.main <excel_path> 202605       # 只处理指定 sheet
"""

from __future__ import annotations

import os
import re
import sys
from pathlib import Path

try:
    from .cdp_client import CdpClient
    from .companies.ningzhi import NingZhiAdapter
    from .companies.yuntuo import YunTuoAdapter
    from .excel_reader import find_company_rows, company_position
    from .excel_writer import write_results
except ImportError:
    from cdp_client import CdpClient
    from companies.ningzhi import NingZhiAdapter
    from companies.yuntuo import YunTuoAdapter
    from excel_reader import find_company_rows, company_position
    from excel_writer import write_results

# 注册所有公司适配器
ADAPTERS = [NingZhiAdapter(), YunTuoAdapter()]


def main():
    if len(sys.argv) < 2:
        print("Usage: python -m src.main <excel_path> [sheet_names]")
        sys.exit(1)

    excel_path = Path(sys.argv[1])
    if not excel_path.exists():
        print(f"Excel file not found: {excel_path}")
        sys.exit(1)

    target_sheets = None
    if len(sys.argv) > 2:
        target_sheets = set(s.strip() for s in sys.argv[2].split(","))

    # 1. 连接 CDP
    cdp_host = os.environ.get("CDP_HOST", "localhost:9222")
    host, _, port_str = cdp_host.partition(":")
    port = int(port_str) if port_str else 9222

    print(f"Connecting to CDP at {host}:{port}...")
    cdp = CdpClient(host=host, port=port)
    try:
        cdp.list_tabs()
    except Exception:
        print("ERROR: Cannot reach Edge CDP. Is Edge running with --remote-debugging-port=9222?")
        print("Double-click 物流网站一键启动.bat to start Edge first.")
        sys.exit(1)

    # 2. 读取 Excel
    print(f"Reading Excel: {excel_path}")
    company_configs = [{"name": a.name, "prefix": a.prefix} for a in ADAPTERS]
    all_rows = find_company_rows(excel_path, company_configs)
    if target_sheets:
        all_rows = [r for r in all_rows if r["sheet"] in target_sheets]

    if not all_rows:
        print("No matching rows found for any company.")
        cdp.close()
        return

    # 按公司分组
    by_company: dict[str, list[dict]] = {}
    for r in all_rows:
        by_company.setdefault(r["company"], []).append(r)

    adapter_map = {a.name: a for a in ADAPTERS}

    # 3. 逐公司查询
    all_results: dict[str, dict[str, str | None]] = {}  # {company: {tn: routing}}

    for company_name, rows in by_company.items():
        adapter = adapter_map[company_name]
        print(f"\n{'='*40}")
        print(f"Company: {company_name} ({adapter.prefix}*)")

        # 切换到公司标签页
        ws_url = adapter.ensure_tab(cdp)
        cdp.connect_tab(ws_url)

        # 检查前置条件
        if not adapter.check_ready(cdp):
            print(f"WARNING: {company_name} not ready. Please check the browser tab.")
            continue

        # 收集唯一单号
        tns: list[str] = []
        seen_tns = set()
        for r in rows:
            for tn in r["tracking_nos"]:
                if tn not in seen_tns:
                    seen_tns.add(tn)
                    tns.append(tn)

        print(f"Found {len(rows)} rows, {len(tns)} unique {adapter.prefix}* numbers.")
        results = adapter.query(cdp, tns)
        all_results[company_name] = {r.tracking_no: r.routing_info for r in results}

    # 4. 合并结果并写入
    print(f"\n{'='*40}")
    print("Merging results...")
    updated_count = 0
    preserved_count = 0
    write_payload = []

    # 跨所有公司的本次新查结果 {单号: 轨迹}
    global_results: dict[str, str] = {}
    for res in all_results.values():
        for tn, routing in res.items():
            if routing:
                global_results[tn] = routing

    for row in all_rows:
        company = row["company"]
        existing = row.get("existing_info") or ""
        old_map = _parse_existing_map(existing)
        # 仅本公司的单号（S 列顺序），每家公司独占一列
        my_tns = row.get("tracking_nos") or []

        blocks = []
        for tn in my_tns:
            fresh = global_results.get(tn)
            if fresh:
                blocks.append(f"{tn}\n{fresh}")
                updated_count += 1
            elif old_map.get(tn):
                blocks.append(f"{tn}\n{old_map[tn]}")
                preserved_count += 1

        row["routing_info"] = "\n".join(blocks)
        # 该公司在 S 列首次出现的次序 → 写入第几个物流轨迹列
        row["track_position"] = company_position(
            row.get("all_tracking_nos") or my_tns, company
        )
        write_payload.append(row)

    # 5. 写回 Excel
    print(f"\nRouting: {updated_count} updated, {preserved_count} preserved.")
    print("Writing results back to Excel...")
    write_summary = write_results(excel_path, write_payload)

    if write_summary.get("locked"):
        print("ERROR: Excel file is open. Please close it and retry.")
    elif write_summary.get("backup"):
        print(f"Backup saved: {write_summary['backup']}")
    print(f"Done: {write_summary['updated']} rows updated, {write_summary['errors']} errors.")

    cdp.close()


_TN_PATTERN = re.compile(r"^[A-Z0-9]{5,25}$")


def _extract_all_tns_from_s_column(y_text: str) -> list[str]:
    """从 Y 列文本中逐行提取所有物流单号（保持顺序）。"""
    tns = []
    if not y_text:
        return tns
    for line in y_text.split("\n"):
        stripped = line.strip()
        if _TN_PATTERN.match(stripped):
            tns.append(stripped)
    return tns


def _extract_routing_for_tn(y_text: str, tn: str) -> str:
    """从 Y 列提取指定单号的路由文本（不含单号行本身）。"""
    if not y_text or not tn:
        return ""
    all_tns = _extract_all_tns_from_s_column(y_text)
    try:
        idx = all_tns.index(tn)
    except ValueError:
        return ""
    # 找该单号在原文中的起始位置
    tn_start = y_text.find(tn)
    if tn_start == -1:
        return ""
    # 找下一个单号的位置作为结束边界
    if idx + 1 < len(all_tns):
        next_tn = all_tns[idx + 1]
        tn_end = y_text.find(next_tn, tn_start + len(tn))
    else:
        tn_end = len(y_text)
    routing = y_text[tn_start + len(tn) : tn_end].strip()
    return routing


def _parse_existing_map(y_text: str) -> dict[str, str]:
    """把旧 Y 列解析为 {单号: 轨迹文本}，用于保留其他公司/未刷新的轨迹。"""
    result: dict[str, str] = {}
    for tn in _extract_all_tns_from_s_column(y_text):
        routing = _extract_routing_for_tn(y_text, tn)
        if routing:
            result[tn] = routing
    return result


def _routing_equal(old: str, new: str) -> bool:
    """比较两条路由信息是否实质相同。"""
    return old.strip() == new.strip()


if __name__ == "__main__":
    main()
