"""物流轨迹查询 - 主流程。

用法:
    python -m src.main <excel_path>              # 处理所有 sheet
    python -m src.main <excel_path> 202605       # 只处理指定 sheet
    python -m src.main <excel_path> 202605,202606  # 处理多个 sheet
"""

from __future__ import annotations

import os
import re
import sys
from pathlib import Path

try:
    from .cdp_client import CdpClient
    from .excel_reader import find_nz_rows
    from .excel_writer import write_results
    from .nzhexp_tracker import check_logged_in, ensure_nzhexp_tab, query_tracking
except ImportError:
    from cdp_client import CdpClient
    from excel_reader import find_nz_rows
    from excel_writer import write_results
    from nzhexp_tracker import check_logged_in, ensure_nzhexp_tab, query_tracking


def main():
    if len(sys.argv) < 2:
        print("Usage: python -m src.main <excel_path> [sheet_names]")
        print("  sheet_names: comma-separated, or omit for all sheets")
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

    # 2. 确保 nzhexp 标签页就绪
    print("Finding nzhexp tab...")
    ws_url = ensure_nzhexp_tab(cdp)
    cdp.connect_tab(ws_url)

    # 3. 检查登录态
    print("Checking login status...")
    if not check_logged_in(cdp):
        print("ERROR: nzhexp is not logged in.")
        print("Please log in to nzhexp in the Edge window, then retry.")
        cdp.close()
        sys.exit(1)

    # 4. 读取 Excel
    print(f"Reading Excel: {excel_path}")
    rows = find_nz_rows(excel_path)
    if target_sheets:
        rows = [r for r in rows if r["sheet"] in target_sheets]

    if not rows:
        print("No 宁致 rows found.")
        cdp.close()
        return

    total_tracking_nos = sum(len(r["tracking_nos"]) for r in rows)
    print(f"Found {len(rows)} rows with {total_tracking_nos} NZ tracking numbers.")

    # 5. 收集所有单号（去重）
    all_nos: list[str] = []
    seen = set()
    for r in rows:
        for tn in r["tracking_nos"]:
            if tn not in seen:
                seen.add(tn)
                all_nos.append(tn)

    # 6. 批量查询 + 增量写入
    print(f"Querying {len(all_nos)} unique tracking numbers (batch size: 5)...")
    results = query_tracking(cdp, all_nos)

    tn_to_routing = {r.tracking_no: r.routing_info for r in results}
    updated_count = 0
    unchanged_count = 0

    write_payload = []
    for row in rows:
        existing = row.get("existing_info") or ""
        all_tns_in_row = _extract_all_tns_from_s_column(existing)

        # 收集该行所有单号的 routing（优先用 CDP 新结果，其次保留原文）
        merged: dict[str, str] = {}

        for tn in all_tns_in_row:
            if tn.startswith("NZ") and tn in tn_to_routing and tn_to_routing[tn]:
                merged[tn] = tn_to_routing[tn]
                if tn in (row.get("tracking_nos") or []):
                    updated_count += 1
            else:
                old_routing = _extract_routing_for_tn(existing, tn)
                merged[tn] = old_routing
                if tn.startswith("NZ"):
                    unchanged_count += 1

        # 追加新出现的 NZ 单号（原 Y 列中完全没有的）
        for tn in row["tracking_nos"]:
            if tn not in merged:
                new_info = tn_to_routing.get(tn)
                if new_info:
                    merged[tn] = new_info
                    updated_count += 1

        row["routing_info"] = "\n".join(f"{tn}\n{routing}" for tn, routing in merged.items())
        write_payload.append(row)

    # 7. 写回 Excel
    print(f"\nRouting: {updated_count} updated, {unchanged_count} unchanged.")
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


def _routing_equal(old: str, new: str) -> bool:
    """比较两条路由信息是否实质相同。"""
    return old.strip() == new.strip()


if __name__ == "__main__":
    main()
