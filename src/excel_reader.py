"""读取发货明细表 Excel，筛选宁致行，提取 NZ 前缀物流单号。"""

from __future__ import annotations

import re
from pathlib import Path

import openpyxl

# 列位 (0-indexed)
COL_SHIP_CHANNEL = 9   # J列-发货渠道 (可能含"宁致")
COL_SHIP_COMPANY = 10  # K列-发货公司 (应含"宁致")
COL_TRACKING_NOS = 18  # S列-物流单号
COL_TRACKING_INFO = 24 # Y列-物流轨迹1

# 宁致物流单号前缀
NZ_PREFIX = "NZ"


def find_nz_rows(excel_path: str | Path) -> list[dict]:
    """扫描 Excel 所有 sheet，返回宁致相关的行和单号。

    Returns:
        [{sheet, row_num, tracking_nos: [str], existing_info: str|None}, ...]
    """
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    results = []

    for sheet_name in wb.sheetnames:
        if not sheet_name.strip().isdigit():
            continue
        ws = wb[sheet_name]
        for row_idx in range(3, ws.max_row + 1):
            ship_channel = _cell_str(ws, row_idx, COL_SHIP_CHANNEL)
            ship_company = _cell_str(ws, row_idx, COL_SHIP_COMPANY)

            if "宁致" not in ship_channel and "宁致" not in ship_company:
                continue

            tracking_str = _cell_str(ws, row_idx, COL_TRACKING_NOS)
            if not tracking_str:
                continue

            nz_numbers = _extract_nz_numbers(tracking_str)
            if not nz_numbers:
                continue

            existing_info = _cell_str(ws, row_idx, COL_TRACKING_INFO)

            results.append({
                "sheet": sheet_name,
                "row_num": row_idx,
                "tracking_nos": nz_numbers,
                "existing_info": existing_info or None,
            })

    wb.close()
    return results


def _extract_nz_numbers(text: str) -> list[str]:
    """从物流单号字符串中提取 NZ 前缀的单号，保持出现顺序去重。"""
    parts = re.split(r"[\n\r]+", text)
    seen = set()
    result = []
    for p in parts:
        p = p.strip()
        if p.startswith(NZ_PREFIX) and p not in seen:
            seen.add(p)
            result.append(p)
    return result


def _cell_str(ws, row: int, col: int) -> str:
    val = ws.cell(row=row, column=col + 1).value
    if val is None:
        return ""
    if isinstance(val, (int, float)):
        return str(int(val)) if val == int(val) else str(val)
    return str(val).strip()
