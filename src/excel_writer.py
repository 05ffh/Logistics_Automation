"""将物流轨迹信息写回 Excel Y 列（物流轨迹1）。"""

from __future__ import annotations

import shutil
import tempfile
from pathlib import Path

import openpyxl

COL_TRACKING_INFO = 24  # Y列-物流轨迹1 (0-indexed)


def write_results(excel_path: str | Path, results: list[dict]) -> dict:
    """将查询结果写回 Excel，写入前自动备份。

    Returns:
        {updated: int, errors: int}
    """
    excel_path = Path(excel_path)

    # 写入前备份
    backup_path = excel_path.with_name(f"{excel_path.stem}_备份{excel_path.suffix}")
    shutil.copy2(excel_path, backup_path)

    # 检查文件是否被占用
    try:
        with tempfile.NamedTemporaryFile(dir=excel_path.parent, delete=False) as f:
            f.write(b"lock_test")
        Path(f.name).unlink()
    except PermissionError:
        return {"updated": 0, "errors": 0, "locked": True}

    wb = openpyxl.load_workbook(excel_path)
    updated = 0
    errors = 0

    by_sheet: dict[str, list[dict]] = {}
    for r in results:
        by_sheet.setdefault(r["sheet"], []).append(r)

    for sheet_name, rows in by_sheet.items():
        if sheet_name not in wb.sheetnames:
            errors += len(rows)
            continue

        ws = wb[sheet_name]
        for r in rows:
            try:
                col_idx = COL_TRACKING_INFO + 1
                ws.cell(row=r["row_num"], column=col_idx).value = r["routing_info"]
                updated += 1
            except Exception:
                errors += 1

    wb.save(excel_path)
    wb.close()
    return {"updated": updated, "errors": errors, "backup": str(backup_path)}
