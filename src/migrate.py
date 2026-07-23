"""将旧格式 Excel 迁移到 2026 发货信息表最终版规范。

用法:
  python -m src.migrate <旧文件路径> [--output <输出路径>]
"""

from __future__ import annotations

import re
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

HEADER_ROW = 2

# 旧列号(1-based) → 新列号(1-based)
_COL_MAP: dict[int, int | None] = {
    1: 1,   2: 2,   3: 3,   4: 4,   5: 5,   6: 6,   7: 7,   8: 8,
    9: 10, 10: 17, 11: 18, 12: 19, 13: 20, 14: 21, 15: 22,
    16: 24, 17: None, 18: 26, 19: 27, 20: 28, 21: 29, 22: 30,
    23: 31, 24: 32, 25: 33, 26: 34, 27: 35, 28: 36,
}

_NEW_HEADERS = [
    "发表日期", "图片", "品名", "asin", "sku", "fnsku", "发货店铺",
    "发货数量", "指定发货渠道", "备注", "箱数", "箱内数量", "箱规(长)",
    "箱规(宽)", "箱规(高)", "重量", "实际发货渠道", "发货公司",
    "仓库发货时间", "发车、发船时间", "时效", "价格", "附加费",
    "条码确认", "仓库", "货件号", "物流单号", "状态", "实发",
    "实发差值", "已收", "实收差值", "物流轨迹1", "物流轨迹2",
    "后台送达时段", "更新日期",
]

NCOL = len(_NEW_HEADERS)

# ── 统一样式 ──
_DATA_FONT = Font(name="等线", size=11)
_DATA_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
_ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

# 左对齐列（长文本，需自动换行）
_LEFT_ALIGN = {10}  # J=备注
# 图片列（DISPIMG，非文本）
_IMAGE_COL = 2


def migrate(old_path: str | Path, output_path: str | Path | None = None) -> Path:
    old_path = Path(old_path)
    if output_path is None:
        output_path = old_path.with_name(f"{old_path.stem}_规范版{old_path.suffix}")
    else:
        output_path = Path(output_path)

    wb_old = openpyxl.load_workbook(old_path)
    wb_new = openpyxl.Workbook()
    first = True

    for sn in wb_old.sheetnames:
        ws_new = wb_new.active if first else wb_new.create_sheet(title=sn)
        if first:
            ws_new.title = sn
            first = False
        _migrate_sheet(wb_old[sn], ws_new)

    wb_new.save(output_path)
    wb_new.close()
    wb_old.close()
    print(f"Done: {output_path}")
    return output_path


def _migrate_sheet(ws_old, ws_new) -> None:
    _write_headers(ws_new)
    _write_section_labels(ws_new)

    for r in range(3, ws_old.max_row + 1):
        _migrate_row(ws_old, ws_new, r)

    _apply_data_format(ws_new)
    _auto_column_widths(ws_new)
    _auto_row_heights(ws_new)


def _write_headers(ws) -> None:
    hdr_font = Font(name="等线", size=11, bold=True, color="FFFFFFFF")
    hdr_fill = PatternFill(patternType="solid", fgColor="FF5A5A5A")
    hdr_align = Alignment(horizontal="center", vertical="center")
    for c, h in enumerate(_NEW_HEADERS, 1):
        cell = ws.cell(row=HEADER_ROW, column=c, value=h)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.border = _DATA_BORDER
        cell.alignment = hdr_align


def _write_section_labels(ws) -> None:
    sec_font = Font(name="等线", size=11, bold=True, color="FFFFFFFF")
    sec_fill = PatternFill(patternType="solid", fgColor="FF5A5A5A")
    sec_align = Alignment(horizontal="center", vertical="center")
    ws.merge_cells("A1:R1")
    c = ws.cell(row=1, column=1, value="运营填写")
    c.font = sec_font; c.fill = sec_fill; c.alignment = sec_align
    ws.merge_cells("S1:AJ1")
    c = ws.cell(row=1, column=19, value="仓库填写")
    c.font = sec_font; c.fill = sec_fill; c.alignment = sec_align


def _migrate_row(ws_old, ws_new, old_row: int) -> None:
    r = old_row
    for old_col, new_col in _COL_MAP.items():
        if new_col is None:
            continue
        old_val = ws_old.cell(row=r, column=old_col).value
        if old_val is None:
            continue

        if old_col == 1:          # 日期
            ws_new.cell(row=r, column=new_col, value=_convert_date(old_val))
        elif old_col == 15:       # 价格拆分 → V(22) + W(23)
            price, sur = _split_price(str(old_val))
            ws_new.cell(row=r, column=new_col, value=price)
            if sur != 0:
                ws_new.cell(row=r, column=23, value=sur)
        else:
            ws_new.cell(row=r, column=new_col, value=old_val)


def _apply_data_format(ws) -> None:
    """统一数据行：字体、细线边框、自动换行、对齐。"""
    for r in range(3, ws.max_row + 1):
        for c in range(1, NCOL + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = _DATA_FONT
            cell.border = _DATA_BORDER
            cell.alignment = _ALIGN_LEFT if c in _LEFT_ALIGN else _ALIGN_CENTER


def _auto_column_widths(ws) -> None:
    """按表头+内容自适应列宽，DISPIMG 公式不参与计算。"""
    # 基础宽度: 表头汉字 × 2.2
    col_widths = {}
    for c, h in enumerate(_NEW_HEADERS, 1):
        col_widths[c] = len(h) * 2.2

    for r in range(3, ws.max_row + 1):
        for c in range(1, NCOL + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            text = str(v)
            if text.startswith("=_xlfn.DISPIMG"):
                continue
            for line in text.split("\n"):
                w = sum(2 if ord(ch) > 127 else 1 for ch in line)
                if w > col_widths.get(c, 8):
                    col_widths[c] = w

    for c, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c)].width = min(w + 2, 40)


def _auto_row_heights(ws) -> None:
    """根据换行数 + 列宽估算折行，设置自适应行高。"""
    col_w_units = {}
    for c in range(1, NCOL + 1):
        cl = get_column_letter(c)
        col_w_units[c] = (ws.column_dimensions[cl].width or 8)

    for r in range(3, ws.max_row + 1):
        max_lines = 1
        for c in range(1, NCOL + 1):
            v = ws.cell(row=r, column=c).value
            if v is None or c == _IMAGE_COL:
                continue
            text = str(v)
            total = 0
            for line in text.split("\n"):
                # 中文字符≈2单位, ASCII≈1单位
                char_w = sum(2 if ord(ch) > 127 else 1 for ch in line)
                cw = col_w_units.get(c, 8) or 8
                total += max(1, -(-char_w // int(cw)))  # ceil(char_w / cw)
            if total > max_lines:
                max_lines = total
        ws.row_dimensions[r].height = max(max_lines * 15, 20)


# ── 数据转换 ──

def _convert_date(val) -> str | None:
    if isinstance(val, (int, float)):
        dt = datetime(1899, 12, 30) + timedelta(days=int(val))
        return f"{dt.month}月{dt.day}日"
    s = str(val).strip()
    if re.match(r"^\d{4}$", s):
        return f"{int(s[:2])}月{int(s[2:])}日"
    return s


def _split_price(raw: str) -> tuple[float, float]:
    m = re.match(r"^([\d.]+)\+(\d+)$", raw.strip())
    if m:
        return float(m.group(1)), float(m.group(2))
    try:
        return float(raw.strip()), 0.0
    except ValueError:
        return 0.0, 0.0


def main():
    import argparse
    parser = argparse.ArgumentParser(description="迁移旧格式 Excel → 最终版规范")
    parser.add_argument("source", help="旧格式 Excel 路径")
    parser.add_argument("--output", "-o", default=None, help="输出路径")
    args = parser.parse_args()
    migrate(args.source, args.output)


if __name__ == "__main__":
    main()
