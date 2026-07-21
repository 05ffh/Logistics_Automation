"""数据录入模块 - 解析半结构化文本并追加写入 Excel。

用户A每天将用户B发送的物流信息粘贴到 OpenClaw，
自动解析并按日期排序插入到指定 Excel 文件。

用法:
    python -m src.data_entry <excel_path> [--stdin]
"""

from __future__ import annotations

import re
import shutil
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HEADER_ROW = 2


# ── 列模板 ──────────────────────────────────────────────────────

def ensure_template(ws) -> None:
    """一次性初始化列模板（幂等：已存在则跳过）。

    插入/重命名列后，列宽会随表头名一并迁移，公式引用也会修正。
    """
    from openpyxl.utils import get_column_letter

    def _find(hdr: str) -> int | None:
        for c in range(1, ws.max_column + 1):
            if str(ws.cell(row=HEADER_ROW, column=c).value or "").strip() == hdr:
                return c
        return None

    # 保存变更前列宽：{表头: 宽度}
    _widths: dict[str, float] = {}
    for c in range(1, ws.max_column + 1):
        h = str(ws.cell(row=HEADER_ROW, column=c).value or "").strip()
        cd = ws.column_dimensions.get(get_column_letter(c))
        if cd and cd.width and h:
            _widths[h] = cd.width

    # 保存插入前合并单元格（排除 Row 1 分区标签）
    _saved_merges: list[dict] = []
    for mc in ws.merged_cells.ranges:
        if mc.min_row == 1 and mc.max_row == 1:
            continue
        _saved_merges.append({
            "min_row": mc.min_row, "max_row": mc.max_row,
            "min_col": mc.min_col, "max_col": mc.max_col,
            "value": ws.cell(row=mc.min_row, column=mc.min_col).value,
        })

    # 记录插入操作 [{at, amount}, ...]
    inserts: list[dict] = []

    # 标准表头样式：等线 11pt 加粗白色 + 深灰底 + 细线边框
    _hdr_font = Font(name="等线", size=11, bold=True, color="FFFFFFFF")
    _hdr_fill = PatternFill(patternType="solid", fgColor="FF5A5A5A")
    _hdr_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    def _style_hdr(col: int) -> None:
        c = ws.cell(row=HEADER_ROW, column=col)
        c.font = _hdr_font
        c.fill = _hdr_fill
        c.border = _hdr_border
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 1. 成品编码后插入"货物名称"
    base = _find("成品编码")
    if base and not _find("货物名称"):
        ws.insert_cols(base + 1, 1)
        ws.cell(row=HEADER_ROW, column=base + 1).value = "货物名称"
        _style_hdr(base + 1)
        inserts.append({"at": base + 1, "amount": 1})

    # 2. 箱数量后插入箱规(长)/箱规(宽)/箱规(高)/重量
    base2 = _find("箱数量")
    if base2 and not _find("箱规(长)"):
        ws.insert_cols(base2 + 1, 4)
        for i, hdr in enumerate(["箱规(长)", "箱规(宽)", "箱规(高)", "重量"]):
            ws.cell(row=HEADER_ROW, column=base2 + 1 + i).value = hdr
            _style_hdr(base2 + 1 + i)
        inserts.append({"at": base2 + 1, "amount": 4})

    # 3-4. 重命名
    _rename(ws, "箱数量", "箱内数量")
    _rename(ws, "发船时间", "发车、发船时间")
    _rename(ws, "配送时段", "发车、发船后配送时段")
    for old_h, new_h in [("箱数量", "箱内数量"), ("发船时间", "发车、发船时间"),
                          ("配送时段", "发车、发船后配送时段")]:
        if old_h in _widths:
            _widths[new_h] = _widths.pop(old_h)

    # 列宽随表头迁移
    for c in range(1, ws.max_column + 1):
        h = str(ws.cell(row=HEADER_ROW, column=c).value or "").strip()
        cl = get_column_letter(c)
        if h and h in _widths:
            ws.column_dimensions[cl].width = _widths[h]
        elif h and cl in ws.column_dimensions:
            del ws.column_dimensions[cl]
    for hdr in ["货物名称", "箱规(长)", "箱规(宽)", "箱规(高)", "重量"]:
        col = _find(hdr)
        if col:
            ws.column_dimensions[get_column_letter(col)].width = 10

    # 修正 insert_cols 导致错位的合并单元格
    if inserts and _saved_merges:
        _fix_merged_cells(ws, inserts, _saved_merges)
    _fix_section_merges(ws)

    # 修正公式引用
    if inserts:
        _fix_formulas(ws, inserts)


def _rename(ws, old: str, new: str) -> None:
    for c in range(1, ws.max_column + 1):
        if str(ws.cell(row=HEADER_ROW, column=c).value or "").strip() == old:
            ws.cell(row=HEADER_ROW, column=c).value = new
            return


def _fix_merged_cells(ws, inserts: list[dict], saved: list[dict]) -> None:
    """修正 insert_cols 后所有合并单元格的列引用。

    使用插入前保存的数据，按列位移重建。
    """
    inserts_sorted = sorted(inserts, key=lambda x: x["at"])
    inserts_orig = []
    for ins in inserts_sorted:
        at_orig = ins["at"]
        for prev in inserts_orig:
            if prev["at"] <= at_orig:
                at_orig -= prev["amount"]
        inserts_orig.append({"at": at_orig, "amount": ins["amount"]})
    inserts_orig.sort(key=lambda x: x["at"])

    def _shift_col(orig_col: int) -> int:
        col = orig_col
        for ins in inserts_orig:
            if orig_col >= ins["at"]:
                col += ins["amount"]
        return col

    # 清除旧合并（排除 Row 1）
    ws.merged_cells.ranges = {
        mc for mc in ws.merged_cells.ranges
        if mc.min_row == 1 and mc.max_row == 1
    }

    # 按新列位重建
    for s in saved:
        new_min = _shift_col(s["min_col"])
        new_max = _shift_col(s["max_col"])
        cl_min = get_column_letter(new_min)
        cl_max = get_column_letter(new_max)
        ws.merge_cells(f"{cl_min}{s['min_row']}:{cl_max}{s['max_row']}")
        if s["value"] is not None:
            ws.cell(row=s["min_row"], column=new_min).value = s["value"]


def _fix_section_merges(ws) -> None:
    """修正 Row 1 的分区合并单元格（运营填写 / 仓库填写）。

    运营填写：A列 ~ 发货公司列（含）
    仓库填写：预计发货时间列 ~ 最后一列
    """
    from openpyxl.utils import get_column_letter
    end_ops = None
    start_wh = None
    for c in range(1, ws.max_column + 1):
        h = str(ws.cell(row=HEADER_ROW, column=c).value or "").strip()
        if h == "发货公司":
            end_ops = c
        if h == "预计发货时间" and start_wh is None:
            start_wh = c

    # 清理旧的 Row 1 合并单元格（直接操作 ranges 集合，避免 unmerge_cells 的 insert_cols 后遗症）
    ws.merged_cells.ranges = {
        mc for mc in ws.merged_cells.ranges
        if not (mc.min_row == 1 and mc.max_row == 1)
    }

    # 重建
    if end_ops:
        ops_end = get_column_letter(end_ops)
        ws.merge_cells(f"A1:{ops_end}1")
        ws.cell(row=1, column=1).value = "运营填写"
    if start_wh:
        wh_start = get_column_letter(start_wh)
        wh_end = get_column_letter(ws.max_column)
        ws.merge_cells(f"{wh_start}1:{wh_end}1")
        ws.cell(row=1, column=start_wh).value = "仓库填写"

    # 设置分区标签样式
    _section_font = Font(name="等线", size=11, bold=True, color="FFFFFFFF")
    _section_fill = PatternFill(patternType="solid", fgColor="FF5A5A5A")
    _section_align = Alignment(horizontal="center", vertical="center")
    for c in [1, start_wh]:
        if c:
            cell = ws.cell(row=1, column=c)
            cell.font = _section_font
            cell.fill = _section_fill
            cell.alignment = _section_align


def _fix_formulas(ws, inserts: list[dict]) -> None:
    """修正所有公式中的列引用，补偿 openpyxl insert_cols 不更新引用的缺陷。

    inserts: [{at, amount}, ...] — 按执行顺序记录的插入操作（at 为当时列号）。
    需要将所有 at 归一化到插入前的原始列号再计算累计位移。
    """
    import re

    # 归一化 at 到原始列号：后续插入的 at 受之前插入影响，要减回去
    inserts_orig = []
    for ins in inserts:
        at_orig = ins["at"]
        # 之前的插入如果在 at 左侧，会把 at 向右推
        for prev in inserts_orig:
            if prev["at"] <= at_orig:
                at_orig -= prev["amount"]
        inserts_orig.append({"at": at_orig, "amount": ins["amount"]})
    inserts_orig.sort(key=lambda x: x["at"])

    def _new_col(orig_col: int) -> int:
        col = orig_col
        for ins in inserts_orig:
            if orig_col >= ins["at"]:
                col += ins["amount"]
        return col

    from openpyxl.utils import get_column_letter, column_index_from_string
    # 列引用边界：左侧不能是字母/下划线，右侧不能是字母数字/下划线
    # 避免误伤 DISPIMG 等函数中嵌入的 GUID（如 "C20967A"）
    cell_ref = re.compile(r"(?<![A-Za-z_])(\$?[A-Z]{1,3})(\$?\d+)(?![A-Za-z0-9_])")

    for row_idx in range(3, ws.max_row + 1):
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if not isinstance(cell.value, str) or not cell.value.startswith("="):
                continue

            def _fix(m):
                cl = m.group(1)
                rn = m.group(2)
                try:
                    old_c = column_index_from_string(cl)
                    new_c = _new_col(old_c)
                except ValueError:
                    return m.group(0)
                return (get_column_letter(new_c) + rn) if new_c != old_c else m.group(0)

            cell.value = cell_ref.sub(_fix, cell.value)


# ── 解析 ────────────────────────────────────────────────────────

# 标签名 → 内部 key（顺序不敏感）
_LABEL_MAP = {
    "货物名称": "product",
    "渠道": "channel",
    "数量": "quantity",
    "箱规": "box_spec",
    "货件编号": "fba_code",
    "配送地址": "address",
    "发车、发船时间": "ship_date_text",
    "发车。发船后提取时间": "delivery_text",
    "发车、发船后提取时间": "delivery_text",
    "单价价格": "price",
    "有无附加（多少）": "surcharge",
}


def parse_entry(text: str) -> dict | None:
    """解析一段物流信息文本，返回结构化字段 dict。"""
    lines = [ln.strip() for ln in text.strip().split("\n") if ln.strip()]
    if len(lines) < 3:
        return None

    entry: dict = {}

    # 第 1 行：日期
    entry["date"] = _parse_date_line(lines[0])

    # 第 2 行：公司名或店铺名（带 - 国家后缀 → 发货店铺）
    line2 = lines[1].strip()
    if re.match(r"^[一-鿿]+-[A-Z]{2}$", line2):
        entry["store"] = line2
    else:
        entry["company"] = line2

    # 余下行：标签键值对
    for ln in lines[2:]:
        for label, key in _LABEL_MAP.items():
            if label in ln:
                # 提取标签后的值
                idx = ln.index(label) + len(label)
                val = ln[idx:].lstrip("：: ").strip()
                if val:
                    entry[key] = val
                break

    if "date" not in entry:
        return None
    return entry


def _parse_date_line(text: str) -> datetime:
    """'7月15日' → datetime(2026, 7, 15)。"""
    m = re.match(r"(\d{1,2})\s*月\s*(\d{1,2})\s*日?", text)
    if m:
        return datetime(2026, int(m.group(1)), int(m.group(2)))
    # 尝试直接解析
    for fmt in ["%Y-%m-%d", "%m/%d", "%m月%d日"]:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    today = datetime.now()
    return datetime(today.year, today.month, today.day)


# ── 格式化 ──────────────────────────────────────────────────────

def _fmt_ship_date(text: str) -> str:
    """'7-23号左右' → '7月23日左右'（固定后缀）"""
    text = text.strip()
    m = re.match(r"(\d{1,2})\s*[-–]\s*(\d{1,2})", text)
    if m:
        return f"{int(m.group(1))}月{int(m.group(2))}日左右"
    return text


def _fmt_delivery(text: str) -> str:
    """'40-50派送' → '40-50自然日'"""
    text = text.strip()
    m = re.match(r"(\d+\s*[-–]\s*\d+)", text)
    if m:
        return f"{m.group(1).replace(' ', '')}自然日"
    return text


def _fmt_price(text: str) -> str:
    """'6.5' → '6.5/KG'"""
    text = text.strip()
    if not text:
        return ""
    if "/KG" not in text.upper() and "/kg" not in text:
        return f"{text}/KG"
    return text


def _fmt_surcharge(text: str) -> str:
    """'0' → '' (跳过), 非0照填"""
    text = text.strip()
    try:
        val = float(text)
        if val == 0:
            return ""
    except ValueError:
        pass
    return text


def _extract_warehouse(address: str) -> str:
    """提取仓库编号。
    'DTM2-Kaltband...' → 'DTM2'
    'Amazon - HAJ1 Zur Alten... (HAJ1)' → 'HAJ1'
    """
    address = address.strip()
    # 末尾括号标注 → 取括号内容
    m = re.search(r"\(([A-Z0-9]{2,6})\)", address)
    if m:
        return m.group(1)
    # 最前面的编号（字母数字开头，后跟 -）
    m = re.match(r"^([A-Z0-9]{2,6})-", address)
    if m:
        return m.group(1)
    # 直接取第一个词
    return address.split()[0].strip("-")


def _extract_box_spec(text: str) -> dict:
    """'59*44*55cm，重21kg' → {box_l: '59', box_w: '44', box_h: '55', weight: '21'}"""
    result = {"box_l": "", "box_w": "", "box_h": "", "weight": ""}
    # 尺寸
    dims = re.match(r"(\d+)\s*\*\s*(\d+)\s*\*\s*(\d+)", text)
    if dims:
        result["box_l"] = dims.group(1)
        result["box_w"] = dims.group(2)
        result["box_h"] = dims.group(3)
    # 重量
    w = re.search(r"重\s*(\d+(?:\.\d+)?)\s*kg", text, re.IGNORECASE)
    if w:
        result["weight"] = w.group(1)
    return result


# ── 写入 ────────────────────────────────────────────────────────

def _col_map(ws) -> dict[str, int]:
    """返回 {内部key: 列号}，基于当前表头。"""
    header_to_key = {
        "货物名称": "product",
        "发货渠道": "channel",
        "箱数": "quantity",
        "箱规(长)": "box_l",
        "箱规(宽)": "box_w",
        "箱规(高)": "box_h",
        "重量": "weight",
        "箱内数量": "box_inner_qty",
        "货件号": "fba_code",
        "仓库": "warehouse",
        "发货店铺": "store",
        "发货公司": "company",
        "发车、发船时间": "ship_date",
        "发车、发船后配送时段": "delivery",
        "价格": "price",
        "附加费": "surcharge",
    }
    mapping = {}
    for c in range(1, ws.max_column + 1):
        h = str(ws.cell(row=HEADER_ROW, column=c).value or "").strip()
        if h in header_to_key:
            mapping[header_to_key[h]] = c
    return mapping


def _find_insert_row(ws, target_date: datetime) -> int:
    """在 A 列中找到 ≤ target_date 的最后一行，返回其后行号（即插入位）。

    空行跳过不中断，因为中间可能有空白分隔行。
    """
    last_match = HEADER_ROW
    for r in range(3, ws.max_row + 1):
        cell = ws.cell(row=r, column=1).value
        dt = None
        if isinstance(cell, datetime):
            dt = cell
        elif isinstance(cell, (int, float)):
            from datetime import timedelta
            dt = datetime(1899, 12, 30) + timedelta(days=int(cell))
        if dt is None:
            continue
        if dt.date() <= target_date.date():
            last_match = r
    return last_match + 1


def insert_entry(excel_path: str | Path, entry: dict) -> dict:
    """将一条解析后的条目插入 Excel，返回结果。

    1. 备份原文件
    2. 确保列模板就绪
    3. 按日期找到插入行
    4. 插入新行，填入数据
    5. 保存
    """
    excel_path = Path(excel_path)
    backup_path = excel_path.with_name(f"{excel_path.stem}_备份{excel_path.suffix}")
    shutil.copy2(excel_path, backup_path)

    wb = openpyxl.load_workbook(excel_path)

    # 取第一个数字命名的 sheet
    sheet_name = _find_data_sheet(wb)
    if not sheet_name:
        wb.close()
        return {"error": "No data sheet found"}
    ws = wb[sheet_name]

    # 列模板
    ensure_template(ws)

    # 重建列映射（模板可能变了）
    cols = _col_map(ws)

    # 解析并格式化字段
    entry_date = entry.get("date")
    if isinstance(entry_date, datetime):
        date_val = entry_date
    else:
        date_val = datetime.now()

    raw_data = _build_row_data(entry)

    # 找插入位置
    insert_row = _find_insert_row(ws, date_val)

    # 插入行
    ws.insert_rows(insert_row, 1)
    _copy_row_format(ws, insert_row - 1, insert_row)

    # ── 填入数据 + 匹配原文件格式 ──
    # 标准格式：等线 11pt，水平垂直居中，自动换行，细线边框
    std_font = Font(name="等线", size=11)
    std_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    std_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    # 发货渠道：等线 11pt，加粗，红色
    channel_font = Font(name="等线", size=11, bold=True, color="FFC00000")
    # 货件号：Arial 10.5pt
    fba_font = Font(name="Arial", size=10.5)

    special_fonts = {
        cols.get("channel"): channel_font,
        cols.get("fba_code"): fba_font,
    }

    def _style_cell(cell, font_=None):
        cell.font = font_ or std_font
        cell.alignment = std_align
        cell.border = std_border

    # 填 A 列日期
    a_cell = ws.cell(row=insert_row, column=1)
    a_cell.value = date_val
    a_cell.number_format = 'm"月"d"日";@'
    _style_cell(a_cell)

    # 填写其他列
    filled = []
    for key, col in cols.items():
        val = raw_data.get(key, "")
        if not val:
            continue
        cell = ws.cell(row=insert_row, column=col)
        cell.value = val
        _style_cell(cell, special_fonts.get(col))
        filled.append(f"{key}={val}")

    # 整行所有单元格补齐边框（含空白列）
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=insert_row, column=col)
        if cell.border.left.style is None:
            cell.border = std_border

    try:
        wb.save(excel_path)
    except (PermissionError, OSError) as e:
        wb.close()
        return {"error": str(e), "backup": str(backup_path)}
    wb.close()

    return {
        "sheet": sheet_name,
        "row": insert_row,
        "date": date_val.strftime("%Y-%m-%d"),
        "filled": filled,
        "backup": str(backup_path),
    }


def _build_row_data(entry: dict) -> dict:
    """从解析结果构建写入数据，含格式化转换。"""
    data = {}

    # 公司 / 店铺（互斥，二选一）
    data["company"] = entry.get("company", "")
    data["store"] = entry.get("store", "")

    # 货物名称
    data["product"] = entry.get("product", "")

    # 渠道
    data["channel"] = entry.get("channel", "")

    # 数量 → 提取数字
    qty = entry.get("quantity", "")
    m = re.search(r"(\d+)", str(qty))
    data["quantity"] = m.group(1) if m else qty

    # 箱规 → 拆成 4 个字段
    spec = _extract_box_spec(entry.get("box_spec", ""))
    data["box_l"] = spec["box_l"] + "cm" if spec["box_l"] else ""
    data["box_w"] = spec["box_w"] + "cm" if spec["box_w"] else ""
    data["box_h"] = spec["box_h"] + "cm" if spec["box_h"] else ""
    data["weight"] = spec["weight"] + "KG" if spec["weight"] else ""

    # 货件编号
    data["fba_code"] = entry.get("fba_code", "")

    # 配送地址 → 仓库
    addr = entry.get("address", "")
    data["warehouse"] = _extract_warehouse(addr) if addr else ""

    # 发车、发船时间
    ship = entry.get("ship_date_text", "")
    data["ship_date"] = _fmt_ship_date(ship) if ship else ""

    # 配送时段
    delivery = entry.get("delivery_text", "")
    data["delivery"] = _fmt_delivery(delivery) if delivery else ""

    # 价格
    price = entry.get("price", "")
    data["price"] = _fmt_price(price) if price else ""

    # 附加费
    surcharge = entry.get("surcharge", "")
    data["surcharge"] = _fmt_surcharge(surcharge) if surcharge else ""

    return data


def _copy_row_format(ws, src_row: int, dst_row: int) -> None:
    """复制行格式（行高、字体对齐等）。"""
    src_dim = ws.row_dimensions.get(src_row)
    if src_dim:
        ws.row_dimensions[dst_row].height = src_dim.height


def _find_data_sheet(wb) -> str | None:
    """找到第一个以数字命名的 sheet。"""
    for sn in wb.sheetnames:
        if sn.isdigit() or any(c.isdigit() for c in sn):
            return sn
    return wb.sheetnames[0] if wb.sheetnames else None


# ── CLI ─────────────────────────────────────────────────────────

def main():
    import sys
    if len(sys.argv) < 2:
        print("Usage: python -m src.data_entry <excel_path> [--stdin]")
        print("       Reads text from stdin and appends to Excel.")
        sys.exit(1)

    excel_path = sys.argv[1]
    text = sys.stdin.read().strip()

    if not text:
        print("No input text provided.")
        sys.exit(1)

    entry = parse_entry(text)
    if not entry:
        print("Failed to parse input text.")
        sys.exit(1)

    result = insert_entry(excel_path, entry)
    if "error" in result:
        print(f"Error: {result['error']}")
        sys.exit(1)

    print(f"Inserted at Sheet={result['sheet']} Row={result['row']} Date={result['date']}")
    print(f"Fields: {', '.join(result['filled'])}")


if __name__ == "__main__":
    main()
