"""缺失单号持久化追踪——记账 + 顽固判重 + 精准补跑。

每次运行后，把没查到轨迹的单号记录到 <excel名>_misses.json；
同一单号多次 MISS 判为"顽固"，用 --retry-stubborn 只查顽固单号。
"""

from __future__ import annotations

import json
import os
import tempfile
from datetime import datetime, timezone
from pathlib import Path

MISS_THRESHOLD = 2  # 顽固 = 缺失 >= 此次数


# ── 文件路径 ──────────────────────────────────────────────────

def get_misses_path(excel_path: Path) -> Path:
    return excel_path.with_name(f"{excel_path.stem}_misses.json")


# ── 读写 ──────────────────────────────────────────────────────

def load_misses(excel_path: Path) -> dict:
    path = get_misses_path(excel_path)
    if not path.exists():
        return _empty()
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if "misses" not in data:
            return _empty()
        return data
    except (json.JSONDecodeError, OSError) as e:
        print(f"WARNING: 无法读取 misses 文件 {path}: {e}，将重建。")
        return _empty()


def save_misses(excel_path: Path, data: dict) -> None:
    path = get_misses_path(excel_path)
    data["last_updated"] = datetime.now(timezone.utc).isoformat()
    with tempfile.NamedTemporaryFile(
        "w", encoding="utf-8", dir=path.parent, delete=False
    ) as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.write("\n")
    os.replace(f.name, path)


# ── 记录 / 更新 / 移除 ────────────────────────────────────────

def _key(entry: dict) -> tuple:
    return (entry["company"], entry["sheet"], entry["row_num"], entry["tn"])


def record_misses(excel_path: Path, entries: list[dict]) -> int:
    """合并新 MISS 条目。已存在的 incremiss_count，新的添加。返回新条目数。"""
    if not entries:
        return 0
    data = load_misses(excel_path)
    existing = {_key(m): m for m in data["misses"]}
    new_count = 0
    now = datetime.now(timezone.utc).isoformat()
    for e in entries:
        k = _key(e)
        if k in existing:
            existing[k]["miss_count"] += 1
            existing[k]["last_missed"] = now
        else:
            e = dict(e)
            e.setdefault("miss_count", 1)
            e.setdefault("first_missed", now)
            e["last_missed"] = now
            data["misses"].append(e)
            existing[k] = e
            new_count += 1
    save_misses(excel_path, data)
    return new_count


def remove_resolved(excel_path: Path, resolved_tns: set[str]) -> int:
    """移除 tn 在 resolved_tns 中的条目（补查成功了的）。返回移除数。"""
    if not resolved_tns:
        return 0
    data = load_misses(excel_path)
    before = len(data["misses"])
    data["misses"] = [m for m in data["misses"] if m["tn"] not in resolved_tns]
    removed = before - len(data["misses"])
    if removed:
        save_misses(excel_path, data)
    return removed


# ── 查询 ──────────────────────────────────────────────────────

def get_stubborn(excel_path: Path, threshold: int = MISS_THRESHOLD) -> list[dict]:
    """返回 miss_count >= threshold 的条目，按缺失次数降序。"""
    data = load_misses(excel_path)
    result = [m for m in data["misses"] if m.get("miss_count", 0) >= threshold]
    result.sort(key=lambda m: m["miss_count"], reverse=True)
    return result


def print_miss_summary(excel_path: Path, new_count: int = 0, removed_count: int = 0) -> None:
    """打印当前缺失追踪概况。"""
    data = load_misses(excel_path)
    all_misses = data["misses"]
    if not all_misses:
        print("\nMiss Tracking: 无缺失单号记录。")
        return
    stubborn = get_stubborn(excel_path)
    print(f"\n{'='*40}")
    print("Miss Tracking:")
    print(f"  本次新 MISS : {new_count}")
    print(f"  本次补回    : {removed_count}")
    print(f"  总计追踪中  : {len(all_misses)}")
    print(f"  顽固(≥{MISS_THRESHOLD}次): {len(stubborn)}")
    if stubborn:
        for s in stubborn[:8]:
            print(f"    [{s['company']}] {s['tn']} ×{s['miss_count']} "
                  f"({s['sheet']} row{s['row_num']})")
        if len(stubborn) > 8:
            print(f"    ... 还有 {len(stubborn)-8} 个")
    print(f"  文件: {get_misses_path(excel_path)}")


# ── helpers ───────────────────────────────────────────────────

def _empty() -> dict:
    return {"last_updated": "", "misses": []}
