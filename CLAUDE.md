# CLAUDE.md

## 项目概述

物流轨迹自动查询系统。读取共享 Excel 中的发货明细，通过 CDP 操控 Edge 在宁致物流网站上批量查询运单轨迹，将最新路由信息写回 Excel。

**对比象往**: 物流项目不需要 MySQL，流程是 Excel → 查网站 → 写回 Excel。

## 核心架构

```
同事的 Windows 电脑:
├── .bat (启动 Edge + CDP 9222 + 8 个物流网站)
└── OpenClaw + logistics-track Skill
    ├── 读 Excel → 筛选宁致行 → 提取 NZ 单号
    ├── CDP → localhost:9222 → nzhexp 查询
    └── 写回 Excel Y 列
```

## 列位映射

| 列 | 内容 | 用途 |
|---|---|---|
| J (index 9) | 发货渠道 | 可能含"宁致"（不规范） |
| K (index 10) | 发货公司 | 应含"宁致"（规范） |
| S (index 18) | 物流单号 | 多单号换行分隔，取NZ前缀 |
| Y (index 24) | 物流轨迹1 | 回写最新路由信息 |

## 项目结构

```
Logistics_Automation/
├── bin/
│   └── 物流网站一键启动.bat
├── src/
│   ├── excel_reader.py      # 读取 Excel，筛选行，提取单号
│   ├── cdp_tracker.py       # CDP 操控 nzhexp 查询轨迹
│   ├── excel_writer.py      # 写回 Excel Y 列
│   └── main.py              # 主流程编排
├── skill/
│   └── logistics-track/
│       └── SKILL.md         # OpenClaw Skill 定义
└── requirements.txt
```

## 筛选逻辑

```
遍历 Sheet 每行:
  J列 (index 9) 或 K列 (index 10) 包含 "宁致" → 命中
  从 S列 (index 18) 拆出所有单号 → 过滤 NZ 前缀
  → 这些是宁致的物流单号
```

## 平台差异

| | WSL (开发) | Windows (生产) |
|---|---|---|
| CDP 地址 | `172.28.190.60:9222` | `localhost:9222` |
| Excel 路径 | `/mnt/c/Users/.../` | `C:\Users\...\` |
| Python | WSL 内置 | OpenClaw 内置 |

`CDP_HOST` 环境变量控制 CDP 地址，默认 `localhost:9222`，WSL 下设为 Windows IP。
