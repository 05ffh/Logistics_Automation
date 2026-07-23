# CLAUDE.md

## 项目概述

物流轨迹自动查询系统。读取共享 Excel 中的发货明细，通过 CDP 操控 Edge 浏览器在物流网站查询运单轨迹，将最新路由信息按公司写回"物流轨迹N"列。

当前支持三家公司：宁致(NZ)、云驼(999)、小满(XM)。采用适配器模式，可扩展。

## 四大模块

| 模块 | 入口 | 功能 |
|------|------|------|
| 物流轨迹查询 | `python -m src.main` | CDP 查物流 → 写回轨迹列 + 缺失追踪 |
| 数据录入 | `python -m src.data_entry` | IM 文本解析 → 按日期插入 Excel |
| ASIN 图片匹配 | `python -m src.image_inserter build/insert` | ASIN→图片库 → 嵌入 B 列 |
| 格式迁移 | `python -m src.migrate` | 旧规范 Excel → 新规范列位映射 |

## 核心架构

```
同事的 Windows 电脑:
├── .bat (启动 Edge + CDP 9222 + 物流网站标签页)
└── OpenClaw + logistics-track Skill
    ├── 读 Excel → 按表头自动匹配列位 → 按前缀归属公司
    ├── CDP → localhost:9222 → 逐公司查询
    │   宁致/小满: fetch API 调内部 JSON 接口 (~0.2s/单号)
    │   云驼: DOM 逐单查询 + 单条回退选运输商
    └── 按 track_position 写回对应物流轨迹N列
```

## 查询策略

| 公司 | 方式 | 说明 |
|------|------|------|
| 宁致 | fetch API | 浏览器内 fetch() 调 `/tracking/app?inajax=1&tracking_number=NZ...` |
| 小满 | fetch API | 同上，调 `xmsdwl.nextsls.com` 同一端点 |
| 云驼 | DOM 逐单 | 17track 无内部 API，保留 DOM 方式 |

fetch API 策略借鉴象往项目：fetch 在浏览器内执行，携带完整 Cookie/会话，
从服务器角度看与页面自身的 AJAX 请求无法区分，零 bot 检测风险。

## 筛选逻辑 (当前)

```
遍历数字命名的 Sheet 每一行:
  按第 2 行表头找到"物流单号"列
  从该列拆出所有单号（换行分隔）
  按前缀归属公司: NZ→宁致, 999→云驼, XM→小满, HY→华洋, HYC→华运昌
  → 不依赖 J/K 列（业务填写不规范）
```

## 列位映射

列位通过第 2 行表头文字自动匹配，不再硬编码索引：
- "物流单号" → 提取单号的来源列
- "物流轨迹1/2/N" → 回写轨迹的目标列

找不到表头时回退到历史默认值（物流单号=S/col19，物流轨迹1=Y/col25）。

每家发货公司独占一个"物流轨迹N"列，N = 该公司单号在物流单号列首次出现的次序。
缺列时紧跟最后一个物流轨迹列后 insert_cols 插入。

## 项目结构

```
Logistics_Automation/
├── bin/
│   └── 物流网站一键启动.bat
├── images/
│   └── products/             # ASIN 图片库
├── src/
│   ├── cdp_client.py        # CDP WebSocket + fetch_api()
│   ├── cdp_util.py           # CDP 工具函数 (val)
│   ├── data_entry.py         # 半结构化物流文本解析 + 自动填入 Excel
│   ├── excel_reader.py       # 读取 + 表头自动匹配 + 前缀归属
│   ├── excel_writer.py       # 按公司写物流轨迹N列 + 备份
│   ├── image_inserter.py     # ASIN 图片库构建 + 自动嵌入图片
│   ├── migrate.py            # 旧格式 → 新规范列位迁移
│   ├── validation.py         # 轨迹数据校验 is_valid_routing
│   ├── miss_tracker.py       # 缺失单号追踪 + 顽固补跑
│   ├── main.py               # 主流程编排 + healthcheck + retry-stubborn
│   └── companies/
│       ├── base.py           # CompanyAdapter 抽象基类
│       ├── ningzhi.py        # 宁致 NZ → fetch API
│       ├── yuntuo.py         # 云驼 999 → DOM 逐单
│       └── xiaoman.py        # 小满 XM → fetch API
├── skill/logistics-track/SKILL.md
└── requirements.txt
```

## 脚本

```bash
# 物流轨迹查询
python -m src.main <excel_path> [sheet_names]
python -m src.main <excel_path> --company 小满,宁致
python -m src.main --healthcheck
python -m src.main <excel_path> --retry-stubborn

# 数据录入
python -m src.data_entry <excel_path>

# ASIN 图片匹配
python -m src.image_inserter build <ASIN映射Excel>
python -m src.image_inserter insert <目标Excel>

# 旧格式迁移
python -m src.migrate <旧格式Excel> -o <输出路径>
```

## 平台差异

| | WSL (开发) | Windows (生产) |
|---|---|---|
| CDP 地址 | `172.28.190.60:9222` | `localhost:9222` |
| Excel 路径 | `/mnt/c/Users/.../` | `C:\Users\...\` |
| Python 命令 | `python3` | `python` |
| 编码 | UTF-8 (原生) | GBK → stdout 强制 UTF-8 |

`CDP_HOST` 环境变量控制 CDP 地址，默认 `localhost:9222`。

## 稳健性设计

- **数据校验**: is_valid_routing 拦截页面改版产生的垃圾数据
- **漏查不覆盖**: merge_preserve 按单号合并新旧，本次未查到保留旧值
- **异常检测**: ≥5 单且成功率 <50% → 跳过写入保护存量
- **金丝雀自检**: --healthcheck 用已知单号预验证各站点
- **缺失追踪**: _misses.json 记录缺失，miss_count≥2 判顽固
- **自动备份**: 每次写入前自动备份 Excel
