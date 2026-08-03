# CLAUDE.md

## 项目概述

物流轨迹自动查询系统。读取共享 Excel 中的发货明细，通过 CDP 操控 Edge 浏览器在物流网站查询运单轨迹，将最新路由信息按公司写回"物流轨迹N"列。

当前支持三家公司：宁致(NZ)、云驼(999)、小满(XM)。采用适配器模式，可扩展。

## 六大模块

| 模块 | 入口 | 功能 |
|------|------|------|
| 物流轨迹查询 | `python -m src.main` | CDP 查物流 → 写回轨迹列 + 缺失追踪 |
| 数据录入 | `python -m src.data_entry` | IM 文本解析 → 按日期插入 Excel |
| ASIN 图片匹配 | `python -m src.image_inserter build/insert` | ASIN→图片库 → 嵌入 B 列 |
| 格式迁移 | `python -m src.migrate` | 旧规范 Excel → 新规范列位映射 |
| 跨表数据填写 | `python -m src.cross_table` | 发货表→统计表 ASIN 关联，扣在采/加在途 |
| 关键词排名 | `python -m src.keyword_rank` | 鼠标驱动 CDP 搜索(搜索框点击+按钮点击+翻页点击) → 写回 Excel。BSR 多语言提取(FR/DE/EN)+自动展开折叠区域。CDP Chrome 需养号(登录Amazon+浏览+加购)否则无Cookie广告不展示 |

## 关键词排名产品配置

| 产品 | --site | --asin | --bsr-asin | 数据起始列 |
|------|--------|--------|------------|-----------|
| 刮水器 | de | B0CLXXD2X4 B0C6TCLHHT B0GSZHYB2T B0H1R1DGKH B0H4MC8STF B0H4LXJ5QG B0H4M6H2GT | B0CLXXD2X4 | 自动检测 |
| 猫砂垫 | fr | B0CH4N8V6P | B0CH4N8V6P | 自动检测 |
| 反光衣 | fr | B0GCDF56DJ B0GCF4T6NM B0GCFNSKDS | B0GCDF56DJ | 自动检测 |

用法:
```bash
# 刮水器 (7个ASIN，自然位和广告位全部追踪)
python -m src.keyword_rank <Excel路径> --site de \
    --asin B0CLXXD2X4 B0C6TCLHHT B0GSZHYB2T B0H1R1DGKH B0H4MC8STF B0H4LXJ5QG B0H4M6H2GT

# 猫砂垫
python -m src.keyword_rank <Excel路径> --site fr --asin B0CH4N8V6P

# 反光衣
python -m src.keyword_rank <Excel路径> --site fr --asin B0GCDF56DJ B0GCF4T6NM B0GCFNSKDS
```

## 核心架构

```
同事的 Windows 电脑:
├── .bat (启动 Chrome + CDP 9222 + 物流网站标签页)
└── logistics-track Skill
    ├── 读 Excel → 按表头自动匹配列位 → 按前缀归属公司
    ├── CDP → localhost:9222 → 逐公司查询
    │   宁致/小满: fetch API 调内部 JSON 接口 (~0.2s/单号)
    │   云驼: DOM 逐单 + 原生鼠标点击查询按钮 + 单条回退选运输商
    └── 按 track_position 写回对应物流轨迹N列
```

所有模块统一 CLI 规范：`--json` 输出结构化运行摘要，顶层异常输出 JSON 到 stderr。
`--keepalive` 启动后台守护线程防 session 过期。

## 查询策略

| 公司 | 方式 | 说明 |
|------|------|------|
| 宁致 | fetch API | 浏览器内 fetch() 调 `/tracking/app?inajax=1&tracking_number=NZ...` |
| 小满 | fetch API | 同上，调 `xmsdwl.nextsls.com` 同一端点 |
| 云驼 | DOM 逐单 + 原生鼠标点击 | 17track SPA，CDP Input.dispatchMouseEvent 触发 React 按钮 |

fetch API 策略借鉴象往项目：fetch 在浏览器内执行，携带完整 Cookie/会话，
从服务器角度看与页面自身的 AJAX 请求无法区分，零 bot 检测风险。

## 筛选逻辑 (当前)

```
遍历数字命名的 Sheet 每一行:
  按第 2 行表头找到"物流单号"列
  从该列拆出所有单号（换行分隔）
  按前缀归属公司: NZ→宁致, 999→云驼, XM→小满, HY→华洋, HYC→华运昌
  → 不依赖发货公司列（业务填写不规范，前缀才是权威标识）
```

## 列位映射

列位通过第 2 行表头文字自动匹配，不再硬编码索引：
- "物流单号" → 提取单号的来源列
- "物流轨迹1/2/N" → 回写轨迹的目标列

找不到表头时回退到新规范默认值（物流单号=col27 AB列，物流轨迹1=col33 AH列）。

每家发货公司独占一个"物流轨迹N"列，N = 该公司单号在物流单号列首次出现的次序。
缺列时紧跟最后一个物流轨迹列后 insert_cols 插入。

## 项目结构

```
Logistics_Automation/
├── bin/
│   ├── 物流网站一键启动.bat        # Edge 版
│   └── 物流网站一键启动-Chrome.bat # Chrome 版 (推荐)
├── images/
│   └── products/             # ASIN 图片库
├── src/
│   ├── cdp_client.py        # CDP WebSocket + fetch_api()
│   ├── cdp_util.py           # CDP 工具函数 (val)
│   ├── cli_utils.py          # CLI 公共: fatal_error / print_json_summary / RunTimer
│   ├── keepalive.py          # 标签页刷新保活守护 TabKeepAlive
│   ├── cross_table.py       # 跨表数据填写 — 发货表→统计表 在采/在途更新
│   ├── data_entry.py         # 半结构化物流文本解析 + 自动填入 Excel
│   ├── excel_reader.py       # 读取 + 表头自动匹配 + 前缀归属
│   ├── excel_writer.py       # 按公司写物流轨迹N列 + 备份 + 空结果守卫
│   ├── image_inserter.py     # ASIN 图片库构建 + 自动嵌入图片
│   ├── migrate.py            # 旧格式 → 新规范列位迁移
│   ├── validation.py         # 轨迹数据校验 is_valid_routing
│   ├── miss_tracker.py       # 缺失单号追踪 + 顽固补跑
│   ├── main.py               # 主流程编排 + healthcheck + retry-stubborn
│   └── companies/
│       ├── base.py           # CompanyAdapter 抽象基类 + TrackingResult
│       ├── fetch_adapter.py  # FetchApiAdapter — 宁致/小满共用基类
│       ├── ningzhi.py        # 宁致 NZ → fetch API
│       ├── yuntuo.py         # 云驼 999 → DOM 逐单
│       └── xiaoman.py        # 小满 XM → fetch API
├── skill/logistics-track/SKILL.md
└── requirements.txt
```

## 脚本

```bash
# 物流轨迹查询
python -m src.main <excel_path> [sheet_names] [--json] [--keepalive]
python -m src.main <excel_path> --company 小满,宁致
python -m src.main --healthcheck [--json]
python -m src.main <excel_path> --retry-stubborn [--json]

# 数据录入
python -m src.data_entry <excel_path> [--json]

# ASIN 图片匹配
python -m src.image_inserter build <ASIN映射Excel>
python -m src.image_inserter insert <目标Excel> [--json]

# 旧格式迁移
python -m src.migrate <旧格式Excel> -o <输出路径> [--json]

# 跨表数据填写
python -m src.cross_table <统计表> <发货表...> [--json]

# 关键词排名查询 (鼠标驱动: 搜索框点击+翻页点击，ease-out轨迹模拟真人)
python -m src.keyword_rank <excel> --site de --asin B0CLXXD2X4 ... [--json]  # 刮水器
python -m src.keyword_rank <excel> --site fr --asin B0CH4N8V6P [--json]      # 猫砂垫
python -m src.keyword_rank <excel> --site fr --asin B0CH4N8V6P --dry-run

# 标签页保活
python -m src.keepalive --status      # CDP 状态
python -m src.keepalive --daemon      # 后台持续保活
```

## 平台差异

| | WSL (开发) | Windows (生产) |
|---|---|---|
| 浏览器 | Chrome (CDP 9222) | Chrome (推荐) / Edge |
| CDP 地址 | Windows 宿主 IP `172.xx.xx.xx:9222`（需设 `CDP_HOST`） | `localhost:9222` |
| Excel 路径 | `/mnt/c/Users/.../` | `C:\Users\...\` |
| Python 命令 | `python3` | `python` |
| 编码 | UTF-8 (原生) | GBK → stdout 强制 UTF-8 |

`CDP_HOST` 环境变量控制 CDP 地址，默认 `localhost:9222`。
WSL 开发时需指向 Windows 宿主 IP（如 `export CDP_HOST=172.28.190.60:9222`）。

## 稳健性设计

- **数据校验**: is_valid_routing 拦截页面改版产生的垃圾数据
- **漏查不覆盖**: merge_preserve 按单号合并新旧，本次未查到保留旧值
- **异常检测**: ≥5 单且成功率 <50% → 跳过写入保护存量
- **金丝雀自检**: --healthcheck 用已知单号预验证各站点
- **缺失追踪**: _misses.json 记录缺失，miss_count≥2 判顽固
- **自动备份**: 每次写入前自动备份 Excel
