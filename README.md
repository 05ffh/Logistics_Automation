# Logistics Automation - 物流轨迹自动查询

从发货明细表 Excel 中按单号前缀识别各家物流单号，通过 CDP 操控 Edge 浏览器查询运单轨迹，将最新路由信息按公司回写到 Excel 的"物流轨迹N"列。

当前支持 **宁致 (NZ)**、**云驼 (999)**、**小满 (XM)** 三家，采用适配器模式，可扩展更多公司。

## 四大模块

| 模块 | 入口 | 功能 |
|------|------|------|
| 物流轨迹查询 | `python -m src.main` | CDP 查物流 → 写回轨迹列 + 缺失追踪 |
| 数据录入 | `python -m src.data_entry` | IM 文本解析 → 按日期插入 Excel |
| ASIN 图片匹配 | `python -m src.image_inserter build/insert` | ASIN→图片库 → 嵌入 B 列 |
| 格式迁移 | `python -m src.migrate` | 旧规范 Excel → 新规范列位映射 |

## 架构

```
┌─ Windows ──────────────────────────────────────────┐
│  .bat (启动 Edge + CDP 9222 + 物流网站)              │
│  OpenClaw Skill → Python → CDP → Edge              │
│  Excel (共享文件)                                   │
└────────────────────────────────────────────────────┘

查询策略 (借鉴象往项目):
  - fetch API (宁致/小满): 浏览器内 fetch() 调内部 JSON API
    → 无 DOM 操作，零 bot 检测风险，20-70x 快于 DOM 抓取
  - DOM 批量 (云驼): 结果页 textarea 批量填入 + 单条回退选运输商
    → 17track 不支持内部 API 直调，保留 DOM 方式
```

## 快速开始

### 前置条件

- Windows 10/11、Microsoft Edge
- Python 3.10+（OpenClaw 内置或独立安装）
- `pip install openpyxl`

### 首次使用

1. 双击 `bin/物流网站一键启动.bat` → Edge 打开物流网站（CDP 端口 9222）
2. 宁致 (nzhexp) 首次需手动登录（账号密码见团队内部文档，登录后 Cookie 持久化无需重复登录）
3. 云驼 (17track) 和小满 (xmsdwl) 无需登录
4. 运行脚本：

```bash
# 物流轨迹查询（全量）
python -m src.main "C:\path\to\发货明细表.xlsx"

# 可选：只处理指定 sheet / 指定公司
python -m src.main "C:\path\to\发货明细表.xlsx" 202606
python -m src.main "C:\path\to\发货明细表.xlsx" --company 小满,宁致

# 健康自检 / 顽固补跑
python -m src.main --healthcheck
python -m src.main "C:\path\to\发货明细表.xlsx" --retry-stubborn

# 数据录入
python -m src.data_entry "C:\path\to\发货明细表.xlsx"

# 图片库构建 + 插入
python -m src.image_inserter build "C:\path\to\ASIN映射.xlsx"
python -m src.image_inserter insert "C:\path\to\发货明细表.xlsx"

# 旧格式迁移
python -m src.migrate "C:\path\to\旧格式.xlsx"
```

### 自然语言调用（OpenClaw Skill）

```
"帮我查202606宁致和云驼的物流轨迹"
```

## 各公司查询方式

| 公司 | 前缀 | 网站 | 查询方式 | 登录 |
|------|------|------|----------|------|
| 宁致 | `NZ` | nzhexp.nextsls.com | **fetch API** — 调内部 JSON 接口 | 需要 |
| 云驼 | `999` | 17track.net | DOM 逐单 + 回退选"愿景征途" | 不需要 |
| 小满 | `XM` | xmsdwl.nextsls.com | **fetch API** — 调内部 JSON 接口 | 不需要 |

**宁致和小满共用同一个 API 端点**：`/tracking/app?inajax=1&tracking_number={tn}`，返回 `{data: {shipment: {traces: [{time, info}, ...]}}}`，traces 时间倒序排列。

## 工作流程

```
读 Excel (仅数字命名的 sheet)
  → 按第 2 行表头文字自动匹配列位 (不再硬编码 S=物流单号)
  → 按单号前缀归属公司 (J/K 公司名填写不规范，前缀才权威)
  → 解析合并单元格
  → 各公司查询最新路由:
      宁致/小满: fetch API 直调 JSON 接口 (~0.2s/单号)
      云驼: DOM 逐单查询 + 单条回退
  → 按公司写入"物流轨迹N"列:
      N = 该公司单号在物流单号列首次出现的次序
      缺列时紧跟最后一个物流轨迹列后插入，列宽统一
  → 迁移清理: 移除查询公司在其他列的残留块
  → 自动备份 → 写回 → 缺失追踪记账
```

## 列位映射

列位通过第 2 行表头文字自动匹配，不再硬编码索引，兼容不同格式 Excel。

| 表头 | 用途 |
|------|------|
| 物流单号 | 提取单号的来源列，多单号换行分隔 |
| 物流轨迹1 | 第 1 家公司的轨迹 |
| 物流轨迹2 | 第 2 家公司的轨迹 |
| 物流轨迹N | 第 N 家，缺列时自动新建 |

**每家发货公司独占一个"物流轨迹N"列**，列号 = 该公司单号在物流单号列首次出现的先后次序。
脚本只写查询到的公司（宁致/云驼/小满）的列，华洋/华运昌等业务手填的列不触碰。

## 项目结构

```
├── bin/
│   └── 物流网站一键启动.bat    # 启动 Edge + CDP
├── images/
│   └── products/               # ASIN 图片库（按 ASIN 命名）
├── src/
│   ├── cdp_client.py           # CDP WebSocket 通信层 + fetch_api()
│   ├── cdp_util.py             # CDP 工具函数 (val)
│   ├── data_entry.py           # 半结构化物流文本解析 + 自动填入 Excel
│   ├── excel_reader.py         # 读取 + 表头自动匹配 + 前缀归属 + 合并单元格
│   ├── excel_writer.py         # 按公司写物流轨迹N列 + 建列 + 迁移清理 + 备份
│   ├── image_inserter.py       # ASIN 图片库构建 + 自动嵌入图片
│   ├── migrate.py              # 旧格式 → 新规范列位迁移
│   ├── validation.py           # 轨迹数据校验 (is_valid_routing)
│   ├── miss_tracker.py         # 缺失单号追踪 + 顽固补跑
│   ├── companies/              # 多公司适配器
│   │   ├── base.py             # CompanyAdapter 抽象基类
│   │   ├── ningzhi.py          # 宁致 (NZ) → nzhexp.nextsls.com (fetch API)
│   │   ├── yuntuo.py           # 云驼 (999) → 17track.net (DOM 逐单)
│   │   └── xiaoman.py          # 小满 (XM) → xmsdwl.nextsls.com (fetch API)
│   └── main.py                 # 主流程编排
├── skill/
│   └── logistics-track/
│       └── SKILL.md            # OpenClaw Skill 定义
└── requirements.txt
```

## 扩展新公司

1. 在 `src/companies/` 新增适配器，继承 `CompanyAdapter`，实现 `ensure_tab` 和 `query`
2. 优先尝试找内部 JSON API 用 `cdp.fetch_api()` 调用（参考宁致/小满），找不到再用 DOM 方式（参考云驼）
3. 在 `src/excel_reader.py` 的 `CARRIER_PREFIXES` 注册前缀 → 公司名
4. 加入 `src/main.py` 的 `ADAPTERS` 列表

## 数据录入

业务人员通过 IM 发来的半结构化物流文本，一键解析并插入 Excel，无需手动复制粘贴。

**输入格式**（业务人员发来的原始消息）：
```
7.15
柘流-US
发货美森正船 第29水
预计开船: 7.17 预计到港: 8.18
美森 CLX++
FBA 编号: FBA18Q888NRNT00028
单号: NZ26071140597
交货仓: ONT8
箱规: 55.9/44/44 重量: 24.29
附加费: 5
配送: 8.22
```

**使用方式**：
```bash
# 直接粘贴文本，程序自动识别并写入
python -m src.data_entry "C:\path\to\发货明细表.xlsx"

# 进入交互模式，逐条粘贴
```

**自动处理**：
- 第 1 行日期 → 按日期排序定位插入行
- 第 2 行 `XX-XX` 格式自动识别为发货店铺（H 列），否则为发货公司（J 列）
- 后续行按关键字匹配到对应列：单号、箱规、重量、FBA 编号、交货仓、附加费等
- 列模板自动初始化：首次运行时新增 货物名称/箱规/重量 列，重命名 箱数量→箱内数量 等
- 所有格式（字体、边框、对齐）与原文件保持一致

## ASIN 图片匹配

从 ASIN 映射 Excel 提取图片库，按目标文件的 ASIN 列自动嵌入产品图片到 B 列。

**映射文件格式**：A=品名, B=asin, C=图片（WPS 单元格内嵌 DISPIMG）

```bash
# 第一步：从映射表构建图片库
python -m src.image_inserter build "C:\path\to\ASIN映射.xlsx"

# 第二步：为目标 Excel 插入图片
python -m src.image_inserter insert "C:\path\to\发货明细表.xlsx"
```

自动处理：
- ASIN 列和图片列按表头自动识别（大小写不敏感）
- 相同 ASIN 多行共享同一图片 ID
- 图片自适应行高
- 图片库随 Git 分发，同事 pull 即用

## 格式迁移

将旧规范 Excel 的一键迁移到 2026 发货信息表最终版规范。

```bash
python -m src.migrate "C:\path\to\旧格式.xlsx" -o "C:\path\to\规范版.xlsx"
```

自动处理：
- 列位映射（旧列→新列）、表头重命名
- 公式引用修正（`=J3*K3` → `=K3*L3`）
- 日期序列号转换
- 价格拆分（`9+2` → 价格 9 + 附加费 2）
- 格式完整保留（或在新建文件时统一 等线/细线边框/自适应列宽行高）

## 缺失追踪

每次正常跑完自动记录 MISS 到 `<excel名>_misses.json`。同一单号多次 MISS 判为"顽固"(miss_count>=2)：

```
正常跑:   查 80 个 → 2 个 MISS → 记入 JSON (miss_count=1)
再次跑:   查 80 个 → 同 2 个 MISS → JSON 递增 (miss_count=2, 顽固)
补跑:     --retry-stubborn → 只查这 2 个 → 补查成功则写回+移除
```
