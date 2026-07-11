# Logistics Automation - 物流轨迹自动查询

从发货明细表 Excel 中按单号前缀识别各家物流单号，通过 CDP 操控 Edge 浏览器在对应物流网站查询运单轨迹，将最新路由信息按公司回写到 Excel 的"物流轨迹N"列。

当前支持 **宁致 (NZ)** 和 **云驼 (999)** 两家，采用适配器模式，可扩展更多公司。

## 架构

```
┌─ Windows ──────────────────────────────────────────┐
│  .bat (启动 Edge + CDP 9222 + 物流网站)              │
│  OpenClaw Skill → Python → CDP → Edge              │
│  Excel (共享文件)                                   │
└────────────────────────────────────────────────────┘

多公司适配器:
  main.py → ADAPTERS = [宁致, 云驼, ...]
    每家公司实现 CompanyAdapter: ensure_tab / query
```

## 快速开始

### 前置条件

- Windows 10/11、Microsoft Edge
- Python 3.10+（OpenClaw 内置或独立安装）
- `pip install openpyxl`

### 首次使用

1. 双击 `bin/物流网站一键启动.bat` → Edge 打开物流网站（CDP 端口 9222）
2. 宁致 (nzhexp) 首次需手动登录；云驼 (17track) 无需登录
3. 运行脚本：

```bash
python -m src.main "C:\path\to\发货明细表.xlsx"
# 可选：只处理指定 sheet
python -m src.main "C:\path\to\发货明细表.xlsx" 202606
```

### 自然语言调用（OpenClaw Skill）

```
"帮我查202606宁致和云驼的物流轨迹"
```

## 各公司查询方式

| 公司 | 前缀 | 网站 | 查询方式 |
|---|---|---|---|
| 宁致 | `NZ` | nzhexp.nextsls.com | 逐单查询（Ant Design，抽屉提取路由；首次需登录） |
| 云驼 | `999` | 17track.net | 结果页原生批量（≤40/批），未命中回退单条（自动选"愿景征途"运输商） |

## 工作流程

```
读 Excel(仅数字命名的 sheet)
  → 按单号前缀归属公司(J/K 公司名填写不规范，前缀才权威)
  → 解析合并单元格(K列发货公司常合并多行)
  → 各公司查询最新路由(时间戳+状态)
  → 按公司写入"物流轨迹N"列:
      N = 该公司单号在 S 列首次出现的次序
      缺列时紧跟物流轨迹2后插入，列宽统一物流轨迹1
  → 迁移清理: 移除查询公司在其他列的残留块
  → 自动备份 → 写回
```

## 列位映射

| 列 | 内容 | 说明 |
|---|---|---|
| S | 物流单号 | 多单号换行分隔，可含多家公司（如 HYC→999→NZ） |
| Y | 物流轨迹1 | 第 1 家公司的轨迹 |
| Z | 物流轨迹2 | 第 2 家公司的轨迹 |
| 物流轨迹3… | 物流轨迹3+ | 第 3 家起自动新增，紧跟物流轨迹2 之后 |

**每家发货公司独占一个"物流轨迹N"列**，列号 = 该公司单号在 S 列首次出现的先后次序。
脚本只写查询到的公司（宁致/云驼）的列，华洋/华运昌等业务手填的列不触碰。

## 项目结构

```
├── bin/
│   └── 物流网站一键启动.bat    # 启动 Edge + CDP
├── src/
│   ├── cdp_client.py           # CDP WebSocket 通信层
│   ├── excel_reader.py         # 读取 + 前缀归属 + 合并单元格 + 位置计算
│   ├── excel_writer.py         # 按公司写物流轨迹N列 + 建列 + 迁移清理 + 备份
│   ├── companies/              # 多公司适配器
│   │   ├── base.py             # CompanyAdapter 抽象基类
│   │   ├── ningzhi.py          # 宁致 (NZ) → nzhexp.nextsls.com
│   │   └── yuntuo.py           # 云驼 (999) → 17track.net（批量查询）
│   └── main.py                 # 主流程编排
├── skill/
│   └── logistics-track/
│       └── SKILL.md            # OpenClaw Skill 定义
└── requirements.txt
```

## 扩展新公司

在 `src/companies/` 新增适配器，继承 `CompanyAdapter`，实现 `ensure_tab` 和 `query`，
在 `src/excel_reader.py` 的 `CARRIER_PREFIXES` 注册前缀→公司名，
并加入 `src/main.py` 的 `ADAPTERS` 列表即可。
