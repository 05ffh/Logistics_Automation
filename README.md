# Logistics Automation - 物流轨迹自动查询

从发货明细表 Excel 中识别宁致物流单号，通过 CDP 操控 Edge 浏览器在 [nzhexp.nextsls.com](http://nzhexp.nextsls.com) 批量查询运单轨迹，将最新路由信息写回 Excel。

## 架构

```
┌─ Windows ──────────────────────────────────────────┐
│  .bat (启动 Edge + CDP 9222 + 8 个物流网站)          │
│  OpenClaw Skill → Python → CDP → Edge              │
│  Excel (共享文件)                                   │
└────────────────────────────────────────────────────┘
```

## 快速开始

### 前置条件

- Windows 10/11
- Microsoft Edge
- Python 3.10+（OpenClaw 内置或独立安装）
- `pip install openpyxl`

### 首次使用

1. 双击 `bin/物流网站一键启动.bat` → Edge 打开 8 个物流网站
2. 在 Edge 中打开 nzhexp 标签页，手动登录（仅首次）
3. 运行脚本：

```bash
python -m src.main "C:\path\to\发货明细表.xlsx"
```

### 自然语言调用（OpenClaw Skill）

```
"帮我查202606宁致的物流轨迹"
```

Skill 自动完成：检查 CDP → 检查登录 → 读 Excel → 筛选单号 → 批量查询 → 对比回写。

## 工作流程

```
Excel → 筛选宁致行(J/K列含"宁致") → 提取 NZ 前缀单号
  → CDP 逐批查询(about:blank重置状态)
  → 提取最新路由信息(时间戳+状态)
  → 对比旧数据(相同保留/不同更新/其他公司轨迹不动)
  → 自动备份 → 写回 Y列
```

## 项目结构

```
├── bin/
│   └── 物流网站一键启动.bat    # 启动 Edge + CDP
├── src/
│   ├── cdp_client.py           # CDP WebSocket 通信层
│   ├── excel_reader.py         # Excel 读取 + 单号提取
│   ├── nzhexp_tracker.py       # nzhexp 查询逻辑
│   ├── excel_writer.py         # 回写 + 备份 + 占用检测
│   └── main.py                 # 主流程编排
├── skill/
│   └── logistics-track/
│       └── SKILL.md            # OpenClaw Skill 定义
└── requirements.txt
```

## 列位映射

| 列 | 内容 | 用途 |
|---|---|---|
| J | 发货渠道 | 可能含"宁致"（不规范） |
| K | 发货公司 | 应含"宁致"（规范） |
| S | 物流单号 | 多单号换行分隔，取 NZ 前缀 |
| Y | 物流轨迹1 | 回写最新路由信息 |
