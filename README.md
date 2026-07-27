# Logistics Automation

物流自动化工具箱，包含五大模块。

## 模块

| 模块 | 命令 | 用途 |
|------|------|------|
| 物流轨迹查询 | `python -m src.main <excel>` | CDP 查物流 → 写回轨迹列 |
| 数据录入 | `python -m src.data_entry <excel>` | IM 文本解析 → 填入 Excel |
| ASIN 图片匹配 | `python -m src.image_inserter build/insert` | ASIN → 图片库 → 嵌入 B 列 |
| 格式迁移 | `python -m src.migrate <旧表>` | 旧规范 → 新规范列位映射 |
| 跨表填写 | `python -m src.cross_table <统计表> <发货表...>` | ASIN 关联，扣在采/加在途 |

## 使用

```bash
# 物流轨迹查询
python -m src.main <excel> [--company 小满,宁致] [--healthcheck] [--retry-stubborn]

# 数据录入 (US/DE/通用)
python -m src.data_entry <excel> --us    # US 规则，复制产品行到各仓库
python -m src.data_entry <excel> --de    # DE 规则，品名+箱数匹配回填

# ASIN 图片
python -m src.image_inserter build <映射表>   # 构建图片库
python -m src.image_inserter insert <目标表>  # 嵌入图片

# 格式迁移
python -m src.migrate <旧表> -o <输出路径>

# 跨表填写 (支持多个发货表)
python -m src.cross_table <统计表> <发货表1> [发货表2] ...
```

## 查询策略

| 公司 | 前缀 | 方式 |
|------|------|------|
| 宁致 | NZ | fetch API |
| 小满 | XM | fetch API |
| 云驼 | 999 | DOM 逐单 |

## 列位匹配

按第 2 行表头文字自动匹配，不硬编码列索引。列位变动不影响运行。

## 环境

| | WSL (开发) | Windows (生产) |
|---|---|---|
| CDP 地址 | 由 `CDP_HOST` 环境变量 | `localhost:9222` |
| Python | `python3` | `python` |
