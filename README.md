# Logistics Automation

物流自动化工具箱，包含六大模块。

## 模块

| 模块 | 命令 | 用途 |
|------|------|------|
| 物流轨迹查询 | `python -m src.main <excel> [--json] [--keepalive]` | CDP 查物流 → 写回轨迹列 |
| 数据录入 | `python -m src.data_entry <excel> [--json]` | IM 文本解析 → 填入 Excel |
| ASIN 图片匹配 | `python -m src.image_inserter build/insert [--json]` | ASIN → 图片库 → 嵌入 B 列 |
| 格式迁移 | `python -m src.migrate <旧表> [--json]` | 旧规范 → 新规范列位映射 |
| 跨表填写 | `python -m src.cross_table <统计表> <发货表...> [--json]` | ASIN 关联，扣在采/加在途 |
| 关键词排名 | `python -m src.keyword_rank <excel> --site de/fr --asin ... [--json]` | 鼠标驱动 CDP 搜索(点击+翻页)，模拟真人操作。BSR多语言自动提取，CDP Chrome需养号 |
| 标签页保活 | `python -m src.keepalive --status/--daemon` | 后台守护防 session 过期 |

## 使用

```bash
# 物流轨迹查询
python -m src.main <excel> [--company 小满,宁致] [--healthcheck] [--retry-stubborn] [--json] [--keepalive]

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

# 关键词排名 (鼠标驱动: 搜索框点击+翻页点击，ease-out轨迹模拟真人)
python -m src.keyword_rank <excel> --site fr --asin B0CH4N8V6P                        # 猫砂垫
python -m src.keyword_rank <excel> --site de --asin B0CLXXD2X4 ...                    # 刮水器
python -m src.keyword_rank <excel> --site fr --asin B0GCDF56DJ B0GCF4T6NM B0GCFNSKDS  # 反光衣

# 标签页保活
python -m src.keepalive --status   # 查看 CDP 状态
python -m src.keepalive --daemon   # 后台保活
```

## 查询策略

| 公司 | 前缀 | 方式 |
|------|------|------|
| 宁致 | NZ | fetch API |
| 小满 | XM | fetch API |
| 云驼 | 999 | DOM 逐单 + 原生鼠标点击 |

## 列位匹配

按第 2 行表头文字自动匹配，不硬编码列索引。列位变动不影响运行。

## 环境

| | WSL (开发) | Windows (生产) |
|---|---|---|
| CDP 地址 | 由 `CDP_HOST` 环境变量 | `localhost:9222` |
| Python | `python3` | `python` |
