---
name: logistics-track
description: 物流轨迹自动查询 - 从发货明细表 Excel 中按单号前缀识别各家物流公司(宁致/云驼/小满)，通过 CDP 操控浏览器查询运单轨迹，按公司分别写回"物流轨迹N"列。宁致/小满使用 fetch API 直调内部 JSON 接口，云驼使用 DOM 批量查询。支持缺失追踪和顽固补查
type: skill
platform: windows
---

# 物流轨迹查询 Skill（多公司版）

## 触发方式

用户通过自然语言调用，例如：

- "帮我查202606宁致和云驼的物流轨迹"
- "查询这个Excel里的物流单号"
- "跑一下物流轨迹查询"
- "补查顽固单号"

## 首次使用（同事拿到 Skill 后只需做一次）

1. 双击 `bin/物流网站一键启动.bat` → Edge 打开物流网站（CDP 端口 9222）
2. 宁致（nzhexp）首次需手动登录（账号密码见团队内部文档，登录后 Cookie 持久化无需重复登录）
3. 云驼（17track）无需登录
4. 小满（xmsdwl）无需登录
5. 登录后关闭 Edge，再双击 `.bat` 确认登录态保持 → 完成

之后每次使用：双击 `.bat` → 告诉 OpenClaw 要查哪个 Excel。

## 每次使用流程

```
你说: "帮我查桌面上测试（云驼、宁致）的物流轨迹"
    ↓
Skill 自动:
  1. 检查 Edge 9222 是否就绪
  2. 自检各站点是否可查询 (--healthcheck)
  3. 读 Excel → 按单号前缀归属公司（不依赖J/K列公司名）→ 解析合并单元格
  4. 报告: "找到 宁致12行/11个单号, 云驼71行/81个单号, 确认开始？"
  5. 逐公司查询 → 宁致/小满 fetch API(~0.2s/单号)、云驼批量(≤5/批) + 单条回退
  6. 按公司分别写回"物流轨迹N"列(N=公司在S列首次出现次序)
  7. 缺失单号记录到 _misses.json，方便后续精准补跑
  8. 自动备份 Excel → 显示运行汇总(每家成功率)
```

## 多公司说明

| 公司 | 前缀 | 网站 | 查询方式 |
|------|------|------|----------|
| 宁致 | NZ | nzhexp.nextsls.com | fetch API（需登录） |
| 云驼 | 999 | 17track.net | DOM 批量(5/批) + 单条回退选"愿景征途" |
| 小满 | XM | xmsdwl.nextsls.com | fetch API（无需登录） |

**单号归属**：按前缀匹配（`999`=云驼、`NZ`=宁致、`XM`=小满），不依赖 J/K 列的公司名（业务填写不规范）。
**轨迹列**：每家公司独占"物流轨迹N"列，N = 该公司单号在 物流单号 列首次出现的次序。缺列自动新增。
**列位匹配**：第 2 行表头文字自动匹配（"物流单号""物流轨迹N"），不再硬编码列索引，兼容不同格式 Excel。

## 脚本

```bash
# 正常查询（全量，自动记账缺失单号）
python -m src.main <excel_path> [sheet_names]

# 只查指定公司
python -m src.main <excel_path> --company 小满
python -m src.main <excel_path> --company 小满,宁致

# 健康自检（跑前先验证各站点是否正常）
python -m src.main --healthcheck

# 精准补跑（只查顽固缺失单号，不全量重跑）
python -m src.main <excel_path> --retry-stubborn
```

| 参数 | 说明 |
|------|------|
| `excel_path` | Excel 文件路径（必需） |
| `sheet_names` | 要处理的 sheet 名称，逗号分隔（可选） |
| `--company 小满,宁致` | 只查指定公司，逗号分隔（可选，默认全查） |
| `--healthcheck` | 金丝雀自检，用已知单号验证各站点结构是否还通 |
| `--retry-stubborn` | 只查 miss_count>=2 的顽固单号，不全量跑 |

环境变量: `CDP_HOST`，默认 `localhost:9222`

## 直接查询指定单号（不用 Excel）

当用户让你查某几个具体单号时，**不要跑 Excel 流程**，用以下 Python 脚本直接在浏览器里查：

### 查云驼 (999) 单号

```bash
cd <项目目录>
python3 -c "
from src.cdp_client import CdpClient
from src.companies.yuntuo import YunTuoAdapter
cdp = CdpClient()
adapter = YunTuoAdapter()
ws = adapter.ensure_tab(cdp)
cdp.connect_tab(ws)
for tn in ['999260706000543', '999260708000910']:
    routing = adapter._query_one(cdp, tn)
    print(f'{tn} → {routing if routing else \"MISS (未查到) — 页面可能需选运输商或单号不存在\"}')
cdp.close()
"
```

### 查宁致 (NZ) 单号

```bash
cd <项目目录>
python3 -c "
from src.cdp_client import CdpClient
from src.companies.ningzhi import NingZhiAdapter
cdp = CdpClient()
adapter = NingZhiAdapter()
ws = adapter.ensure_tab(cdp)
cdp.connect_tab(ws)
results = adapter.query(cdp, ['NZ2605063839'])
for r in results:
    print(f'{r.tracking_no} → {r.routing_info if r.routing_info else \"MISS (未查到) — 可能页面未登录或单号不存在\"}')
cdp.close()
"
```

### 查小满 (XM) 单号

```bash
cd <项目目录>
python3 -c "
from src.cdp_client import CdpClient
from src.companies.xiaoman import XiaoManAdapter
cdp = CdpClient()
adapter = XiaoManAdapter()
ws = adapter.ensure_tab(cdp)
cdp.connect_tab(ws)
results = adapter.query(cdp, ['XM26070315932', 'XM26070358194'])
for r in results:
    print(f'{r.tracking_no} → {r.routing_info if r.routing_info else \"MISS (未查到) — 页面可能需手动输入单号或单号不存在\"}')
cdp.close()
"
```

**重要**: 查询前确保 Edge 已通过 `.bat` 启动、对应标签页已打开（云驼=17track、宁致=nzhexp 且已登录、小满=xmsdwl）。查多个云驼单号时**优先一次全查**，不要每个单号单独开一个 python3 进程——在同一进程里循环更快。

## 列位映射

列位通过第 2 行表头文字自动匹配，不再硬编码索引，兼容不同格式 Excel。常见布局参考：

| 表头 | 说明 |
|------|------|
| 物流单号 | 多单号换行分隔，按前缀归属公司的依据 |
| 物流轨迹1 | 第 1 家公司的轨迹 |
| 物流轨迹2 | 第 2 家公司的轨迹 |
| 物流轨迹N | 第 N 家，缺列时自动新增（列宽对齐物流轨迹1） |

J/K 列（发货渠道/发货公司）填写不规范，脚本不依赖此项——前缀才是权威标识。

## 缺失追踪 + 精准补跑

每次正常跑完自动记录 MISS 到 `<excel名>_misses.json`。同一单号多次 MISS 判为"顽固"(miss_count>=2)：

```
正常跑:   查 81 个 → 2 个 MISS → 记入 JSON (miss_count=1)
再次跑:   查 81 个 → 同 2 个 MISS → JSON 递增 (miss_count=2, 顽固)
补跑:     --retry-stubborn → 只查这 2 个 → 补查成功则写回+移除
```

如果某单号在后续正常跑中被查到 → 自动从 JSON 删除，不再追踪。

## 互动确认点

1. **启动前**: 确认 Edge 已启动、nzhexp 已登录、xmsdwl 和 17track 标签页正常
2. **筛选后**: 报告各公司行数和单号数，确认开始查询
3. **异常情况**: 某公司成功率异常(<50%) → 醒目警告 + 跳过写入保护存量，提示人工核查
4. **完成后**: 显示运行汇总 + 缺失追踪概况

## 错误处理

| 情况 | 处理 |
|------|------|
| Edge 9222 不通 | 提示 "请先双击 物流网站一键启动.bat 启动 Edge" |
| nzhexp 未登录 | 提示 "请在 Edge 中打开 nzhexp 页面并登录" |
| Excel 文件被占用 | 提示 "请关闭 Excel 后重试" |
| 单号查询无结果 | 保留旧轨迹不覆盖，记入 misses JSON 供后续补查 |
| 某公司成功率异常低 | ⚠️ 告警 + 跳过写入该公司（保护存量数据不被覆盖） |
| 页面结构变化 | 金丝雀自检可提前发现；异常检测在跑时兜底 |
