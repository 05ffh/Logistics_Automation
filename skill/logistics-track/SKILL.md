---
name: logistics-track
description: 物流轨迹自动查询 - 从发货明细表 Excel 中按单号前缀识别各家物流公司(宁致/云驼/小满)，通过 CDP 操控浏览器查询运单轨迹，按公司分别写回"物流轨迹N"列。宁致/小满使用 fetch API 直调内部 JSON 接口，云驼使用 DOM 逐单查询（原生 CDP 鼠标点击按钮 + 限流预警）。支持缺失追踪、顽固补查、数据录入、ASIN图片匹配、格式迁移、跨表数据填写、Amazon关键词排名查询(刮水器/猫砂垫/反光衣)
type: skill
platform: windows,linux
---

# 物流轨迹查询 Skill（多公司版）

## 触发方式

用户通过自然语言调用，例如：

- "帮我查202606宁致和云驼的物流轨迹"
- "查询这个Excel里的物流单号"
- "跑一下物流轨迹查询"
- "帮我把发货表更新到备货计划里"
- "查今天的刮水器排名"
- "查猫砂垫排名"

## 首次使用（同事拿到 Skill 后只需做一次）

1. 双击 `bin/物流网站一键启动-Chrome.bat`（或 `bin/物流网站一键启动.bat`）→ Chrome 打开物流网站（CDP 端口 9222）
2. 宁致（nzhexp）首次需手动登录（账号密码见团队内部文档，登录后 Cookie 持久化无需重复登录）
3. 云驼（17track）无需登录
4. 小满（xmsdwl）无需登录
5. 登录后关闭 Chrome，再双击 `.bat` 确认登录态保持 → 完成

之后每次使用：双击 `.bat` → 告诉 Claude 要查哪个 Excel。

Edge 版本保留（`物流网站一键启动.bat`），Chrome 版本（`物流网站一键启动-Chrome.bat`）推荐使用。

## 每次使用流程

```
你说: "帮我查桌面上测试（云驼、宁致）的物流轨迹"
    ↓
Skill 自动:
  1. 检查 Chrome/Edge 9222 是否就绪
  2. 自检各站点是否可查询 (--healthcheck)
  3. 读 Excel → 按单号前缀归属公司（不依赖发货公司列）→ 解析合并单元格
  4. 报告: "找到 宁致12行/11个单号, 云驼71行/81个单号, 确认开始？"
  5. 逐公司查询 → 宁致/小满 fetch API(~0.2s/单号)、云驼 DOM 逐单 + 单条回退
  6. 按公司分别写回"物流轨迹N"列(N=公司在S列首次出现次序)；写入前自动备份，空结果拒绝写入保护存量
  7. 缺失单号记录到 _misses.json（含 error 字段区分 MISS/ERROR），方便后续精准补跑
  8. 显示运行汇总 + --json 结构化摘要（成功率/各公司统计/耗时/误差详情）
```

## 多公司说明

| 公司 | 前缀 | 网站 | 查询方式 |
|------|------|------|----------|
| 宁致 | NZ | nzhexp.nextsls.com | fetch API（需登录） |
| 云驼 | 999 | 17track.net | DOM 逐单 + 原生鼠标点击查询按钮 + 单条回退选"愿景征途" |
| 小满 | XM | xmsdwl.nextsls.com | fetch API（无需登录） |

**单号归属**：按前缀匹配（`999`=云驼、`NZ`=宁致、`XM`=小满），不依赖发货公司列（业务填写不规范，前缀才是权威标识）。
**轨迹列**：每家公司独占"物流轨迹N"列，N = 该公司单号在 物流单号 列首次出现的次序。缺列自动新增。
**列位匹配**：第 2 行表头文字自动匹配（"物流单号""物流轨迹N"），不再硬编码列索引，兼容不同格式 Excel。

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

# 关键词排名查询 (首页使用搜索框输入，模拟人工搜索保证广告布局一致)
python -m src.keyword_rank <excel> --site de --asin B0CLXXD2X4 ... [--json]   # 刮水器 DE
python -m src.keyword_rank <excel> --site fr --asin B0CH4N8V6P [--json]        # 猫砂垫 FR
python -m src.keyword_rank <excel> --site fr --asin B0GCDF56DJ B0GCF4T6NM B0GCFNSKDS [--json]    # 反光衣 FR

# 标签页保活 (防 session 过期)
python -m src.keepalive --status      # 查看 CDP 状态
python -m src.keepalive --daemon      # 后台持续保活
```

| 参数 | 说明 |
|------|------|
| `excel_path` | Excel 文件路径（必需） |
| `sheet_names` | 要处理的 sheet 名称，逗号分隔（可选） |
| `--company 小满,宁致` | 只查指定公司，逗号分隔（可选，默认全查） |
| `--healthcheck` | 金丝雀自检，用已知单号验证各站点结构是否还通 |
| `--retry-stubborn` | 只查 miss_count>=2 的顽固单号，不全量跑 |
| `--json` | 输出结构化 JSON 运行摘要（适用于所有模块），方便 agent 自动解析 |
| `--keepalive` | 后台守护线程定期刷新标签页，防止长时间运行时 session 过期 |

环境变量: `CDP_HOST`，默认 `localhost:9222`（WSL 开发时设为 Windows 宿主 IP）

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

## 数据录入

三种模式，通过 stdin 传入文本。

### US 规则（`--us`）

复制原文件产品行到各仓库，ZIP XML 直写（不经 openpyxl，WPS 内容完整保留）。

```bash
python -m src.data_entry <excel> --us
```

输入格式（共享头部 + 编号货件块）：
```
发货公司：小满
发货店铺：稳再-US
指定发货渠道：海运
箱规：60*40*40cm
重量：21
开船时间：7月30日

1、货件号：FBA19J3PTGXG-仓库：FWA4-SKU:4-箱数：5
US美东纽约海卡专线
价格：7.29
发车、发船后配送时段：开船到签收38-52天
```

自动处理：品名取原文件 / DISPIMG 自动匹配 / 箱数留空 / 全行边框 / 品名列黄底 / 新行替换原产品行。

### DE 规则（`--de`）

品名+箱数匹配已有行回填，ZIP XML 直写。

```bash
python -m src.data_entry <excel> --de
```

输入格式：`品名：1、xxx-N箱 2、yyy-M箱` + 渠道/时效/价格等字段。

### 通用单条/批量

```bash
python -m src.data_entry <excel>          # 单条
python -m src.data_entry <excel> --batch  # 批量（空行分隔）
```

## 跨表数据填写

根据发货信息表自动更新统计表。按 ASIN 关联两表，从"在采"扣减发货数量，根据发货店铺（稳再/柘流）累加到对应的"在途"列。

触发示例：
- "帮我把发货表更新到备货计划里"
- "把这些发货表的数据同步到统计表"

```bash
python -m src.cross_table <统计表> <发货表1> [发货表2] ...
```

逻辑：
1. 按表头匹配两表列位（asin、发货店铺、发货数量、在采、稳再在途、柘流在途）
2. 在采 -= 发货数量
3. 发货店铺含"稳再" → 稳再在途 += 发货数量；含"柘流" → 柘流在途 += 发货数量
4. ASIN 找不到 → 报警；在采扣至负数 → 报警
5. 自动备份统计表后写回

## 列位映射

列位通过第 2 行表头文字自动匹配，不再硬编码索引，兼容不同格式 Excel。常见布局参考：

| 表头 | 说明 |
|------|------|
| 物流单号 | 多单号换行分隔，按前缀归属公司的依据 |
| 物流轨迹1 | 第 1 家公司的轨迹 |
| 物流轨迹2 | 第 2 家公司的轨迹 |
| 物流轨迹N | 第 N 家，缺列时自动新增（列宽对齐物流轨迹1） |

发货公司列填写不规范，脚本不依赖此项——前缀才是权威标识。

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

## 关键词排名查询

每天自动查 Amazon 搜索结果中的自然位排名 + 广告位页码 + BSR 大类排名。

**交互策略**：鼠标驱动搜索（Input.dispatchMouseEvent，ease-out 轨迹 + 微抖动模拟真人）。首页→点击搜索框→输入→点击搜索按钮，翻页→鼠标点击 .s-pagination-next。Amazon 广告布局与人工搜索一致。广告识别分两路：卡片子树搜索 Sponsored 文字（覆盖商品/视频广告）+ 卡片祖先检测 Brand 容器（覆盖品牌横幅广告），零误判零跨卡污染。

**CDP Chrome 养号**：CDP Chrome 使用独立空白 Profile（`--user-data-dir`），无浏览历史/Cookie 时 Amazon 不投放广告或广告布局异常。首次使用需在 CDP Chrome 中**登录 Amazon 账号**并浏览几分钟（搜索、浏览商品详情页、加购），正常关闭 Chrome 后 Cookie 持久化，后续查询广告布局才与人工搜索一致。

**站点支持**：Amazon DE (`--site de`) / Amazon FR (`--site fr`)，`.bat` 文件中需包含对应站点标签页。

### 产品配置

| 产品 | 站点 | ASIN | 广告 ASIN | BSR ASIN | Excel 路径 |
|------|------|------|-----------|----------|-----------|
| 刮水器 | de | B0CLXXD2X4 B0C6TCLHHT ... (7个) | B0CLXXD2X4 B0H1R1DGKH | B0CLXXD2X4 | 刮水器关键词.xlsx |
| 猫砂垫 | fr | B0CH4N8V6P | 同 ASIN | B0CH4N8V6P | 猫砂垫关键词7.29.xlsx |
| 反光衣 | fr | B0GCDF56DJ B0GCF4T6NM B0GCFNSKDS | 同 ASIN | B0GCDF56DJ | 反光衣关键词.xlsx |

### 结果格式

`9` = 自然位第 9；`/` = 未找到；`/（广告1）` = 未找到自然位、广告在第 1 页；`9（广告1）` = 自然位第 9、同时广告在第 1 页。

### 特性

- 进度条 + 断点恢复（`.progress.json`）
- 模拟人工浏览节奏（随机延迟、分段滚动）
- WebSocket 断线自动重连
- BSR 自动检测数据起始列 + 多语言提取（FR/DE/EN）+ 自动展开折叠区域
- CDP Chrome 需养号（登录 Amazon + 浏览 + 加购），否则无 Cookie 会导致广告不展示
- `--dry-run` 干跑验证
- `--reset` 忽略断点重新开始
- `--ad-asin` 指定广告位追踪 ASIN（默认追踪全部 `--asin`）
