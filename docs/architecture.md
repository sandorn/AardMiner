# 架构设计

## 概述

AardMiner 是一个单文件 Aardio 桌面应用，通过 COM 接口操作 Excel，通过 HTTP 调用 DeepSeek API。架构清晰，按模块化设计。

## 技术选型

| 层级 | 技术 | 理由 |
|------|------|------|
| UI | Aardio winform（原生 Win32） | Aardio 内置，无需额外依赖 |
| Excel 操作 | COM（`com.CreateObject("Excel.Application")`） | Windows 原生，稳定可靠 |
| HTTP 客户端 | Aardio `inet.http` | 内置库 |
| AI API | DeepSeek Chat Completions API | 兼容 OpenAI 格式 |
| 配置 | TOML（手写解析器） | 人类可读，手动编辑友好 |
| 公司列表 | JSON（Aardio 内置解析） | 结构化数据 |
| 版本管理 | Aardio 工程文件 `.aproj` | Aardio 标准 |

## 系统架构图

```
┌──────────────────────────────────────────────────┐
│                   main.aardio                     │
│  ┌─────────┐  ┌─────────┐  ┌─────────────────┐  │
│  │ UI 布局  │  │ 配置加载 │  │ 三阶段编排      │  │
│  └─────────┘  └─────────┘  └─────────────────┘  │
├──────────────────────────────────────────────────┤
│                    lib/                           │
│  ┌────────────┐  ┌────────────┐  ┌────────────┐ │
│  │ config/    │  │ models/    │  │ excel      │ │
│  │ TOML 解析   │  │ 数据模型    │  │ COM 封装   │ │
│  └────────────┘  └────────────┘  └────────────┘ │
│  ┌────────────┐  ┌────────────┐  ┌────────────┐ │
│  │ common     │  │report_copy │  │ deepseek_  │ │
│  │ 工具函数    │  │ 报表复制    │  │ api        │ │
│  └────────────┘  └────────────┘  └────────────┘ │
│  ┌────────────┐  ┌────────────┐  ┌────────────┐ │
│  │ insurance  │  │ commercial │  │ hotel      │ │
│  │ 保险汇总    │  │ 商写汇总    │  │ 酒店汇总   │ │
│  └────────────┘  └────────────┘  └────────────┘ │
│  ┌────────────────────┐  ┌────────────────────┐ │
│  │ sector_analysis    │  │ company_analysis   │ │
│  │ 板块 AI 分析        │  │ 公司 AI 分析        │ │
│  └────────────────────┘  └────────────────────┘ │
├──────────────────────────────────────────────────┤
│                  resources/                       │
│  ┌────────────────┐  ┌────────────────────────┐  │
│  │ companies.json │  │ prompts/*.md (4 个)    │  │
│  └────────────────┘  └────────────────────────┘  │
└──────────────────────────────────────────────────┘
```

## 模块职责

### `main.aardio` — 入口与编排

**职责**：UI 创建、配置加载、三阶段编排、日志路由。

```aardio
// 关键职责
1. 创建 winform UI（9 个控件）
2. 将全局 print 重定向到 UI 日志框
3. 多级回退查找 project.toml
4. 编排 _runPhase1 → _runPhase2 → _runPhase3
5. 配置双向同步：TOML ↔ UI 编辑框
```

### `lib/config/_.aardio` — 配置加载

**职责**：TOML 解析、公司列表加载、API Key 读取、Prompt 加载。

```
loadAll(configPath) → AppConfig
  ├── loadProjectConfig()    手动 TOML 解析（按 section → key 匹配）
  ├── loadCompanies()        JSON 解析 companies.json
  ├── loadApiKey()           读取 %USERPROFILE%\.dskey
  └── loadPrompts()          加载 4 个 .md prompt 文件
```

配置加载策略：`string.load` 优先读磁盘文件，找不到自动从内嵌资源加载（适用于发布版）。

### `lib/models/_.aardio` — 数据模型

```
Company        — 公司（name, sector, isBhrt, sourceFile, targetCol）
SectorResult   — 板块分析结果（content, success, errorMsg）
CompanyResult  — 公司分析结果（content, qualityScore, 5 维度标志, retriesUsed）
AppConfig      — 完整配置容器（所有配置项 + 公司分组 + prompts）
```

### `lib/excel.aardio` — Excel COM 封装

**职责**：所有 Excel COM 操作的一层薄封装。

```
初始化:  init() → com.CreateObject("Excel.Application")
读写:    openWorkbook / readCell / writeCell / readRange / writeRange
高级:    copyRangeValues / writeFormula / writeNA / setNumberFormat
清理:    closeWorkbook / close (含跟踪列表，逆序关闭避免索引问题)
```

设计要点：
- Visible=False, DisplayAlerts=False, ScreenUpdating=False 最大化静默性能
- 维护 `openedWbs` 跟踪列表，确保 close() 时全部清理
- `writeCellRaw` 对 null 写空字符串（保留空值语义）
- `writeNA` 写 `=NA()` 公式（表示"不适用"）

### `lib/common.aardio` — 公共工具

**职责**：数字提取、YTD 求和、月份/年份检测、月份标签解析。

```
ParseNumeric         — Variant → 纯数字（三层回退）
ExtractSingleNumber  — 文本中提取首个连续数字（酒店业态）
ExtractNumberFromCell — 增强版提取（含 "+" 求和逻辑）
SafeRead             — 安全读取（基于 ParseNumeric）
SumAchievementCols   — YTD 达成列求和（col = 2*m + offset）
GetTargetMonth       — 月份检测（填写页 → 文件夹名 → 系统月份）
GetYearFromConfig    — 年份检测（填写页 → 系统年份）
GetProgressFromConfig — 序时进度（填写页 → month/12）
ParseMonthLabel      — "1月"→1, "12月"→12
```

### `lib/report_copy.aardio` — 经营报表复制

**职责**：从 `经营报表/{公司}.xlsx` 的 `指标统计!D4:O20` 纯值复制到输出文件同名 Sheet 的 `G2:R18`。

### `lib/insurance.aardio` — 保险数据汇总

**职责**：从 `活动量/{公司}.xlsx` 的 `保险类` Sheet 汇总 2 家公司，写入输出文件 `保险类` Sheet C/D 列。

汇总指标：
- 人力：期初、入职、离职、净增、月末、平均、开单人数
- 保费：新单规模保费、期交规模保费
- 续期：13月/25月应收/实收
- 承保件数
- 月度规模保费（12 个月，未来月份写 =NA()）

自动写入公式：活动率、件均保费、人均保费。

### `lib/commercial.aardio` — 商写数据汇总

**职责**：从 `活动量/{公司}.xlsx` 的 `写字楼和商业综合体类` Sheet 汇总 5 家公司，写入输出文件 `商写类` Sheet C-G 列。

汇总指标：
- 面积：期初、新增签约、本期退租、月末
- 渠道：带客、成交、签约面积、转化率
- 自营：带客、成交、签约面积、转化率
- 到期面积、续签面积

自动写入公式：渠道转化率、自营转化率。

### `lib/hotel.aardio` — 酒店数据汇总

**职责**：从双数据源汇总 2 家酒店。

- **营销活动**（来源：`活动量/{公司}.xlsx` `酒店类`）：
  - 伯豪瑞廷：三行分组（官微/抖音/OTA）× 三指标（投放/受众/成交），达成列 E 起始
  - 重庆瑞尔：单行 × 三指标，达成列 D 起始
- **业务指标**（来源：`经营报表/{公司}.xlsx` `指标统计`）：
  - 月均入住率（第 15 行）、OTA 网络评价（第 17 行），逐月写入
  - 伯豪瑞廷 → J/F 列，重庆瑞尔 → K/G 列

自动写入公式：营销转化率。

### `lib/deepseek_api.aardio` — DeepSeek API 客户端

**职责**：封装 DeepSeek Chat Completions API 调用。

```
CallDeepSeekAPI     — 核心调用（构建 HTTP 请求、指数退避重试、状态码判断）
BuildJSONRequest    — 构造请求 JSON
BuildDataPrompt     — 将二维数组转为格式化文本表格
ParseAPIResponseSimple — 解析 JSON 响应（支持 content/reasoning_content）
CleanAnalysisResult — 去空行、统一换行符
WriteAnalysisResult — 写入 Excel 单元格并设置格式
```

重试策略：
- 网络错误：指数退避（2^1, 2^2, 2^3 秒）
- HTTP 429/5xx：指数退避
- HTTP 401/402/403/422：不重试

### `lib/sector_analysis.aardio` — 板块 AI 分析

**职责**：读取输出文件的业态数据，调用 DeepSeek API，写入分析结果。

```
runAll → runCommercial → runInsurance → runHotel
```

三个业态流程一致：
1. 从目标 Sheet 读取数据区域
2. 构建数据提示词（静态前缀 + BuildDataPrompt 表格）
3. 调用 DeepSeek API（timeout=30s, tokens=1000, retries=3）
4. 解析响应，清理格式，写入结果单元格

保险板块特殊处理：`BuildInsuranceDataPrompt` 过滤"规模保费"区段中超出当前月份的行。

### `lib/company_analysis.aardio` — 公司 AI 分析

**职责**：逐公司读取填写页数据，调用 DeepSeek API，质量检查 + 自动重试。

```
runAll → [batch] → AnalyzeCompany → AnalyzeWithQualityCheck
                                           ├─ CallDeepSeekAPI
                                           └─ CheckAnalysisQuality
```

质量检查维度（共 10 分）：
- 经营综述（2 分）：是否有非指标关键词的综述段落
- 营业收入（2 分）
- EBITDA/扣非利润（2 分）
- 经营活动净现金流（2 分）
- 经营支出（2 分）

低于阈值（默认 8 分）自动重试，最多 `company_max_retries` 次。

批次控制：每处理 `company_batch_size` 家后暂停 `company_batch_delay_ms` 毫秒，避免 API 限流。

## 三阶段执行流程

```
阶段1: 数据汇总
  report_copy    → 9 家公司，经营报表纯值复制
  insurance      → 2 家公司，保险数据汇总 → 写入公式
  commercial     → 5 家公司，商写数据汇总 → 写入公式
  hotel          → 2 家公司，营销活动 + 业务指标 → 写入公式

阶段2: 板块 AI 分析
  runCommercial  → 读取商写类 A1:G18 → API → 写入 L14
  runInsurance   → 读取保险类 F1:H25 → API → 写入 L14
  runHotel       → 读取三区域 → API → 写入 M14

阶段3: 公司 AI 分析
  逐家（9 家公司，分批）:
    → 读取 C1:R5 → 构建提示词（含年份/月份/序时进度）
    → API 调用 + 质量检查 + 重试
    → 写入 C61
```

## 数据流

```
原始 Excel (经营报表/*.xlsx, 活动量/*.xlsx)
          │
          ▼ COM Read
  [main.aardio 三阶段编排]
          │
          ▼ COM Write
输出 Excel ({月份}经营数据.xlsx)
          │ COM Read (阶段2/3)
          ▼
  [DeepSeek API]
          │
          ▼ COM Write
输出 Excel (同一文件，写入分析结果)
```

## 错误处理策略

| 层级 | 策略 |
|------|------|
| 顶层按钮事件 | try/catch 包裹，异常显示为错误日志，确保按钮重新启用 |
| 阶段编排 | 每个汇总模块独立 try/catch（一个失败不影响其他） |
| 公司级别 | 逐家处理，单家失败继续下一家 |
| API 调用 | 指数退避重试（网络错误/5xx），认证错误直接返回 |
| Excel 关闭 | 窗口 onClose 中 finally 式清理 |
