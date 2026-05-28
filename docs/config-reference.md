# 配置参考

## 配置文件位置

`project.toml` — 与 `AardMiner.exe` 同级目录。

程序启动时按以下顺序查找配置：
1. 应用程序根目录 `/project.toml`（IDE 中=项目目录，发布后=EXE 目录）
2. 当前工作目录 `./project.toml`
3. 上级目录 `../project.toml`

## 完整配置项

### `[data]` — 数据路径

```toml
[data]
source_path = "C:\Users\Administrator\Desktop\2026年4月分析"
output_file = "C:\Users\Administrator\Desktop\2026年4月分析\【2026年4月】经营数据.xlsx"
```

| 配置项 | 类型 | 说明 |
|--------|------|------|
| `source_path` | string | 数据源根目录，应包含 `经营报表/` 和 `活动量/` 子文件夹 |
| `output_file` | string | 汇总分析输出 xlsx 文件路径 |

### `[report]` — 报告参数

```toml
[report]
target_month = 4
target_year = 2026
```

| 配置项 | 类型 | 默认值 | 说明 |
|--------|------|--------|------|
| `target_month` | int | `0` | 报告月份（1-12）。设为 `0` 自动检测 |
| `target_year` | int | `0` | 报告年份。设为 `0` 自动检测 |

自动检测优先级：
1. 输出文件 `填写页` sheet 的 A2（月份）、A4（年份）
2. 从 `source_path` 目录名解析（如 `2026年4月`）
3. 系统当前日期

### `[analysis]` — AI 分析参数

```toml
[analysis]
model = "deepseek-chat"
sector_timeout_ms = 30000
sector_max_tokens = 1000
company_timeout_ms = 60000
company_max_tokens = 1500
company_max_retries = 2
company_batch_size = 3
company_batch_delay_ms = 1000
quality_score_threshold = 8
```

| 配置项 | 类型 | 默认值 | 说明 |
|--------|------|--------|------|
| `model` | string | `deepseek-chat` | DeepSeek 模型名。也支持 `deepseek-reasoner` 等 |
| `sector_timeout_ms` | int | `30000` | 板块分析单次 API 请求超时（毫秒） |
| `sector_max_tokens` | int | `1000` | 板块分析单次 API 最大输出 token |
| `company_timeout_ms` | int | `60000` | 公司分析单次 API 请求超时（毫秒） |
| `company_max_tokens` | int | `1500` | 公司分析单次 API 最大输出 token |
| `company_max_retries` | int | `2` | 公司分析质量不合格时的最大重试次数 |
| `company_batch_size` | int | `3` | 每批处理公司数，批次间会暂停 |
| `company_batch_delay_ms` | int | `1000` | 批次间暂停时间（毫秒），用于 API 限流控制 |
| `quality_score_threshold` | int | `8` | 公司分析质量评分阈值（0-10），低于此值触发重试 |

## API Key 配置

API Key **不**在 `project.toml` 中配置，而是从 `%USERPROFILE%\.dskey` 读取。

### 文件位置

```
C:\Users\<用户名>\.dskey
```

### 文件格式

```
# 注释行以 # 开头
EXCEL=sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
```

- 大小写：`EXCEL=` 和 `excel=` 均可识别
- 仅读取第一个匹配行
- 支持 `#` 开头的注释行

### 多程序共享

如果安装了多个使用 DeepSeek API 的工具，可在一行写一段：

```
EXCEL=sk-xxx
OTHER_KEY=abc
```

AardMiner 只读取 `EXCEL=` 段。

## 公司列表配置

公司列表来自 `resources/companies.json`（发布时嵌入 exe，也可放在 exe 同级目录覆盖内嵌版本）。

```json
{
    "insurance": [
        {"name": "盛唐融信"},
        {"name": "君康经纪"}
    ],
    "commercial": [
        {"name": "北京中言"},
        {"name": "大连凯丹"},
        {"name": "福建钱隆"},
        {"name": "春夏秋冬"},
        {"name": "重庆宜新"}
    ],
    "hotel": [
        {"name": "伯豪瑞廷", "is_bhrt": true},
        {"name": "重庆瑞尔", "is_bhrt": false}
    ]
}
```

各业态公司按此 JSON 中的顺序处理。酒店业态中 `is_bhrt: true` 表示伯豪瑞廷（使用 E 列起始的达成列偏移）。

## Prompt 模板

位于 `resources/prompts/` 目录：

| 文件 | 用途 |
|------|------|
| `保险分析.md` | 保险板块系统提示词 |
| `商写分析.md` | 商写板块系统提示词 |
| `酒店分析.md` | 酒店板块系统提示词 |
| `经营分析.md` | 公司核心指标分析系统提示词 |

发布时嵌入 exe，也可放在 exe 同级 `resources/prompts/` 目录覆盖内嵌版本。

## 配置保存

在界面中修改配置项后，点击任何执行按钮都会自动保存回 `project.toml`，同时会将月份/年份写入输出文件的 `填写页` sheet（A2/A4 单元格）。
