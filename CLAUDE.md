# AardMiner — 经营数据汇总分析工具

基于 Aardio 的 Windows 桌面应用，从 Excel 数据源汇总各业态经营数据，调用 DeepSeek API 进行 AI 分析。

## 技术栈

- **语言/框架**: Aardio，编译为 Windows x86 原生 exe
- **项目文件**: `scr/default.aproj`（Aardio 工程文件，版本号在此定义）
- **配置**: `project.toml`（运行时加载），API key 从 `%USERPROFILE%\.dskey` 的 `EXCEL=` 段读取
- **输出**: Excel xlsx 工作簿，通过 COM 操作

## 目录结构

```
scr/
  main.aardio          — 入口，UI 布局，三阶段编排
  project.toml          — 运行配置（数据路径、月份、模型参数）
  default.aproj         — Aardio 工程文件（版本号）
  lib/
    config/_.aardio     — TOML 解析、公司列表加载、API key 加载、prompt 加载
    models/_.aardio     — AppConfig/Company 数据模型
    excel.aardio        — Excel COM 操作封装
    common.aardio       — 公共工具（月份/年份检测等）
    report_copy.aardio  — 经营报表复制
    insurance.aardio    — 保险数据汇总
    commercial.aardio   — 商写数据汇总
    hotel.aardio        — 酒店数据汇总
    deepseek_api.aardio — DeepSeek API 客户端
    sector_analysis.aardio — 板块 AI 分析
    company_analysis.aardio — 公司 AI 分析
  resources/
    companies.json      — 公司列表（按业态分组）
    prompts/            — AI prompt 模板（保险/商写/酒店/经营分析四个 .md）
  Publish/AardMiner.exe — 发布产物
```

## 三阶段执行流程

1. **数据汇总**: report_copy → insurance → commercial → hotel
2. **板块 AI 分析**: 商写 → 保险 → 酒店
3. **公司 AI 分析**: 9 家公司逐家分析

## 配置项（project.toml）

`[data]`: source_path（数据源目录）, output_file（输出 xlsx 路径）
`[report]`: target_month, target_year（0=自动检测）
`[analysis]`: model, sector_timeout_ms, sector_max_tokens, company_timeout_ms, company_max_tokens, company_max_retries, company_batch_size, company_batch_delay_ms, quality_score_threshold

## 注意事项

- API key 文件路径 `%USERPROFILE%\.dskey`，格式: `EXCEL=sk-xxx`
- Aardio 中 print 被重定向到 UI 日志编辑框
- `config.loadAll()` 优先读磁盘文件，找不到自动从内嵌资源加载（适用于发布版）
- 版本号在 `scr/default.aproj` 的 `FileVersion` / `ProductVersion` 属性中
