# AardMiner

经营数据汇总分析工具 — Windows 桌面应用，从 Excel 数据源汇总各业态（保险/商写/酒店）经营数据，调用 DeepSeek API 进行 AI 分析，生成分析报告。

## 运行环境

- Windows 10+ (x64/x86)
- 安装有 Microsoft Excel（用于数据读写）
- [Aardio](https://www.aardio.com/) 运行环境（或直接使用 Publish 目录的 exe）

## 快速开始

1. 编辑 `project.toml`，指定数据源目录和输出文件路径
2. 在 `%USERPROFILE%\.dskey` 中配置 DeepSeek API key，格式：
   ```
   EXCEL=sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
   ```
3. 在 Aardio IDE 中打开 `scr/default.aproj`，点击运行；或直接运行 `scr/Publish/AardMiner.exe`

## 功能

- **数据汇总**: 从各公司原始 Excel 报表中提取汇总数据
- **业态分析**: 按保险/商写/酒店三个板块进行 AI 分析
- **公司分析**: 逐公司进行核心经营指标 AI 分析

## 配置说明

`project.toml` 关键配置项：

| 配置项 | 说明 |
|--------|------|
| `data.source_path` | 各公司原始 Excel 报表所在目录 |
| `data.output_file` | 汇总分析结果 xlsx 输出路径 |
| `report.target_month` | 报告月份（0=自动检测） |
| `report.target_year` | 报告年份（0=自动检测） |
| `analysis.model` | DeepSeek 模型名，如 `deepseek-chat` |

## 项目结构

```
scr/
  main.aardio          — 程序入口与 UI
  project.toml          — 运行配置
  lib/                  — 业务逻辑库
  resources/            — 公司列表与 AI prompt 模板
  Publish/              — 编译产物
```
