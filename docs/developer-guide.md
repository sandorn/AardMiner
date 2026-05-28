# 开发者文档

## 环境搭建

### 必需工具

| 工具 | 用途 | 下载 |
|------|------|------|
| [Aardio](https://www.aardio.com/) | IDE + 编译器 | aardio.com |
| Git | 版本控制 | — |
| Microsoft Excel | 运行测试（COM 调用） | — |

### 克隆与打开

```bash
git clone <repo-url>
cd AardMiner
```

在 Aardio IDE 中打开 `scr/default.aproj`。

## 项目结构

```
AardMiner/
├── README.md              # 项目简介
├── CLAUDE.md              # AI 助手上下文
├── LICENSE                # 许可证
├── docs/                  # 文档
│   ├── README.md          # 文档索引
│   ├── user-guide.md      # 用户手册
│   ├── developer-guide.md # 本文档
│   ├── architecture.md    # 架构设计
│   ├── api-reference.md   # API 参考
│   ├── config-reference.md # 配置参考
│   └── CHANGELOG.md       # 变更记录
└── scr/
    ├── default.aproj      # Aardio 工程文件（版本号）
    ├── main.aardio        # 程序入口 + UI
    ├── project.toml        # 运行时配置
    ├── lib/               # 业务逻辑库
    │   ├── config/_.aardio       # 配置加载
    │   ├── models/_.aardio       # 数据模型
    │   ├── excel.aardio          # Excel COM 封装
    │   ├── common.aardio         # 公共工具
    │   ├── report_copy.aardio    # 经营报表复制
    │   ├── insurance.aardio      # 保险数据汇总
    │   ├── commercial.aardio     # 商写数据汇总
    │   ├── hotel.aardio          # 酒店数据汇总
    │   ├── deepseek_api.aardio   # DeepSeek API 客户端
    │   ├── sector_analysis.aardio # 板块 AI 分析
    │   └── company_analysis.aardio # 公司 AI 分析
    ├── resources/
    │   ├── app.ico               # 应用图标
    │   ├── companies.json        # 公司列表
    │   └── prompts/              # AI prompt 模板
    │       ├── 保险分析.md
    │       ├── 商写分析.md
    │       ├── 酒店分析.md
    │       └── 经营分析.md
    └── Publish/
        └── AardMiner.exe         # 发布产物
```

## 核心设计原则

### print 重定向

Aardio 中 `print` 输出到控制台。本项目的 `lib/*.aardio` 使用 `..print()` 写日志。`main.aardio` 在启动时覆盖全局 `print`，将所有输出重定向到 UI 日志框：

```aardio
..print = function(msg) {
    log(tostring(msg));
}
```

因此库文件中的 `..print("xxx")` 实际输出到 UI 日志并自动添加时间戳。

### 配置加载策略

所有资源文件均采用"磁盘优先，资源回退"策略：

```aardio
// string.load 优先读磁盘，找不到自动从内嵌资源加载
var content;
try { content = ..string.load("/resources/prompts/保险分析.md"); }
catch(e) { content = null; }
```

这使得：
- IDE 开发时可直接编辑磁盘上的 `resources/` 文件
- 发布后文件嵌入 exe，无需携带额外资源文件
- 用户可放置同名文件到 exe 同级目录覆盖内嵌版本

### Excel COM 管理

- `excel.init()`: 无头模式打开 Excel（Visible=false, DisplayAlerts=false, ScreenUpdating=false）
- `openWorkbook / closeWorkbook`: 维护 `openedWbs` 跟踪列表
- `close()`: 逆序关闭所有工作簿，最后 Quit Excel
- 窗口 `onClose` 回调中确保 Excel 清理

### 错误处理约定

- 顶层（按钮事件）：try/catch 防止崩溃
- 中层（汇总模块 run）：检查前置条件（文件存在/Sheet 存在）
- 底层（单公司处理）：失败返回 false，不影响其他公司

### 命名约定

```
模块函数：  _privateFunction（局部变量）
公开 API：   PublicFunction（命名空间成员）
配置常量：  _INS_SOURCE_SHEET（下划线前缀大写）
API 常量：   _CONTENT_PREFIX（下划线前缀大写）
```

## 添加新业态

1. 在 `resources/companies.json` 的对应分组中添加公司
2. 数据汇总：参考 `insurance.aardio` / `commercial.aardio` 模式新建 lib 文件
3. 导入 `main.aardio` 并在 `_runPhase1` 中调用
4. AI 分析：在 `sector_analysis.aardio` 中添加 `runXxx` 函数
5. 在 `resources/prompts/` 中添加对应的提示词 `.md` 文件
6. 在 `config/_.aardio` 的 `loadPrompts` 中注册新的 prompt key

## 编译与发布

### IDE 发布

1. 在 Aardio IDE 中打开 `scr/default.aproj`
2. 工具栏点击 **发布** 按钮
3. 产物输出到 `scr/Publish/AardMiner.exe`

### 版本号管理

版本号在 `scr/default.aproj` 的 XML 属性中：

```xml
<project FileVersion="0.0.2.0" ProductVersion="0.0.2.0" ...>
```

发布前更新这两个属性。

### 发布注意事项

- `local="false"` — 不在本地保存依赖库副本
- `libEmbed="true"` — 库文件嵌入 exe
- `embed="true"` — 资源文件嵌入 exe
- Publish 目录不应提交到 Git（已在 `.gitignore` 中）

## 调试

### 本地调试

在 Aardio IDE 中 F5 运行，可直接查看 UI 窗口和控制台输出。

### 调试 API 调用

在 `CallDeepSeekAPI` 中设置 `debugMode=true`，可在日志中看到：
- 请求体长度
- 每次重试的请求时间
- API 响应长度
- 错误详情（前 500 字符）

### 日志解读

日志格式：`HH:MM:SS  内容`

关键日志标记：
```
=== 阶段1: 数据汇总 ===      ← 阶段开始
  [1/9] 盛唐融信 ...          ← 单公司处理
    [跳过] ...                 ← 非致命跳过（文件不存在等）
    完成                       ← 成功
  [失败] ...                   ← 单公司失败
错误: ...                      ← 致命错误
```

## XLSX 模板结构

输出文件是一个 xlsx 工作簿，包含以下 Sheet：

| Sheet | 用途 |
|-------|------|
| 填写页 | 全局配置（A2=月份, A4=年份, A8=序时进度, C2:C10=公司名, C20=系统提示词） |
| 盛唐融信 ~ 重庆瑞尔 | 各公司数据（经营报表 + G2:R18 汇总数据 + C61 分析结果） |
| 保险类 | 保险汇总（C/D 列=公司, G/H 列=月度保费）+ L14 分析结果 |
| 商写类 | 商写汇总（C-G 列=公司） + L14 分析结果 |
| 酒店类 | 酒店汇总（C/D=营销, F/G=OTA, J/K=入住率） + M14 分析结果 |

## 常见开发问题

### 为什么用 string.indexOf 而不是 == 比较字符串？

aardio 的 `==` 对字符串数组中的元素行为不一致。代码中使用 `..string.indexOf(line, "关键词")` 来检测字符串包含关系，返回非零表示匹配。

### 为什么 ExtractSingleNumber 用 byte 比较？

aardio 中 `s[i]` 返回字节值（number 类型），不能用字符字面量 `'0'`（string 类型）比较，必须使用 ASCII 码值（48-57 对应 `0`-`9`）。

### 为什么保险汇总的续期保费用 writeCellRaw？

续期保费可能为空（该月无数据），此时需保留 null 语义以确保 Excel 公式正确处理。`writeCell` 会将其转换为 0，而 `writeCellRaw` 写入空字符串。
