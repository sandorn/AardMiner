# API 参考

## DeepSeek API 客户端

封装文件：[`scr/lib/deepseek_api.aardio`](../scr/lib/deepseek_api.aardio)

### 核心函数

#### `CallDeepSeekAPI`

调用 DeepSeek Chat Completions API。

```aardio
CallDeepSeekAPI(
    prompt,          // string — 用户提示词
    apiKey,          // string — API Key
    apiBaseUrl,      // string — API 基础 URL，默认 "https://api.deepseek.com"
    model,           // string — 模型名，默认 "deepseek-chat"
    temperature,     // number — 温度参数，默认 0.3
    maxTokens,       // int — 最大输出 token，默认 1000
    timeoutMs,       // int — 超时时间（毫秒），默认 30000
    contextName,     // string — 上下文名称（用于日志），默认 ""
    debugMode,       // bool — 调试模式，默认 false
    maxRetries,      // int — 最大重试次数，默认 3
    systemMsg        // string — 系统提示词，默认 ""
)
```

**返回值**：
- 成功：返回 API 原始 JSON 响应字符串
- 失败：返回 `[contextName] 错误类别：详情` 格式的错误信息

**错误分类**：
| HTTP 状态码 | 是否重试 | 说明 |
|-------------|----------|------|
| 200 | N/A | 成功 |
| 401, 402, 403 | 否 | 认证/授权错误 |
| 422 | 否 | 请求参数错误 |
| 429 | 是 | 限流 |
| 5xx | 是 | 服务器错误 |
| 网络错误 | 是 | 连接超时/DNS/SSL |

**重试策略**：指数退避，等待时间 = 2^retryCount 秒（2s, 4s, 8s）。

#### `ParseAPIResponseSimple`

从 Chat Completions JSON 响应中提取 content 文本。

```aardio
ParseAPIResponseSimple(apiResponse)  // string → string
```

提取顺序：
1. `choices[0].message.content` — 常规模型
2. `choices[0].message.reasoning_content` — 推理模型（如 deepseek-reasoner）
3. 若两者皆空，返回 `"未找到content字段"`

若响应包含 `error` 对象，返回 `"API错误: {message}"`。

#### `BuildJSONRequest`

构造 Chat Completions API 请求 JSON。

```aardio
BuildJSONRequest(systemMsg, prompt, model, temperature, maxTokens)  // → string (JSON)
```

请求结构：
```json
{
    "model": "deepseek-chat",
    "messages": [
        {"role": "system", "content": "..."},
        {"role": "user", "content": "..."}
    ],
    "temperature": 0.3,
    "max_tokens": 1000
}
```

#### `BuildDataPrompt`

将二维数组转为带表头的格式化文本表格，用于嵌入用户提示词。

```aardio
BuildDataPrompt(dataArray)  // table[][] → string
```

输入：`{{"指标", "盛唐", "君康"}, {"保费", 123, 456}}`
输出：
```
数据表格如下：
 指标   盛唐   君康
 保费   123    456
```

#### `CleanAnalysisResult`

清理 API 输出中的空行，统一换行符为 `\r\n`。

```aardio
CleanAnalysisResult(result)  // string → string
```

#### `WriteAnalysisResult`

将分析结果写入 Excel 单元格并设置格式。

```aardio
WriteAnalysisResult(ws, cellAddr, result, debugMode=false, contextName="")
```

写入后设置：
- WrapText = true（自动换行）
- HorizontalAlignment = xlHAlignLeft
- VerticalAlignment = xlVAlignTop

### 板块分析 API 调用参数

| 业态 | Sheet | 数据区域 | 结果单元格 | Timeout | Max Tokens | System Msg |
|------|-------|----------|-----------|---------|------------|------------|
| 商写 | 商写类 | A1:G18 | L14 | 30s | 1000 | `prompts.commercial` |
| 保险 | 保险类 | F1:H25 | L14 | 30s | 1000 | `prompts.insurance` |
| 酒店 | 酒店类 | 三区域* | M14 | 30s | 1000 | `prompts.hotel` |

*酒店三区域：B1:D5（营销活动）、E1:G13（OTA 评价）、I1:K13（入住率）

### 公司分析 API 调用参数

| 参数 | 配置项 | 默认值 |
|------|--------|--------|
| 超时 | `company_timeout_ms` | 60s |
| Max Tokens | `company_max_tokens` | 1500 |
| Temperature | 硬编码 | 0.3 |
| System Msg | `prompts.financial` | 经营分析.md |
| 质量阈值 | `quality_score_threshold` | 8 |
| 最大重试 | `company_max_retries` | 2 |
| 批次大小 | `company_batch_size` | 3 |
| 批次延迟 | `company_batch_delay_ms` | 1s |

数据提示词格式：
```
公司名称：{公司名}
年份：{2026}
当前月份：{4月}
序时进度：{33.33}%
数据单位：万元
请按系统提示词要求输出指定格式。
数据表格如下：
...（C1:R5 区域数据）...
```

### 质量检查算法

`CheckAnalysisQuality` 函数按五个维度评分：

| 维度 | 分值 | 检测方式 |
|------|------|----------|
| 经营综述 | 2 | 首行不含特定指标关键词 |
| 营业收入 | 2 | 包含"营业收入" |
| EBITDA/扣非利润 | 2 | 包含"EBITDA"或"扣非利润" |
| 现金流 | 2 | 包含"经营活动净现金流"或"现金流" |
| 经营支出 | 2 | 包含"经营支出" |

总分 0-10。低于阈值触发质量重试，最终仍低于阈值时在结果末尾附加质量提示。
