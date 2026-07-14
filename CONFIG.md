# AutoLinker AI 配置指南

> **重要：强烈推荐使用易语言 5.95 版本，以获得 AutoLinker 的最佳兼容性和完整功能体验。其他版本仅建议用于兼容性测试。**
>
> 面向易语言开发者的 AI 功能配置说明，零基础也能看懂。

---

## 目录

- [基础概念：搞清楚这几个词](#基础概念搞清楚这几个词)
- [配置界面各项说明](#配置界面各项说明)
- [GPT-5.6 模型家族](#gpt-56-模型家族)
- [各平台配置方法](#各平台配置方法)
- [源码编辑模式](#源码编辑模式)
- [Tavily 联网搜索（可选）](#tavily-联网搜索可选)
- [外部 Agent MCP 配置](#外部-agent-mcp-配置)
- [常见问题](#常见问题)

---

## 基础概念：搞清楚这几个词

在配置 AI 之前，先了解几个关键概念，后面看配置项就一目了然了。

### 🔑 什么是 API 密钥（API Key）？

API Key 就是一串密钥，相当于你访问 AI 服务的"账号密码"。  
你在某个 AI 平台注册并充值后，平台会给你一个 Key，拿这个 Key 就能调用 AI。

> 💡 **Key 要保管好，不要泄露给别人，否则别人会消耗你的余额。**

---

### 🌐 什么是中转站（代理接口）？

中转站本质上是一个"二道贩子"，但在 AI 领域是合理且常见的存在。目前主流中转站的来源有两种：

**① 转发型：解决国内访问问题**

```
你的程序 → 中转站服务器（国内可访问）→ 海外 AI（OpenAI / Claude 等）
```

由于 OpenAI、Claude 等海外服务在国内访问困难，中转站用自己的服务器做一层转发，让你不用 VPN 就能调用。

**② 逆向型：把"订阅会员用量"转成 API**

这类中转站更有意思。Codex、Claude Code、Gemini CLI 等 AI 编程工具向开发者提供订阅会员，会员内包含大量的模型调用额度。  
中转站通过逆向分析这些工具的通信协议，将会员订阅的用量包装成标准 API 对外出售。

```
中转站购买大批会员订阅 → 逆向协议 → 转为 API 对外售卖
        ↓
你按 token 付费，成本极低（相比官方 API 可低 10 倍以上）
```

这也是为什么同样是 GPT-5.6 / Claude Opus 4.8，通过某些中转站调用的价格会比官方 API 便宜很多，本质上是在消耗会员订阅的额度。

> 📌 **协议注意：** 逆向 Codex 等工具的中转站通常使用 **OpenAI Responses** 协议（即 `/responses` 端点），但不同平台的兼容方式并不相同。优先使用“使用预设站点新建”，让界面自动选择协议；当前 OpenAI 官方预设使用 `OpenAI Responses`，Right 预设使用 `OpenAI Chat`。手动配置时必须以服务商文档为准。

**使用中转站的好处：**
- 国内网络直接可用，无需 VPN
- 聚合多个 AI 平台，一个 Key 可调用多种模型
- 价格远低于官方 API（尤其是逆向型中转站）

> ⚠️ 中转站是第三方服务，质量参差不齐，建议选口碑好、有一定用户量的平台。逆向型中转站也存在随时被封禁的风险，请酌情使用。

---

### 🤖 什么是接口地址（Base URL）？

接口地址就是调用 AI 时的服务地址，也常被称为 Base URL。

- 官方 OpenAI 的地址是 `https://api.openai.com/v1`
- 中转站会有自己的地址，如 `https://right.codes/codex`
- 不同平台的地址不同，但通过“**使用预设站点新建**”创建配置组后会自动填入，无需手动查找

---

### 💬 什么是模型？

模型就是具体的 AI，不同模型能力和价格不同。例如：

| 模型名 | 特点 |
|---|---|
| `gpt-5.6` | OpenAI 最新系列的默认别名，自动指向旗舰版 `gpt-5.6-sol` |
| `gpt-5.6-terra` | 兼顾能力与成本 |
| `gpt-5.6-luna` | 适合高并发、高用量场景 |
| `deepseek-chat` | 国内模型，价格实惠 |
| `claude-sonnet-4-6` | 擅长理解和写作 |

> 💡 不知道选哪个？可以点击“使用预设站点新建”，从站点当前提供的模型中选择。

---

### 🧠 什么是思考等级？

部分高级模型支持深度思考，会在给出答案前先进行推理分析。AutoLinker 当前提供七档设置：

| 界面选项 | 配置值 | 说明 |
|---|---|---|
| 关闭 | `off` | 关闭或尽量不使用额外推理，速度最快 |
| 低 | `low` | 轻度思考，适合简单任务 |
| 中 | `medium` | 能力、速度与费用较均衡，推荐作为 GPT-5.6 的起点 |
| 高 | `high` | 增加推理量，适合较复杂任务 |
| 超高 | `xhigh` | 更深入地分析和验证，耗时与费用更高 |
| 最大 | `max` | 面向最困难、质量优先的任务 |
| 极限 | `ultra` | AutoLinker 的最高预设；对 GPT-5.6 发出的推理等级仍为 `max` |

> 💡 GPT-5.6 官方支持 `none`、`low`、`medium`、`high`、`xhigh` 和 `max`。不同协议或模型支持的等级不同，AutoLinker 会按模型能力转换参数；部分模型会把“超高”“最大”“极限”统一降为其支持的最高档。若接口报参数不支持，请降低等级或选择“关闭”。

---

## 配置界面各项说明

打开“AutoLinker AI 设置”后，会看到“模型服务配置组”和“其他设置”两个区域：

| 字段 | 说明 | 是否必填 |
|---|---|---|
| **配置组** | 保存并切换多套模型服务配置，可新建、使用预设站点新建、重命名或删除 | — |
| **接口协议** | 可选 `OpenAI Chat`、`OpenAI Responses`、`Gemini` 或 `Claude` | ✅ |
| **接口地址** | 模型服务地址；界面下方会预览最终请求地址 | ✅ |
| **API 密钥** | 从模型平台获取的密钥 | ✅ |
| **模型** | 使用的模型名称；可直接填写、从候选列表选择，或点击 `↻` 从接口获取模型列表 | ✅ |
| **上下文长度** | 达到该长度的 95% 时自动压缩历史；留空时按模型名自动判断 | 可选 |
| **思考等级** | 控制模型的推理强度，详见上方七档说明 | 可选 |
| **系统提示词** | 附加到 AutoLinker 内置提示词后的自定义要求 | 可选 |
| **自定义请求头** | 每行填写一个“请求头名称: 值”，可覆盖内置请求头 | 可选 |
| **源码编辑模式** | 全局设置，可选“真实页优先”或“解包镜像基准（测试）” | 可选 |
| **Tavily API 密钥** | 全局联网搜索密钥，独立于模型服务配置组 | 可选 |

### 配置组（多套配置）

配置组功能让你可以保存多套模型服务配置，方便切换。例如：
- 一套用国内的 DeepSeek，日常省钱用
- 一套用海外的 GPT，处理复杂任务用

点击“**新建**”可创建空白配置组；点击“**使用预设站点新建**”并选择具体的“站点（模型）”，会自动创建配置组并填好接口协议、接口地址和模型，只需再填写 API 密钥。还可以对当前配置组执行“**重命名**”或“**删除**”（至少保留一个配置组）。

> 💡 “使用预设站点新建”取代了旧版界面的“平台预设”字段。切换配置组不会改变“源码编辑模式”和“Tavily API 密钥”，因为这两项是全局设置。

---

## GPT-5.6 模型家族

GPT-5.6 是 OpenAI 当前最新模型家族。AutoLinker 的 OpenAI 预设已提供以下四个模型名，并会把它们的上下文长度自动识别为 **1,050,000 tokens**：

| 模型 | 适用场景 |
|---|---|
| `gpt-5.6` | 默认别名，当前自动指向 `gpt-5.6-sol` |
| `gpt-5.6-sol` | 旗舰能力，适合复杂编程、分析和高质量生产任务 |
| `gpt-5.6-terra` | 在能力与成本之间取得平衡 |
| `gpt-5.6-luna` | 更高效率，适合高并发和大用量任务 |

建议使用 **OpenAI Responses** 协议，并从“中”思考等级开始。复杂任务可逐步提高到“高”“超高”或“最大”；“极限”对 GPT-5.6 发出的官方推理等级也是 `max`，不是一个独立的 OpenAI 推理档位。

> 📖 官方资料：[Using GPT-5.6](https://developers.openai.com/api/docs/guides/latest-model)

---

## 各平台配置方法

> 💡 **推荐流程：** 点击“使用预设站点新建” → 选择“站点（模型）” → 填写 API 密钥 → 选择思考等级 → 点击“测试连通性” → 保存

---

<details>
<summary><b>🌟 Right（推荐中转站，国内直连，聚合多平台模型）</b></summary>

**简介：** Right.codes 是一个支持多平台模型的中转站，国内可直接访问，支持 GPT、Claude 等主流模型。

**注册地址：** https://right.codes/register

**配置步骤：**

1. 点击“**使用预设站点新建**”，选择一个 Right 预设：
   - `Right(gpt-5.3-codex)` — 代码任务
   - `Right(gpt-5.2)` — 通用任务
2. 界面会自动填入：
   - 接口地址：`https://right.codes/codex`
   - 接口协议：`OpenAI Chat`
3. 前往 Right 官网注册并充值，在控制台复制 API 密钥，填入“**API 密钥**”
4. 按需设置“**思考等级**”，“**上下文长度**”一般留空即可
5. 点击“**测试连通性**”确认配置正确
6. 点击“**保存并继续**”

</details>

---

<details>
<summary><b>🇨🇳 DeepSeek（国内直连，性价比高）</b></summary>

**简介：** DeepSeek 是国内领先的 AI 公司，模型能力强，价格相比海外模型有明显优势。

**官网：** https://platform.deepseek.com

**配置步骤：**

1. 点击“**使用预设站点新建**”，选择一个 Deepseek 模型
2. 自动填入：
   - 接口地址：`https://api.deepseek.com`
   - 接口协议：`OpenAI Chat`
3. 前往 DeepSeek 平台注册，在 API 管理页面创建密钥，填入“**API 密钥**”
4. 当前预设模型：
   - `deepseek-v4-flash` — 快速版
   - `deepseek-v4-pro` — 增强版
   - `deepseek-chat` — 通用对话
   - `deepseek-reasoner` — 深度推理
5. 测试并保存

> 💡 使用推理模型时可将“思考等级”设为“中”或“高”；若服务端不接受对应参数，请改为“关闭”。

</details>

---

<details>
<summary><b>🤖 智谱 GLM（国内直连，适合中文场景）</b></summary>

**简介：** 智谱 AI 出品的 GLM 系列模型，中文理解能力强。

**官网：** https://open.bigmodel.cn

**配置步骤：**

1. 点击“**使用预设站点新建**”，选择一个智谱模型
2. 自动填入：
   - 接口地址：`https://open.bigmodel.cn/api/paas/v4`
   - 接口协议：`OpenAI Chat`
3. 在智谱官网注册，前往“API 密钥”页面创建密钥，填入“**API 密钥**”
4. 当前预设模型：
   - `glm-5.2`
   - `glm-5-turbo`
   - `glm-4.7`
   - `glm-4.5-air`
5. 测试并保存

</details>

---

<details>
<summary><b>🌊 千问 / 通义（阿里云，国内直连）</b></summary>

**简介：** 阿里云旗下的通义千问，国内直连，有丰富的模型可选。

**官网：** https://dashscope.console.aliyun.com

**配置步骤：**

1. 点击“**使用预设站点新建**”，选择一个千问模型
2. 自动填入：
   - 接口地址：`https://dashscope.aliyuncs.com/compatible-mode/v1`
   - 接口协议：`OpenAI Chat`
3. 在阿里云 DashScope 控制台创建 API 密钥，填入“**API 密钥**”
4. 当前预设模型：
   - `qwen3.7-plus`
   - `qwen3.7-max`
   - `qwen3.6-flash`
   - `qwen3-coder-next`
   - `qwen3-coder-plus`
5. 测试并保存

</details>

---

<details>
<summary><b>🌙 Kimi（月之暗面，国内直连，长上下文）</b></summary>

**简介：** 月之暗面旗下 Kimi 模型，支持超长上下文，擅长文档分析。

**官网：** https://platform.moonshot.cn

**配置步骤：**

1. 点击“**使用预设站点新建**”，选择一个 Kimi 模型
2. 自动填入：
   - 接口地址：`https://api.moonshot.cn/v1`
   - 接口协议：`OpenAI Chat`
3. 在 Moonshot 平台创建 API 密钥，填入“**API 密钥**”
4. 当前预设模型：
   - `kimi-k2.7-code`
   - `kimi-k2.7-code-highspeed`
   - `kimi-k2.6`
   - `kimi-k2.5`
5. 测试并保存

</details>

---

<details>
<summary><b>🫘 豆包（字节跳动，国内直连）</b></summary>

**简介：** 字节跳动旗下豆包大模型，国内访问稳定。

**官网：** https://console.volcengine.com/ark

**配置步骤：**

1. 点击“**使用预设站点新建**”，选择一个豆包模型
2. 自动填入：
   - 接口地址：`https://ark.cn-beijing.volces.com/api/v3`
   - 接口协议：`OpenAI Chat`
3. 在火山引擎控制台创建 API 密钥，填入“**API 密钥**”

   > ⚠️ 注意：部分豆包账户或接口需要先创建“推理接入点”，并在“模型”中填写接入点 ID。请以火山引擎控制台的实际接入方式为准。

4. 当前预设模型：
   - `doubao-seed-2.0-pro`
   - `doubao-seed-2.0-code`
   - `doubao-seed-2.0-lite`
   - `doubao-seed-1.8`
5. 测试并保存

</details>

---

<details>
<summary><b>🔵 MiniMax（国内直连）</b></summary>

**简介：** MiniMax 旗下大模型，国内直连。

**官网：** https://platform.minimax.chat

**配置步骤：**

1. 点击“**使用预设站点新建**”，选择一个 MiniMax 模型
2. 自动填入：
   - 接口地址：`https://api.minimax.chat/v1`
   - 接口协议：`OpenAI Chat`
3. 在 MiniMax 平台注册并获取 API 密钥，填入“**API 密钥**”
4. 当前预设模型：
   - `MiniMax-M3`
   - `MiniMax-M2.7`
   - `MiniMax-M2.7-highspeed`
   - `MiniMax-M2.5`
5. 测试并保存

</details>

---

<details>
<summary><b>🌐 aihubmix（中转站，聚合主流海外模型）</b></summary>

**简介：** aihubmix 是支持 GPT、Claude、Gemini 等多个顶级模型的中转站，国内可访问，按量计费。

**注册地址：** https://aihubmix.com

**配置步骤：**

1. 点击“**使用预设站点新建**”，选择一个 aihubmix 模型
2. 自动填入：
   - 接口地址：`https://aihubmix.com/v1`
   - 接口协议：`OpenAI Chat`
3. 在 aihubmix 控制台充值并创建 API 密钥，填入“**API 密钥**”
4. 当前预设模型：
   - `gpt-5.5`
   - `claude-opus-4-8`
   - `claude-sonnet-4-6`
   - `deepseek-v4-pro`
   - `deepseek-v4-flash`
   - `gemini-3.1-pro-preview`
5. 测试并保存

</details>

---

<details>
<summary><b>🔥 硅基流动（中转站，聚合开源模型为主）</b></summary>

**简介：** 硅基流动以国产开源模型为主，提供 DeepSeek、Qwen 等模型的 API 接入，部分模型有免费额度。

**注册地址：** https://cloud.siliconflow.cn

**配置步骤：**

1. 点击“**使用预设站点新建**”，选择一个硅基流动模型
2. 自动填入：
   - 接口地址：`https://api.siliconflow.cn/v1`
   - 接口协议：`OpenAI Chat`
3. 在硅基流动控制台创建 API 密钥，填入“**API 密钥**”
4. 当前预设模型：
   - `deepseek-ai/DeepSeek-V4-Flash`
   - `deepseek-ai/DeepSeek-V4-Pro`
   - `Pro/zai-org/GLM-5`
   - `zai-org/GLM-5.1`
   - `Qwen/Qwen3.5-397B-A17B`

   > 💡 注意：硅基流动的模型名带有路径格式（如 `deepseek-ai/DeepSeek-V4-Pro`），填写时要完整。

5. 测试并保存

</details>

---

<details>
<summary><b>🌍 OpenAI 官方（需能访问海外网络）</b></summary>

**简介：** OpenAI 官方 API，需要海外网络或自备代理。按 token 计费，价格较高但模型质量有保证。

**官网：** https://platform.openai.com

**配置步骤：**

1. 点击“**使用预设站点新建**”，选择一个 OpenAI 模型
2. 自动填入：
   - 接口地址：`https://api.openai.com/v1`
   - **接口协议：`OpenAI Responses`**
3. 在 OpenAI Platform 充值后创建 API 密钥，填入“**API 密钥**”
4. 当前预设模型：
   - `gpt-5.6` — 默认指向 `gpt-5.6-sol`
   - `gpt-5.6-sol` — 旗舰能力
   - `gpt-5.6-terra` — 能力与成本平衡
   - `gpt-5.6-luna` — 高效率、高用量
   - `gpt-5.5`
   - `gpt-5.4`
   - `gpt-5.4-mini`
   - `gpt-5.3-codex`
5. 测试并保存

> ⚠️ 直接用官方 OpenAI 在国内需要能访问海外网络，如果访问困难建议使用上方的中转站。

</details>

---

<details>
<summary><b>🤍 Claude 官方（Anthropic，需能访问海外网络）</b></summary>

**简介：** Anthropic 官方 API，Claude 模型在理解和写作方面表现优秀，也擅长代码。国内访问需要代理。

**官网：** https://console.anthropic.com

**配置步骤：**

1. 点击“**使用预设站点新建**”，选择一个 Claude 模型
2. 自动填入：
   - 接口地址：`https://api.anthropic.com`
   - **接口协议：`Claude`**
3. 在 Anthropic 控制台获取 API 密钥，填入“**API 密钥**”
4. 当前预设模型：
   - `claude-opus-4-8`
   - `claude-sonnet-4-6`
   - `claude-haiku-4-5`
5. 测试并保存

> 💡 如果你在国内访问有问题，可以通过 aihubmix 或其他中转站使用 Claude 模型。

</details>

---

<details>
<summary><b>✨ Gemini 官方（Google，需能访问海外网络）</b></summary>

**简介：** Google 旗下 Gemini 模型，多模态能力强。国内访问需要代理。

**官网：** https://aistudio.google.com

**配置步骤：**

1. 点击“**使用预设站点新建**”，选择一个 Gemini 模型
2. 自动填入：
   - 接口地址：`https://generativelanguage.googleapis.com`
   - **接口协议：`Gemini`**
3. 在 Google AI Studio 获取 API 密钥，填入“**API 密钥**”
4. 当前预设模型：
   - `gemini-3.1-pro-preview`
   - `gemini-3.1-pro-preview-customtools`
   - `gemini-3.5-flash`
   - `gemini-3.1-flash-lite`
   - `gemini-2.5-pro`
5. 测试并保存

</details>

---

## 源码编辑模式

“源码编辑模式”位于“其他设置”区域，是所有模型服务配置组共用的全局设置：

| 模式 | 说明 |
|---|---|
| **真实页优先** | 编辑前读取易语言 IDE 中的真实页面，推荐日常使用 |
| **解包镜像基准（测试）** | 按 `read_file` 读取的镜像源码进行匹配并写回页面，仅建议测试或排查特定问题时使用 |

> 💡 不确定时保留默认的“真实页优先”。

---

## Tavily 联网搜索（可选）

Tavily 是一个 AI 搜索 API，配置后 AutoLinker 的 `search_web_tavily` 工具就能实时联网搜索信息。

**官网：** https://tavily.com（注册后有免费额度）

**配置步骤：**

1. 前往 Tavily 官网注册
2. 在控制台复制你的 API 密钥
3. 在配置界面的“其他设置”区域，填入“**Tavily API 密钥**”
4. 保存即可

> 💡 Tavily API 密钥是全局设置，不随模型服务配置组切换；它和主模型的 API 密钥没有关系，两者互不影响。
> 不需要联网搜索功能的话，这里留空也没关系。

---

## 外部 Agent MCP 配置

AutoLinker 在易语言 IDE 启动后会自动开启一个 **MCP（Model Context Protocol）服务**，监听本地端口 `19207`。

你可以将这个 MCP 服务配置到支持 MCP 的外部 AI 编码工具中（如 Cursor、Claude Code、Antigravity CLI 等），让这些外部 Agent 也能直接操作易语言 IDE，实现代码读写、项目浏览等功能。

> 💡 **前提条件：** 需要先打开易语言 IDE 并加载 AutoLinker，MCP 服务才会启动。外部 Agent 发送请求时 IDE 必须处于运行状态。

> ⚠️ **编码注意：** 通过 MCP 传递中文参数（如页面名称）时，请确保使用 **UTF-8** 编码，否则可能出现乱码或识别失败。

---

### MCP 服务信息

| 项目 | 值 |
|---|---|
| 协议 | Streamable HTTP；协商支持 MCP `2025-11-25`、`2025-03-26`、`2024-11-05` |
| 本地地址 | `http://127.0.0.1:19207/mcp` |
| 传输类型 | `http`（streamable-http） |

> 安全说明：服务只绑定 `127.0.0.1`，拒绝携带非空 `Origin` 的浏览器脚本请求，不开放通配 CORS；当前不提供 Bearer Token，因此请只连接可信的本机原生 MCP 客户端，不要通过端口转发暴露到局域网或公网。

> 会话说明：客户端应保存 `initialize` 响应中的 `Mcp-Session-Id` 并在后续请求中回传。每个会话第一次调用源码读取/编辑工具前必须先成功调用 `refresh_workspace_mirror`；工程切换后需要重新刷新。结束时可发送 `DELETE /mcp` 释放会话。

> 高风险工具：内置 AI 对话中的 `run_powershell_command` 和写入操作仍会走交互确认；通过本地 19207 MCP 端口发起的外部调用不弹审批窗口。外部调用仍受工具白名单、参数 Schema、工作区刷新、源码哈希 CAS、超时和取消机制约束，PowerShell 超时、取消或关闭 IDE 时会终止整个进程树。

> 源码写入：外部 Agent 必须先用 `read_real_file` 获取当前真实页的 SHA-256 `code_hash`，再把它作为 `edit_file` / `multi_edit_file` / `write_file` / `diff_file` 的 `expected_base_hash`；恢复快照时传 `expected_current_hash`。大文件读取可用 `next_source_byte_offset` → `byte_offset` 继续，并在续页时回传 `mirror_generation`。

---

### 各工具配置方法

<details>
<summary><b>🖱️ Cursor</b></summary>

在 Cursor 的配置文件（`.cursor/mcp.json` 或全局 MCP 配置）中添加以下内容：

```json
{
  "mcpServers": {
    "AutoLinker": {
      "url": "http://127.0.0.1:19207/mcp",
      "type": "http"
    }
  }
}
```

**步骤：**

1. 打开 Cursor，进入 `Settings` → `MCP` 页面
2. 点击 `Add MCP Server`，填写名称 `AutoLinker`
3. 类型选 `HTTP`，URL 填 `http://127.0.0.1:19207/mcp`
4. 保存后，在 Agent 对话框中即可调用易语言相关工具

> 💡 也可以直接编辑项目根目录下的 `.cursor/mcp.json` 文件，粘贴上方 JSON 配置。

</details>

---

<details>
<summary><b>🤖 Codex CLI（OpenAI 官方 CLI）</b></summary>

Codex CLI（Rust 版，当前维护版本）使用 **`config.toml`** 格式管理配置，HTTP 类型的 MCP 服务器需手动编辑配置文件添加。

**手动编辑 `config.toml`**

打开或创建 `%USERPROFILE%\.codex\config.toml`（Windows），添加以下内容：

```toml
[mcp_servers.AutoLinker]
url = "http://127.0.0.1:19207/mcp"
```

> ⚠️ 注意：格式是 `[mcp_servers.<名称>]`（点号分隔的单表），**不是** `[[mcp_servers]]`（数组表）。`url` 的存在自动表示使用 Streamable HTTP 传输，无需额外声明 `type` 字段。

**步骤：**

1. 确保已安装 Codex CLI（参考官方安装说明）
2. 编辑 `%USERPROFILE%\.codex\config.toml`，添加上方配置
3. 重新启动 Codex CLI，AutoLinker 提供的易语言工具即可在对话中使用

</details>

---

<details>
<summary><b>🤍 Claude Code（Anthropic 官方 CLI）</b></summary>

在 Claude Code 的项目 MCP 配置文件（`.mcp.json`，位于项目根目录）中添加：

```json
{
  "mcpServers": {
    "AutoLinker": {
      "type": "http",
      "url": "http://127.0.0.1:19207/mcp"
    }
  }
}
```

**步骤：**

1. 在易语言项目根目录下找到或创建 `.mcp.json` 文件
2. 将上方 JSON 内容合并到 `mcpServers` 节点下
3. 重启 Claude Code 或重新加载配置，AutoLinker 工具即可在对话中使用

也可以使用命令行快速添加（用户级别）：

```bash
claude mcp add --transport http AutoLinker http://127.0.0.1:19207/mcp
```

</details>

---

<details>
<summary><b>✨ Antigravity CLI（Google DeepMind）</b></summary>

在 Antigravity CLI 的全局 MCP 配置文件中添加：

```json
{
  "mcpServers": {
    "AutoLinker": {
      "type": "http",
      "serverUrl": "http://127.0.0.1:19207/mcp",
      "disabled": false
    }
  }
}
```

**步骤：**

1. 打开 Antigravity CLI 的全局 MCP 配置文件（路径：`%USERPROFILE%\.gemini\antigravity\mcp_config.json`）
2. 在 `mcpServers` 下添加上方配置
3. 重启 Antigravity CLI，在对话中即可使用 AutoLinker 提供的易语言工具

> 💡 也可以在 Antigravity CLI 的 Agent 面板中通过 UI 操作添加：点击 `⋯` 菜单 → `MCP Servers`。

</details>

---

<details>
<summary><b>🛠️ 其他支持 MCP 的工具（通用配置）</b></summary>

任何支持 **MCP Streamable HTTP** 传输协议的工具，均可按如下信息配置：

| 字段 | 值 |
|---|---|
| 服务名称 | `AutoLinker`（可自定义） |
| 传输类型 | `http`（streamable-http） |
| URL / Endpoint | `http://127.0.0.1:19207/mcp` |

通用 JSON 模板：

```json
{
  "mcpServers": {
    "AutoLinker": {
      "url": "http://127.0.0.1:19207/mcp",
      "type": "http"
    }
  }
}
```

> 📌 不同工具的配置文件路径和字段名称略有差异，请参考对应工具的官方文档。核心信息只有两点：**传输类型为 Streamable HTTP**，**地址为 `http://127.0.0.1:19207/mcp`**。

</details>

---

### 验证连接

配置完成后，可以通过以下方式验证 MCP 是否正常工作：

1. 确保易语言 IDE 已打开并加载了 AutoLinker
2. 在外部 Agent 中发送一条测试指令，例如："列出当前易语言项目的所有页面"
3. 如果能正常返回结果，说明 MCP 连接成功

**排查连接失败：**

- 检查易语言 IDE 是否已启动且 AutoLinker 已加载
- 确认端口 `19207` 未被防火墙或其他程序占用
- 在浏览器访问 `http://127.0.0.1:19207/`（GET请求），如能看到 JSON 健康检查响应则服务正常
- 如果收到 HTTP 503，表示连接队列已满；稍后重试，避免同时发起大量长耗时工具调用
- 浏览器网页脚本携带 `Origin` 时会被拒绝，这是预期的安全策略；地址栏直接打开健康检查不等于允许网页调用 MCP

---

## 常见问题

**Q：点“测试连通性”提示失败，怎么排查？**

1. 检查“API 密钥”是否填写正确（有没有多余的空格）
2. 检查“接口地址”和“接口协议”是否正确（使用预设站点新建时会自动填写）
3. 检查账户余额是否充足
4. 如果是海外服务（OpenAI、Claude），检查是否能正常访问海外网络

---

**Q：模型名填错了会怎样？**

会在调用时报错，提示模型不存在。直接修改“模型”字段，或点击右侧的 `↻` 获取服务端模型列表。

---

**Q：配置完成后在哪里使用 AI 功能？**

配置保存后，回到易语言 IDE，通过 AutoLinker 菜单或快捷方式即可调用 AI 功能（代码优化、注释、翻译等）。

---

**Q：配置文件保存在哪里？**

配置保存在易语言安装目录下的 `AutoLinker/AIConfig.json` 文件中。
