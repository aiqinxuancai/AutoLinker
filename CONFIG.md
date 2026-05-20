# AutoLinker AI 配置指南

> 面向易语言开发者的 AI 功能配置说明，零基础也能看懂。

---

## 目录

- [基础概念：搞清楚这几个词](#基础概念搞清楚这几个词)
- [配置界面各项说明](#配置界面各项说明)
- [各平台配置方法](#各平台配置方法)
- [Tavily 联网搜索（可选）](#tavily-联网搜索可选)
- [外部 Agent MCP 配置](#外部-agent-mcp-配置)
- [常见问题](#常见问题)

---

## 基础概念：搞清楚这几个词

在配置 AI 之前，先了解几个关键概念，后面看配置项就一目了然了。

### 🔑 什么是 API Key？

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

这也是为什么同样是 GPT-5.5 / Claude Opus 4.7，通过某些中转站调用的价格会比官方 API 便宜很多，本质上是在消耗会员订阅的额度。

> 📌 **协议注意：** 逆向 Codex 等工具的中转站，其 GPT 接口通常走的是 **OpenAI Responses** 协议（即 `/responses` 端点），而非普通的 Chat 协议（`/chat/completions`）。配置时 Protocol 应选 `OpenAI Responses`，而不是默认的 `OpenAI Chat`，否则可能调用失败或功能不完整。Right 等主流逆向中转站已在平台预设中正确配置，直接选预设即可。

**使用中转站的好处：**
- 国内网络直接可用，无需 VPN
- 聚合多个 AI 平台，一个 Key 可调用多种模型
- 价格远低于官方 API（尤其是逆向型中转站）

> ⚠️ 中转站是第三方服务，质量参差不齐，建议选口碑好、有一定用户量的平台。逆向型中转站也存在随时被封禁的风险，请酌情使用。

---

### 🤖 什么是 Base URL？

Base URL 是你调用 AI 时的"服务地址"。

- 官方 OpenAI 的地址是 `https://api.openai.com/v1`
- 中转站会有自己的地址，如 `https://right.codes/codex`
- 不同平台的地址不同，但**选了平台预设后会自动填入**，无需手动查找

---

### 💬 什么是模型（Model）？

模型就是具体的 AI，不同模型能力和价格不同。例如：

| 模型名 | 特点 |
|---|---|
| `gpt-5.5` | 长上下文能力强，适合编程 |
| `deepseek-chat` | 国内模型，价格实惠 |
| `claude-sonnet-4-6` | 擅长理解和写作 |

> 💡 不知道选哪个？可以先用平台预设自动填入的推荐模型。

---

### 🧠 什么是 Thinking Level（思考等级）？

部分高级模型支持"深度思考"功能，会在给出答案前先进行推理分析。

| 等级 | 说明 |
|---|---|
| 关闭（Off） | 不思考，速度最快，适合简单任务 |
| 低（Low） | 轻度思考 |
| 中（Medium） | 适中思考 |
| 高（High） | 深度思考，最准确但最慢、最贵 |

> 💡 如果你用的模型不支持思考功能，选"关闭"即可，选其他无效也不报错。

---

## 配置界面各项说明

打开 AutoLinker AI 设置后，你会看到以下配置项：

| 字段 | 说明 | 是否必填 |
|---|---|---|
| **配置组** | 可以保存多套配置，方便在不同平台之间快速切换 | — |
| **平台预设** | 选择一个平台，自动填入协议和地址 | 可选 |
| **Protocol** | 接口协议，一般跟随平台预设自动设置 | ✅ |
| **Base URL** | 服务地址，选平台预设后自动填入 | ✅ |
| **API Key** | 从平台获取的密钥 | ✅ |
| **Model** | 使用的模型名称 | ✅ |
| **Thinking Level** | 思考等级，不确定选"关闭" | 可选 |
| **System Prompt** | 附加的系统提示词，一般无需填写 | 可选 |
| **Custom Headers** | 自定义请求头，高级用法，一般无需填写 | 可选 |
| **Tavily API Key** | 联网搜索的独立密钥，见下方说明 | 可选 |

### 配置组（多套配置）

配置组功能让你可以保存多套 API 配置，方便切换。例如：
- 一套用国内的 DeepSeek，日常省钱用
- 一套用海外的 GPT，处理复杂任务用

点击"**新建**"即可添加新配置组，"**删除**"可删除当前组（至少保留一个）。

---

## 各平台配置方法

> 💡 **推荐流程：** 选择平台预设 → 自动填入地址和协议 → 填写 API Key → 选择模型 → 点击"测试连通性" → 保存

---

<details>
<summary><b>🌟 Right（推荐中转站，国内直连，聚合多平台模型）</b></summary>

**简介：** Right.codes 是一个支持多平台模型的中转站，国内可直接访问，支持 GPT、Claude 等主流模型。

**注册地址：** https://right.codes/register

**配置步骤：**

1. 在"**平台预设**"下拉框中选择 `Right`
2. Base URL 和 Protocol 会自动填入：
   - Base URL: `https://right.codes/codex`
   - Protocol: `OpenAI Chat`
3. 前往 Right 官网注册并充值，在控制台复制你的 **API Key**，填入"**API Key**"
4. 在"**Model**"下拉中选择想用的模型：
   - `gpt-5.5` — 最新模型
   - `gpt-5.4` — （推荐）
   - `gpt-5.2` — 通用模型
5. 点击"**测试连通性**"确认配置正确
6. 点击"**保存并继续**"

</details>

---

<details>
<summary><b>🇨🇳 DeepSeek（国内直连，性价比高）</b></summary>

**简介：** DeepSeek 是国内领先的 AI 公司，模型能力强，价格相比海外模型有明显优势。

**官网：** https://platform.deepseek.com

**配置步骤：**

1. 在"**平台预设**"中选择 `DeepSeek`
2. 自动填入：
   - Base URL: `https://api.deepseek.com`
   - Protocol: `OpenAI Chat`
3. 前往 DeepSeek 平台注册，在 API 管理页面创建密钥，填入"**API Key**"
4. 选择模型：
   - `deepseek-chat` — 通用对话（价格实惠）
   - `deepseek-reasoner` — 深度推理（更准确，适合复杂代码分析）
   - `deepseek-v4-flash` — 快速版
   - `deepseek-v4-pro` — 增强版
5. 测试并保存

> 💡 DeepSeek 支持"思考"功能，使用 `deepseek-reasoner` 时可以把 Thinking Level 调为"中"或"高"。

</details>

---

<details>
<summary><b>🤖 智谱 GLM（国内直连，适合中文场景）</b></summary>

**简介：** 智谱 AI 出品的 GLM 系列模型，中文理解能力强。

**官网：** https://open.bigmodel.cn

**配置步骤：**

1. 平台预设选 `智谱`
2. 自动填入：
   - Base URL: `https://open.bigmodel.cn/api/paas/v4`
3. 在智谱官网注册，前往"API 密钥"页面创建密钥，填入"**API Key**"
4. 选择模型：
   - `glm-5.1` — 最新旗舰
   - `glm-5` — 标准版
   - `glm-5-turbo` — 快速版
   - `glm-4.7` — 上一代旗舰
5. 测试并保存

</details>

---

<details>
<summary><b>🌊 千问 / 通义（阿里云，国内直连）</b></summary>

**简介：** 阿里云旗下的通义千问，国内直连，有丰富的模型可选。

**官网：** https://dashscope.console.aliyun.com

**配置步骤：**

1. 平台预设选 `千问`
2. 自动填入：
   - Base URL: `https://dashscope.aliyuncs.com/compatible-mode/v1`
3. 在阿里云 DashScope 控制台创建 API Key，填入"**API Key**"
4. 选择模型：
   - `qwen3-max` — 最强版本
   - `qwen3.6-plus` — 增强版
   - `qwen3-coder-plus` — 代码专向（适合编程场景）
5. 测试并保存

</details>

---

<details>
<summary><b>🌙 Kimi（月之暗面，国内直连，长上下文）</b></summary>

**简介：** 月之暗面旗下 Kimi 模型，支持超长上下文，擅长文档分析。

**官网：** https://platform.moonshot.cn

**配置步骤：**

1. 平台预设选 `Kimi`
2. 自动填入：
   - Base URL: `https://api.moonshot.cn/v1`
3. 在 Moonshot 平台创建 API Key，填入"**API Key**"
4. 选择模型：
   - `kimi-k2.5` — 最新旗舰
   - `kimi-k2-thinking` — 深度思考版
   - `moonshot-v1-128k` — 超长上下文
5. 测试并保存

</details>

---

<details>
<summary><b>🫘 豆包（字节跳动，国内直连）</b></summary>

**简介：** 字节跳动旗下豆包大模型，国内访问稳定。

**官网：** https://console.volcengine.com/ark

**配置步骤：**

1. 平台预设选 `豆包`
2. 自动填入：
   - Base URL: `https://ark.cn-beijing.volces.com/api/v3`
3. 在火山引擎控制台创建 API Key，填入"**API Key**"

   > ⚠️ 注意：豆包平台需要先创建"推理接入点"，Model 填写接入点 ID，而不是普通模型名。请参考豆包平台文档。

4. 选择模型（推荐参考）：
   - `doubao-seed-1.8` — 最新版
   - `doubao-seed-1.6` — 标准版
   - `doubao-seed-1.6-thinking` — 思考版
5. 测试并保存

</details>

---

<details>
<summary><b>🔵 MiniMax（国内直连）</b></summary>

**简介：** MiniMax 旗下大模型，国内直连。

**官网：** https://platform.minimax.chat

**配置步骤：**

1. 平台预设选 `MiniMax`
2. 自动填入：
   - Base URL: `https://api.minimax.chat/v1`
3. 在 MiniMax 平台注册并获取 API Key，填入"**API Key**"
4. 选择模型：
   - `MiniMax-M2.7` — 旗舰版
   - `MiniMax-M2.7-highspeed` — 高速版
5. 测试并保存

</details>

---

<details>
<summary><b>🌐 aihubmix（中转站，聚合主流海外模型）</b></summary>

**简介：** aihubmix 是支持 GPT、Claude、Gemini 等多个顶级模型的中转站，国内可访问，按量计费。

**注册地址：** https://aihubmix.com

**配置步骤：**

1. 平台预设选 `aihubmix`
2. 自动填入：
   - Base URL: `https://aihubmix.com/v1`
   - Protocol: `OpenAI Chat`
3. 在 aihubmix 控制台充值并创建 API Key，填入"**API Key**"
4. 选择模型（该中转站聚合了大量模型）：
   - `gpt-5.4` — 最新 GPT
   - `claude-opus-4-7` — Claude 旗舰
   - `claude-sonnet-4-6` — Claude 高性价比
   - `gemini-3.1-pro-preview` — Gemini 旗舰
   - `grok-4` — xAI 模型
5. 测试并保存

</details>

---

<details>
<summary><b>🔥 硅基流动（中转站，聚合开源模型为主）</b></summary>

**简介：** 硅基流动以国产开源模型为主，提供 DeepSeek、Qwen 等模型的 API 接入，部分模型有免费额度。

**注册地址：** https://cloud.siliconflow.cn

**配置步骤：**

1. 平台预设选 `硅基流动`
2. 自动填入：
   - Base URL: `https://api.siliconflow.cn/v1`
3. 在硅基流动控制台创建 API Key，填入"**API Key**"
4. 选择模型：
   - `deepseek-ai/DeepSeek-V3.2` — DeepSeek 最新版
   - `deepseek-ai/DeepSeek-R1` — DeepSeek 推理版
   - `Qwen/Qwen3-Coder-480B-A35B-Instruct` — 千问大型代码模型

   > 💡 注意：硅基流动的模型名带有路径格式（如 `deepseek-ai/DeepSeek-V3.2`），填写时要完整。

5. 测试并保存

</details>

---

<details>
<summary><b>🌍 OpenAI 官方（需能访问海外网络）</b></summary>

**简介：** OpenAI 官方 API，需要海外网络或自备代理。按 token 计费，价格较高但模型质量有保证。

**官网：** https://platform.openai.com

**配置步骤：**

1. 平台预设选 `OpenAI`
2. 自动填入：
   - Base URL: `https://api.openai.com/v1`
   - **Protocol: `OpenAI Responses`**（注意，OpenAI 官方使用新接口）
3. 在 OpenAI Platform 充值后创建 API Key，填入"**API Key**"
4. 选择模型：
   - `gpt-5.5` — 最新旗舰
   - `gpt-5.4` — 上一代旗舰
   - `gpt-5.4-mini` — 经济版
   - `gpt-5.3-codex` — 代码专向
5. 测试并保存

> ⚠️ 直接用官方 OpenAI 在国内需要能访问海外网络，如果访问困难建议使用上方的中转站。

</details>

---

<details>
<summary><b>🤍 Claude 官方（Anthropic，需能访问海外网络）</b></summary>

**简介：** Anthropic 官方 API，Claude 模型在理解和写作方面表现优秀，也擅长代码。国内访问需要代理。

**官网：** https://console.anthropic.com

**配置步骤：**

1. 平台预设选 `Claude`
2. 自动填入：
   - Base URL: `https://api.anthropic.com`
   - **Protocol: `Claude`**（Claude 有独立协议）
3. 在 Anthropic 控制台获取 API Key，填入"**API Key**"
4. 选择模型：
   - `claude-opus-4-7` — 旗舰版（最强）
   - `claude-sonnet-4-6` — 平衡版（推荐）
   - `claude-3-5-haiku-latest` — 快速轻量版
5. 测试并保存

> 💡 如果你在国内访问有问题，可以通过 aihubmix 或其他中转站使用 Claude 模型。

</details>

---

<details>
<summary><b>✨ Gemini 官方（Google，需能访问海外网络）</b></summary>

**简介：** Google 旗下 Gemini 模型，多模态能力强。国内访问需要代理。

**官网：** https://aistudio.google.com

**配置步骤：**

1. 平台预设选 `Gemini`
2. 自动填入：
   - Base URL: `https://generativelanguage.googleapis.com`
   - **Protocol: `Gemini`**（Google 有独立协议）
3. 在 Google AI Studio 获取 API Key，填入"**API Key**"
4. 选择模型：
   - `gemini-3.1-pro-preview` — 旗舰版（最强推理，适合复杂任务）
   - `gemini-3-flash-preview` — Pro 级智能 + Flash 速度（推荐）
   - `gemini-3.1-flash-lite` — 轻量经济版（高吞吐量场景）
   - `gemini-2.5-pro` — 上一代旗舰
   - `gemini-2.5-flash` — 上一代轻量版
5. 测试并保存

</details>

---

## Tavily 联网搜索（可选）

Tavily 是一个 AI 搜索 API，配置后 AutoLinker 的 `search_web_tavily` 工具就能实时联网搜索信息。

**官网：** https://tavily.com（注册后有免费额度）

**配置步骤：**

1. 前往 Tavily 官网注册
2. 在控制台复制你的 **API Key**
3. 在配置界面右侧"Tavily 联网搜索"区域，填入"**Tavily API Key**"
4. 保存即可

> 💡 Tavily Key 是独立的，和主模型的 API Key 没有关系，两者互不影响。  
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
| 协议 | Streamable HTTP（MCP 2025-03-26 规范） |
| 本地地址 | `http://127.0.0.1:19207/mcp` |
| 传输类型 | `http`（streamable-http） |

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

---

## 常见问题

**Q：点"测试连通性"提示失败，怎么排查？**

1. 检查 API Key 是否填写正确（有没有多余的空格）
2. 检查 Base URL 是否正确（用了平台预设一般不会错）
3. 检查账户余额是否充足
4. 如果是海外服务（OpenAI、Claude），检查是否能正常访问海外网络

---

**Q：模型名填错了会怎样？**

会在调用时报错，提示模型不存在。直接修改 Model 字段为正确的模型名即可。

---

**Q：配置完成后在哪里使用 AI 功能？**

配置保存后，回到易语言 IDE，通过 AutoLinker 菜单或快捷方式即可调用 AI 功能（代码优化、注释、翻译等）。

---

**Q：配置文件保存在哪里？**

配置保存在易语言安装目录下的 `AutoLinker/AIConfig.json` 文件中。
