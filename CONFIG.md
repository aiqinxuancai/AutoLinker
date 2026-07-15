# AutoLinker AI 配置指南

> **重要：** 使用易语言 5.95+ 版本以获得最佳体验

快速上手，零基础也能看懂。

---

## 📋 目录
- [基础概念](#基础概念)
- [配置界面说明](#配置界面说明)
- [平台配置方法](#各平台配置方法)
- [源码编辑模式](#源码编辑模式)
- [Tavily 联网搜索](#tavily-联网搜索)
- [外部 Agent MCP 配置](#外部-agent-mcp-配置)
- [常见问题](#常见问题)

---

## 🎯 基础概念

### 🔑 API 密钥（API Key）
访问 AI 服务的"账号密码"。请妥善保管，避免泄露。

### 🌐 中转站（代理接口）
解决国内访问问题或将会员订阅转换为 API 的第三方服务。
- **转发型：** 国内服务器转发请求到海外 AI
- **逆向型：** 逆向分析 AI 工具协议，转为 API 出售（价格通常更低）

> ⚠️ 中转站为第三方服务，质量参差不齐，建议选择口碑好的平台。

### 🔗 接口地址（Base URL）
调用 AI 时的服务地址。
- 官方 OpenAI：`https://api.openai.com/v1`
- 中转站：各平台提供（推荐使用"使用预设站点新建"，自动填写）

### 🤖 模型
具体的 AI，不同模型能力和价格不同。
- `gpt-5.6-sol`：OpenAI 旗舰模型
- `claude-opus-4-8`：顶级写作模型
- `deepseek-v4-pro`：性价比高

### 🧠 思考等级
部分模型支持深度思考，调节推理强度：

| 界面选项 | 配置值 | 说明 |
|---|---|---|
| 关闭 | `off` | 速度最快 |
| 低 | `low` | 简单任务 |
| 中 | `medium` | 推荐起点 |
| 高 | `high` | 复杂任务 |
| 超高 | `xhigh` | 质量优先 |
| 最大 | `max` | 最高等级 |

> ⚠️ 若接口报参数不支持，请降低等级。

---

## 🛠️ 配置界面说明

### 关键字段

| 字段 | 说明 | 必填 |
|---|---|---|
| **配置组** | 管理多套模型配置，可新建/重命名/删除 | — |
| **接口协议** | `OpenAI Chat`、`OpenAI Responses`、`Gemini`、`Claude` | ✅ |
| **接口地址** | 模型服务地址 | ✅ |
| **API 密钥** | 模型平台提供的密钥 | ✅ |
| **模型** | 使用的模型名称（可点 `↻` 获取列表） | ✅ |
| 上下文长度 | 达到 95% 时自动压缩历史；留空按模型默认 |
| 思考等级 | 控制推理强度，详见上表 | 可选 |
| 系统提示词 | 附加到内置提示词后的自定义要求 | 可选 |
| 自定义请求头 | 每行填写 `请求头名称: 值` | 可选 |
| 源码编辑模式 | 全局设置：真实页优先 / 解包镜像基准 | 可选 |
| Tavily API 密钥 | 全局联网搜索密钥 | 可选 |

---

## 🌍 各平台配置方法

> 💡 **推荐流程：** 点击"使用预设站点新建" → 选择站点（模型） → 填写 API 密钥 → 测试连通性 → 保存

### 🌟 Right（推荐中转站）
- **简介：** 国内直连，聚合多平台模型
- **官网：** https://right.codes/register
- **预设：**
  - Right(gpt-5.6-sol) — 代码任务
  - Right(gpt-5.6) — 通用任务
- **地址：** `https://right.codes/codex`（自动填写）
- **协议：** `OpenAI Chat`（自动填写）

### 🇨🇳 DeepSeek（性价比高）
- **简介：** 国内领先，价格优势明显
- **官网：** https://platform.deepseek.com
- **预设模型：** deepseek-v4-flash / pro / chat / reasoner
- **地址：** `https://api.deepseek.com`

### 🤖 智谱 GLM（中文场景）
- **简介：** 中文理解能力强
- **官网：** https://open.bigmodel.cn
- **地址：** `https://open.bigmodel.cn/api/paas/v4`
- **预设模型：** glm-5.2 / glm-5-turbo / glm-4.7 / glm-4.5-air

### 🌊 千问 / 通义（阿里云）
- **简介：** 国内直连，模型丰富
- **官网：** https://dashscope.console.aliyun.com
- **地址：** `https://dashscope.aliyuncs.com/compatible-mode/v1`
- **预设模型：** qwen3.7-plus / qwen3.7-max / qwen3-coder-next

### 🌙 Kimi（长上下文）
- **简介：** 支持超长上下文，长文档分析
- **官网：** https://platform.moonshot.cn
- **地址：** `https://api.moonshot.cn/v1`
- **预设模型：** kimi-k2.7-code / kimi-k2.6

### 🔵 MiniMax
- **简介：** 国内直连
- **官网：** https://platform.minimax.chat
- **地址：** `https://api.minimax.chat/v1`
- **预设模型：** MiniMax-M3 / M2.7 / M2.5

> 💡 更多平台（豆包、OpenAI、Claude、Gemini 等）均可通过"使用预设站点新建"选择，或手动填写接口地址与协议。

---

## ⚙️ 源码编辑模式
"源码编辑模式" 位于"其他设置"区域，为全局配置：

| 模式 | 说明 |
|---|---|
| **真实页优先** | 编辑前读取真实页面，推荐日常使用 |
| **解包镜像基准（测试）** | 仅用于测试或排查问题 |

> 💡 不确定时保留默认的"真实页优先"。

---

## 🌐 Tavily 联网搜索（可选）
Tavily 是 AI 搜索 API。配置后 `search_web_tavily` 工具可实时联网搜索信息。

- **官网：** https://tavily.com（注册后有免费额度）
- **配置：** 在"其他设置"区域填入 **Tavily API 密钥** 即可

---

## 🔌 外部 Agent MCP 配置
AutoLinker 在易语言 IDE 启动后自动开启本地 **MCP 服务**（端口 `19207`）。

任何支持 **MCP Streamable HTTP** 的外部 AI 工具（如 Cursor、Claude Code、Codex CLI、Antigravity CLI）均可通过该地址配置。

### 🔑 MCP 服务信息
| 项目 | 值 |
|---|---|
| 协议 | Streamable HTTP；支持协商 MCP 多版本 |
| 本地地址 | `http://127.0.0.1:19207/mcp` |
| 传输类型 | `http` |

> 🔒 **安全说明：** 仅 `127.0.0.1` 可访问，浏览器脚本携带 `Origin` 会被拒绝，请勿暴露到局域网/公网。

### 📋 工具配置模板
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

<details>
<summary><b>📖 各平台配置路径（点击展开）</b></summary>

**Cursor:**
- 编辑 `.cursor/mcp.json` 或通过 UI "Add MCP Server"
- 示例：
```json
{
  "mcpServers": {
    "AutoLinker": { "type": "http", "url": "http://127.0.0.1:19207/mcp" }
  }
}
```

**Codex CLI (Rust版):**
- 编辑 `%USERPROFILE%\.codex\config.toml`
- 示例：
```toml
[mcp_servers.AutoLinker]
url = "http://127.0.0.1:19207/mcp"
```

**Claude Code:**
- 编辑项目根目录的 `.mcp.json`

**Antigravity CLI:**
- 编辑 `%USERPROFILE%\.gemini\antigravity\mcp_config.json`

</details>

> 💡 格式和路径有差异，请参考对应工具的官方文档以确保准确配置。

---

## ❓ 常见问题

**Q：点"测试连通性"提示失败，怎排查？**
1. 检查 API 密钥是否填写正确（有无多余空格）
2. 检查接口地址和协议是否正确（使用预设站点新建时会自动填写）
3. 检查账户余额是否充足
4. 如为海外服务（OpenAI、Claude），检查是否能正常访问海外网络

**Q：找不到"使用预设站点新建"按钮？**
该按钮在配置界面顶部区域。如未见到，请检查 AutoLinker 版本是否为最新。

**Q：配置完成后如何使用 AI 功能？**
回到易语言 IDE，通过 AutoLinker 菜单或快捷方式即可调用 AI 功能（代码优化、注释、翻译等）。

**Q：配置文件保存在哪里？**
保存在易语言安装目录下的 `AutoLinker/AIConfig.json` 文件中。

**Q：e-packager 下载不下来怎办？**
若程序自动下载 e-packager（镜像解包器）失败（网络问题等），可前往 [e-packager Releases](https://github.com/aiqinxuancai/e-packager/releases) 手动下载，解压缩到 **易语言安装目录\tools** 目录中即可。
