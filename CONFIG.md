# AutoLinker AI 配置指南

> **重要：** 使用易语言 5.95+ 版本以获得最佳体验

零基础也能看懂。第一次使用请直接看下面的「配置入口」，三步就能用起来。

---

## 📋 目录
- [配置入口](#-配置入口)
- [基础概念](#-基础概念)
- [配置界面说明](#-配置界面说明)
- [平台配置方法](#-各平台配置方法)
- [源码编辑模式](#-源码编辑模式)
- [Tavily 联网搜索](#-tavily-联网搜索可选)
- [AI SKILL 技能](#-ai-skill-技能)
- [外部 Agent MCP 配置](#-外部-agent-mcp-配置)
- [常见问题](#-常见问题)

---

## 🚀 配置入口

**在哪里设置？** 打开易语言 IDE，点击顶部 **工具菜单 → AutoLinker 设置**，在左侧选择 **AI 接口** 页。

> 找不到工具菜单里的 AutoLinker？说明支持库还没启用。先把 `AutoLinker.fne` 放进易语言的 `lib` 目录，然后在 IDE 的支持库设置里勾选启用并重启 IDE。

### 三步跑起来

1. 打开 **工具菜单 → AutoLinker 设置 → AI 接口**
2. 点 **使用预设站点新建** → 选一个站点（例如 Right、DeepSeek）→ 填入该平台的 **API 密钥**
3. 点 **测试连通性**，通过后点 **保存并继续**

保存后回到 IDE，左侧的 **AutoLinker AI 对话** 页签和代码编辑器里的右键 AI 菜单就可以用了。

> 💡 没配置就直接用 AI 功能时，AutoLinker 会自动弹出这个设置页，填完即可继续当次操作。

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

> 📊 不确定如何选择时，可查看[易语言大模型基准评分](https://e-language-bench.apptest.dev)，对比各模型的格式、编译率与总分。

### 🧠 思考等级
部分模型支持深度思考，调节推理强度：

| 界面选项 | 配置值 | 说明 |
|---|---|---|
| 关闭 | `off` | 速度最快 |
| 低 | `low` | 简单任务 |
| 中 | `medium` | 推荐起点 |
| 高 | `high` | 复杂任务 |
| 超高 | `xhigh` | 质量优先 |
| 最大 | `max` | 高等级 |
| 极限 | `ultra` | 最高等级，仅部分模型支持 |

> ⚠️ 若接口报参数不支持，请降低等级。

---

## 🛠️ 配置界面说明

位置：**工具菜单 → AutoLinker 设置 → AI 接口**。

### 关键字段

| 字段 | 说明 | 必填 |
|---|---|---|
| **配置组** | 管理多套模型配置，可新建/重命名/删除 | — |
| **接口协议** | `OpenAI Chat`、`OpenAI Responses`、`Gemini`、`Claude` | ✅ |
| **接口地址** | 模型服务地址 | ✅ |
| **API 密钥** | 模型平台提供的密钥 | ✅ |
| **模型** | 使用的模型名称（可点 `↻` 获取列表） | ✅ |
| 上下文长度 | 达到 95% 时自动压缩历史；留空按模型默认 | 可选 |
| 思考等级 | 控制推理强度，详见上表 | 可选 |
| 系统提示词 | 附加到内置提示词后的自定义要求 | 可选 |
| 自定义请求头 | 每行填写 `请求头名称: 值` | 可选 |
| 源码编辑模式 | 位于「其他设置」区域，全局生效：真实页优先 / 解包镜像基准 | 可选 |
| Tavily API 密钥 | 位于「其他设置」区域，全局联网搜索密钥 | 可选 |

填完点底部的 **测试连通性** 验证，再点 **保存并继续**。

---

## 🌍 各平台配置方法

> 💡 **推荐流程：** 点击"使用预设站点新建" → 选择站点（模型） → 填写 API 密钥 → 测试连通性 → 保存

### 🌟 Right（推荐中转站）
- **简介：** 国内直连，聚合多平台模型
- **官网：** https://www.rightapi.ai/register
- **预设：**
  - Right Codex：`gpt-5.6-sol` / `gpt-5.6-terra` / `gpt-5.6-luna`
  - Right Grok：`grok-4.6`
  - Right Claude AWS：`claude-sonnet-5`
- **地址：** Codex 使用 `https://www.rightapi.ai/codex`，Grok 使用 `https://www.rightapi.ai/grok`，Claude AWS 使用 `https://www.rightapi.ai/claude-aws`（均自动填写）
- **协议：** Codex/Grok 使用 `OpenAI Chat`，Claude AWS 使用 `Claude`（自动填写）

### 🇨🇳 DeepSeek（性价比高）
- **简介：** 国内领先，价格优势明显
- **官网：** https://platform.deepseek.com
- **预设模型：** `deepseek-v4-flash` / `deepseek-v4-pro`
- **版本说明：** `deepseek-v4-flash` 当前对应 DeepSeek-V4-Flash-0731；思考等级默认 `high`，并支持 `low` / `high` / `max`
- **地址：** `https://api.deepseek.com`

### 🤖 智谱 GLM（中文场景）
- **简介：** 中文理解能力强
- **官网：** https://open.bigmodel.cn
- **地址：** `https://open.bigmodel.cn/api/paas/v4`
- **预设模型：** glm-5.2

### 🌊 千问 / 通义（阿里云）
- **简介：** 国内直连，模型丰富
- **官网：** https://dashscope.console.aliyun.com
- **地址：** `https://dashscope.aliyuncs.com/compatible-mode/v1`
- **预设模型：** qwen3.7-plus / qwen3.7-max / qwen3-coder-next

### 🌙 Kimi（长上下文）
- **简介：** 支持超长上下文，长文档分析
- **官网：** https://platform.moonshot.cn
- **地址：** `https://api.moonshot.cn/v1`
- **预设模型：** kimi-k3

### 🔵 MiniMax
- **简介：** 国内直连
- **官网：** https://platform.minimaxi.com
- **地址：** `https://api.minimax.chat/v1`
- **预设模型：** MiniMax-M3

> 💡 更多平台（豆包、OpenAI、Claude、Gemini 等）均可通过"使用预设站点新建"选择，或手动填写接口地址与协议。

---

## ⚙️ 源码编辑模式
位置：**AutoLinker 设置 → AI 接口 → 其他设置** 区域。该项独立于模型服务配置组，全局生效：

| 模式 | 说明 |
|---|---|
| **真实页优先** | 编辑前读取真实页面，推荐日常使用 |
| **解包镜像基准（测试）** | 仅用于测试或排查问题 |

> 💡 不确定时保留默认的"真实页优先"。

---

## 🌐 Tavily 联网搜索（可选）
Tavily 是 AI 搜索 API。配置后 `search_web_tavily` 工具可实时联网搜索信息。

- **官网：** https://tavily.com（注册后有免费额度）
- **配置：** 在 **AutoLinker 设置 → AI 接口 → 其他设置** 中填入 **Tavily API 密钥** 即可

---

## 🧩 AI SKILL 技能

AutoLinker 内部 AI 对话可以通过 `SKILL.md` 加载专用工作流程。打开 **工具菜单 → AutoLinker 设置 → AI SKILL** 可管理技能，也可以从 AI 对话顶部的 **SKILL** 按钮直接进入。

### 技能目录

| 作用域 | 路径 | 优先级 |
|---|---|---|
| 全局 | `{易语言目录}\AutoLinker\Skills\<技能名>\SKILL.md` | 普通 |
| 当前项目 | `<工程目录>\.agents\skills\<技能名>\SKILL.md` | 同名时优先 |

设置页支持启用、停用、更新、卸载和打开技能目录。无效的 `SKILL.md` 会显示解析错误，不会提供给 AI。

### 安装方式

- 在 **本地添加** 标签中输入或选择技能目录、直接的 `SKILL.md`，也可选择 `.zip` 文件。父目录或 ZIP 中有多个技能时，会列出候选项并逐个添加。
- 对本地目录或 `SKILL.md`，可选择 **复制安装** 或 **引用原路径**。复制安装会复制完整技能目录；引用不会复制文件，原路径修改会在刷新设置或下一轮 AI 对话时生效。
- 本地复制和原路径引用都支持全局、当前项目作用域。项目引用按当前工程目录记录在本机 `AISkillsConfig.json`，不会把绝对路径写入工程文件。
- ZIP 仅支持复制安装，使用系统 PowerShell/.NET 解压，不依赖额外压缩库。安装前会检查压缩包大小、条目数、声明解压总量和路径穿越。
- 在 **skills.sh** 标签中搜索并选择安装范围。
- 搜索结果可打开详情查看技能介绍、来源和安装量，再决定是否安装。
- 在 **GitHub 安装** 标签中粘贴公开仓库、`tree` 子目录或指向 `SKILL.md` 的 `blob` 地址。
- 一个仓库包含多个技能时，需要明确选择要安装的技能。
- 当前版本只支持公开 GitHub 仓库，不读取或保存 GitHub Token。

移除原路径引用只会注销配置，不会删除源目录。引用目标缺失或 `SKILL.md` 失效时仍会显示为无效技能，便于修复路径或手动移除。

### 在对话中使用

- 输入 `$技能名` 可在当前轮明确激活，例如 `$pdf 分析这份文档`。
- 未显式点名但任务明显符合技能描述时，AI 会按需读取对应 `SKILL.md`。
- 技能引用的 `scripts`、`references`、`assets` 和模板会随技能一起安装；读取被限制在技能目录内。
- 安装技能不会执行其中的脚本。AI 后续如需运行本机命令，仍会触发现有的命令确认流程。

> ⚠️ SKILL 是会影响 AI 行为的指令包。仅安装可信来源，并在更新后检查其 `SKILL.md` 和脚本内容。

---

## 🔌 外部 Agent MCP 配置
AutoLinker 在易语言 IDE 启动后自动开启本地 **MCP 服务**，无需手动开启。`19207` 是固定网关端口，各个 IDE 实例使用 `19208` 之后的端口作为内部后端。在 **工具菜单 → AutoLinker 设置 → MCP 服务** 中可查看和管理相关配置。

任何支持 **MCP Streamable HTTP** 的外部 AI 工具（如 Cursor、Claude Code、Codex CLI、Antigravity CLI）均可通过该地址配置。

### 🔑 MCP 服务信息
| 项目 | 值 |
|---|---|
| 协议 | Streamable HTTP；支持协商 MCP 多版本 |
| 本地地址 | `http://127.0.0.1:19207/mcp` |
| 传输类型 | `http` |

### 多实例切换

- 外部工具始终配置 `http://127.0.0.1:19207/mcp`，不要改成实例后端端口。
- 同时打开多个易语言 IDE 时，调用 `list_instances` 查看实例、工程路径和当前页面，再使用 `select_instance` 的 `instance_id` 选择目标。
- 选择结果按 `Mcp-Session-Id` 隔离，不会影响其他 MCP 客户端。
- 选中的实例关闭后，AutoLinker 会明确要求重新选择，不会自动把修改操作切到其他工程。
- 持有 `19207` 的 IDE 退出后，其他存活实例会自动接管固定网关；客户端重新连接后仍使用原地址。

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

**Q：工具菜单里没有"AutoLinker 设置"？**
支持库没有加载。确认 `AutoLinker.fne` 已放入易语言 `lib` 目录，并在 IDE 的支持库设置中启用，然后重启 IDE。

**Q：找不到"使用预设站点新建"按钮？**
该按钮在 **AI 接口** 页的"模型服务配置组"区域，紧挨着"新建"。如未见到，请检查 AutoLinker 版本是否为最新。

**Q：配置完成后如何使用 AI 功能？**
回到易语言 IDE：左侧的 **AutoLinker AI 对话** 页签可直接对话；在代码编辑器中右键可使用 AI 优化函数、添加注释、翻译等菜单项。

**Q：配置文件保存在哪里？**
保存在 `{易语言安装目录}\AutoLinker\AIConfig.json`。可通过 **工具菜单 → 打开 AutoLinker 配置目录** 快速定位。

**Q：e-packager 下载不下来怎办？**
若程序自动下载 e-packager（镜像解包器）失败（网络问题等），可前往 [e-packager Releases](https://github.com/aiqinxuancai/e-packager/releases) 手动下载，解压缩到 **易语言安装目录\tools** 目录中即可。
