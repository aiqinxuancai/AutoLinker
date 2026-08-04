---
name: autolinker-usage
description: >-
  介绍如何使用 AutoLinker，即易语言 AI Agent 支持库和本地 MCP 服务器。内容涵盖安装、
  AI Agent 对话页、计划模式、目标模式、右键 AI 菜单操作、各项目的 .AGENTS.md 规范、
  按源码切换链接器、动态/静态 ec 自动切换、e-packager 工程镜像及 EC 模块/支持库依赖
  信息、本地 MCP HTTP 服务器、协议方法及完整工具参数、外部 Agent（Claude Code /
  Codex / Cursor）的 MCP 配置、多实例路由、无界面命令行编译、核心库 C++ 重写，以及
  AI 服务商和配置（CONFIG.md）。当任务涉及运行、配置或驱动 AutoLinker，通过它将
  外部 AI Agent 连接
  到易语言 IDE，了解其代码读取技术原理，或排查版本兼容性、闪退、响应慢、对话失败、
  MCP 连接异常等问题时使用此技能。
---

# AutoLinker 使用指南

AutoLinker（`AutoLinker.fne`）是易语言支持库（易语言 IDE 启动时加载的 DLL）。它在
IDE 内直接提供 AI 辅助编码功能，并开放本地 **MCP Streamable HTTP** 服务器，使外部
Agent（Claude Code、Codex、Gemini/Antigravity CLI、Cursor、Windsurf）能够通过 IDE
读取和编辑已打开的 `.e` 工程。

`.e` 源码是加密的单文件格式，无法在 IDE 外部直接编辑。AutoLinker 会从 IDE 内存导出
包含未保存修改的临时快照，再由 e-packager 解包成文本镜像供读取和搜索；修改则按镜像
路径映射回 IDE 真实程序项。整个过程不会直接改写磁盘上的原始 `.e` 文件。

仓库内的权威资料：`README.md`、`CONFIG.md`、`AGENTS.md`。

## 安装

1. 首次安装时下载**最新版** Release，将 `AutoLinker.fne` 放入易语言的 `lib` 目录，
   然后在 IDE 中启用该支持库。
2. 使用易语言 **5.95**。其他易语言版本即使能够运行，也只属于勉强兼容，不提供支持。
3. 不支持 Windows 7。
4. IDE 启动时，AutoLinker 会自动启动本地 MCP 服务器，并将日志写入 IDE 输出窗口和
   `autolinker.log`。

已经安装 AutoLinker 后，优先通过“插件菜单 → AutoLinker 设置 → 关于 → 组件更新”
检查和更新 AutoLinker 或 e-packager，不必重新执行首次安装流程。AutoLinker 更新需要
退出当前 IDE 后替换正在加载的 `AutoLinker.fne`；按页面提示完成即可。

配置文件位于 `{易语言安装目录}\AutoLinker\AIConfig.json`，日志位于
`{易语言安装目录}\AutoLinker\Log`。

## 功能概览

### 1. AI Agent 对话页（“AutoLinker AI 对话”）

位于左侧工作区。直接告诉 AI 你的需求，它会自动搜索、读取、编辑当前打开的工程，并
插入代码。输入 `$skillname` 可在当前轮次强制激活指定 SKILL（例如
`$pdf 分析这份文档`）。

#### 计划模式

需要先调查工程、比较方案或审批高影响修改时，点击对话页底部的“计划”进入计划模式，
再发送需求。

1. 计划模式下，AI 只能探索、读取、搜索和设计实现方案，不能写入或回滚源码、编译工程，
   也不能执行本地命令。
2. 存在会实质影响方案的关键歧义时，AI 可以显示结构化问题；回答后它会继续规划，回答
   问题不等于批准最终方案。
3. AI 提交方案后，模式进入“待批准”。检查方案后选择批准，AI 才会按方案实施；如方案
   需要调整，提供修改意见并等待它重新提交。
4. 不想继续规划时，退出计划模式。退出会清除尚未批准的计划，不会自动实施修改。

计划模式适合需求边界不清、涉及多个文件、可能改变工程结构或需要先评审实现路径的任务；
明确且局部的小修改通常直接对话即可。

#### 目标模式

需要 AI 跨多轮持续推进一个长期任务时，在底部输入框写清楚最终目标，再点击“目标”开启。
目标创建后会持久保存在当前对话会话中，AI 不会因为一次普通回复就停止，而会在上一轮结束
后继续处理尚未完成的工作；新到达的用户消息会被优先处理，然后继续原目标。

- 点击已有目标可打开管理菜单：活动目标可“暂停 Goal”；已暂停或阻塞的目标可“恢复
  Goal”；还可以“标记完成”或“清除 Goal”。创建新目标前必须先清除已有目标。
- 只有任务确实全部完成时才标记完成。只有缺少必要用户输入或必须等待外部状态变化、已经
  无法继续推进时，才将目标标记为阻塞；条件具备后再恢复。
- 目标模式会记录累计 Token 和有效执行时间。暂停、阻塞或进入计划模式期间不会继续累计
  活动时间。
- 计划模式和目标模式同时存在时，以计划模式为先：目标的自动推进和计时暂时挂起，批准
  计划或退出计划模式后再恢复。

目标模式适合较长的实现、排错和迁移任务。只需一次回答、简单查询或局部修改时不要开启，
以免 AI 在回答后继续自动发起后续轮次。

### 2. 右键 AI 菜单（代码编辑器内）

- **AI优化函数**：在保持行为一致的前提下优化当前函数
- **AI为当前函数添加注释**：为当前函数生成注释
- **AI翻译当前函数+变量名**：将名称改为英文 `lowerCamelCase`
- **AI翻译选中文本**：翻译选中的文本
- **AI按当前页类型添加代码**：根据当前页面类型和上下文生成新代码

### 3. 各工程的规范文件 `{文件名}.AGENTS.md`

在 `.e` 文件旁创建同名的 `.AGENTS.md`（例如 `test_a.e` 对应
`test_a.AGENTS.md`）。文件内容会作为工程规范注入所有 AI 功能的系统提示词，相当于
易语言工程中的 `CLAUDE.md`。

也可以打开“插件菜单 → AutoLinker 设置 → 当前项目 AGENTS.md”，直接查看和编辑当前
已打开工程对应的规范文件。

### 4. 按源码切换链接器

通过“插件菜单 → AutoLinker 设置 → 链接器”添加和编辑多个 `link.ini` 配置，然后在
“主菜单 → 编译”中切换当前源码使用的链接器，无需再手动来回替换。

### 5. 调试/编译时自动切换动态和静态 ec

将同一模块的一对 ec 文件（动态版和静态版）放在**同一文件夹**。AutoLinker 会自动
切换导入模块：“开始编译”时使用静态版，“开始调试”时使用动态版，常用于 VMP SDK、
ExDui。通过“插件菜单 → AutoLinker 设置 → EC 模块切换”维护规则。必须先引用其中任意
一个 ec。

### 6. 核心库 C++ 重写（性能、反破解、防御性免杀）

使用现代 C++（`.lib`、32 位）替换核心库函数。通过“插件菜单 → AutoLinker 设置 →
核心库函数重写”配置需要强制链接的 `.lib` 列表（Lib 路径排在
`krnln_static.lib` 之前；可选择按链接器名称子串匹配；每条规则可单独启用）。此功能会
启用链接器 `/FORCE`；自有 Lib 必须禁用 `/GL`。参考实现见 `TestCore`（C++20 /
VC2022）。

### 7. 将当前 `.e` 解包到目录（e-packager）

上下文/菜单项“将[<file>.e]反编译到目录”会通过 e-packager 反编译当前打开的 `.e`。
只有打开 `.e` 源码时才会启用此功能。

### 8. 鼠标后退键撤销上一次修改

跳转到上一个修改位置，行为与其他 IDE 类似。

### 9. 统一设置窗口（“AutoLinker 设置”）

统一入口是“插件菜单 → AutoLinker 设置”。回答“某功能在哪里”或指导用户操作时，优先
给出下表中的完整路径，不要只描述配置文件或旧版独立窗口：

| 设置页签 | 位置和用途 |
| --- | --- |
| AI 接口 | “AutoLinker 设置 → AI 接口”：管理模型服务配置组、API Key、协议、模型、源码编辑模式和 Tavily 联网搜索 |
| MCP 服务 | “AutoLinker 设置 → MCP 服务”：查看和管理外部 Agent 的 MCP 配置 |
| AI SKILL | “AutoLinker 设置 → AI SKILL”：安装、启停、更新、打开目录或卸载全局/工程级技能 |
| AI 对话配色 | “AutoLinker 设置 → AI 对话配色”：选择内置配色，或新建、复制、重命名、删除和编辑自定义配色 |
| 当前项目 AGENTS.md | “AutoLinker 设置 → 当前项目 AGENTS.md”：编辑当前 `.e` 工程的同名规范文件 |
| 链接器 | “AutoLinker 设置 → 链接器”：管理 `link.ini` 配置；实际切换在“主菜单 → 编译” |
| EC 模块切换 | “AutoLinker 设置 → EC 模块切换”：维护调试/编译时动态、静态 ec 自动切换规则 |
| 核心库函数重写 | “AutoLinker 设置 → 核心库函数重写”：维护需要强制链接的 `.lib` 规则 |
| 日志优化 | “AutoLinker 设置 → 日志优化”：调整编译内部日志 Hook 和调试输出快速路径 |
| 关于 | “AutoLinker 设置 → 关于”：查看版本、项目链接，并在“组件更新”中检查或更新 AutoLinker 与 e-packager |

“AI 对话配色”页提供默认配色和内置深色配色。内置配色不可直接修改，可先复制再编辑；
自定义配色可先调整表面、文字、主色、强调、成功、警告、危险 7 个主色，再按需展开高级
颜色覆盖。右侧可实时预览，保存后会立即应用到已经打开的 AI 对话界面。

插件菜单还包含“打开项目目录”、“打开 AutoLinker 配置目录”、“打开易语言目录”。

## AI 服务商配置

完整说明见 `CONFIG.md`。在“插件菜单 → AutoLinker 设置 → AI 接口”中配置。推荐流程：
**使用预设站点新建** → 选择站点/模型 → 填写 API Key → 测试连通性 → 保存。

| 配置项 | 使用说明 |
| --- | --- |
| 配置组 | 保存多套站点和模型配置，可新建、重命名、删除或切换，避免反复覆盖参数 |
| 接口协议 | 选择 `OpenAI Chat`、`OpenAI Responses`、`Gemini` 或 `Claude`，必须与服务商接口匹配 |
| 接口地址 | 填写服务商提供的 Base URL；优先使用预设，让 AutoLinker 自动填写地址和协议 |
| API 密钥与模型 | 两者均必填；妥善保管密钥，不要写入工程或公开日志；可点击 `↻` 获取模型列表 |
| 上下文长度 | 达到 95% 时自动压缩历史；留空使用模型默认值 |
| 思考等级 | 支持 `off` / `low` / `medium` / `high` / `xhigh` / `max`；接口拒绝参数时降低等级 |
| 系统提示词 | 作为附加要求追加到 AutoLinker 内置提示词之后，不会替换内置规则 |
| 自定义请求头 | 可选，每行填写一个 `Name: Value`，按中转站文档配置 |

日常使用保持“源码编辑模式”为默认的**真实页优先**；“解包镜像基准”仅用于测试或排查。
需要 AI 联网搜索时，在其他设置中填写 Tavily API 密钥以启用 `search_web_tavily`。

预设站点包括 Right（推荐中转）、DeepSeek、智谱 GLM、千问/通义、Kimi、MiniMax；
豆包、OpenAI、Claude、Gemini 可通过“使用预设站点新建”添加。

手动配置时不要混用其他平台的协议、Base URL 或模型名。保存前先测试连通性；测试成功只
表示当前配置当时可用，中转站或具体模型线路后续仍可能发生网络、限流或上游故障。

## AI SKILL 系统（AutoLinker 内置对话）

内置 AI 对话可加载 `SKILL.md` 工作流。通过“AutoLinker 设置 → AI SKILL”管理，或
点击对话页顶部的 SKILL 按钮。

- 全局：`{易语言目录}\AutoLinker\Skills\<name>\SKILL.md`
- 工程级：`<工程目录>\.agents\skills\<name>\SKILL.md`（重名时优先）

可以从 **skills.sh** 页签或公开的 **GitHub** 仓库/tree/blob URL 安装。安装 skill
不会运行其中的脚本；执行本地命令仍会触发正常的确认流程。只安装可信来源。

## 本地 MCP 服务器

- 监听 `http://127.0.0.1:19207/mcp`。`19207` 是**固定网关端口**；各 IDE 实例使用
  从 `19208` 开始的内部后端端口。外部工具始终连接 `19207`，不要连接后端端口。
- 启动日志：
  `[AutoLinker][LocalMCP] 本地 MCP 服务已启动：http://127.0.0.1:19207/mcp`
- 协议：JSON-RPC 2.0、Streamable HTTP。`initialize` 可协商 `2025-11-25` /
  `2025-03-26` / `2024-11-05`，未知版本回退到 `2025-11-25`。支持的方法：
  `initialize`、`notifications/initialized`、`ping`、`tools/list`、`tools/call`、
  `DELETE /mcp`。
- 安全：仅绑定 `127.0.0.1`；拒绝包含非空 `Origin` 的浏览器请求；不支持 CORS /
  Bearer token。外部调用不会弹出审批窗口，但仍会执行工具白名单、参数结构校验、
  工作区刷新和源码 SHA-256 CAS。服务器使用 4 个固定工作线程，连接队列最多 32 个，
  过载时返回 HTTP 503。不要将其暴露到局域网或互联网。

### 外部 Agent 配置

**Claude Code**（`~/.claude.json` 或工程 `.mcp.json`）/ **Gemini CLI**
（`~/.gemini/settings.json`）：

```json
{ "mcpServers": { "AutoLinker": { "type": "http", "url": "http://127.0.0.1:19207/mcp" } } }
```

**Codex**（`~/.codex/config.toml`）：

```toml
[mcp_servers.AutoLinker]
url = "http://127.0.0.1:19207/mcp"
```

**Cursor / Windsurf**：MCP 设置 → 名称 `AutoLinker`，类型 `http` /
`streamable_http`，URL `http://127.0.0.1:19207/mcp`。

**Antigravity CLI**：`%USERPROFILE%\.gemini\antigravity\mcp_config.json`。

### 多实例路由

同时打开多个 IDE 时，先调用 `list_instances` 查看实例（工程路径和当前页面），再使用
`instance_id` 调用 `select_instance`。选择结果按 `Mcp-Session-Id` 隔离。如果选中的
实例关闭，AutoLinker 会要求重新选择，而不会静默切换到其他工程。如果占用 `19207`
的 IDE 退出，仍存活的实例会接管网关，客户端继续重连同一地址即可。

### 代码读取技术原理

AutoLinker 不直接解析或修改加密的 `.e` 文件。准备完整工作区时，它先从当前 IDE 内存
工程导出临时快照（因此能包含尚未保存的修改），再调用 e-packager 的 `unpack` 命令，
将快照解包到 `%TEMP%/AutoLinker/workspace-mirror/` 下的实例专用目录。`list_files`、
`search_code`、`read_file` / `read_files` 和 `read_code_item` 读取的都是这个文本镜像。

完整镜像不只包含当前工程源码：

| 镜像目录 | 可读取的内容 |
| --- | --- |
| `src/` | 当前 `.e` 工程解包后的程序集、类、固定表和窗口相关文件 |
| `ecom/` | 当前工程引用的 EC 模块解包后的源码，可用于查找模块命令的实现和调用方式 |
| `elib/` | 当前工程所用支持库的公开信息，包括支持库通过 `GetNewInf` 等公开描述提供的命令、数据类型和接口 |
| `header/` | e-packager 生成的其他声明或头部参考信息 |

`ecom/`、`elib/` 和 `header/` 只用于读取和搜索，不能通过源码写入工具修改。`elib/` 展示
的是支持库公开接口，不是对 `.fne` 内部 C/C++ 实现的反编译。修改当前工程时，AutoLinker
只允许把可写的 `src/` 文本映射回 IDE 程序项，并通过真实页哈希防止覆盖较新的修改。

### 读写模型（外部 Agent 必须了解）

- 内置 AI 每轮都会以 `full` 模式自动准备镜像。**外部 MCP 会话在首次读写前必须调用
  一次 `refresh_workspace_mirror`**。`mode` 可取 `auto` / `main_only` / `full`。镜像
  由 e-packager 解包到 `%TEMP%/AutoLinker/workspace-mirror/`，包含未保存的修改，且
  不会触碰源码目录。
- 读取工具使用镜像相对路径（`list_files`、`search_code`、`read_file`、
  `read_files`、`read_code_item`）。大文件会返回 `next_source_byte_offset` 供继续分页；
  后续请求应传回上一页非零的 `mirror_generation`。
- 编辑前，调用 `read_real_file` 获取分页视图和 `code_hash`（CAS 基线）。写入工具通过
  `file_path` 定位目标，并映射回 IDE 程序项后直接写入 IDE，不会在写入时重新编译。
- 写入操作必须携带 SHA-256 `expected_base_hash`（恢复时使用
  `expected_current_hash`），以防止基于过期版本覆盖。
- `src/*.xml` 是易语言原生窗口 UI 文件，仅支持读取/搜索。当前工具不支持修改窗口或控件
  的位置、大小、层级和属性，不支持添加、删除控件，也不支持新增、删除或修改控件的事件
  绑定。不要尝试写入窗口 XML，也不要声称已经完成这些界面修改。
- 窗口程序集代码仍可通过对应的 `src/*.txt` 编辑，包括修改已有控件事件子程序的代码实现；
  但新增一个事件子程序不代表已经建立控件事件绑定。固定表（常量、全局变量、DLL 声明、
  数据类型）可通过对应路径编辑，程序集变量会以 IDE 可接受的形式写回。
- 需要了解模块或支持库时，使用 `list_files` / `search_code` / `read_files` 检索 `ecom/`
  和 `elib/`。不要因为它们出现在镜像中就将其当作当前工程的可写源码。

### 工具集（`tools/list`）

需要查询、生成或排查 MCP 请求时，必须读取
[完整 MCP API 参考](references/mcp-api.md)。该参考以当前代码中的公开目录和参数校验为准，
包含协议级方法、HTTP 会话操作、21 个 `tools/list` 工具、所有参数及嵌套参数、默认值、
取值范围、调用前置条件、关键返回字段和错误处理。

工具按用途分为：工作区刷新与镜像读取、真实页 CAS 编辑、IDE 状态与编译、公网页面访问、
多实例路由。标准调用顺序是 `initialize` → 必要时 `list_instances` / `select_instance` →
`get_current_eide_info` → `refresh_workspace_mirror` → 读取/编辑/编译工具。不要根据本节摘要
猜测参数；实际调用前查阅子文档中的对应工具条目。

## 无界面命令行编译

推荐方式：通过 `AutoLinkerTest headless-compile` 启动 `e.exe`。该命令会关闭启动弹窗、
隐藏 IDE、调用 `compile_with_output_path`，并输出 JSON 结果：

```powershell
.\bin\fne_release\AutoLinkerTest.exe headless-compile `
  "C:\path\to\e571.exe" "D:\demo\demo.e" "D:\demo\build\demo.exe" `
  --target auto --static --result "D:\demo\build\compile-result.json" --timeout 120
```

`--target` 可取 `auto` | `win_exe` | `win_console_exe` | `win_dll` | `ecom`。
`--static` 仅适用于 EXE/DLL。结果还会写入
`{易语言目录}\AutoLinker\Log\headless_compile_last.json`。

也可以直接驱动主程序（仅限无界面编译；处理早期弹窗时仍优先使用启动器）：

```powershell
"C:\path\to\e571.exe" "D:\demo\demo.e" `
  --autolinker-headless-compile `
  --autolinker-output "D:\demo\build\demo.exe" `
  --autolinker-target auto `
  --autolinker-result "D:\demo\build\compile-result.json"
```

`AutoLinkerTest.exe` 还提供以下服务商测试和自检子命令：
`headless-compile`、`mock-mcp-stdio`、`mcp-self-test`、
`gameanalytics-self-test`、`version-text`、`version-compare`、
`deepseek-model-test`、`openai-chat-test`、`openai-responses-test`、
`gemini-model-test`、`claude-model-test`、`linker-out` / `linker-krnln` /
`between-dashes`。

## 构建和测试 AutoLinker（贡献者）

构建 `fne_release` / x86（VS2026 VC++、ISO C++20）：

```powershell
MSBuild.exe ..\AutoLinker.vcxproj /t:Build "/p:Configuration=fne_release;Platform=Win32" /m
```

`MSBuild.exe` 的位置因计算机和 VS 版本而异，请先定位可执行文件。

测试编译后的支持库：先关闭易语言主程序以释放 `AutoLinker.fne` 文件锁，再将
`AutoLinker.fne` 复制到 IDE 的 `lib` 目录；然后打开 `eproj/` 下的 `.e` 文件并测试
相关功能。加载时 MCP 会在端口 19207 上启动。传递中文页面名称时使用 **UTF-8**。
不依赖 IDE 的逻辑（例如模块本地解析）可使用 `AutoLinkerTest` 工程测试。排查真实场景
问题时，使用 `e5.95.exe` 打开目标文件、操作 AI 对话，并结合日志目录分析。

## 兼容性和问题反馈前提

回答或排查 AutoLinker 问题前，按以下顺序确认环境：

1. 确认操作系统不是 Windows 7。
2. 确认使用易语言 **5.95**；其他版本不在支持范围内。
3. 确认已经更新到**最新版 AutoLinker**。
4. 确认没有同时加载其他第三方插件或支持库。遇到闪退、异常行为或疑似冲突时，先卸载
   或禁用其他插件，仅保留 AutoLinker 后重新测试。
5. 只有在上述条件全部满足且问题仍可复现时，再按 AutoLinker 问题或缺陷继续排查，并
   提供易语言版本、AutoLinker 版本、复现步骤和相关日志。

不要在易语言版本不受支持、AutoLinker 不是最新版或插件冲突尚未排除时，直接将现象
判定为 AutoLinker 缺陷。

## 常见问题

### 为什么 AI 对话响应很慢？

“慢”主要是模型、API 服务商或网络响应速度问题。优先更换模型或接口站点，并检查网络
和服务商状态，不要先按 AutoLinker 性能问题排查。

### 为什么对话因网络错误失败、中断或突然不可用？

出现连接失败、超时、断流、HTTP 5xx 或请求一度可用后又失败时，优先检查**中转站网络**
和**中转站当前模型线路的可用性**，不要先判定为 AutoLinker 缺陷。中转站本身可以访问，
不代表所选模型线路一定可用。

1. 在“AutoLinker 设置 → AI 接口”中重新执行连通性测试，并查看中转站公告或服务状态。
2. 在同一中转站切换到另一个确定可用的模型测试，判断是否仅当前模型线路故障。
3. 切换到另一个可用中转站测试，判断是否为原中转站的网络、限流或上游故障。
4. 再检查本机网络、代理、API 密钥、余额、Base URL 和接口协议。
5. 只有多个独立可用的中转站和模型均能稳定复现同一问题时，再结合 AutoLinker 日志继续
   排查程序问题。

### 为什么外部 Agent 调不通 MCP，或用一会后又无法调用？

此类现象优先按外部 Agent 工具的 MCP 配置或会话权限问题排查：

1. 确认 AutoLinker 启动日志已经显示本地 MCP 服务运行在
   `http://127.0.0.1:19207/mcp`。
2. 确认外部工具的 MCP 配置指向固定网关端口 `19207`，而不是 `19208` 及以上的内部
   后端端口。
3. 配置 MCP 后，检查外部工具是否还要求在**当前会话**中单独启用或授权 AutoLinker
   MCP。有些工具只保存服务器配置，不会自动为每个会话开放 MCP 权限。
4. 确认外部工具能够直接列出并调用 AutoLinker MCP 工具。如果它转而编写 `.py`、
   PowerShell 脚本或 CMD 命令来请求 MCP，通常说明其原生 MCP 没有配置好、没有启用，
   或当前会话没有权限；先按该外部工具的文档修正配置和授权。
5. 外部工具种类和版本众多，应自行查阅对应工具的 MCP 使用说明；AutoLinker 只能保证
   本地 MCP 服务和协议行为，不能逐一负责所有外部工具的安装、配置和会话授权。

### 为什么 AI 服务连通性测试失败？

优先确认中转站网络和所选模型线路当前可用，再检查 API 密钥是否带有多余空格、Base URL
和协议是否匹配、账户余额，以及访问 OpenAI/Claude 所需的海外网络是否连通。

### e-packager 无法下载或正确解压怎么办？

先在“AutoLinker 设置 → 关于 → 组件更新”中检查并安装 e-packager。如果自动下载失败、
下载后无法正确解压，或长时间卡住，不要反复等待，建议手动获取工具包：

1. 打开 https://github.com/aiqinxuancai/e-packager（或其 Releases 页面）下载最新 Release。
2. 如果无法访问 GitHub，也可以从 AutoLinker 相关交流群的群共享中获取 e-packager。
3. 将下载包解压到 `{易语言安装目录}\tools`，并确认最终文件位于
   `{易语言安装目录}\tools\e-packager.exe`，不要多套一层压缩包目录。
4. 重启易语言 IDE，再调用或重试 `refresh_workspace_mirror`。

更新 `AutoLinker.fne` 时，如果自动解压、PowerShell 或退出后替换仍失败，先关闭易语言 IDE，
从 AutoLinker Releases 或相关交流群群共享获取 `AutoLinker.fne`，再手动替换
`{易语言安装目录}\lib\AutoLinker.fne`。

### 配置文件在哪里？

配置文件位于 `{易语言安装目录}\AutoLinker\AIConfig.json`。

## 参考资料

- 白皮书：https://github.com/aiqinxuancai/Awesome-E-Agent
- e-packager：https://github.com/aiqinxuancai/e-packager
- 核心库静态实现参考：https://github.com/zhongjianhua163/BlackMoonKernelStaticLib
