---
name: autolinker-usage
description: >-
  How to use AutoLinker — the 易语言 (E-language) AI Agent support library and
  local MCP server. Covers install, the AI Agent chat tab, right-click AI menu
  actions, per-project .AGENTS.md conventions, per-source linker switching,
  dynamic/static ec auto-swap, the e-packager workspace mirror and EC module /
  support library dependency information, the local MCP HTTP server and its tool set,
  external agent (Claude Code / Codex / Cursor) MCP config, multi-instance
  routing, headless command-line compile, core-library C++ rewrite, and AI
  provider/config setup (CONFIG.md). Use this whenever a task involves running,
  configuring, or driving AutoLinker, understanding its code-reading model, or
  connecting an external AI agent to an 易语言 IDE through it.
---

# AutoLinker usage

AutoLinker (`AutoLinker.fne`) is an 易语言 support library (a DLL loaded by the
E-language IDE at startup). It adds AI-assisted coding directly inside the IDE
and exposes a local **MCP Streamable HTTP** server so external agents (Claude
Code, Codex, Gemini/Antigravity CLI, Cursor, Windsurf) can read and edit the
open `.e` project through the IDE.

`.e` sources are an encrypted single-file format that can't be edited from
outside the IDE. AutoLinker exports a temporary snapshot from the live IDE
project, including unsaved changes, and lets e-packager unpack it into a text
mirror for reads and searches. Writes map mirror paths back to real IDE program
items; they do not directly rewrite the original `.e` file on disk.

Source of truth in-repo: `README.md`, `CONFIG.md`, `AGENTS.md`.

## Install

1. For a first install, download the latest Release, put `AutoLinker.fne` in the
   易语言 `lib` directory, then enable the support library in the IDE.
2. 易语言 **5.95+** is recommended for the best experience.
3. On IDE launch, AutoLinker auto-starts the local MCP server and logs to the
   IDE output window and `autolinker.log`.

For an existing install, use 工具菜单 → AutoLinker 设置 → 关于 → 组件更新 to
check and update AutoLinker or e-packager. Updating AutoLinker replaces the
loaded `AutoLinker.fne` after the current IDE exits; follow the page prompts.

Config lives at `{易语言安装目录}\AutoLinker\AIConfig.json`; logs at
`{易语言安装目录}\AutoLinker\Log`.

## Feature map (what a user can do)

### 1. AI Agent chat tab ("AutoLinker AI 对话")
A left work-area tab. Tell the AI what you want; it searches, reads, edits,
inserts code in the currently open project automatically. Type `$skillname` to
force-activate a SKILL for that turn (e.g. `$pdf 分析这份文档`).

### 2. Right-click AI menu (in the code editor)
- **AI优化函数** — equivalence-preserving optimize of the current function
- **AI为当前函数添加注释** — generate comments for the current function
- **AI翻译当前函数+变量名** — rename to English `lowerCamelCase`
- **AI翻译选中文本** — translate the selected text
- **AI按当前页类型添加代码** — generate new code by current page type + context

### 3. Per-project spec file `{文件名}.AGENTS.md`
Create a same-named `.AGENTS.md` next to the `.e` file (e.g. `test_a.e` →
`test_a.AGENTS.md`). Its content is injected into the system prompt of all AI
features as the project convention — the E-language analogue of `CLAUDE.md`.
The same file can be viewed and edited at 工具菜单 → AutoLinker 设置 → 当前项目
AGENTS.md for the currently open project.

### 4. Per-source linker switching
Add and edit `link.ini` configs at 工具菜单 → AutoLinker 设置 → 链接器, then
switch the active linker for the current source under 主菜单 → 编译.

### 5. Dynamic/static ec auto-swap on debug/compile
Keep a pair of ec files (dynamic + static) for one module in the **same
folder**. AutoLinker swaps the imported module automatically: static on
"开始编译", dynamic on "开始调试" (typical for VMP SDK, ExDui). Configure via
工具菜单 → AutoLinker 设置 → EC 模块切换. You must first reference either ec.

### 6. Core-library C++ rewrite (perf / anti-crack / AV-evasion-for-defense)
Replace core-lib functions with modern C++ (`.lib`, 32-bit). Configure the
force-linked `.lib` list under 工具菜单 → AutoLinker 设置 → 核心库函数重写
(Lib path ordered before `krnln_static.lib`; optional linker-name substring
match; per-rule enable). Enables linker `/FORCE`; own Lib must disable `/GL`.
See `TestCore` (C++20 / VC2022) for a reference implementation.

### 7. Unpack current `.e` to a directory (e-packager)
Context/menu item "将[<file>.e]反编译到目录" — decompiles the open `.e` via
e-packager. Only enabled when a `.e` source is open.

### 8. Mouse back-button = undo last change
Navigates to the previous modification, like other IDEs.

### 9. Unified settings window ("AutoLinker 设置")

The unified entry is 工具菜单 → AutoLinker 设置. When explaining where a
feature lives, give the complete path from this table instead of referring only
to a config file or an obsolete standalone dialog:

| Settings tab | Location and purpose |
| --- | --- |
| AI 接口 | AutoLinker 设置 → AI 接口: provider profiles, API Key, protocol, model, source editing mode, and Tavily search |
| MCP 服务 | AutoLinker 设置 → MCP 服务: inspect and manage MCP configuration for external agents |
| AI SKILL | AutoLinker 设置 → AI SKILL: install, enable, disable, update, open, or remove global/project skills |
| AI 对话配色 | AutoLinker 设置 → AI 对话配色: choose a built-in theme or create, copy, rename, delete, and edit custom themes |
| 当前项目 AGENTS.md | AutoLinker 设置 → 当前项目 AGENTS.md: edit the convention file for the open `.e` project |
| 链接器 | AutoLinker 设置 → 链接器: manage `link.ini`; select the active linker under 主菜单 → 编译 |
| EC 模块切换 | AutoLinker 设置 → EC 模块切换: maintain dynamic/static ec switching rules |
| 核心库函数重写 | AutoLinker 设置 → 核心库函数重写: maintain force-linked `.lib` rules |
| 日志优化 | AutoLinker 设置 → 日志优化: compilation-log Hook and debug-output fast-path options |
| 关于 | AutoLinker 设置 → 关于: version and project links; check/update AutoLinker and e-packager under 组件更新 |

The AI 对话配色 page includes the default and built-in dark themes. Built-in
themes are read-only; copy one before editing. A custom theme exposes seven
primary colors (surface, text, primary, accent, success, warning, danger),
advanced per-color overrides, and a live preview. Saving applies it immediately
to any open AI chat view.

The add-in menu also contains 打开项目目录, 打开 AutoLinker 配置目录, and
打开易语言目录.

## AI provider configuration (see CONFIG.md for full detail)

Configure in 工具菜单 → AutoLinker 设置 → AI 接口. Recommended flow: **使用预设站点新建**
→ pick site/model → fill API Key → 测试连通性 → 保存.

Key fields: 接口协议 (`OpenAI Chat` / `OpenAI Responses` / `Gemini` / `Claude`),
接口地址 (Base URL), API 密钥, 模型 (`↻` fetches list), 上下文长度 (auto-compacts
history at 95%; blank = model default), 思考等级 (`off`/`low`/`medium`/`high`/
`xhigh`/`max` — lower it if the API rejects the param), 系统提示词, 自定义请求头
(`Name: Value` per line), 源码编辑模式 (global: 真实页优先 [default] / 解包镜像基准
[testing only]), Tavily API 密钥 (enables `search_web_tavily`).

Preset sites include Right (recommended relay), DeepSeek, 智谱 GLM, 千问/通义,
Kimi, MiniMax, plus 豆包/OpenAI/Claude/Gemini via 使用预设站点新建.

## AI SKILL system (inside AutoLinker's own chat)
`SKILL.md` workflows load into the internal AI chat. Manage via
AutoLinker 设置 → AI SKILL (or the SKILL button atop the chat).
- Global: `{易语言目录}\AutoLinker\Skills\<name>\SKILL.md`
- Per-project: `<工程目录>\.agents\skills\<name>\SKILL.md` (wins on name clash)
Install from the **skills.sh** tab or a public **GitHub** repo/tree/blob URL.
Installing a skill does not run its scripts; running local commands still
triggers the normal confirmation flow. Only install trusted sources.

## Local MCP server

- Listens on `http://127.0.0.1:19207/mcp`. `19207` is the **fixed gateway
  port**; each IDE instance uses an internal backend port from `19208` up.
  External tools always point at `19207` — never a backend port.
- Startup log line:
  `[AutoLinker][LocalMCP] 本地 MCP 服务已启动：http://127.0.0.1:19207/mcp`
- Protocol: JSON-RPC 2.0, Streamable HTTP. `initialize` negotiates
  `2025-11-25` / `2025-03-26` / `2024-11-05` (unknown → `2025-11-25`).
  Methods: `initialize`, `notifications/initialized`, `ping`, `tools/list`,
  `tools/call`, `DELETE /mcp`.
- Security: binds `127.0.0.1` only; rejects browser requests with a non-empty
  `Origin`; no CORS / Bearer token. External calls skip the approval popup but
  still enforce the tool allowlist, param schema, workspace refresh, and
  SHA-256 CAS on source. 4 fixed worker threads, max 32 queued connections,
  overload → HTTP 503. Do not expose to LAN/internet.

### External agent config
**Claude Code** (`~/.claude.json` or project `.mcp.json`) / **Gemini CLI**
(`~/.gemini/settings.json`):
```json
{ "mcpServers": { "AutoLinker": { "type": "http", "url": "http://127.0.0.1:19207/mcp" } } }
```
**Codex** (`~/.codex/config.toml`):
```toml
[mcp_servers.AutoLinker]
url = "http://127.0.0.1:19207/mcp"
```
**Cursor / Windsurf**: MCP settings → Name `AutoLinker`, Type `http` /
`streamable_http`, URL `http://127.0.0.1:19207/mcp`.
**Antigravity CLI**: `%USERPROFILE%\.gemini\antigravity\mcp_config.json`.

### Multi-instance routing
With several IDEs open, call `list_instances` to see instances (project path +
current page), then `select_instance` with an `instance_id`. Selection is
isolated per `Mcp-Session-Id`. If the selected instance closes, AutoLinker
requires re-selection rather than silently retargeting another project. If the
IDE holding `19207` exits, a surviving instance takes over the gateway; clients
reconnect to the same address.

### How code reading works

AutoLinker does not directly parse or modify the encrypted `.e` file. To build
a complete workspace, it exports the current in-memory IDE project to a
temporary snapshot (including unsaved changes), then runs e-packager `unpack`
into an instance-specific directory under `%TEMP%/AutoLinker/workspace-mirror/`.
`list_files`, `search_code`, `read_file` / `read_files`, and `read_code_item`
read this text mirror.

The complete mirror contains more than the current project's own source:

| Mirror path | Readable content |
| --- | --- |
| `src/` | Unpacked assemblies, classes, fixed tables, and window files from the current `.e` project |
| `ecom/` | Unpacked source of EC modules referenced by the project, useful for finding command implementations and usage |
| `elib/` | Public information for referenced support libraries, including commands, data types, and interfaces described through public metadata such as `GetNewInf` |
| `header/` | Other declarations or header reference information generated by e-packager |

Treat `ecom/`, `elib/`, and `header/` as read/search-only reference material.
`elib/` exposes the support library's public interface, not a decompilation of
its internal native `.fne` implementation. Only writable `src/` text maps back
to IDE program items, with real-page hashes protecting newer changes.

### Read/write model (important for external agents)
- The internal AI auto-prepares the mirror in `full` mode each turn. An
  **external MCP session must call `refresh_workspace_mirror` once** before its
  first read/write. `mode`: `auto` / `main_only` / `full`. The mirror is
  unpacked by e-packager to `%TEMP%/AutoLinker/workspace-mirror/` (includes
  unsaved changes) and does not touch the source directory.
- Reads use mirror-relative paths (`list_files`, `search_code`, `read_file`,
  `read_files`, `read_code_item`). Large files return `next_source_byte_offset`
  for continuation; pass back the previous page's non-zero `mirror_generation`.
- Before editing, call `read_real_file` for a paginated view + `code_hash`
  (the CAS baseline). Write tools target `file_path`, map back to the IDE
  program item, and write straight into the IDE (no recompile-on-write).
- Writes must carry a SHA-256 `expected_base_hash` (`expected_current_hash`
  for restore) to prevent stale-baseline overwrites. `src/*.xml` are window UI
  files — read/search only; fixed tables (constants, globals, DLL decls, data
  types) are editable via their paths. Program-set variables are written back in
  IDE-acceptable form.
- To inspect modules or support libraries, search and read `ecom/` and `elib/`
  with the mirror tools. Their presence in the mirror does not make them
  writable current-project source.

### Tool set (`tools/list`)
| Category | Tool | Purpose |
| --- | --- | --- |
| Read | `refresh_workspace_mirror` | Refresh mirror from IDE memory (`auto`/`main_only`/`full`) |
| Read | `list_files` | Glob-list files in mirror |
| Read | `search_code` | Per-file search: batch patterns, glob, context, paging |
| Read | `read_file` / `read_files` | Read single / multiple files or ranges, with line numbers |
| Read | `read_code_item` | Read a full subprogram / decl block by top-level item name |
| Read | `read_real_file` | Paginated view + `code_hash` from the IDE real page (write baseline) |
| Edit | `edit_file` / `multi_edit_file` | Exact text replace (single / batch) |
| Edit | `write_file` | Full-source overwrite of the real page, with `expected_base_hash` |
| Edit | `diff_file` | Preview structured diff, no write |
| Edit | `restore_file_snapshot` | Restore the pre-write snapshot |
| Edit | `add_new_file` | Create a new 程序集 / class, optionally write full source + refresh mirror |
| Current page | `get_current_page_info` | Current page name, type, parse source |
| Current page | `get_current_eide_info` | Source path, IDE process path, MCP port, etc. |
| Compile | `compile_with_output_path` | `target` default `auto`; verifies success by artifact fingerprint |
| Web | `search_web_tavily` | Web search (needs Tavily key) |
| Web | `fetch_url` | Fetch public HTTP(S) text; blocks loopback/private/redirect |
| Web | `extract_web_document` | Extract page body + absolute-link summary |
| Gateway | `list_instances` / `select_instance` | Enumerate / route to a specific IDE instance |

## Headless command-line compile

Preferred: launch `e.exe` via `AutoLinkerTest headless-compile` (closes
startup popups, hides the IDE, calls `compile_with_output_path`, prints result
JSON):
```powershell
.\bin\fne_release\AutoLinkerTest.exe headless-compile `
  "C:\path\to\e571.exe" "D:\demo\demo.e" "D:\demo\build\demo.exe" `
  --target auto --static --result "D:\demo\build\compile-result.json" --timeout 120
```
`--target`: `auto` | `win_exe` | `win_console_exe` | `win_dll` | `ecom`.
`--static` applies to EXE/DLL only. Result also written to
`{易语言目录}\AutoLinker\Log\headless_compile_last.json`.

Or drive the main program directly (headless compile only; still prefer the
launcher for early popups):
```powershell
"C:\path\to\e571.exe" "D:\demo\demo.e" `
  --autolinker-headless-compile `
  --autolinker-output "D:\demo\build\demo.exe" `
  --autolinker-target auto `
  --autolinker-result "D:\demo\build\compile-result.json"
```

`AutoLinkerTest.exe` also exposes provider/self-test subcommands:
`headless-compile`, `mock-mcp-stdio`, `mcp-self-test`, `gameanalytics-self-test`,
`version-text`, `version-compare`, `deepseek-model-test`, `openai-chat-test`,
`openai-responses-test`, `gemini-model-test`, `claude-model-test`,
`linker-out` / `linker-krnln` / `between-dashes`.

## Building & testing AutoLinker itself (contributors)

Build `fne_release` / x86 (VS2026 VC++, ISO C++20):
```powershell
MSBuild.exe ..\AutoLinker.vcxproj /t:Build "/p:Configuration=fne_release;Platform=Win32" /m
```
MSBuild.exe location varies by machine/VS version — locate it first.

To test the built library: close the e main program (to release the
`AutoLinker.fne` lock), copy `AutoLinker.fne` into the IDE `lib` dir, then open
an `.e` under `eproj/` and exercise the feature. MCP starts on port 19207 on
load. Pass Chinese page names as **UTF-8**. For logic that doesn't need the IDE
(e.g. module local-parse), use the `AutoLinkerTest` project. For real-scenario
triage, launch `e5.95.exe` on a target file, drive the AI chat, and correlate
with the log directory.

## Common issues
- 连通性测试失败 → check API key (stray spaces), Base URL + protocol, account
  balance, and overseas-network reachability for OpenAI/Claude.
- e-packager download fails or stalls → first retry at AutoLinker 设置 → 关于 →
  组件更新. If it still fails or remains stuck, download the latest Release
  from https://github.com/aiqinxuancai/e-packager/releases and extract it into
  `{易语言安装目录}\tools`. Confirm the executable is exactly
  `{易语言安装目录}\tools\e-packager.exe`, without an extra archive directory
  level. Users in an AutoLinker community group can alternatively obtain
  e-packager from the group shared files and extract it to the same location.
- Config file location → `{易语言安装目录}\AutoLinker\AIConfig.json`.

## References
- White paper: https://github.com/aiqinxuancai/Awesome-E-Agent
- e-packager: https://github.com/aiqinxuancai/e-packager
- Core-lib static impl reference: https://github.com/zhongjianhua163/BlackMoonKernelStaticLib
