<div align="center">
  <img src="res/icon128.png" width="128" alt="AutoLinker" />
  <h1>AutoLinker</h1>
</div>

AutoLinker 是易语言的 AI Agent 支持库，通过逆向让 AI 全自动编写代码：根据你的需求自动查找、读取相关代码，并直接编辑、修改、插入。

同时提供本地 MCP 服务，可将易语言 IDE 接入 Codex、Claude Code、Gemini CLI 等工具，让它们调用易语言全自动编码（见下方 MCP 说明）。

> 📖 [易语言 × AI Agent 实践白皮书](https://github.com/aiqinxuancai/Awesome-E-Agent) ·
> 📖 [e-packager：拆解 .e 为 txt 修改后再打包](https://github.com/aiqinxuancai/e-packager)

## 快速开始

下载 Release 后，将 `AutoLinker.fne` 放入易语言 `lib` 目录并启用该支持库。

> 🔧 [AI / MCP 功能配置指南（API Key、中转站、各平台配置）](./CONFIG.md)

## 核心功能

### ⭐ AI Agent 会话页签
告诉 AI 你要做什么，它会自动搜索、读取、编辑当前工程源码。

<img width="611" alt="AI Agent 会话页签" src="https://github.com/user-attachments/assets/301f5ecf-5078-4c30-a9d6-c23284c2a22e" />

### ⭐ 右键菜单 AI 功能
- **AI 优化函数** — 对当前函数做等价优化
- **AI 添加注释** — 为当前函数生成注释
- **AI 翻译函数 + 变量名** — 重命名为英文 `lowerCamelCase`
- **AI 翻译选中文本**
- **AI 按当前页类型添加代码** — 依据当前页类型、上下文和你的需求生成新增代码

### ⭐ 项目规范文件 `{文件名}.AGENTS.md`
在 `.e` 源文件同目录下创建同名 `.AGENTS.md`，其内容会自动注入所有 AI 功能的系统提示词，作为项目规范（类似 Claude Code 的 `CLAUDE.md`、Codex 的 `AGENTS.md`）。

| 源文件 | 规范文件（同目录） |
| --- | --- |
| `test_a.e` | `test_a.AGENTS.md` |
| `MyProject.e` | `MyProject.AGENTS.md` |

### ⭐ 不同 .e 源文件使用不同链接器
再也不用手动来回切换链接器。在工具菜单中添加多组 `link.ini` 配置，并在【主菜单 → 编译】下切换当前源文件所用的链接器。

### ⭐ 调试 / 编译时动静态 ec 模块自动切换
为同一模块准备一对 ec 文件（动态版 / 静态版），AutoLinker 会在**开始调试**和**开始编译**时自动替换工程中已导入的对应模块。典型场景如 VMP SDK、ExDui：编译需 Lib 声明的静态版，调试只能用 Dll 声明的动态版。

在 IDE「工具」菜单打开 **AutoLinker EC 模块自动切换设置** 维护规则。注意：需先自行引用动/静任一 ec，且成对的两个 ec 必须放在**同一文件夹**中（替换时只改文件名）。

## 本地 MCP 服务

将易语言 MCP 接入 Codex、Claude Code、Gemini CLI 等工具，需客户端支持 **MCP Streamable HTTP**。

### 服务地址
- 默认监听 `http://127.0.0.1:19207/mcp`，端口占用时自动向 `19208` 起顺延。
- 启动成功后 IDE 输出窗口和 `autolinker.log` 会记录：
  ```text
  [AutoLinker][LocalMCP] 本地 MCP 服务已启动：http://127.0.0.1:19207/mcp
  ```

### 协议
- `JSON-RPC 2.0`，`initialize` 协商 `2025-11-25` / `2025-03-26` / `2024-11-05`（未知版本回退 `2025-11-25`）。
- 支持方法：`initialize`、`notifications/initialized`、`ping`、`tools/list`、`tools/call`、`DELETE /mcp`。
- **安全边界**：仅绑定 `127.0.0.1`，拒绝带非空 `Origin` 的浏览器请求，无 CORS / Bearer Token。外部调用不弹审批窗，但仍执行工具白名单、参数 Schema、工作区刷新与源码哈希 CAS 校验。原生客户端每个会话须先成功调用一次 `refresh_workspace_mirror`。固定 4 个工作线程，最多排队 32 个连接，过载返回 HTTP 503。

### 客户端配置示例

**Claude Code** — `~/.claude.json`；**Gemini CLI** — `~/.gemini/settings.json`
```json
{
  "mcpServers": {
    "AutoLinker": {
      "transport": "streamable_http",
      "url": "http://127.0.0.1:19207/mcp"
    }
  }
}
```

**Codex** — `~/.codex/config.toml`
```toml
[mcp_servers.AutoLinker]
url = "http://127.0.0.1:19207/mcp"
```

**Cursor / Windsurf / IDE** — MCP 设置页添加：Name `AutoLinker`，Type `http` / `streamable_http`，URL `http://127.0.0.1:19207/mcp`。

### 工程源码读写模型
- 内置 AI 每轮请求前以 `full` 模式自动准备镜像；外部 MCP 会话首次读写前须调用 `refresh_workspace_mirror`。镜像由 e-packager 解包到 `%TEMP%/AutoLinker/workspace-mirror/`（含未保存改动），不污染源码目录。`mode` 支持 `auto` / `main_only` / `full`。
- 读取统一走镜像相对路径（`list_files`、`search_code`、`read_file`、`read_files`、`read_code_item`）；大文件返回 `next_source_byte_offset` 用于续读，分页建议回传 `mirror_generation`，旧代次游标会被拒绝。
- 编辑前用 `read_real_file` 取分页视图和 `code_hash` 作为 CAS 基线。写工具（`edit_file`、`multi_edit_file`、`write_file` 等）以 `file_path` 为目标，映射回 IDE 程序项后直接写回 IDE，不回包编译。
- 写入须带 SHA-256 `expected_base_hash`（恢复用 `expected_current_hash`）防止旧基线覆盖新改动；结果仅返回哈希、快照、验证与变更统计，完整结果在 `structuredContent`。
- `src/*.xml` 为窗口界面文件，仅供读取搜索；固定表（常量、全局变量、DLL 声明、数据类型）可经对应路径编辑。程序集变量写回会按 IDE 可接受格式处理。

### 公开工具（`tools/list`）

| 类别 | 方法 | 说明 |
| --- | --- | --- |
| 读取 | `refresh_workspace_mirror` | 从 IDE 内存工程刷新镜像（`auto` / `main_only` / `full`） |
| 读取 | `list_files` | 按 glob 列出镜像内文件 |
| 读取 | `search_code` | 镜像内逐文件搜索，支持批量 patterns、glob、上下文、分页 |
| 读取 | `read_file` / `read_files` | 读取单个 / 批量文件或区间，带行号 |
| 读取 | `read_code_item` | 按顶层代码项名读取完整子程序 / 声明块 |
| 读取 | `read_real_file` | 从 IDE 真实页返回分页视图与 `code_hash`（写前基线） |
| 编辑 | `edit_file` / `multi_edit_file` | 精确文本替换（单个 / 批量） |
| 编辑 | `write_file` | 完整源码覆盖真实页，支持 `expected_base_hash` |
| 编辑 | `diff_file` | 预览结构化差异，不写回 |
| 编辑 | `restore_file_snapshot` | 恢复写入前快照 |
| 当前页 | `get_current_page_info` | 当前页名称、类型与解析来源 |
| 当前页 | `get_current_eide_info` | 源码路径、IDE 进程路径、MCP 端口等 |
| 编译 | `compile_with_output_path` | `target` 默认 `auto`，以产物指纹验证成功 |
| 交互 | `run_powershell_command` | 经确认后执行 PowerShell，超时终止进程树 |
| 联网 | `search_web_tavily` | 联网搜索网页 |
| 联网 | `fetch_url` | 抓取公网 HTTP(S) 文本，拦截回环 / 私网 / 重定向 |
| 联网 | `extract_web_document` | 提取网页正文与绝对链接摘要 |

## 其他功能

### ⭐ 无头命令行编译
推荐用 `AutoLinkerTest headless-compile` 启动 e.exe：自动关闭启动期弹窗、隐藏 IDE、调用 `compile_with_output_path`，并把结果 JSON 输出到控制台。

```powershell
.\bin\fne_release\AutoLinkerTest.exe headless-compile `
  "C:\path\to\e571.exe" "D:\demo\demo.e" "D:\demo\build\demo.exe" `
  --target auto --static --result "D:\demo\build\compile-result.json" --timeout 120
```

`target` 支持 `auto`、`win_exe`、`win_console_exe`、`win_dll`、`ecom`；`--static` 仅适用于 EXE/DLL。结果默认同时写入 `e\AutoLinker\Log\headless_compile_last.json`。

也可直接启动主程序（仅负责无头编译，早期弹窗仍建议用启动器）：

```powershell
"C:\path\to\e571.exe" "D:\demo\demo.e" `
  --autolinker-headless-compile `
  --autolinker-output "D:\demo\build\demo.exe" `
  --autolinker-target auto `
  --autolinker-result "D:\demo\build\compile-result.json"
```

### ⭐ 重写核心库函数
用现代 C++ 替换核心库函数，可大幅**提升性能**，并具备**防 e 函数特征破解**与**免杀**效果，同样适用于第三方库。

**使用方法**
1. 在 IDA 中确定正确的 `函数签名`（核心库函数大多在 `krnln_static.lib` 的 `LibFn.obj`），可参考[黑月核心库开源实现](https://github.com/zhongjianhua163/BlackMoonKernelStaticLib)。
2. 在自己的 Lib 中实现函数（参考 `TestCore` 项目），Release 编译为 32 位 lib。示例：

   ```c
   // C++20（需 VC2022）重写核心库「寻找文本」，约为核心库 300% 速度
   extern "C" void __cdecl krnln_fnInStr(PMDATA_INF pRetData, INT nArgCount, PMDATA_INF pArgInf) {
       std::string_view inputString = pArgInf[0].m_pText;
       std::string_view searchString = pArgInf[1].m_pText;
       if (inputString.empty() || searchString.empty()) { pRetData->m_int = -1; return; }
       size_t start = (pArgInf[2].m_dtDataType == _SDT_NULL || pArgInf[2].m_int <= 1) ? 0 : pArgInf[2].m_int - 1;
       auto r = (pArgInf[3].m_bool)
           ? std::search(inputString.begin() + start, inputString.end(), searchString.begin(), searchString.end(),
                         [](char a, char b) { return std::tolower(a) == std::tolower(b); })
           : std::search(inputString.begin() + start, inputString.end(), searchString.begin(), searchString.end());
       pRetData->m_int = (r != inputString.end()) ? std::distance(inputString.begin(), r) + 1 : -1;
   }
   ```

3. 在 IDE「工具」菜单打开 **AutoLinker 核心库函数重写设置** 维护要强制链接的 `.lib` 列表：
   - **Lib 路径**：追加链接的 `.lib`（排在 `krnln_static.lib` 之前以覆盖同名符号）
   - **链接器匹配**：留空对所有链接器生效；填文本则仅当所选链接器名称**包含**该文本时生效（如 `vc2022+pf`）
   - **启用开关**：可临时停用某条规则
4. 保存后进行静态编译即生效，会出现 `LNK4006` / `LNK4088` 等告警，不影响运行。

**注意**
- 此方法自动开启链接器 `/FORCE` 忽略冲突。
- 自己的 Lib 需关闭 `/GL`，否则二次链接会忽略 AutoLinker 调整的 Lib 顺序。
- 示例 `TestCore` 使用 C++20，须用 VC2022 链接器；设置页基于 WebView2，请确保已安装运行时。

### ⭐ 鼠标后退键回退到上个修改
其他 IDE 常见的功能，个人喜好。
