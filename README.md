# AutoLinker

AutoLinker支持库，通过逆向实现易语言上的AI Agent，即代码全自动编写，会自动根据你的需求，查找、获取相关代码，并直接对代码进行编辑、修改、插入功能。

同时提供MCP服务，你可以将易语言IDE接入到Codex、Claude Code、Gemini Cli等工具中，其也可以自行调用易语言全自动编写代码，具体参考下方MCP相关内容。

> 📖 [易语言 × AI Agent 实践白皮书](https://github.com/aiqinxuancai/Awesome-E-Agent)
> 
> 📖 [e-packager，拆解.e为txt修改后再打包，兼容现代Agent](https://github.com/aiqinxuancai/e-packager)  

## 使用方法

下载Release后将AutoLinker.fne放在易语言的lib目录中，并启用AutoLinker支持库。

> 🔧 [AI、MCP 功能配置指南（API Key / 中转站 / 各平台配置方法）](./CONFIG.md)

### ⭐AI Agent 会话页签

AI可以全自动搜索、读取、编辑当前工程源码，你只需告诉他你要做什么即可。

<img width="611" height="702" alt="image" src="https://github.com/user-attachments/assets/301f5ecf-5078-4c30-a9d6-c23284c2a22e" />


### ⭐右键菜单 AI 功能
1. `AI优化函数` 对当前函数做等价优化。
2. `AI为当前函数添加注释`
3. `AI翻译当前函数+变量名` 将函数名、参数名、局部变量名翻译/重命名为英文 `lowerCamelCase`。
4. `AI翻译选中文本`
5. `AI按当前页类型添加代码` 根据“当前页类型 + 你的需求 + 当前页上下文”生成新增代码。

### ⭐项目规范文件 `{文件名}.AGENTS.md`

你可以在 `.e` 源文件的同目录下，创建一个与其同名的 `.AGENTS.md` 文件，AutoLinker 会自动将其内容注入到所有 AI 功能的系统提示词中，作为"项目规范"，效果类似 Claude Code 的 `CLAUDE.md` / Codex 的 `AGENTS.md`。

**命名规则**

| 源文件 | 规范文件 |
| --- | --- |
| `test_a.e` | `test_a.AGENTS.md`（同目录） |
| `MyProject.e` | `MyProject.AGENTS.md`（同目录） |

---

### ⭐不同的.e源文件使用不同的链接器

此功能实现针对不同.e源文件使用不同链接器，**再也不用来回手动切换链接器了**。

**使用方法**

在工具菜单中进行设置，可添加多组link.ini配置，并可在【主菜单->编译】下切换当前代码源文件所使用的链接器。

### ⭐调试、编译时动态/静态 ec 模块自动切换

为同一模块准备一对 ec 文件（动态版 / 静态版），AutoLinker 会在**开始调试**和**开始编译**时自动替换工程中已导入的对应模块，**再也不用来回手动切换了**。

典型场景如 VMP 的 SDK、ExDui 等模块：

- **编译**时需要用 Lib 声明的静态版，才能把模块链接进最终 exe；
- **调试**时只能用 Dll 声明的动态版（易语言调试环境无法加载 Lib 版）。

以往每次在调试与编译之间切换都要手动改导入，非常繁琐，此功能将这一步完全自动化。

**工作原理**

AutoLinker 挂接了 IDE 的“开始调试 / 开始编译”入口：触发时遍历当前工程已导入的 ec 模块，按文件名匹配配置表，命中后就地把该模块替换为对应版本（移除旧的、导入新的）。

**使用方法**

在易语言 IDE 的「工具」菜单中打开 **AutoLinker EC模块自动切换设置**，在设置页中维护切换规则。

注意：你需要自己先引用动、静任意一个ec，成对的两个 ec 必须放在**同一个文件夹**中（替换时只改文件名，沿用原模块所在目录）。

---

## AutoLinker 本地 MCP 文档

你可以通过配置把易语言MCP接入到Codex、Claude Code、Gemini Cli等工具。

### ⭐服务地址
- 默认监听：`http://127.0.0.1:19207/mcp`
- 如果端口 `19207` 被占用，会自动尝试 `19208` 起的后续端口。
- 启动成功后，IDE 输出窗口和 `autolinker.log` 会记录类似日志：
  ```text
  [AutoLinker][LocalMCP] 本地 MCP 服务已启动：http://127.0.0.1:19207/mcp
  ```

### ⭐协议说明
- 协议：`JSON-RPC 2.0`
- MCP `initialize` 会协商并支持：`2025-11-25`、`2025-03-26`、`2024-11-05`（未知版本回退到 `2025-11-25`）
- 当前支持的方法：
  - `initialize`
  - `notifications/initialized`
  - `ping`
  - `tools/list`
  - `tools/call`
  - `DELETE /mcp`（释放当前 `Mcp-Session-Id`）

安全边界：服务只绑定 `127.0.0.1`，拒绝带非空 `Origin` 的浏览器请求；暂不提供浏览器 CORS 或 Bearer Token。19207 外部 MCP 调用不弹交互审批窗口，但仍执行工具白名单、参数 Schema、工作区刷新和源码哈希 CAS 校验。原生 MCP 客户端每个会话必须先成功调用一次 `refresh_workspace_mirror`，刷新凭据会绑定当前工程路径和镜像代次，工程切换或其他会话改写工程后需重新刷新。服务使用 4 个固定连接工作线程，最多排队 32 个连接，过载时返回 HTTP 503。

### ⭐客户端接入配置示例

AutoLinker 提供本地 HTTP MCP 服务，你可以在其他 AI Agent 中使用，请确保客户端支持 MCP Streamable HTTP。

#### 1. Claude Code
配置文件：`~/.claude.json` (JSON)
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

#### 2. Gemini CLI
配置文件：`~/.gemini/settings.json` (JSON)
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

#### 3. Codex
配置文件：`~/.codex/config.toml` (**TOML**)
```toml
[mcp_servers.AutoLinker]
url = "http://127.0.0.1:19207/mcp"
```

#### 4. Cursor / Windsurf / IDE
在 IDE 的 MCP 设置页面（Settings -> Features -> MCP）中添加：
- **Name**: `AutoLinker`
- **Type**: `http` / `streamable_http`
- **URL**: `http://127.0.0.1:19207/mcp`

### ⭐工程源码读写模型

- 内置 AI 对话会在每轮请求开始前以 `full` 模式自动准备工程镜像；外部 MCP 会话首次读取或编辑源码前必须先成功调用 `refresh_workspace_mirror`。AutoLinker 会用 e-packager 将当前 IDE 工程（包含未保存改动）解包到源文件目录下的 `.temp/al_<pid>_*` 临时镜像目录，源码目录不可写时回退到系统临时目录。
- `mode` 支持 `auto`（默认）、`main_only`（仅主工程源码）和 `full`（完整全量刷新）。不同 IDE 进程的镜像目录互不清理，只有确认所有者进程退出后才回收。
- 镜像定位统一使用 `list_files`、`search_code`、`read_file`、`read_files`、`read_code_item`，路径均为解包镜像内的相对路径；搜索支持最多 16 个 pattern 的单次逐文件流式扫描，可命中 1 MiB 之后的内容。
- 大文件读取返回 `next_source_byte_offset`，后续把它作为 `byte_offset` 传回即可继续；分页时建议同时回传 `mirror_generation`，工程刷新或写入后会明确拒绝旧代次游标。
- 编辑前使用 `read_real_file` 读取 IDE 真实页的分页编号视图和 `code_hash`；该工具不再重复返回完整整页 `code`。
- `edit_file`、`multi_edit_file`、`write_file`、`diff_file`、`restore_file_snapshot` 以 `file_path` 为目标。内部会把文件路径映射回易语言程序项，再读取 IDE 真实整页源码，修改后写回 IDE；不会使用 e-packager 回包编译。
- 写工具支持 SHA-256 `expected_base_hash`（恢复工具使用 `expected_current_hash`）以阻止旧基线覆盖新修改；外部 MCP 调用必须提供该哈希。写回成功后会优先原子增量更新对应镜像文件，无法增量更新时才将镜像标记为过期。
- 写入、恢复成功结果只返回哈希、快照、验证和变更统计，不再重复返回完整整页源码；MCP 完整结果位于 `structuredContent`，`content.text` 仅提供摘要。
- `src/*.xml` 是窗口界面描述文件，只用于读取和搜索，不作为普通代码编辑目标。固定表文件（如常量、全局变量、DLL 声明、数据类型）可以通过对应文件路径编辑，内部会映射到 IDE 的真实表页。
- 写入程序集变量时会兼容易语言 IDE 的插入问题，将需要写回的 `.程序集变量` 行按 IDE 可接受格式处理为 `.局部变量`。

### ⭐tools/list 返回的当前公开工具

| 类别 | 方法 | 说明 |
| --- | --- | --- |
| 文件读取 | `refresh_workspace_mirror` | 从 IDE 当前内存工程刷新解包镜像；`mode` 支持 `auto`、`main_only`、`full` |
| 文件读取 | `list_files` | 按 glob 模式列出当前工程解包镜像内的文件 |
| 文件读取 | `search_code` | 在解包镜像内逐文件搜索；支持批量 patterns、glob、上下文和分页 |
| 文件读取 | `read_file` | 读取解包镜像内指定文件，按带行号的 `cat -n` 风格返回，可指定 offset/limit |
| 文件读取 | `read_files` | 批量读取多个文件或多个区间，部分失败时返回 `status=partial` |
| 文件读取 | `read_code_item` | 按易语言顶层代码项名称读取完整子程序/声明块 |
| 文件读取 | `read_real_file` | 从 IDE 真实页返回分页编号视图和 `code_hash`，供写入前建立 CAS 基线 |
| 文件编辑 | `edit_file` | 对指定 `file_path` 做精确文本替换；写回前读取 IDE 真实整页源码 |
| 文件编辑 | `multi_edit_file` | 对指定 `file_path` 批量执行多个精确文本替换 |
| 文件编辑 | `write_file` | 用完整源码覆盖真实 IDE 页面，可带 `expected_base_hash` 防止覆盖新修改 |
| 文件编辑 | `diff_file` | 基于真实 IDE 页预览结构化差异，不写回，可校验 `expected_base_hash` |
| 文件编辑 | `restore_file_snapshot` | 恢复写入前快照，可用 `expected_current_hash` 防止覆盖新修改 |
| 当前页 | `get_current_page_info` | 获取当前页名称、类型与解析来源 |
| 当前页 | `get_current_eide_info` | 获取当前源码路径、IDE 进程路径、本地 MCP 端口等实例信息 |
| 编译 | `compile_with_output_path` | `target` 默认 `auto`；以编译前后高精度产物指纹验证成功，不只相信 IDE 返回值 |
| 本地交互 | `run_powershell_command` | 经确认后执行 PowerShell；授权按内部对话/外部 MCP Session 隔离，超时会终止整个进程树 |
| 联网 | `search_web_tavily` | 联网搜索网页结果 |
| 联网 | `fetch_url` | 抓取公网 HTTP/HTTPS 原始文本；阻止回环、私网、链路本地和重定向到私网 |
| 联网 | `extract_web_document` | 提取网页正文与已解析为绝对地址的链接摘要 |


## 其他功能

### ⭐无头命令行编译

推荐使用 `AutoLinkerTest headless-compile` 启动 e.exe。启动器会写入一次性请求文件、枚举并关闭 IDE 启动期 `MessageBox`，AutoLinker 加载后隐藏 IDE、自动调用 `compile_with_output_path`，并把成功/失败、IDE 输出、错误位置和结果 JSON 输出到控制台。

```powershell
.\bin\fne_release\AutoLinkerTest.exe headless-compile `
  "C:\Users\aiqin\OneDrive\e5.6\e571.exe" `
  "D:\demo\demo.e" `
  "D:\demo\build\demo.exe" `
  --target auto --static `
  --result "D:\demo\build\compile-result.json" `
  --timeout 120
```

`target` 支持 `auto`、`win_exe`、`win_console_exe`、`win_dll`、`ecom`；`--static` 仅适用于 EXE/DLL，易模块只支持普通编译。结果默认同时写到 `e\AutoLinker\Log\headless_compile_last.json`。FNE 内部只能处理 AutoLinker 加载后的弹窗；启动器的父进程窗口枚举用于捕获 `.e` 加载失败、缺少支持库、缺少易模块等更早期错误，并会分别输出 `support_libraries` 和 `list_items`。IDE 编译链路里的输出目标选择会被自动抑制，并以 `compile_dialogs` 输出。其他后续 MsgBox 会自动关闭并以 `kind=info` 记录。

也可以直接启动易语言主程序，无需 `AutoLinkerTest.exe`：

```powershell
"C:\Users\aiqin\OneDrive\e5.6\e571.exe" `
  "D:\demo\demo.e" `
  --autolinker-headless-compile `
  --autolinker-output "D:\demo\build\demo.exe" `
  --autolinker-target auto `
  --autolinker-result "D:\demo\build\compile-result.json"
```

直接传参只负责无头编译；如果还要处理启动早期弹窗，继续用 `AutoLinkerTest headless-compile`。


### ⭐重写核心库函数
此功能可用现代C++方法替换核心库函数，大幅**提升函数性能**，同时也可**防针对e函数特征的破解**以及**免杀**的效果。同理你也可以覆盖第三方库的函数，比如特殊功能支持库等等。


**使用方法**
* 在IDA中查找正确的`函数签名`并声明，打开`krnln_static.lib`中的`LibFn.obj`，大部分核心库函数在此实现
  ![image](https://github.com/aiqinxuancai/AutoLinker/assets/4475018/33d718a7-1a36-4973-b7a6-ee22860879d8)

* 可以在黑月核心库的开源中参考其代码进行新的实现
  https://github.com/zhongjianhua163/BlackMoonKernelStaticLib

* 在自己的Lib中实现函数，实现请参考TestCore项目。
  ```c
  /// <summary>
  /// 使用C++20（需VC2022）方法替代核心库的寻找文本，大约是核心库的300%速度，未仔细测试，仅为覆盖实现的例子
  /// </summary>
  /// <param name="pRetData"></param>
  /// <param name="nArgCount"></param>
  /// <param name="pArgInf"></param>
  /// <returns></returns>
  extern "C" void __cdecl krnln_fnInStr(PMDATA_INF pRetData, INT nArgCount, PMDATA_INF pArgInf) {
  
      //TODO 你还可以在这里添加VMP标志
  
      // 获取输入字符串
      std::string_view inputString = pArgInf[0].m_pText;
      std::string_view searchString = pArgInf[1].m_pText;
  
      // 空字符串检查
      if (inputString.empty() || searchString.empty()) {
          pRetData->m_int = -1;
          return;
      }
  
      // 确定开始搜索的位置
      size_t searchStartPos = (pArgInf[2].m_dtDataType == _SDT_NULL || pArgInf[2].m_int <= 1) ? 0 : pArgInf[2].m_int - 1;
  
      // 搜索指定字符串
      auto searchResult = (pArgInf[3].m_bool) ?
          std::search(
              inputString.begin() + searchStartPos, inputString.end(),
              searchString.begin(), searchString.end(),
              [](char c1, char c2) { return std::tolower(c1) == std::tolower(c2); }
          ) :
          std::search(
              inputString.begin() + searchStartPos, inputString.end(),
              searchString.begin(), searchString.end()
          );
  
      // 设置返回位置，如果找到则返回位置，否则返回-1
      pRetData->m_int = (searchResult != inputString.end()) ? std::distance(inputString.begin(), searchResult) + 1 : -1;
  }
  ```
  * 使用Release编译为32位lib，然后在易语言 IDE 的「工具」菜单中打开 **AutoLinker 核心库函数重写设置**，在设置页中维护要强制链接的 .lib 列表：
    * **Lib 路径**：要追加链接的 `.lib` 文件路径（会被排在 `krnln_static.lib` 之前以覆盖同名符号）。
    * **链接器匹配**：留空表示对所有链接器生效；填写文本则仅当当前源码所选链接器名称**包含**该文本时才使用此 Lib（例如填 `vc2022+pf` 表示仅 VC2022 链接器生效）。
    * **启用开关**：可临时停用某条规则而不必删除。
    * 点击保存即生效，下次静态编译时自动按规则追加链接。

    **注意**：例子TestCore使用C++20，必须使用VC2022链接器才可链接；该设置页基于 WebView2，请确保系统已安装 WebView2 运行时。
    
  * 进行静态编译既可，会提示如下内容，不会影响代码运行
    ```
    krnln_static.lib(Libfn.obj) : warning LNK4006: _krnln_fnBAnd 已在 TestCore.lib(a1.obj) 中定义；已忽略第二个定义
    krnln_static.lib(Libfn.obj) : warning LNK4006: _krnln_fnSin 已在 TestCore.lib(a1.obj) 中定义；已忽略第二个定义
    krnln_static.lib(Libfn.obj) : warning LNK4006: _krnln_fnInStr 已在 TestCore.lib(a1.obj) 中定义；已忽略第二个定义
    nafxcw.lib(afxmem.obj) : warning LNK4006: "void * __cdecl operator new(unsigned int)" (??2@YAPAXI@Z) 已在 LIBCMT.lib(new_scalar.obj) 中定义；已忽略第二个定义
    nafxcw.lib(afxmem.obj) : warning LNK4006: "void __cdecl operator delete(void *)" (??3@YAXPAX@Z) 已在 LIBCMT.lib(delete_scalar.obj) 中定义；已忽略第二个定义
    C:\Users\aiqin\OneDrive\Code\模块\测试空.exe : warning LNK4088: 因 /FORCE 选项生成了映像；映像可能不能运行
    正在写出可执行文件
    ```


**注意**
* 此方法自动开启链接器的/FORCE忽略冲突
* 自己写的LIB文件需要关闭链接器/GL参数，否则会触发二次链接导致忽略由AutoLinker调整的Lib顺序。
  
  
### ⭐按实现鼠标后退键后退到上个修改
没啥可说的，其他IDE都有的功能，个人喜好。  
