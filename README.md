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
- 启动成功后，E 输出窗口会打印类似日志：
  ```text
  [AutoLinker][LocalMCP] listening on http://127.0.0.1:19207/mcp
  ```

### ⭐协议说明
- 协议：`JSON-RPC 2.0`
- MCP `initialize` 返回的 `protocolVersion` 为：`2024-11-05`
- 当前支持的方法：
  - `initialize`
  - `notifications/initialized`
  - `ping`
  - `tools/list`
  - `tools/call`

### ⭐客户端接入配置示例

AutoLinker 提供本地 HTTP MCP 服务，你可以在其他AI Agent中使用，请确保客户端支持 `streamable_http` 或 SSE 协议。

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
- **Type**: `SSE` (或 `http`)
- **URL**: `http://127.0.0.1:19207/mcp`

### ⭐工程源码读写模型

- 会话开始或首次源码工具调用时，AutoLinker 会用 e-packager 将当前 IDE 工程（包含未保存改动）解包到源文件目录下的 `.temp/al_*` 临时镜像目录；源码目录不可写时会回退到系统临时目录。
- AI 侧不再区分“真实源码 / 工程缓存 / IDE 搜索结果 / 模块公开代码 / 支持库公开代码”。定位源码统一使用 `list_files`、`search_code`、`read_file`，路径均为解包镜像内的相对路径。
- `edit_file`、`multi_edit_file`、`write_file`、`diff_file`、`restore_file_snapshot` 以 `file_path` 为目标。内部会把文件路径映射回易语言程序项，再读取 IDE 真实整页源码，修改后写回 IDE；不会使用 e-packager 回包编译。
- 写回成功后，当前解包镜像会被标记为过期；下一次 `list_files`、`search_code`、`read_file` 会重新解包，确保读到最新源码。
- `src/*.xml` 是窗口界面描述文件，只用于读取和搜索，不作为普通代码编辑目标。固定表文件（如常量、全局变量、DLL 声明、数据类型）可以通过对应文件路径编辑，内部会映射到 IDE 的真实表页。
- 写入程序集变量时会兼容易语言 IDE 的插入问题，将需要写回的 `.程序集变量` 行按 IDE 可接受格式处理为 `.局部变量`。

### ⭐tools/list 返回的当前公开工具

| 类别 | 方法 | 说明 |
| --- | --- | --- |
| 文件读取 | `list_files` | 按 glob 模式列出当前工程解包镜像内的文件 |
| 文件读取 | `search_code` | 在解包镜像内做内容搜索，支持文件过滤、上下文、大小写、数量限制等参数 |
| 文件读取 | `read_file` | 读取解包镜像内指定文件，按带行号的 `cat -n` 风格返回，可指定 offset/limit |
| 文件编辑 | `edit_file` | 对指定 `file_path` 做精确文本替换；写回前读取 IDE 真实整页源码 |
| 文件编辑 | `multi_edit_file` | 对指定 `file_path` 批量执行多个精确文本替换 |
| 文件编辑 | `write_file` | 用完整源码覆盖指定 `file_path` 对应的真实 IDE 页面，可带 `expected_base_hash` |
| 文件编辑 | `diff_file` | 预览指定 `file_path` 的编辑差异，不写回 |
| 文件编辑 | `restore_file_snapshot` | 恢复最近一次写入前保存的页面快照 |
| 当前页 | `get_current_page_info` | 获取当前页名称、类型与解析来源 |
| 当前页 | `get_current_eide_info` | 获取当前源码路径、IDE 进程路径、本地 MCP 端口等实例信息 |
| 编译 | `compile_with_output_path` | 指定输出路径发起编译/静态编译 |
| 本地交互 | `run_powershell_command` | 经过确认后执行 PowerShell 命令 |
| 联网 | `search_web_tavily` | 联网搜索网页结果 |
| 联网 | `fetch_url` | 抓取指定 URL 原始文本响应 |
| 联网 | `extract_web_document` | 提取网页正文与链接摘要 |


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
