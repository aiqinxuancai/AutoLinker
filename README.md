# AutoLinker

AutoLinker支持库，通过逆向实现易语言上的AI Agent，即代码全自动编写，会自动根据你的需求，查找、获取相关代码，并直接对代码进行编辑、修改、插入功能。

同时提供MCP服务，你可以将易语言IDE接入到Codex、Claude Code、Gemini Cli等工具中，其也可以自行调用易语言全自动编写代码，具体参考下方MCP相关内容。

> 📖 [易语言 × AI Agent 实践白皮书](https://github.com/aiqinxuancai/Awesome-E-Agent)
> 
> 📖 [e-packager，拆解.e为txt修改后再打包，兼容现代Agent](https://github.com/aiqinxuancai/e-packager)  

## 使用方法
编译后将AutoLinker.fne放在易语言的lib目录中，并启用AutoLinker支持库。

### ⭐右键菜单 AI 功能
1. `AI优化函数` 对当前函数做等价优化。
2. `AI为当前函数添加注释`
3. `AI翻译当前函数+变量名` 将函数名、参数名、局部变量名翻译/重命名为英文 `lowerCamelCase`。
4. `AI翻译选中文本`
5. `AI按当前页类型添加代码` 根据“当前页类型 + 你的需求 + 当前页上下文”生成新增代码。

### ⭐AI Agent 会话页签

<img width="989" height="644" alt="QQ20260401-134616" src="https://github.com/user-attachments/assets/7a08de65-5107-4d0e-a6f1-3314cb813d67" />

- 支持本地 MCP 工具调用（如读取当前页代码、搜索代码、搜索支持库内容、搜索模块内容）。
- 当前工程源码相关场景已支持 `project_cache` 路线：搜索与读代码优先走内存序列化后的工程源码缓存，真正写回前再抓取真实页源码。

### ⭐项目规范文件 `{文件名}.AGENTS.md`

你可以在 `.e` 源文件的同目录下，创建一个与其同名的 `.AGENTS.md` 文件，AutoLinker 会自动将其内容注入到所有 AI 功能的系统提示词中，作为"项目规范"，效果类似 Claude Code 的 `CLAUDE.md` / Codex 的 `AGENTS.md`。

**命名规则**

| 源文件 | 规范文件 |
| --- | --- |
| `test_a.e` | `test_a.AGENTS.md`（同目录） |
| `MyProject.e` | `MyProject.AGENTS.md`（同目录） |

**生效范围**

- ⭐AI Agent 会话页签（聊天对话）
- ⭐右键菜单 AI 功能（优化、注释、翻译等）
- 所有内嵌 AI 调用

**文件格式**

- 编码：`UTF-8 with BOM`（推荐）或 `UTF-8 without BOM`
- 换行：CRLF 或 LF 均可
- 内容：标准 Markdown，写清楚项目背景、技术栈约定、命名规范、禁止事项等

**示例文件内容**

```markdown
# 项目规范

## 技术背景
本项目是一个用于 XXX 的易语言程序，目标运行环境为 Windows XP+，

## 项目约定
请在完成修改后执行编译，如果报错请先修复。

## 命名约定
- 子程序名使用中文，格式：`动词_名词`
- 局部变量使用小驼峰英文命名

## 禁止事项
- 不得引入新的第三方支持库
- 不得修改 `_启动子程序` 子程序

```

> 文件不存在时不影响任何功能，AutoLinker 会静默忽略。

---

### ⭐不同的.e源文件使用不同的链接器

此功能实现针对不同.e源文件使用不同链接器，**再也不用来回手动切换链接器了**。

**使用方法**
  * 在`e\AutoLinker\Config`中，存放`link.ini`的各种版本，命名如：`vc6.ini`、`vc2017.ini`、`vc2022.ini`等。
    ![image](https://github.com/aiqinxuancai/AutoLinker/assets/4475018/f73e8188-2011-469a-aee2-f9d5f4e2af01)

  * `当前源文件所对应的链接器`。**！！！此位置新版本已改到编译菜单下！！！**<br>
    ![image](https://github.com/aiqinxuancai/AutoLinker/assets/4475018/a4ab4cea-2b1d-4532-9c43-5175f298e2b9)

### ⭐调试、编译时动态静态ec自动切换

此功能实现一对ec文件，在调试、编译时，自动进行切换，如VMP的SDK、ExDui模块，编译要用Lib声明链接进exe，调试用Dll声明（因为无法e无法在调试时使用Lib），来回切换模块非常麻烦，所以实现了此功能。

支持版本`5.71`、`5.95` （其他的未测试但应该行）

**使用方法**
* 在`e\AutoLinker\ModelManager.ini`中添加内容既可，格式如下，一行一个，**注意俩ec需要放在同一个文件夹中**，左侧为调试模块、右侧为编译模块
  ```
  VMPSDK.ec=VMPSDK_LIB.ec
  ```

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

AutoLinker 提供本地 HTTP MCP 服务，请确保客户端支持 `streamable_http` 或 SSE 协议。

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

### ⭐tools/list 返回的当前公开工具

| 类别 | 方法 | 说明 |
| --- | --- | --- |
| 当前页 | `get_current_page_code` | 获取当前页完整源码与页面信息 |
| 当前页 | `get_current_page_info` | 获取当前页名称、类型与解析来源 |
| 当前页 | `get_current_eide_info` | 获取当前源码路径、IDE 进程路径、本地 MCP 端口等实例信息 |
| 多实例 | `list_local_mcp_instances` | 列出本机已启动的 AutoLinker MCP 实例 |
| 多实例 | `call_local_mcp_instance_tool` | 转发调用其他本机实例上的 MCP 工具 |
| 模块 | `list_imported_modules` | 列出当前项目导入模块 |
| 模块 | `search_available_module_public_code` | 搜索易语言 `ecom` 目录下所有可用 `.ec` 模块的公开声明文本 |
| 模块 | `add_module_to_project` | 按路径或模块名将 `.ec` 模块加入当前工程 |
| 支持库 | `list_support_libraries` | 列出当前已选支持库 |
| 支持库 | `get_support_library_info` | 获取单个支持库公开信息 |
| 完全搜索 | `search_public_code` | 默认先刷新并搜索 `project_cache`，再补充当前 IDE 工程源码命中、模块公开声明、支持库公开声明，支持多关键字或正则 |
| 支持库 | `search_support_library_public_code` | 按行搜索支持库公开声明文本，支持多关键字或正则 |
| 支持库 | `read_support_library_public_code` | 读取支持库公开声明文本指定行范围 |
| 模块 | `get_module_public_info` | 获取模块公开接口信息 |
| 模块 | `search_module_public_code` | 按行搜索模块公开声明文本，支持多关键字或正则 |
| 模块 | `read_module_public_code` | 读取模块公开声明文本指定行范围 |
| 程序树 | `list_program_items` | 列出程序树项目，可选附带参考代码 |
| 程序树缓存源码 | `get_program_item_project_cache_code` | 从当前工程源码缓存读取指定程序树页面，不切换 IDE 页面 |
| 程序树真实源码 | `get_program_item_real_code` | 获取指定程序树页面真实整页源码 |
| 程序树真实源码 | `read_program_item_real_code` | 读取真实源码缓存或分页视图 |
| 程序树真实源码 | `edit_program_item_code` | 按精确文本编辑真实源码页 |
| 程序树真实源码 | `multi_edit_program_item_code` | 批量编辑真实源码页 |
| 程序树真实源码 | `write_program_item_real_code` | 用整页源码覆盖写回 |
| 程序树真实源码 | `diff_program_item_code` | 生成真实源码页差异预览 |
| 程序树真实源码 | `restore_program_item_code_snapshot` | 恢复真实源码页快照 |
| 程序树真实源码 | `search_program_item_real_code` | 在真实源码页内搜索，支持单关键字或正则 |
| 程序树真实源码 | `list_program_item_symbols` | 解析并列出页面符号 |
| 程序树真实源码 | `get_symbol_real_code` | 获取符号对应真实源码 |
| 程序树真实源码 | `edit_symbol_real_code` | 按符号替换真实源码 |
| 程序树真实源码 | `insert_program_item_code_block` | 按位置或符号插入代码块 |
| 程序树 | `switch_to_program_item_page` | 切换到指定程序树页面 |
| 工程源码缓存 | `refresh_project_source_cache` | 强制刷新当前工程源码缓存，仅使用内存二进制序列化 |
| 工程源码缓存 | `search_project_source_cache` | 仅搜索当前工程源码缓存，搜索前会先刷新缓存 |
| 工程源码缓存 | `read_project_source_cache_code` | 基于工程源码缓存命中读取指定行附近代码，通常不切页 |
| 搜索 | `search_project_keyword` | 仅搜索当前 IDE 工程源码命中，仅支持单关键字 |
| 搜索 | `read_project_search_result_code` | 基于 IDE 搜索命中读取真实页代码行范围 |
| 搜索 | `jump_to_search_result` | 跳转到指定搜索结果 |
| 编译 | `compile_with_output_path` | 指定输出路径发起编译/静态编译 |
| 本地交互 | `run_powershell_command` | 经过确认后执行 PowerShell 命令 |
| 联网 | `search_web_tavily` | 联网搜索网页结果 |
| 联网 | `fetch_url` | 抓取指定 URL 原始文本响应 |
| 联网 | `extract_web_document` | 提取网页正文与链接摘要 |

### 无头命令行编译

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

`target` 支持 `auto`、`win_exe`、`win_console_exe`、`win_dll`、`ecom`；`--static` 仅适用于 EXE/DLL，易模块只支持普通编译。结果默认同时写到 `e\AutoLinker\Log\headless_compile_last.json`。FNE 内部只能处理 AutoLinker 加载后的弹窗；启动器的父进程窗口枚举用于捕获 `.e` 加载失败、缺少支持库、缺少易模块等更早期错误，并会分别输出 `support_libraries` 和 `list_items`。其他后续 MsgBox 会自动关闭并以 `kind=info` 记录。

### ⭐搜索参数示例

`search_project_source_cache`、`search_public_code`、`search_available_module_public_code`、`search_module_public_code`、`search_support_library_public_code` 这几个工具使用相同的 `keyword / keywords / keyword_mode / regex` 搜索规则，下面是可直接照抄的示例：

单关键字字面量搜索：

```json
{"keyword":".子程序 初始化"}
```

多关键字 AND，同一行同时包含全部关键字：

```json
{"keywords":[".子程序","初始化"],"keyword_mode":"all"}
```

多关键字 OR，命中任意一个关键字即可：

```json
{"keywords":["初始化","创建"],"keyword_mode":"any"}
```

使用正则 OR：

```json
{"regex":"初始化|创建"}
```

注意：`{"keyword":"初始化|创建"}` 不会按“或”处理，而是去搜索字面量 `初始化|创建`。如果要做 OR，请使用 `keywords + keyword_mode=\"any\"` 或 `regex`。

其他搜索工具的差异：

- `search_project_keyword` 走 IDE 自带隐藏搜索，只支持单关键字，不支持多关键字或正则。
- `search_program_item_real_code` 支持 `keyword + use_regex=false` 的单关键字搜索，或 `keyword + use_regex=true` 的单正则搜索；它不支持 `keywords` 数组。

---

## 其他功能

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
  * 使用Release编译为32位lib，将lib的路径添加到`e\AutoLinker\ForceLinkLib.txt`中，你可以指定链接器才使用此Lib，也可以单独一行指定在所有链接器中使用
    ```
    vc2022+pf=D:\git\X\Release\TestCore.lib
    D:\git\X\Release\TestCore.lib
    ```
    **注意**：例子TestCore使用C++20，必须使用VC2022链接器才可链接
    
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
