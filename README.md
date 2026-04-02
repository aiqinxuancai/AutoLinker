# AutoLinker

AutoLinker支持库，通过逆向实现易语言上的AI Agent，即代码全自动编写，会自动根据你的需求，查找、获取相关代码，并直接对代码进行编辑、修改、插入功能。

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

### ⭐客户端接入 JSON 示例

#### 通用 HTTP MCP 配置
- 如果你的 MCP 客户端支持基于 HTTP 的 MCP，可直接配置为，默认端口19207，如果占用则顺延+1：
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
  
- AutoLinker 当前提供的是本地 HTTP MCP 服务，不是 `stdio` 型 MCP。

### ⭐tools/list 返回的当前公开工具

| 类别 | 方法 | 说明 |
| --- | --- | --- |
| 当前页 | `get_current_page_code` | 获取当前页完整源码与页面信息 |
| 当前页 | `get_current_page_info` | 获取当前页名称、类型与解析来源 |
| 模块 | `list_imported_modules` | 列出当前项目导入模块 |
| 支持库 | `list_support_libraries` | 列出当前已选支持库 |
| 支持库 | `get_support_library_info` | 获取单个支持库公开信息 |
| 支持库 | `search_support_library_info` | 搜索支持库公开信息 |
| 模块 | `get_module_public_info` | 获取模块公开接口信息 |
| 模块 | `search_module_public_info` | 搜索模块公开接口信息 |
| 程序树 | `list_program_items` | 列出程序树项目，可选附带参考代码 |
| 程序树真实源码 | `get_program_item_real_code` | 获取指定程序树页面真实整页源码 |
| 程序树真实源码 | `read_program_item_real_code` | 读取真实源码缓存或分页视图 |
| 程序树真实源码 | `edit_program_item_code` | 按精确文本编辑真实源码页 |
| 程序树真实源码 | `multi_edit_program_item_code` | 批量编辑真实源码页 |
| 程序树真实源码 | `write_program_item_real_code` | 用整页源码覆盖写回 |
| 程序树真实源码 | `diff_program_item_code` | 生成真实源码页差异预览 |
| 程序树真实源码 | `restore_program_item_code_snapshot` | 恢复真实源码页快照 |
| 程序树真实源码 | `search_program_item_real_code` | 在真实源码页内搜索 |
| 程序树真实源码 | `list_program_item_symbols` | 解析并列出页面符号 |
| 程序树真实源码 | `get_symbol_real_code` | 获取符号对应真实源码 |
| 程序树真实源码 | `edit_symbol_real_code` | 按符号替换真实源码 |
| 程序树真实源码 | `insert_program_item_code_block` | 按位置或符号插入代码块 |
| 程序树 | `switch_to_program_item_page` | 切换到指定程序树页面 |
| 搜索 | `search_project_keyword` | 调用 IDE 整体搜索 |
| 搜索 | `jump_to_search_result` | 跳转到指定搜索结果 |
| 编译 | `compile_with_output_path` | 指定输出路径发起编译/静态编译 |
| 本地交互 | `run_powershell_command` | 经过确认后执行 PowerShell 命令 |
| 联网 | `search_web_tavily` | 联网搜索网页结果 |
| 联网 | `fetch_url` | 抓取指定 URL 原始文本响应 |
| 联网 | `extract_web_document` | 提取网页正文与链接摘要 |

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
