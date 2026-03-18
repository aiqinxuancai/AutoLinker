# AutoLinker

AutoLinker支持库，通过各种方法实现以下功能：
* 不同的.e源文件使用不同的链接器
* 调试、编译时动静态ec自动切换
* 重写核心库函数

## 使用方法
编译后将AutoLinker.fne放在e的lib目录中，并启用AutoLinker支持库。

### ⭐不同的.e源文件使用不同的链接器
此功能实现针对不同.e源文件使用不同链接器，再也不用来回切换链接器了。

支持E版本全部

**使用方法**
  * 在`e\AutoLinker\Config`中，存放`link.ini`的各种版本，命名如：`vc6.ini`、`vc2017.ini`、`vc2022.ini`等。
    ![image](https://github.com/aiqinxuancai/AutoLinker/assets/4475018/f73e8188-2011-469a-aee2-f9d5f4e2af01)

  * `当前源文件所对应的链接器`。**！！！此位置新版本已改到编译菜单下！！！**<br>
    ![image](https://github.com/aiqinxuancai/AutoLinker/assets/4475018/a4ab4cea-2b1d-4532-9c43-5175f298e2b9)

### ⭐调试、编译时动静态ec自动切换
此功能实现一对ec文件，在调试、编译时，自动进行切换。

如VMP的SDK、ExDui模块，编译要用Lib声明链接进exe，调试用Dll声明（因为无法e无法在调试时使用Lib），来回切换模块非常烦人，所以实现了此功能，调试和编译时实现成对模块的自动切换。

支持版本`5.71`、`5.95` （其他的未测试但应该行）

**使用方法**
* 在`e\AutoLinker\ModelManager.ini`中添加内容既可，格式如下，一行一个，**注意俩ec需要放在同一个文件夹中**，左侧为调试模块、右侧为编译模块
  ```
  VMPSDK.ec=VMPSDK_LIB.ec
  ```
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

## AutoLinker AI 功能使用说明

### ⭐入口位置
- 代码编辑区右键菜单：`AutoLinker` 子菜单（函数类操作）。
- 工具菜单：`AutoLinker AI接口设置`（配置接口地址、Key、模型等）。
- 输出区页签：`AutoLinker AI 会话`（常驻会话面板，支持异步返回）。

未配置完整时，调用 AI 功能会自动弹出配置窗口。

### ⭐右键菜单 AI 功能
1. `AI优化函数`
- 对当前函数做等价优化。
- 返回后会弹出可复制的结果窗口，并提供 `替换（不稳定）` 选项。

2. `AI为当前函数添加注释`
- 为当前函数补充函数注释和关键行注释，不改业务逻辑。

3. `AI翻译当前函数+变量名`
- 将函数名、参数名、局部变量名翻译/重命名为英文 `lowerCamelCase`。
- 不修改以 `.` 开头的系统指令行（如 `.子程序`、`.参数`、`.局部变量`、`.如果` 等）。

4. `AI翻译选中文本`
- 仅翻译当前选中文本，不改代码。
- 未选中文本时该菜单项会自动禁用。

5. `AI按当前页类型添加代码`
- 根据“当前页类型 + 你的需求 + 当前页上下文”生成新增代码。
- 仅返回新增内容，不返回整页；确认后追加到当前页。

### ⭐AI 会话页签（内置 TAB）
<img width="701" height="268" alt="image" src="https://github.com/user-attachments/assets/9933305f-ad8d-4c47-bb85-316a25efea0d" />

- 支持本地 MCP 工具调用（如读取当前页代码、弹出代码修改对话框）。

## AutoLinker 本地 MCP 文档

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
- 如果你的 MCP 客户端支持基于 HTTP 的 MCP，可直接配置为：
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

#### 某些客户端的简化写法
- 有些客户端不要求显式写 `transport`，可写成：
  ```json
  {
    "mcpServers": {
      "AutoLinker": {
        "url": "http://127.0.0.1:19207/mcp"
      }
    }
  }
  ```

#### 如果端口被自动顺延
- 当 `19207` 被占用时，AutoLinker 会自动尝试后续端口，所以客户端配置中的 `url` 需要与 E 输出窗口中的实际监听地址保持一致，例如：
  ```json
  {
    "mcpServers": {
      "AutoLinker": {
        "transport": "streamable_http",
        "url": "http://127.0.0.1:19208/mcp"
      }
    }
  }
  ```

#### 说明
- AutoLinker 当前提供的是本地 HTTP MCP 服务，不是 `stdio` 型 MCP。
- 如果某个客户端只支持 `stdio` 方式而不支持 HTTP/Streamable HTTP，则不能直接接入当前版本的 AutoLinker MCP。

### ⭐HTTP 访问说明
- `GET /` 或 `GET /mcp`
  - 返回服务健康信息和当前 `mcp_endpoint`
- `POST /mcp`
  - 发送 JSON-RPC 请求
- `OPTIONS /mcp`
  - CORS 预检

### ⭐tools/list 返回的当前公开工具

#### 1. 当前页相关
- `get_current_page_info`
  - 获取当前 IDE 页名称、页类型以及名称解析来源。
- `get_current_page_code`
  - 获取当前 IDE 页完整代码，同时返回当前页名称和页类型。

#### 2. 模块相关
- `list_imported_modules`
  - 列出当前项目导入的易模块/ECOM 路径。
- `get_module_public_info`
  - 获取指定模块的公开接口信息。
  - 说明：这里返回的是模块公开接口/伪代码参考，不是正常 IDE 源码页。
- `search_module_public_info`
  - 在模块公开接口信息中搜索关键词。
  - 说明：搜索结果同样属于公开接口/伪代码参考。

#### 3. 支持库相关
- `list_support_libraries`
  - 列出当前 IDE 已选支持库。
  - 返回支持库基本信息、文件路径解析结果以及原始信息文本。
- `get_support_library_info`
  - 获取单个支持库公开信息。
  - 优先通过支持库文件 `GetNewInf/lib2.h` 结构解析命令、常量、数据类型等；无法定位文件时退回 IDE 文本。
- `search_support_library_info`
  - 在支持库公开信息中搜索关键词。

#### 4. 程序树/项目结构相关
- `list_program_items`
  - 列出程序树中的程序集、类模块、全局变量、自定义数据类型、DLL 命令、窗口、资源等项目。
  - 支持按 `kind`、名称过滤，并可选附带代码。
  - 说明：这里通过程序树抓到的代码是伪代码参考，结构可能与正常 IDE 页略有差异。
- `get_program_item_code`
  - 按精确名称获取某个程序树项目的整页代码。
  - 说明：返回代码属于伪代码参考。
- `switch_to_program_item_page`
  - 按精确名称切换到某个程序集/类/资源页面。
  - 只负责切页，不返回代码。

#### 5. 项目搜索相关
- `search_project_keyword`
  - 调用 IDE 整体搜索，在项目内搜索关键词。
  - 返回匹配页名、页类型、行号、显示文本和 `jump_token`。
- `jump_to_search_result`
  - 使用 `search_project_keyword` 返回的 `jump_token` 精确跳转到某一条结果。

#### 6. 本地交互相关
- `request_code_edit`
  - 弹出本地代码编辑确认窗口，返回用户确认后的代码。

### ⭐tools/call 参数说明
- 调用格式：
  ```json
  {
    "jsonrpc": "2.0",
    "id": 1,
    "method": "tools/call",
    "params": {
      "name": "get_current_page_info",
      "arguments": {}
    }
  }
  ```
- `params.name`
  - 工具名称
- `params.arguments`
  - 工具参数对象

### ⭐常用调用示例

#### 列出工具
```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "tools/list",
  "params": {}
}
```

#### 获取当前页信息
```json
{
  "jsonrpc": "2.0",
  "id": 2,
  "method": "tools/call",
  "params": {
    "name": "get_current_page_info",
    "arguments": {}
  }
}
```

#### 获取当前页代码
```json
{
  "jsonrpc": "2.0",
  "id": 3,
  "method": "tools/call",
  "params": {
    "name": "get_current_page_code",
    "arguments": {}
  }
}
```

#### 搜索项目关键词
```json
{
  "jsonrpc": "2.0",
  "id": 4,
  "method": "tools/call",
  "params": {
    "name": "search_project_keyword",
    "arguments": {
      "keyword": "subWinHwnd",
      "limit": 20
    }
  }
}
```

#### 跳转到某条搜索结果
```json
{
  "jsonrpc": "2.0",
  "id": 5,
  "method": "tools/call",
  "params": {
    "name": "jump_to_search_result",
    "arguments": {
      "jump_token": "v1:1:1224804111:315:0:0"
    }
  }
}
```

### ⭐会改变 IDE 当前页的工具
- 以下工具具有副作用，会导致当前 IDE 页面或光标位置发生变化：
  - `get_program_item_code`
  - `switch_to_program_item_page`
  - `jump_to_search_result`
- 使用这些工具前后，不要假定当前页仍保持不变。

### ⭐关于“伪代码参考”的说明
- 以下来源拿到的代码/结构，不等同于 IDE 中用户正在编辑的原始页内容：
  - `get_module_public_info`
  - `search_module_public_info`
  - `list_program_items` 中附带的代码
  - `get_program_item_code`
  - `search_project_keyword` 的结果文本及其后续关联代码
- 这些内容适合做定位、检索、接口参考，但不应当当作 100% 原样源码。

