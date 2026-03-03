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

  * 启动e后可在顶部菜单栏看到`当前源文件所对应的链接器`。<br>
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

### 1. 入口位置
- 代码编辑区右键菜单：`AutoLinker` 子菜单。
- 工具菜单：`AutoLinker AI接口设置`（用于配置接口地址、Key、模型等）。

### 2. 首次使用配置
在 `AutoLinker AI接口设置` 中填写：
- `Base URL`：兼容 OpenAI Chat Completions 的接口地址。
- `API Key`
- `Model`
- `Timeout(ms)`：默认 `120000`。
- `Temperature`：默认 `0.2`。
- `额外系统提示词`（可选）。

未配置完整时，调用 AI 功能会自动弹出配置窗口。

### 3. AI 功能列表（右键 -> AutoLinker）
1. `AI优化函数`
- 作用：对当前函数做优化（保持行为等价），返回完整可替换函数。
- 使用：光标放在函数内，输入优化要求（可留空），确认后预览并替换。

2. `AI为当前函数添加注释`
- 作用：为当前函数补充函数注释与关键行注释，不改逻辑。
- 使用：光标放在函数内，确认预览后替换当前函数。

3. `AI翻译当前函数+变量名`
- 作用：将函数名/参数名/局部变量名翻译或重命名为英文 `lowerCamelCase`。
- 约束：不会翻译或修改以 `.` 开头的系统指令（如 `.子程序`、`.参数`、`.局部变量`、`.如果`、`.返回` 等）。
- 使用：光标放在函数内，确认预览后替换当前函数。

4. `AI翻译文本`
- 作用：翻译当前选中文本，只输出翻译结果到输出窗口，不改代码。
- 注意：未选中文本时该菜单项会自动禁用（灰色）。

5. `AI为当前函数补全API声明`
- 作用：根据当前函数，补全缺失的 `.DLL命令/.参数/.数据类型/.成员` 声明。
- 输出：仅声明，不包含函数实现。
- 应用：预览确认后插入到当前页（会尽量跳过已存在声明）。

6. `AI按当前页类型添加代码`
- 作用：根据“当前页类型 + 你的需求 + 当前页完整代码”生成新增代码。
- 支持页类型：程序集/类、DLL 命令声明、数据类型（以及通用代码页）。
- 输出约束：仅返回“新增内容”，不返回整页、不附加解释。
- 应用：预览确认后追加到当前页底部。

### 4. 统一执行流程
1. 读取当前上下文（当前函数/选中文本/当前页代码）。
2. 后台请求模型（不会阻塞主线程）。
3. 弹出结果预览。
4. 你确认后再执行替换/插入；取消则不改动代码。

### 5. 使用建议与排查
- 进行“函数类”操作时，确保光标位于目标函数内部。
- 若日志提示“无法获取当前函数代码”，先点击函数体任意行再重试。
- 若接口返回异常，先检查 `Base URL / API Key / Model` 是否可用。
- 推荐保持源码换行为 `CRLF`，避免不同编辑器换行差异影响替换定位。
- AI 请求与应用过程会输出 `[AutoLinker][AI]` 日志，可用于排查问题。