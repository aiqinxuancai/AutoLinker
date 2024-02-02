# AutoLinker

AutoLinker支持库，通过各种方法实现以下功能：
* 不同的.e源文件使用不同的链接器
* 调试、编译时动静态ec自动切换
* 强制将Lib链接到核心库之前（并开启/FORCE忽略冲突），用现代方法替换核心库函数

## 使用方法
编译后将AutoLinker.fne放在e的lib目录中，并启动候选AutoLinker支持库。

### ⭐不同的.e源文件使用不同的链接器
此功能实现针对不同.e源文件使用不同链接器，再也不用来回切换链接器了。

支持E版本全部

**使用方法**
  * 在`e\AutoLinker\Config`中，存放`link.ini`的各种版本，命名如：`vc6.ini`、`vc2017.ini`、`vc2022.ini`等。
    ![image](https://github.com/aiqinxuancai/AutoLinker/assets/4475018/a32fbc89-8ea1-4ccc-8f08-b245019ca81d)

  * 启动e后可在顶部菜单栏看到`当前源文件所对应的链接器`。<br>
    ![image](https://github.com/aiqinxuancai/AutoLinker/assets/4475018/a4ab4cea-2b1d-4532-9c43-5175f298e2b9)

### ⭐调试、编译时动静态ec自动切换
此功能实现一对ec文件，在调试、编译时，自动进行切换。

如VMP的SDK、ExDui模块，编译要用Lib声明，调试用Dll声明，来回切换模块非常烦人，所以实现了此功能，调试和编译实现成对模块的自动切换。算是类似C++的 #if DEBUG 的功能吧。

支持版本`5.71`、`5.95` （其他的未测试但应该行）

**使用方法**
* 在`e\AutoLinker\ModelManager.ini`中添加内容既可，格式如下，一行一个，**注意俩ec需要放在同一个文件夹中**，左侧为调试模块、右侧为编译模块
  ```
  VMPSDK.ec=VMPSDK_LIB.ec
  ```
### ⭐制将Lib链接到核心库之前（并开启/FORCE忽略冲突）
此功能可用现代C++方法替换核心库函数，大幅提升部分函数性能，同时也可防止专门针对e函数特征码的一些破解。同理你也可以覆盖第三方库的函数，比如特殊功能支持库等等。


**使用方法**
* 在IDA中查找正确的`函数签名`并声明，打开`krnln_static.lib`中的`LibFn.obj`，大部分核心库函数在此实现
  ![image](https://github.com/aiqinxuancai/AutoLinker/assets/4475018/33d718a7-1a36-4973-b7a6-ee22860879d8)

* 可以在黑月核心库的开源中参考其代码进行新的实现
  https://github.com/zhongjianhua163/BlackMoonKernelStaticLib

* 在自己的Lib中实现函数，实现请参考TestCore项目。
  ```c
  /// <summary>
  /// 使用现代C++方法实现寻找文本，大约是核心库的300%速度，未仔细测试，仅为覆盖实现的例子
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
  * 静态编译既可，会提示如下内容，不会影响代码运行
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

自己写的LIB文件需要关闭链接器/GL参数，否则会触发二次链接导致忽略由AutoLinker调整的Lib顺序。
  
  
### ⭐按实现鼠标后退键后退到上个修改
没啥可说的，其他IDE都有的功能，个人喜好。  

