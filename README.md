# AutoLinker

AutoLinker支持库，通过各种方法实现以下功能：
* 不同的.e源文件使用不同的链接器
* 调试、编译时动静态ec自动切换

## 使用方法
编译后将AutoLinker.fne放在e的lib目录中，并启动候选AutoLinker支持库。

## 功能介绍

* **不同的.e源文件使用不同的链接器**
  * 由于我一部分源码无法使用最新的链接器，所以全换的话，又过于麻烦，又不想手动切换器切来切去，所以实现了此功能。
  * 支持E版本全部
  * 使用方法
    * 在`e\AutoLinker\Config`中，存放`link.ini`的各种版本，命名如：`vc6.ini`、`vc2017.ini`、`vc2022.ini`等。
      ![image](https://github.com/aiqinxuancai/AutoLinker/assets/4475018/a32fbc89-8ea1-4ccc-8f08-b245019ca81d)
  
    * 启动e后可在顶部菜单栏看到`当前源文件所对应的链接器`。<br>
      ![image](https://github.com/aiqinxuancai/AutoLinker/assets/4475018/a4ab4cea-2b1d-4532-9c43-5175f298e2b9)

* **调试、编译时动静态ec自动切换**
  * 由于一部分源码需要使用VMP的SDK，但VMP编译要用Lib声明，调试要用Dll声明，来回切换模块非常烦人，所以实现了此功能，调试和编译实现成对模块的自动切换。算是类似C++的 #if DEBUG 的功能吧。
  * 支持版本`5.71`、`5.95` （其他的未测试但应该行吧）
  * 使用方法
    * 在`e\AutoLinker\ModelManager.ini`中添加内容既可，格式如下，一行一个，注意俩ec需要放在同一个文件夹中，左侧为调试模块、右侧为编译模块
      ```
      VMPSDK.ec=VMPSDK_LIB.ec
      ```
  
* **按实现鼠标后退键后退到上个修改**
  * 没啥可说的，其他IDE都有的功能。  

