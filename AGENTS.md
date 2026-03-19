## 项目概要

本项目是一个易语言支持库（插件类型），最终编译产物为AutoLinker.fne（实际是dll类型），在易语言主进程启动时加载，用于简化、帮助易语言的开发流程，提供AI Agent来自动化完成部分工作。

### 项目规范

* 项目的依赖要精简，基于Win32+MFC实现，避免引入vcpkg包。
* 注重项目的可维护性，对于可独立工作的工具类、管理类，应单独封装。
* 避免单文件代码量过于庞大难以维护。
* 必要时使用mcp来对e.exe进行逆向来实现功能。
* 对.h头文件始终添加简单的中文注释。
* 源代码文件保持UTF8+BOM的格式保存，注意中文是否存在乱码并进行修正。
* 源代码文件使用windows下的换行符（CR LF）。
* 项目使用ISO C++20 标准 (/std:c++20)，可使用其特性开发。

### 编译测试

请在完成修改后验证fne_release+x86是否可正常编译，

请参考此命令行方法进行编译，不同设备和VS版本的MSBuild.exe的位置可能不同，请注意寻找及确认。

```powershell
'C:\Program Files\Microsoft Visual Studio\18\Professional\MSBuild\Current\Bin\MSBuild.exe' ..\AutoLinker.vcxproj /t:Build "/
  │ p:Configuration=fne_release;Platform=Win32"
```

如果编译存在报错，请进行修正，确保功能实现正常的情况下可正常编译。

