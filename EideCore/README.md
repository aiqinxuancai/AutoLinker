# EideCore：易语言 5.95 工程与编译核心

已提取的原生机器码以可重定位 x86 COFF 对象直接链接进 DLL、静态 LIB 和 Console。**普通构建与运行不需要 e5.95.exe，不使用旁置核心 DLL 或工作进程。**

这是可运行的研究实现，尚未完成全部裁剪和指针审计。原 IDE 资源和主要启动/界面入口已删除，仍有未裁净的 MFC、编辑器代码及全局数据。详见 [研究结论与边界](docs/research.md)。

## 共同核心和三个工程

```text
assets/e595/native-core.obj + src/native + src/sdk
    ├─ EideCore         → EideCore.dll + EideCore.lib（导入库）
    └─ EideCore.Static  → EideCoreStatic.lib（真正静态库）
              └─ EideCore.Console → EideCore.exe
```

DLL 和静态库链接同一份 native-core.obj；Console 链接静态库，避免重复维护实现。静态调用方的 EXE 内包含原生核心代码，不加载 EideCore.dll。三个工程均仅支持 **Win32 / x86**；x64 宿主不能直接加载此核心，主解决方案的 x64 配置也只构建 x86 EideCore 产物。

| 目录 | 内容 |
| --- | --- |
| include | 公开 C ABI |
| projects | 三个 VC++ 工程和共享配置 |
| src/native、src/sdk、src/console | 原生适配、API、命令行 |
| assets/e595 | 原生 OBJ、裁剪映像、引用证据和清单 |
| scripts | 日常构建与资产校验 |
| tools/extraction、tools/research | 再提取工具、IDA 调研与探针 |
| examples、tests、docs | 示例、回归、设计与验证记录 |

## 构建

Windows + VS2026 x86 VC 工具链 + Windows SDK。C++20；Release 使用 /MT，Debug 使用 /MTd。普通构建不需要 Python 或原 IDE。

```powershell
./EideCore/scripts/build.ps1 -BuildExamples
MSBuild.exe EideCore/EideCore.sln /t:Build "/p:Configuration=Release;Platform=Win32" /m
```

默认产物在 bin/EideCore/Release；脚本支持 -OutputDirectory、-Configuration 和 -VisualStudioDirectory。三个工程已接入 AutoLinker.sln。

推送发布标签时，GitHub Actions 比较构建输入指纹。输入未变化且上一可达 Release 的清单与附件校验通过时，原字节复用 EXE 和 ZIP，不重新编译或压缩；否则构建 Release / Win32 控制台，将当前 tag 写入命令行和 Windows 版本资源。ZIP 名为 `EideCore-<当前发布版本>-win-x86.zip`，复用时 EXE 和组件版本保留原构建 tag。

ZIP 根目录仅含 `EideCore.exe`，外部依赖仍由用户提供。`update.json` 的 `components.eidecore` 记录组件版本、ZIP/EXE 大小与 SHA-256、构建输入指纹及复用来源。CI 检查 `--help`、`--version` 和版本资源一致性；完整编译回归仍需易语言外部依赖。普通本地构建版本默认为 `0.0.0-dev`。

AutoLinker 普通启动时缺少 `tools/EideCore.exe` 会后台尝试一次下载，已有安装不检查更新；关于页可主动下载安装或手动更新。SHA-256 仅作完整性校验。详见 [组件分发与验证](../docs/eidecore-distribution.md)。

## 控制台与外部依赖

链接器、支持库、静态库、模块和包装模板不随核心分发。未指定 --e-dir（API 的 e_directory）时，从**调用时当前工作目录**开始逐级检查父目录，使用最近一个同时包含 **ecom、lib、static_lib 三个文件夹**的目录。当前目录符合条件时优先使用；查至根目录仍未找到则保留当前目录，由原有加载逻辑报告依赖缺失。显式指定的目录不参与向上查找，分别指定的依赖目录仍优先：

```powershell
$core = 'D:/git/AutoLinker/bin/EideCore/Release/EideCore.exe'
Set-Location 'C:/Users/aiqin/OneDrive/e5.6'
& $core validate --project D:/work/input.e
& $core snapshot --project D:/work/input.e --output D:/work/snapshot.e
& $core compile --project D:/work/input.e --output D:/work/app.exe
& $core compile --project D:/work/input.e --output D:/work/static.exe --static
& $core compile --e-dir D:/elang --project D:/work/input.e --output D:/work/app2.exe
```

`--target`（JSON `target`）默认 `auto`，按工程本身的类型编译。编译操作可指定 `win_exe`、`win_console_exe`、`win_dll`，只临时覆盖内存模型，本次请求结束恢复。适用普通和静态编译；转换到 DLL 仍须满足原编译器的公开子程序要求。`validate` / `snapshot` 不接受非 auto 目标；核心尚不发布 EC 模块。AutoLinker 工具集成见 [接入说明](../docs/eidecore-tool-compile.md)。

| 选项 | JSON 字段 | 默认位置 |
| --- | --- | --- |
| --e-dir | e_directory | 当前目录或最近符合条件的祖先目录；均不符合则回退当前目录 |
| --lib-dir | library_directory | e_directory/lib |
| --static-lib-dir | static_library_directory | e_directory/static_lib |
| --tools-dir | tools_directory | e_directory/tools |
| --linker-dir | linker_directory | e_directory/linker |
| --module-dir | module_directory | e_directory/ecom |
| --package-dir | package_directory | e_directory/epackage |

linker_directory 指向包含 VC6linker 等后端的 linker 目录，后端由 tools/link.ini 决定。单独目录通过临时目录联接适配原生路径约定，请求结束后删除联接，不复制外部依赖。

请求和响应均为 UTF-8 JSON。原生 ANSI 路径仍受系统代码页约束，已在 CP936 下验证中文路径。--output 和 --result 要求新路径；--request request.json 可直接执行请求。退出码：0 成功、1 操作失败、2 参数或结果分配失败。

## DLL / 静态 LIB

公开头文件：[include/EideCore.h](include/EideCore.h)。C ABI、__cdecl：

```cpp
#include "EideCore.h"
char* result = nullptr;
int status = EideCoreExecute(
    R"({"operation":"compile","project_path":"D:/work/input.e","output_path":"D:/work/app.exe","e_directory":"D:/elang"})",
    &result);
if (result) EideCoreFree(result);
```

- DLL：调用方定义 EIDE_CORE_IMPORTS，链接 EideCore.lib，部署 EideCore.dll。
- 静态：链接 EideCoreStatic.lib、bcrypt.lib、user32.lib，设置 /SAFESEH:NO，CRT 与库一致。无需公开 DLL 或其他核心文件。
- Console：只需部署 EideCore.exe 和用户配置的外部依赖。

EideCoreVersion 返回 ABI **2**，结果 schema_version 为 **3**。同步返回 0=成功、1=有 JSON 的失败、2=无可用结果；必须用 EideCoreFree 释放响应。

当前进程内执行，互斥锁将并发调用串行化。原生 CRT/MFC 和全局上下文初始化一次，保留到进程结束；首次执行固定承载模块，**不支持卸载 DLL、取消编译或强制超时**。严重 SEH 异常后上下文作废，后续返回 native_context_poisoned；普通语义错误后可继续加载和编译。旧 runtime_directory、timeout_ms 字段明确拒绝。

```powershell
./EideCore/examples/build.ps1 -RuntimeDirectory ./bin/EideCore/Release
./bin/EideCore/Release/sdk-examples/client-static-x86.exe request.json
./bin/EideCore/Release/sdk-examples/client-dll-x86.exe request.json
```

示例脚本的 RuntimeDirectory 是产物目录参数，不是 API 请求字段。

## 原生诊断

数字直接读取原生错误管理器，未根据文本猜测：

| 情况 | code | raw_code |
| --- | ---: | ---: |
| 文本赋给整数 | 10044 | 16787260 |
| 未知命令 / 子程序 | 37 | 16777253 |

返回 severity、message 和 error_location 的 page、symbol、symbol_id、section、row、column。新增行号和源码字段同时用于 Console、DLL、静态 LIB；无需新参数，仍为 ABI 2 / schema_version 3 的兼容扩展。

| 字段 | 含义 |
| --- | --- |
| line、line_basis | 从 1 开始的行号；基准为返回的 source_text，值为 native_page_text_1based |
| code、code_excerpt | 错误所在完整代码行，以及前后最多两行上下文，保留缩进和空行 |
| excerpt_start_line、focus_line | 片段首行在代码页中的行号；错误行在片段内从 1 开始的行号 |
| source_text | 从原生模型还原的当前代码页，UTF-8、LF 换行，含 .版本 2；折叠子程序按展开形式输出 |
| location_verified、source_status | 是否通过原生前缀、单行及完整代码页文本的交叉核验，以及未映射原因 |
| location_kind | statement、subprogram_declaration、parameter_declaration 或 local_declaration |
| mapped_section、mapped_row、native_code_offset | 实际映射使用的位置；原 section/row/column 保持原导航语义 |
| diagnostic_index | 唯一对应错误在 diagnostics 数组中的索引；不能明确关联时为 null |

diagnostics[i].location 只附到可明确关联的单条错误；其他诊断为 null。没有有效位置的加载/链接错误返回 line=null、location_verified=false，不猜测代码行。部分加载错误仅有文本，code/raw_code=0；链接器消息也不保证有结构化错误码。

实测类型错误 10044：原生 section=3,row=4 对应返回代码页第 **18** 行，code 为 r ＝ “TYPE_ERROR_EXPECTED”。原生折叠子程序可能只报告声明位置，核心现保留编译器语句偏移，因此仍可定位其内部语句。

**line 不是外部解包 TXT 的承诺行号，column 仍是原生坐标。** 原生格式化器可能在流程结束位置增加空行，.e 也不是按行保存的文本文件。调用方应使用返回的 source_text 打开代码页并定位 line，不能直接拿原生 row 或 line 跳转到未经核对的导出 TXT。详细契约及边界见 [源码诊断](docs/source-diagnostics.md)。

execution_mode=in_process_coff、process_id 和 core_address 用于核对直接执行。self_integrity_verified 表示原生算法验证内嵌裁剪映像成功；core_code_intact 比较适配完成后的代码，不表示原始字节完全未改。

## 无窗口与产物保护

不运行原 WinMain / InitInstance，不加载原 IDE 资源。界面入口接入适配器，原生及 link.dll 的相关窗口 API 被拒绝；当前线程 CBT hook 在创建阶段拒绝窗口，外部链接器使用 CREATE_NO_WINDOW。需要交互时返回错误，不自动确认。

回归通过 PID 检查 WinEvent 创建、显示、销毁事件；已覆盖用例未观察到窗口。这不是对所有第三方支持库、自建线程和未知执行分支的绝对保证。

产物在目标同卷临时目录生成，验证后不覆盖地重命名发布。编译结果检查基本 x86 PE 结构；snapshot 在**同一进程**重载，明确返回 snapshot_reload_isolated=false。需要崩溃隔离时，调用方应显式运行 Console 验证。AutoLinker 的 [原生窗口编辑事务](../docs/native-window-editing.md) 已接入 XML 回包、独立进程 `validate` 预检及 IDE 内部备份恢复；`validate` 验证原生加载，源码编译检查仍使用 `compile`。

## 提取与回归

[资产说明](assets/e595/README.md)记录来源和重新生成方法。仅再提取需要精确原版和 Python；机器码来自原版二进制，不是取得原版 C++ 源代码。

```powershell
python EideCore/tests/verify.py --runtime bin/EideCore/Release --e-dir D:/elang --output-dir artifacts/eidecore-projects
python EideCore/tests/verify_direct.py --runtime bin/EideCore/Release --e-dir D:/elang --fixtures artifacts/eidecore-projects --output-dir artifacts/eidecore-direct-test
python EideCore/tests/verify_linkage.py --runtime bin/EideCore/Release --result artifacts/eidecore-linkage.json
python EideCore/tests/verify_source_diagnostics.py --runtime bin/EideCore/Release --e-dir D:/elang --output-dir artifacts/eidecore-source-diagnostics
```

第二条需先构建示例，独占锁定原 EXE 以验证执行不读取它。锁定仅用于测试，不是产品依赖。测试不启动原 IDE，也不运行窗口类产物。

编译产物 DLL 的函数调用测试另记进程 PID；其易语言运行库可创建 OLE/隐藏运行时窗口，相关事件保留在 compiled_artifact_window_events，不计作编译核心窗口。

最新裁剪记录：[dead-code-20261105.md](docs/dead-code-20261105.md)及[验证摘要](docs/validation/dead-code-20261105/summary.json)。直接链接架构基线保留在 [direct-20261104/summary.json](docs/validation/direct-20261104/summary.json)；此前 extraction/sdk 目录是历史架构记录。

## 尚未完成

- 全调用图裁剪和所有重定位候选的类型审计；仍有 GUI 死代码，.native 当前含读写执行权限。
- 全局状态及加载模块的完整销毁；多个独立上下文或进程内真正并行编译。
- 完整窗口工程的 WF_MENU 模块更新目前返回 headless_interaction_required；EC 编译、自定义 CompileProcessor、密码输入和任意第三方支持库未支持或未全面验证。
- 长期宿主和多进程并发下的原生目录、临时文件行为仍需扩展验证；串行锁仅覆盖单个链接实例。

