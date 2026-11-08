# AutoLinker 错误原因与排查参考

原始排查记录来自 2026-10-19；2026-11-08 补充跨版本恢复、窗口事务、EideCore 分发及日志路由。错误文案会随版本变化；先匹配完整 `error` 和 `trace`，再在对应版本源码中核对。“检查方向”不表示已证实该用户的根因。不是所有错误都会带 trace。

## 按失败阶段导航

- [收集证据](#收集证据)：用户反馈、版本不明、跨电脑复现。
- [内存快照与序列化](#内存快照与序列化)：`memory_entry_unavailable`、`resolve_failed`、`native_abi_mismatch`、`invalid_magic`。
- [已检查的 EXE](#已检查的-exe)：入口兼容性与静态检查边界。
- [工作区镜像与解包](#工作区镜像与解包)：刷新失败、e-packager、文件路径、分页。
- [MCP 会话与调度](#mcp-会话与调度)：HTTP、JSON-RPC、实例选择、响应超时。
- [源码编辑与校验](#源码编辑与校验)：CAS、文本替换、写入回读、快照恢复。
- [项目编译](#项目编译)：项目快照、子进程、结果文件、前后置动作。
- [备份与工程恢复](#备份与工程恢复)：备份落盘、IDE 重新加载、身份恢复。
- [AI 接口与组件安装](#ai-接口与组件安装)：网络/模型线路、下载与解压。

## 收集证据

先记录错误原文、工具名或菜单操作、时间、`trace`、`hint`、请求结果里的 `stage` / `status`。若有 `execution_state`、`retry_safe`、`may_still_be_running`，据此决定能否重试；HTTP 200 不保证业务成功，还需看 `isError`、`structuredContent.ok` 和批量结果的逐项错误。

多实例环境记录 PID、`instance_id`、当前工程完整路径，查 `{易语言目录}\AutoLinker\Log\autolinker_pid<PID>.log` 及同 PID 的启动日志。不要混用另一实例或旧日期的错误。只索取相关时间段日志，避免收集 API Key 和无关源码。

版本不明时，取得实际启动的 IDE EXE、AutoLinker.fne 版本及文件哈希。快捷方式名、补丁文件夹名、EXE 文件名和 Windows 文件版本都不能替代实际二进制身份。在持有该文件的机器上可执行只读检查：

```powershell
$ideExePath = 'C:\实际安装目录\e.exe'
(Get-Item -LiteralPath $ideExePath).VersionInfo |
    Select-Object FileVersion, ProductVersion, FileDescription
Get-FileHash -LiteralPath $ideExePath -Algorithm SHA256
Get-FileHash -LiteralPath $ideExePath -Algorithm MD5
Get-CimInstance Win32_Process |
    Where-Object { $_.ExecutablePath -eq $ideExePath } |
    Select-Object ProcessId, ExecutablePath, CommandLine
```

MD5 仅用于关联既有样本，不作为可信性判断。磁盘上的 FNE 若启动后被替换，不能据此认定旧进程已加载新版；结合启动时间和加载模块核对。

## 内存快照与序列化

源码入口：`src/EideProjectBinarySerializer.cpp` 的 `SerializeCurrentProject`、`PopulateSerializerAddresses`，以及 `src/EideProjectMemoryArchive.cpp` 的 `ResolveApi` / `Capture`。镜像、导出、备份、按配置编译均可能经由同一快照链路失败。

| 错误原文 / trace 关键词 | 实际触发条件 | 检查方向与处理 |
| --- | --- | --- |
| `project memory serializer entry unavailable` / `memory_entry_unavailable` | 已知地址和后备扫描均未得到通过完整验证的入口；尚未解析工程对象或读取工程内容 | 按下文入口诊断。不能从这条错误推断工程损坏、未保存、密码保护、内存不足或网络故障 |
| `current project binary serializer is only available on x86 IDE builds` / `unsupported_platform`；`full memory project archive is only available on x86 IDE builds` | 编译时不是 `_M_IX86` 路径 | 核对插件构建平台和宿主进程位数；本项目要求 `fne_release|Win32` |
| `resolve active editor object failed` / `resolve_active_editor_failed` | 已通过入口阶段，但获取活动编辑器对象失败，且其他上下文路线不可用 | 确认选中实例、已打开工程、当前页与活动编辑器 trace；只读切换到代码页对照。不要归类为入口特征码失败 |
| `active editor serializer field resolution failed` / `field_read_failed` / `field_bad_object` | 从活动编辑器第 23 个指针字段读取失败或对象检查不通过 | 核对 `owner`、对象/字段地址、raw/inner 路线及 EXE 布局；不要猜测新偏移后直接调用 |
| `resolve_failed\|context=serializer_this_missing` | 未记录经保存 Hook 验证的对象，且后续替代路线也失败 | 缺少保存上下文本身不等于必须保存：代码还尝试全局文档对象和活动编辑器字段；查看完整 `field=` 失败原因 |
| `serializer_source_mismatch` | 记录的保存上下文路径与当前工程路径不一致 | 查看 `current=` / `recorded=`，检查切换工程或实例；该片段单独出现不证明整体失败，后备路线可能成功 |
| `resolved serializer object is invalid` / `invalid_serializer_object` | 已解析的对象未通过地址范围/可读性/头部检查 | 核对对象生命周期、当前工程和内部布局；不是磁盘文件写入失败 |
| `full memory project archive ABI is unsupported by this IDE build` / `native_abi_mismatch` | 实际导出前再次校验入口失败，或传入 serializer 为空 | 区别于首次入口发现；检查缓存地址与当前内存字节、页面属性及对象。运行时被修改只是待证实方向 |
| `full memory project serialization failed` | 原生调用有异常、返回值不为 1、句柄为空或流长度为 0 | 保留 `native_abi_verified`、`serialize_result`、`exception`、`bytes`；区分异常与普通失败，不能仅靠 `exception=0x0` 宣称成功 |
| `full memory project archive handle is invalid` | HGLOBAL 容量小于流长度，或 `GlobalLock` 失败 | 已进入内存流获取阶段；检查句柄、长度及内部接口契约，不能靠放宽入口特征解决 |
| `encrypted_source_password_required`，附带 `invalid_magic` | 导出字节以 `WTLE` 开头，被识别为密码保护快照 | 当前不支持解密；需要时由用户在副本上移除源码密码、保存并重新打开，不改动唯一原件 |
| `serialized bytes do not start with CNWTEPRG` / `invalid_magic` | 导出内容没有预期文件头，且不符合已识别的加密分支 | 保存头部和导出 trace，检查对象、格式、兼容性；不可一律解释成密码保护 |
| `snapshot output path is empty` / `output_path_empty`；`snapshot path is empty` | 输出路径为空 | 核对调用方路径生成；前者出现在 serializer，后者也可能出现在 e-packager 集成层 |
| `open output snapshot file failed` / `open_output_failed` | 已生成快照字节，但输出文件无法打开 | 检查实际父目录、访问权限和文件占用；路径问题不能反推入口不可用 |
| `write output snapshot file failed` / `write_output_failed` | 文件写入或关闭后流状态失败 | 检查磁盘空间、设备/同步目录状态和占用；不能将已有文件视为完整快照 |
| `write output snapshot exception: ...` / `write_output_exception` | 创建目录或写出过程中发生 C++ 异常 | 保留异常原文及路径，按具体文件系统错误排查 |
| `snapshot capture requires IDE main thread` / `source query requires IDE main thread` | 内存快照/当前工程查询从非 IDE 主线程调用 | 修正调度到主线程的调用链；调整路径或保存工程无法修复线程调用错误 |

### `memory_entry_unavailable` 的具体诊断

当前实现先验证 image base `0x400000` 下的已知 VA：5.95 的 `0x464050`、5.71 的 `0x45B640`，运行时按模块基址换算。均不通过才扫描短特征码，要求恰好一个候选。随后 `ResolveApi` 校验：

1. 地址属于主 EXE 的已提交可执行内存区域，范围足够且不带 guard。
2. 完整 282 字节（`0x11A`）指令模式匹配，仅放宽已定义的地址字段。
3. 固定偏移提取的 10 个相对 call 目标都在主 EXE 可执行区域内。
4. 成功/失败分支的两个析构目标分别一致。

任何一项失败都会拒绝入口；当前错误没有逐项说明失败位置。查一次性日志：

```text
[ProjectBinarySerializer] memory_entry=0x...|route=...|found=false
```

- `route=pattern_scan` 且 `memory_entry=0x0`：短特征没有唯一命中（零个或多个），当前日志不能区分数量。
- `route=pattern_scan` 且 `memory_entry` 非零、`found=false`：已有唯一短特征候选，但完整验证未通过；并不表示函数根本不存在。
- `route=e5.95` / `route=e571` 且 `found=true`：当次解析已通过。若仍见 unavailable，先核对 PID、时间和 FNE 版本，避免拼接不同进程的日志。

首次解析结果（包括失败）按进程缓存，`ClearVerifiedSerializerContext` 只清对象上下文，不清入口地址缓存。相同进程中反复刷新不会重新扫描。修正环境后可以在保留未保存工作的前提下重启对照；重启不能修复真实 ABI 差异。

需要进一步定位时，先离线比较用户 EXE 中的短特征候选数、完整模式、相对调用目标与析构分支；若磁盘匹配，再只读检查运行时内存和已加载 FNE。第三方 Hook、其他构建或页面属性不同都只是待验证原因。不要把新地址加进表就宣称兼容，也不要为了通过校验删掉参数/栈布局验证：入口匹配用于保护旧版 MFC 对象尺寸与调用约定。

## 已检查的 EXE

以下为 2026-10-19 本机**静态检查**，地址按 image base `0x400000` 表示。文件名仅帮助辨认样本，匹配以哈希为准；通过单个入口不能证明整套功能。当前支持下限为 5.6 / x86；2026-11-08 窗口运行验证已覆盖 5.6、5.71、5.95 和增强补丁 v1.1，正版补丁仅完成 ABI 验证，运行测试受加密狗授权限制。

| 本机样本 | MD5 | 序列化入口检查 |
| --- | --- | --- |
| `e5.95.exe` | `6b0ec5c5a72d242c18a2b47d4a27defe` | `0x464050` 完整模式、可执行段目标及析构分支通过 |
| `e593_to_595_patch/e.exe` | `44d207232edf1b95dac0b2289b151ff7` | 同上，应直接命中 5.95 已知入口；未对该样本进行本次运行时回归 |
| `e571.exe` | `1fb19da090aaca06f8abb1da056e7f9d` | `0x45B640` 通过 |
| 安装目录 `e.exe`（文件版本 5.6） | `d97a7873d87cee2a567af19c00ebbedf` | 唯一候选 `0x45B130`，完整验证通过 |
| `e_正版.exe`（文件版本 5.0） | `13b90a0dde2f3427d60e361ca068485b` | 唯一短特征候选 `0x452970`，完整模式失败；首个非通配差异为 `+0x58` |

5.0 样本的内层序列化 call 之前只有 6 个栈参数，当前 5.95 布局为 7 个，属于实际调用契约差异；不能简单放宽模式。补丁目录样本的 Windows FileVersion / ProductVersion 均为 `5, 9, 0, 0`，不能仅凭该字段确定小版本。其 SHA-256 为 `f37e8adbd861cc4fca8c0be619e40dfbea343529000d5acb9125e05cb61d559a`。

仓库 `docs/test-results/restore-abi-compatibility-20261108/README.md` 记录最新映像身份与测试边界。不要将已验证样本的结果套用到用户未知 EXE，仍需现场日志和实际二进制身份。

## 工作区镜像与解包

源码：`src/WorkspaceMirror.cpp`、`src/WorkspaceFileTools.cpp`、`src/EPackagerIntegration.cpp`。读取失败先定位“生成内存快照 → e-packager 解包 → 解析元数据 → 发布镜像 → 读取”的哪一步。

| 错误原文 / 关键词 | 触发条件与处理 |
| --- | --- |
| `no_source_open` / `当前没有打开易语言源码文件` / `当前打开的不是 .e 源码文件` | 当前目标实例没有合适的源码工程；用 `get_current_eide_info` 核对实例和路径 |
| `workspace_refresh_required` | MCP 会话尚未成功准备工作区；调用 `refresh_workspace_mirror`，如失败则诊断其内层错误，不循环重试读文件 |
| `workspace mirror is not prepared` / `workspace mirror snapshot is empty` | 无可用镜像状态/传入空快照；核对之前刷新是否成功及当前工程是否变更 |
| `workspace_changed` | 后台候选发布时，工程路径、上下文 epoch 或镜像 generation 已改变；确认最终目标后重新刷新 |
| `e-packager 未安装` | 指定工具位置没有可用程序；按主 skill 的手动安装说明放到 `{易语言目录}\tools\e-packager.exe` |
| `CreateProcessW failed, error=...` | 启动解包工具失败；结合 Win32 数值、工具路径和可执行文件状态判断，不视为解包内容错误 |
| `cannot establish e-packager process cleanup job` / `cannot resume e-packager process` | 子进程清理 Job 建立失败或挂起进程不能恢复；检查进程启动阶段日志 |
| `e-packager failed: <退出码>` / `reference unpack failed: ...` | 解包失败；保留退出码和进程输出，区分快照、依赖文件、格式与工具自身问题 |
| `读取 e-packager 元数据失败` / `解析 e-packager 元数据失败` | 解包输出元数据缺失、不可读或 JSON 解析失败；检查本次实际输出目录和 e-packager 版本 |
| `reference unpack produced no files` | 参考文件解包后未得到常规文件；检查输入 `.e/.ec` 与输出，不宣称导入成功 |
| `file_path must be a relative path inside the workspace mirror` / `absolute file_path is not allowed` / `file_path must not contain . or .. segments` | 路径不符合镜像相对路径约束；使用 `list_files` 返回的路径，不能直接传磁盘工程绝对路径 |
| `file_path resolves outside workspace mirror` / `mirror contains unsupported symbolic link` | 路径越界或发现不支持的符号链接；使用正常镜像文件，不绕过检查 |
| `file not found in workspace mirror: ...` / `mirror file not found` | 当前镜像中目标不存在；重新列出当前 generation 文件，不沿用另一工程路径 |
| `open mirror file for write failed` / `write mirror file failed` / `replace mirror file failed` / `mirror copy failed` | 镜像文件创建、写入、原子替换或复制失败；核对具体路径、系统错误和占用 |
| `stale mirror_generation; refresh or restart pagination with current generation ...` | 分页跨越镜像刷新；丢弃旧 offset，从第一页重读并保留新 `mirror_generation` |
| `invalid regex pattern` / `regex pattern is too complex` | 正则语法或复杂度超限；修正表达式或使用字面量搜索 |
| `code item name matched multiple declarations` | 同名声明不唯一；依据实际工具 schema 指定更精确的目标或 occurrence |

`workspace_changed` 时先确认最终目标实例/工程，再重新刷新。`cancelled` 表示任务已取消，不是版本不兼容。`main_only` 解包失败时当前实现已有一次 `full` 后备尝试；快照阶段失败不会因切换这两种模式而绕过。不要用磁盘旧 `.e` 替代含未保存修改的快照并称为当前代码。

## MCP 会话与调度

源码：`src/LocalMcpServer.cpp`、`src/LocalMcpProxyTransport.cpp`、`src/AIChatTooling.cpp`。调用参数见 [MCP API 参考](mcp-api.md)。

| 错误 / 状态 | 原因与处理 |
| --- | --- |
| `-32700` / `parse error` | 请求不是可解析的 JSON；核对 UTF-8、JSON 转义和传输内容 |
| `-32600` / `invalid Mcp-Session-Id header` | 会话头格式不合法；使用初始化实际返回的会话 ID |
| `-32602` / `unknown or unavailable external tool` | 工具名、参数类型、未知字段或 CAS 等参数校验失败；按当前 `tools/list` 修正。内部工具不必然对外公开 |
| HTTP 404、`-32001` / `mcp_session_not_found` | 会话不存在或已过期；重新 initialize，保存新会话头，再选择实例/准备工作区 |
| `-32002` / `selected_instance_unavailable`；`instance_not_found` / `instance_unreachable` | 目标实例不存在、不可达或选择失败；查看 `detail` / `discovery`，重新 list_instances 并显式 select_instance |
| `-32003` / `tool_response_timeout` / `tool_response_unavailable` | 请求已完整发出，但响应超时、断开或无效；执行状态未知，按下文先核查再决定是否重试 |
| `browser_origin_not_allowed` | 请求携带非空 Origin；使用原生 MCP 客户端，不绕过浏览器来源限制 |
| HTTP 503 / `server busy` | 服务连接队列满；降低并发，先核对进行中的操作后再决定重试 |
| `synchronous tool dispatch cannot wait on IDE thread` | 在 IDE 主线程同步等待需要该线程处理的工具；属于调度调用错误 |
| `main thread tool request was cancelled` | 主线程请求已取消；结合取消和任务日志确认状态 |
| `main thread tool execution failed` / `tool execution transport failed` | 调度/执行传输失败的包装错误；向前找具体工具错误，不能单凭包装文案判定工程或网络损坏 |

响应未知时保留 `execution_state=unknown`、`may_still_be_running=true`、`retry_safe=false`，先只读核对 IDE、日志、文件/编译产物与结果文件，禁止自动重发写入、删除、恢复或编译。重连成功不等于上一次请求未执行。连接固定网关 `19207`，不要把后端 `19208+` 当成稳定入口。

## 源码编辑与校验

源码：`src/AIChatToolingX86.cpp`、`src/RealPageCodeToolSupport.cpp`。精确参数和支持的写入目标见 [MCP API 参考](mcp-api.md)。

| 错误 / trace | 原因与处理 |
| --- | --- |
| `expected_base_hash is required; call read_real_file first` | 缺少真实页 CAS 基线；先 read_real_file，不能用镜像文件哈希代替 |
| `expected_base_hash mismatch` | 实际页内容与编辑基线不同；重新读取，基于新内容重做修改，不替换哈希后盲目提交旧补丁 |
| `file_path is not mapped to an editable current-project source page` | 核对当前镜像映射和 `src/` 路径；依赖镜像不可写。已有窗口 XML 使用支持窗口事务的写工具，新窗口用 `write_window`，不是普通 TXT 页 |
| `old_text not found` / `old_text not found in active edit base` | 替换前文本在真实编辑基线不存在；核对页面、换行、空白及最新代码 |
| `old_text matched multiple times` / `old_text matched multiple times in active edit base` | 匹配不唯一；增大有效上下文或按工具支持方式消除歧义 |
| `whole page write verify failed` / `verify_mismatch` | 写入后的验证不符合预期；读取最终真实页与回滚日志，不能假定完全没写入后重复执行 |
| `verify_mismatch_but_page_changed` / `verify_copy_failed_final_code=write_target` | 内部诊断片段可能位于整体成功路线；结合顶层 ok/status 和最终验证结果判断，不能看到 mismatch 就宣称失败 |
| `snapshot_id is required when restore_latest is false` | 页级恢复没有指定目标快照；按 schema 提供 ID 或选择 restore_latest |
| `snapshot not found` | 页级快照不存在；重新列出当前目标可用快照，区分页级快照与整工程 `.BAK` 备份 |

恢复需提供当前真实页的 `expected_current_hash`；不以旧快照的哈希代替。写入后若返回明确成功和结构校验结果，按 API 契约处理，不因孤立 trace 反复重写。

## 项目编译

源码：`src/ProjectBuildPipeline.cpp`；编译诊断字段见 [MCP API 参考](mcp-api.md)。

| 错误 / 状态 | 原因与处理 |
| --- | --- |
| `stage=snapshot`，附序列化错误 | 编译前工程快照失败，尚未进入该任务的独立编译进程；回到快照章节，不按源码语法错误排查 |
| `cannot resolve system temp directory` / `cannot create project build task directory` | 无法解析临时目录或创建任务目录；检查实际目录与权限 |
| `CreateProcessW failed, win32=...` / `ResumeThread failed, win32=...` | 独立编译进程创建/恢复失败；保留 Win32 数值，核对 IDE 路径、参数和进程状态 |
| `headless compile timed out` | 独立编译等待超过时限；检查任务日志与结果文件，区别于 MCP 响应超时 |
| `wait for headless compile failed, win32=...` | 进程等待 API 失败；不能仅凭该文案认定编译器报语法错误 |
| `headless result file is missing (process exit=...)` | 子进程结束后缺少预期 JSON；检查命令行结果路径、实际加载 FNE、进程退出码和启动日志 |
| `invalid headless result: ...` | 结果文件无法按约定解析；检查本次完整 JSON 和异常，不混用旧任务结果 |
| `headless compile failed` | 结果 JSON 未表示成功且缺少更具体 error 时的兜底；读取其全部诊断和编译输出 |
| `action contains unknown variables` | 前/后置动作含未定义的模板变量；按项目配置支持变量修正 |
| `action failed (exit=...)` | 前/后置脚本返回失败；检查对应阶段标准输出、标准错误与退出码，不直接判定 IDE 编译失败 |

`queued` 仅表示已排队，不是编译成功。并发任务使用各自 result 路径和 invocation/job ID，不用共享 `headless_compile_last.json` 判断所有任务。编译错误定位需看 `location_verified`、`code_source` 和诊断文本，当前页标题或光标不足以证明出错文件。黑月与普通编译的参数冲突按返回 hint 处理，不隐式切换编译器。

## 备份与工程恢复

源码：`src/AIChatSourceBackupStore.cpp`、`src/AIChatSourceBackup.cpp`、`src/BackupRestore.cpp`、`src/EideProjectRestore.cpp`。

| 错误 | 原因与处理 |
| --- | --- |
| `source_file_not_saved_to_disk` | 备份源路径不是绝对路径或磁盘上不存在常规文件；确认工程已建立磁盘文件，此项与内存快照入口兼容性无关 |
| `unsupported_source_extension` | 备份存储层仅接受 `.e` / `.ec` 后缀；不代表所有调用方都支持打开 `.ec` |
| `snapshot_path_not_representable` | 路径不能无损转换为 IDE 使用的本地编码；检查临时目录/路径字符，不把损坏编码交给 IDE |
| `empty_snapshot` | 快照回调结束后文件不存在或大小为零；追查快照阶段，不发布空文件 |
| `publish_snapshot_failed: <Win32码>` | 临时备份无法移动/复制到最终 `.BAK` 目录；检查目录权限、磁盘和具体系统码 |
| `retention warning`（日志关键词） | 新备份可能已成功发布，仅旧备份清理异常；核对实际输出路径，不能一概说备份失败 |
| `unsupported_ide_restore_abi` | 恢复 ABI 的锚点、调用链或运行时入口未通过验证；核对 x86、磁盘映像和 Hook trace，不能仅按版本名放行 |
| `unsupported_ide_requires_e5.95`（旧版本） | 历史版本的固定布局检查失败；现版本使用动态 ABI 校验，先确认实际加载的 FNE |
| `current_project_identity_unavailable` | 无法读取当前文档状态和路径字段；确认工程已打开、路径和对象有效 |
| `restore_identity_missing` | 恢复时缺少原工程身份路径；停止该恢复流程，追查身份捕获 |
| `restore_identity_mismatch` | native reload 后工程路径字段与原身份不一致；检查恢复及安全备份回滚结果，不继续叠加恢复 |
| `native_reload_exception` | 原生加载流程返回负值；当前实现可能是异常或未成功完成流程，保留现场与回滚日志 |
| `native_reload_failed` | 原生加载/收尾未返回成功；核对备份文件与加载诊断 |
| `source_path_changed_after_reload` / `reload_exception` | 外层恢复发现源路径改变或捕获异常；检查后续自动回滚是否成功，不将首次失败等同最终状态 |

工程恢复会替换 IDE 内存内容；调查请求不自动授权覆盖用户工作。已有明确恢复授权时按既定流程执行，先保留安全备份并检查回滚结果，不循环恢复未知状态。

## 原生窗口事务

`[AutoLinker][AbiResolver]` 的已知入口及 pattern_scan 诊断仅写入 `autolinker_pid<PID>.log`，不显示在 IDE 输出中。合法入口 Hook 可通过校验并保留原调用入口；循环跳转、不可执行目标、指令损坏或歧义布局会拒绝操作，不能为消除报错跳过校验。

| 错误 / 字段 | 处理 |
| --- | --- |
| `expected_base_hash mismatch` | XML 重新 `read_real_file`；`code_files` 刷新后重新 `read_file`，合并改动后再提交 |
| `window_isolated_validation_failed` | EideCore 加载预检失败；查看 `source_validation`，此阶段未修改当前 IDE |
| `source_validation.performed=false` | 未安装 EideCore，跳过独立加载预检；不代表执行过编译校验 |
| `window_event_handler_cannot_be_shared` | 每个事件绑定使用独立处理器，共用逻辑抽为普通子程序 |
| `window_event_symbol_ambiguous` | 自动解绑改名无法确定符号归属，通过 `code_files` 显式提交方法和引用修改 |
| `window_project_changed_after_reload` / `rollback_skipped` | 加载后存在外部修改，不能覆盖；先查真实状态及 `safety_snapshot`，不要盲目重发或恢复 |

失败检查 `rolled_back`，成功检查 `verified`；XML 撤回需重交旧 XML 和相关代码，`restore_file_snapshot` 仅适用于 TXT 页。超时不能证明事务尚未执行。

## AI 接口与组件安装

EideCore 普通启动缺失时只尝试下载安装一次，已安装后不检查更新；失败后在关于页手动重试。旧 Release 无 `components.eidecore` 或校验失败会拒绝安装，不回退到无哈希校验下载。文件被编译进程占用时保留旧文件，等待任务结束后重试，不要求管理员权限来绕过占用。显示旧 tag 可能是发布复用原附件，不表示安装失败。

此类错误先参考主 skill 中的接口与组件安装说明。与本地序列化/镜像错误分开归因：修改 API Key 无法修复 `memory_entry_unavailable`。

- 网络失败、超时、断流、HTTP 5xx：保留协议、状态码和服务商错误体；同站点其他模型与另一可用站点可作为对照。故障可能在网络、中转或模型线路，不能仅凭状态码确认根因，也不要因单次连通性测试通过就排除线路问题。
- 密钥、权限、余额、限流、模型或参数不支持：以响应体和当前配置为依据，不泄露 Key；检查 Base URL、协议、模型及服务商明确返回的字段，不把所有失败都当成网络问题。
- `无法下载 e-packager` / `解压后未找到 e-packager.exe` / `复制工具文件失败`：分别对应下载、包内容/目录结构和安装文件复制；按主 skill 的手动安装路径处理，不无限重试同一失败下载。

新增错误条目时记录准确原文、触发函数、直接条件、仍不确定的原因、最小验证和重试边界。未在源码/日志中证实的原因标为推测；版本样本必须附指纹和“静态检查/运行时实测”范围。
