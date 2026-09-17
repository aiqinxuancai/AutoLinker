# 黑月编译参数与验证

`compiler` 默认为 `eide`；非空 `blackmoon_mode` 不会隐式切换编译器。
普通编译携带该参数会返回 `blackmoon_mode_requires_blackmoon`，`hint` 提示移除模式，
或者显式选择 `compiler=blackmoon` 并移除 `static_compile`。

黑月调用示例：

```json
{"compiler":"blackmoon","blackmoon_mode":"vcxx","target":"auto","output_path":"D:\\build\\demo.exe"}
```

黑月模式默认 `vcxx`。参数类型错误、未知编译器、未知模式、静态参数（包括 false）和
目标类型覆盖均明确报错。黑月仅支持当前 EXE/DLL 工程；模块工程使用 `eide`。
兼容旧模式别名，但返回的 `blackmoon_mode` 始终为 `asm/cxx/vcxx`。
能力查询根据真实配置档和 `Normal.BMType` 返回可用模式，不再固定报告三种模式。

`BlackMoonCompileCoordinator` 以规范化配置路径生成 Global 命名互斥体，读取、恢复、
切换和产物确认处于同一会话中。配置和恢复记录采用同目录临时文件、刷盘、原子替换。
进程异常退出后，下次会话仅在当前配置等于记录中的原值或临时值时自动恢复；
发现外部修改时保留配置与恢复记录并返回冲突，不盲目覆盖。
恢复记录位于 `Config.json.autolinker-recovery.json`，记录两份完整原始字节数组。
此互斥只约束 AutoLinker 调用；黑月设置窗口等外部写入者不遵守该锁。

输出路径通过 `GetSaveFileNameA` 注入，作用域结束自动取消请求，有效期 120 秒。
实际编译仍调用 `BMCompile(nullptr, nullptr)`。黑月返回的 `static_compile` 为 null，
并报告 `backend_invoked`、`output_path_injected`、`artifact_verified` 和 `config_restored`。
`backend_invoked` 不代表编译成功，恢复失败也不能返回成功。

验证命令（先构建并按项目说明安装新支持库）：

```powershell
.\bin\fne_release\AutoLinkerTest.exe blackmoon-self-test
.\scripts\test_blackmoon_compile.ps1
```

独立自检使用唯一临时目录，验证参数纠错、三种模式、配置逐字节恢复、锁竞争、
恢复记录及外部修改保护。IDE 回归使用工程副本，覆盖三种黑月模式、普通静态编译、
五类错误传参和两个不同工程并发编译；逐项核对配置哈希和产物更新。
原始日志及结果保存在 `.tmp/blackmoon-regression-<id>/`。

2026-09-17 验证：`fne_release | Win32` 构建通过；17 项独立检查通过。
真实 e5.95.exe / BlackMoon 5.1 的 12 个回归场景通过：三种模式、默认 VC++ 模式和普通静态编译成功，
五类错误传参均返回预期错误及 hint，两个不同工程并发编译成功。
黑月输出窗口确认实际配置与请求模式一致，所有测试后配置 SHA256 保持原值。
恢复测试使用独立临时文件模拟崩溃后的持久记录，没有终止用户的 IDE 进程。
本次最终结果：`.tmp/blackmoon-regression-e344cab28c6c41fcb2d56bfac9a1dd36/summary.json`。

## 实际 MCP HTTP 验证

2026-09-17 补充使用 `http://127.0.0.1:19207/mcp` 完成完整协议流程：
`initialize`、`notifications/initialized`、`list_instances`、`select_instance`、
`tools/list`、`get_current_eide_info` 及 `tools/call`。本轮未使用无头启动器或内部调用接口。

13 个场景全部通过：普通静态编译、asm/cxx/vcxx、默认黑月模式与中文输出路径，
普通编译误传黑月模式、未指定编译器却传黑月模式、黑月携带静态参数 true/false、
目标类型不符、未知模式、未知编译器及错误参数类型。

参数冲突返回 MCP 工具错误 `isError=true`；`structuredContent` 和 `content[].text`
都包含明确 `error` 和纠正 `hint`，仅消费文本的客户端也能获得正确指示。
未知枚举值及错误类型被 MCP Schema 校验提前拒绝，返回 JSON-RPC `-32602` 并标识
错误字段，因而不会进入编译处理函数；这与内部工具调用的错误封装不同。
黑月模式通过 IDE 实际输出确认，配置 SHA256 在每项测试后均保持原值且无恢复记录残留。

复用脚本（先在 IDE 中打开临时控制台工程副本）：

```powershell
python scripts/test_blackmoon_mcp.py --project <工程副本.e> --output <独立结果目录>
```

原始请求、HTTP 状态及响应头、完整响应和汇总保存在：
`.tmp/blackmoon-mcp-45ac1f44b0f9492495c3aaa6e00a8d90/final-results/`。
该目录 `012-response.json` 是普通编译误传模式的实际 MCP 响应。
