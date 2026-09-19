静态 JSON 优化验证（2026-09-19）

本次扫描生产支持库的 `src/*.cpp`、`src/*.h`，将有效固定 JSON 定义和固定子结构改为文本解析缓存，保留运行时请求、响应、会话数据的动态构造。不替换 nlohmann/json，不启用 PCH，也不新增第三方依赖。

实现使用 `src/StaticJson.h` 中的 `AUTOLINKER_STATIC_JSON`：每个调用点拥有独立的函数局部 `static const nlohmann::json`，第一次调用时解析，后续返回只读引用。C++ 保证并发首次初始化安全；赋给普通 `json` 变量或按值返回时取得独立副本。此处没有按字符串查询的全局缓存，也没有为每份文本生成字符数组模板实例。

覆盖内容：

| 类别 | 主要文件与处理方式 |
|---|---|
| 31 个公共工具定义 | 移至 `AIChatPublicToolCatalog.inc`，由 `AIService.cpp` 按需解析一次，再返回副本 |
| Goal、图片、用户提问工具 | `AIService.cpp`、`AIChatUserInputRequest.cpp`；固定 schema 缓存，提问约束仍使用原来的唯一常量 |
| MCP 握手和工具目录 | `AIChatMcpClient.cpp`、`LocalMcpServer.cpp`；HTTP/stdio 共用固定握手数据，版本及工具名继续从原常量赋值 |
| MCP 历史默认服务器 | `AIChatMcpConfig.cpp`；5 份历史定义保留内容和顺序，用只读 JSON 数组替代 vector 中的嵌套初始化 |
| 浅色/深色主题 | `AIChatThemeManager.cpp`；各 40 项颜色缓存，调用方得到副本 |
| 固定协议片段、错误响应和自检样例 | 会话适配、AIProviderCache、技能、工作区、GameAnalytics、MCP、黑月及内部测试代码等 |
| 设置页清单及链接 | `AutoLinkerSettingsDialog.cpp`；两份固定数组缓存 |

共修改 20 个既有源文件（含一个头文件），新增一个通用头文件及一个独立数据文件。生产源文件中的固定 JSON 文本检查覆盖 **181 处**（其中公共目录按一个整体计数，内含 31 个工具）。简单空对象/空数组、标量赋值和非 JSON 的 C++ 表不需要文本解析缓存。含运行时变量的初始化仍按原逻辑构造；设置页故意生成非法 UTF-8 的测试样例也保留原测试目的。

JSON 文本中的非 ASCII 字符使用 Unicode 转义，避免 `/utf-8` 与默认执行字符集混用。两个图片为空的本地编码提示不能直接改成 UTF-8 JSON 解析，改用局部静态值加字段赋值，保留原窄字符串字节。

**构建实测**

与此前调查使用相同的 VS 2026 / MSVC 14.51.36231、HostX64/x86、`fne_release|Win32`、`/O2 /MT /MP4 /FS`，顺序执行完整 Rebuild，不清空系统文件缓存：

| 配置 | 墙钟耗时 |
|---|---:|
| 优化前，三次原始构建 | 127.331 / 119.958 / 132.305 秒 |
| 静态定义优化后，第 1 次 | 94.067 秒 |
| 静态定义优化后，第 2 次 | 94.595 秒 |
| 最终代码完整 Rebuild | 91.234 秒 |

三轮优化后为 91.234～94.595 秒，三轮中位数由 127.331 秒降至 94.067 秒，约快 **26.1%**。前两轮之后仅补齐设置页的两个固定数组及说明注释；以最后一行作为最终代码的构建结果。所有发布构建均成功，0 警告、0 错误。没有将插件复制到易语言安装目录。

**内容与功能验证**

- 用相同 MSVC x86 编译器编译旧初始化表达式和新缓存表达式，**210 项序列化输出逐字节一致**。原始对照、日志及结果保存在 `.temp/json-compile/static-{old,new}.*` 和 `static-equivalence.json`。
- 对 5 组完整定义进一步比较：公共工具目录、历史 MCP 默认配置、request_user_input 定义、MCP 实例工具、MCP 握手参数。全部相同，并验证修改公共目录返回副本不污染缓存。
- `python scripts/test_static_json.py` 读取实际源文件中的固定文本，用 MSVC x86 编译并运行验证器，181 处解析结果均与预期一致；验证 16 线程首次初始化只执行一次、缓存引用稳定、副本修改隔离，以及中文、emoji、内嵌 NUL、空对象/数组、布尔值和 uint64 最大值。
- AutoLinkerTest 的 `gameanalytics-self-test`、`ai-endpoint-config-self-test`、`blackmoon-self-test`、`plan-mode-self-test`、`goal-mode-self-test`、`image-tool-batch-self-test` 均返回成功且 `ok=true`。
- 完整 `mcp-self-test` 的历史 `0xC0000409` 已定位并修复：`AIService::BuildAgentOptimizationSelfTestJson()` 将默认执行字符集下的“技能”字节直接写入 UTF-8 JSON，严格 `dump()` 抛出未捕获异常。Windows 调试事件及符号化堆栈确认落点为分页样例序列化。现在先用 `LocalToUtf8` 构造样例单元，并依据实际 UTF-8 字节数填充内容，不再硬编码长度。修复后总自检退出码 0、`ok=true`，50 个顶层检查全部通过，包括 HTTP/stdio 往返、会话/重试/取消、工作区分页、技能及配置不变量。

固定正整数从 C++ 初始化列表改成 JSON 解析后，nlohmann 的内部数字标签可能从 signed 变为 unsigned；JSON 文本、数值、`is_number_integer()` 及现有转换语义一致。上述内容对照和协议自检验证实际使用，不承诺所有内部 `value_t` 标签相同。

新增或修改的源代码保持 UTF-8 BOM、CRLF，通过 `git diff --check`。最终普通配置产物为 `bin/fne_release/AutoLinker.fne`，PE 架构为 x86。

复现：

```powershell
python scripts/test_static_json.py
& 'C:\Program Files\Microsoft Visual Studio\18\Professional\MSBuild\Current\Bin\amd64\MSBuild.exe' AutoLinker.vcxproj /t:Rebuild '/p:Configuration=fne_release;Platform=Win32' /m
& 'C:\Program Files\Microsoft Visual Studio\18\Professional\MSBuild\Current\Bin\amd64\MSBuild.exe' AutoLinkerTest/AutoLinkerTest.vcxproj /t:Build '/p:Configuration=fne_release;Platform=Win32' /m
& .\bin\fne_release\AutoLinkerTest.exe image-tool-batch-self-test
& .\bin\fne_release\AutoLinkerTest.exe ai-endpoint-config-self-test
```

此前 `json-compile-investigation.md` 记录的是提交 `99e9663` 的优化前实验；其初始化列表微基准应针对该历史源码复现，不能把本次优化后的缓存函数当成原来的初始化列表。

后续排错验证（2026-09-19）：

- 新旧插件均可复现的异常已确认来自自检样例编码，不是缓存生命周期、内存破坏或 JSON 静态初始化。
- 失败堆栈：`serializer::dump_escaped` → `json::dump` → `AIService::BuildAgentOptimizationSelfTestJson` → `AutoLinkerTest_RunAIChatMcpSelfTest`。调试器捕获未处理 C++ 异常 `0xE06D7363`，非调试运行最终表现为 `0xC0000409`。
- `read_skill_resource_context_preserves_full_pages_and_cursors` 检查已执行并通过，保留原来的 11,989 / 262,241 字节数据、4,098 / 262,144 字节分页及三次重复读取覆盖。
- 再次运行普通 smoke、GameAnalytics、配置迁移、黑月、Plan、Goal、图片批次，均成功。181 处静态 JSON 的并发、Unicode 和副本隔离测试再次通过。
- 原始证据：`.temp/json-compile/crash-stack.txt`、`mcp-fixed.json`、`final-offline-results.json`、`fixed-final-build.log`。这些测试覆盖本次改动及可离线运行的回归项目；没有将离线通过表述为真实 IDE 和外部付费模型的全部端到端场景已经验证。

排错后的最终 AutoLinkerTest 及依赖支持库完整 Rebuild：`fne_release|Win32`，101.129 秒，0 警告、0 错误，产物确认为 x86。完整重建后另一次 MCP 总自检结果记录于 `.temp/json-compile/mcp-final-rebuilt.json`。
