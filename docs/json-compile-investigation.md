JSON 编译耗时调查（2026-09-18）

本报告记录优化前提交 `99e9663` 的实验。后续已实施静态定义优化，最终结果见 [静态 JSON 优化验证](static-json-optimization.md)；下面的初始化列表基准需使用该历史源码复现。

当前使用 `thirdparty/json.hpp`，nlohmann/json 3.11.3，文件大小 944,740 字节。实际 MSVC 依赖记录显示，100 个编译单元中有 52 个直接或间接依赖它。`fne_release|Win32` 使用 `/O2 /MT /std:c++20 /MP4 /FS`，原配置没有 JSON 预编译头。

结论：JSON 的重复头文件解析及大型嵌套初始化列表是显著的编译瓶颈。yyjson 能降低调用方编译成本，但目前没有完成全项目 API 迁移，不能把单文件提速比例当作整库换库收益。最明确、改动范围最小的优化点是 `src/AIService.cpp:3001` 的 `BuildPublicToolCatalog()`：保留现有库，先把静态工具定义改成数据。

测试基线提交：`99e9663`，开始时工作区干净。环境：Windows，Intel Genuine Intel(R) 0000，14 核/20 逻辑处理器；VS 2026 Professional，MSVC 14.51.36231，HostX64/x86，MSBuild 18.10.1。所有计时任务顺序执行，没有并行跑其他本次实验。未清空操作系统文件缓存；全量构建使用 Rebuild，不能称为冷缓存构建。

**单文件对照**

下面是三轮中位数，单位秒。使用同一 MSVC x86 编译器和 `/O2 /MT /EHsc /std:c++20 /utf-8 /DNDEBUG /Gy /Oi`，测量编译器进程墙钟时间，不含链接；微基准未启用项目的 `/Zi`，因此不能直接和整库日志中各文件的耗时相加。轮次之间反转测试顺序。

| 测试 | 当前 nlohmann | yyjson 0.12.0 | yyjson 降幅 |
|---|---:|---:|---:|
| 仅包含库头文件 | 1.445 | 0.087 | 94.0% |
| 解析输入、构造包含 payload 的对象、输出 JSON | 2.030 | 0.534 | 73.7% |
| 实际 CompileConcurrencyManager.cpp | 2.668 | 1.728 | 35.2% |

空编译单元为 0.042 秒。真实文件实验只将 `BuildSelfTestJson()` 的 JSON 构造替换成 yyjson，锁和线程实现保留原样。yyjson 的 `yyjson.c` 每次全新编译中位数为 **9.534 秒**，不是零成本；增量编译未修改库实现时才可以复用其对象文件。以这个真实文件为例，仅迁移一个文件节省不到一秒，不足以抵消全量构建中新增加的库实现编译成本。

保留 nlohmann 并使用 JSON PCH 时，同一解析/输出样例为 0.676 秒，真实并发管理文件为 1.333 秒；PCH 创建另需 1.550 秒。这表明小文件主要受公共头文件处理成本影响。

**工具目录热点**

完整基线的 `/Bt+` 日志中，`AIService.cpp` 前端 34.985 秒、后端 12.399 秒，总计 47.385 秒，是最慢文件。其次为 `AIChatFeature.cpp` 9.584 秒、`AIChatToolingX86.cpp` 9.378 秒。PCH 实验中 `AIService.cpp` 仍需 42.611 秒。

将 `BuildPublicToolCatalog()` 单独提取为可执行程序，比较以下两个等价实现，三轮均验证解析后的输出完全相同，包含全部 31 个工具：

| 实现 | 编译及链接中位数 |
|---|---:|
| 原来的约 480 行嵌套 JSON 初始化列表 | 29.828 秒 |
| 同一 nlohmann 库，解析等价 JSON 文本 | 2.016 秒 |

降幅 **93.2%**。首轮原实现前端 23.708 秒、后端 5.974 秒；文本实现分别为 1.431 秒和 0.452 秒。这个实验隔离了工具定义的构造方式，是比简单统计 include 数量更直接的证据。实验统一使用 UTF-8，省略项目的本地编码转换；整库临时副本则使用 ASCII JSON 转义，解析后还原 UTF-8，避免改变文件其他部分的执行字符集。

整库实验中的 `AIService_catalog.cpp` 前端为 4.537 秒、后端为 7.467 秒，总计 12.004 秒。只改变该函数就使这个真实编译单元由 47.385 秒降至 12.004 秒，进一步验证热点确实在工具目录构造代码。

**全量构建对照**

| 方案 | 全量墙钟时间 | 结果 |
|---|---:|---|
| 原配置，首次普通 Rebuild | 127.331 秒 | 0 警告、0 错误 |
| 原配置，加 /Bt+ 记录阶段耗时 | 119.958 秒 | 0 警告、0 错误 |
| 原配置，最终普通 Rebuild 恢复产物 | 132.305 秒 | 0 警告、0 错误 |
| 保留 nlohmann，为 52 个文件建立 JSON PCH，第 1 次 | 114.467 秒 | 0 警告、0 错误 |
| 同一 PCH 方案，第 2 次 | 119.454 秒 | 0 警告、0 错误 |
| 保留 nlohmann，仅替换临时副本中的工具目录构造 | 103.854 秒 | 0 警告、0 错误 |

PCH 使用三个变体，分别匹配原项目默认字符集、`/utf-8` 和 `/execution-charset:utf-8`；创建成本已包含在全量耗时中。两次复测说明其整库收益有限且存在明显波动，第二次几乎等于 119.958 秒基线，不能把首次约 10% 的改善作为稳定收益承诺。

仅调整工具目录构造，整库相对三次基线快 **13.4%～21.5%**（相对基线中位数 127.331 秒快 18.4%），同时也有独立三轮测试和真实文件前后端计时支持。全量目录替换只测了一次，应将该比例视为本机观察值；未测运行时解析性能，也未据此直接修改生产实现。没有证据支持“JSON 库本身占用了全部或过半编译时间”。

完整构建日志同时包含编译器前后端计时。多个 `/MP4` 编译进程的阶段耗时不能直接当作整库墙钟时间。PCH 会改变编译批次组织；工具目录实验使用独立临时源文件及 include 路径，也会改变批次安排。因此这些数字是实测构建结果，不是对 JSON 占比的精确分摊。

**可靠性和迁移代价**

候选选用 [yyjson 0.12.0](https://github.com/ibireme/yyjson/tree/0.12.0)，其[官方说明](https://github.com/ibireme/yyjson/blob/0.12.0/README.md)介绍了单 `.h` 加单 `.c`、ANSI C、UTF-8 校验和 64 位整数支持。本次直接下载固定版本并校验 SHA256，没有引入 vcpkg，也没有把候选库加入生产项目。实际通过 MSVC x86 编译、链接及运行。

功能对照覆盖 7 组合法输入，包括中文、emoji、内嵌 NUL、转义字符、数组、对象、空值、布尔值、int64 最小值、uint64 最大值和浮点数；另验证 4 组非法输入被拒绝。两种并发管理实现的自检报告一致且 `ok=true`，31 个工具定义输出也全部一致。

这些测试不构成全应用迁移验证。现有代码依赖 `value/get/items`、异常、`error_handler_t::replace` 和直接在公共类型中存放 `nlohmann::json`。yyjson 的文档生命周期、可变/不可变文档、重复键、键顺序及非法 UTF-8 策略需要逐处确认；尤其 `AIChatToolPolicy.cpp` 使用 `parse().dump()` 构造签名，不能忽略当前对象键排序语义。没有证据表明当前库“不可靠”，问题主要是本项目的编译成本和使用方式。

建议按以下顺序推进：

1. 将静态工具 schema 作为 JSON 数据维护，嵌入资源或由构建步骤生成文本，初始化时解析并缓存；业务调用获取副本。避免手工保留两份 schema，也避免每次调用重新解析。
2. 收紧公共头文件依赖，对适合前置声明的接口使用匹配版本的 `json_fwd.hpp`，评估统一 PCH；已有按值 JSON 成员的类型不能仅靠前置声明解决。
3. 若还需要降低编译耗时，在独立、以字符串为边界的模块逐步评估 yyjson。不要为兼容全部 nlohmann API 自建庞大适配层，也不要只凭微基准就整库替换。

**复现与交付**

新增三个脚本，所有候选源码、临时项目导入文件、可执行程序和原始结果都在 `.temp/json-compile/`，生产 `.vcxproj` 和业务源文件没有修改：

```powershell
# 首先按项目要求执行一次普通 fne_release/Win32 构建，生成依赖记录。
python scripts/benchmark_json_compile.py --rounds 3
python scripts/benchmark_json_catalog.py
python scripts/benchmark_json_full_build.py --mode baseline --label baseline-recheck
python scripts/benchmark_json_full_build.py --mode pch --label pch-recheck
python scripts/benchmark_json_full_build.py --mode baseline --catalog-text --label catalog-recheck

# 全量实验复用常规 bin 目录；最后恢复普通生产配置产物。
& 'C:\Program Files\Microsoft Visual Studio\18\Professional\MSBuild\Current\Bin\amd64\MSBuild.exe' AutoLinker.vcxproj /t:Rebuild '/p:Configuration=fne_release;Platform=Win32' /m
```

脚本会通过 vswhere 定位 VS；候选下载使用标准库并校验固定 SHA256。请顺序运行计时命令。原始结果包括 `micro-results.json`、`catalog-results.json`、`*-time.json` 和相应 `.log`。实验没有向易语言安装目录复制插件，也没有运行 IDE 场景测试。

最终已完成普通 `fne_release|Win32` Rebuild，耗时 132.305 秒，0 警告、0 错误。`bin/fne_release/AutoLinker.fne` 已恢复为原始生产源码及配置生成的 x86 产物；最终构建命令没有 PCH、临时工具目录替换或 `/Bt+` 参数。三个 Python 脚本也通过语法检查；新增文件均为 UTF-8 BOM、CRLF。
