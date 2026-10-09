# 备份与恢复功能验证（2026-10-09）

设置新增“备份与恢复”页，列出当前工程自动备份的时间、文件名/摘要和大小，可刷新并恢复选中项。恢复操作仅支持已验证内部地址的 e5.95：签名不匹配时在页面提示“无法恢复，此功能需要 e5.95 版本。”并禁用恢复。不会仅凭 exe 名称识别版本。

恢复事务先发布“恢复前安全备份”，包含未保存的修改，即使自动备份开关关闭也必须成功写出。此时不清理旧备份；所选快照先暂存，成功后才执行保留数量清理，优先保留安全备份。失败时尝试从安全快照回滚，所有安全备份保留供手工恢复。

内部流程采用关闭旧视图、文件反序列化、post-load 重建、界面刷新；在 post-load 之前恢复原工程路径及相关身份字段，并将内容标记为未保存。原始 .e 磁盘文件不改写。暂存加载文件放在原工程目录，保持相对模块引用的解析目录，结束后清除。恢复前后使工作区 epoch 失效并清理页代码/序列化缓存；AI 或 MCP 正在执行时拒绝恢复，恢复期间暂停新工具和 AI 请求。

实现分为 EideProjectRestore（版本和 ABI）、BackupRestore（快照与回滚事务）、AutoLinkerSettingsDialog / backup_restore_settings.html（WebView2 页面）。现有自动备份设置继续位于 AI 其他设置。

## 实测

使用独立复制的 IDE 安装目录和 baseline.e，PID 68148，未替换用户正在运行的插件。

- 通过实际设置页面选择快照并点击恢复：真实代码页由 UnsavedBeforeRestore 更新为 RestoreFixtureA。
- 恢复前安全快照包含 UnsavedBeforeRestore；原始磁盘文件与 eproj/e-window-exe-full.e 逐字节一致。
- 当前工程路径仍为 baseline.e；内存修改标记为 1，恢复结果待用户手动保存。
- 加载破坏校验的快照：页面提示恢复失败且已回滚；随后真实代码 hash 与恢复前一致，原始文件未变。
- 在空闲专用测试进程暂时更改一个入口签名字节，刷新页面：恢复按钮禁用，显示 e5.95 限制。测试 finally 中恢复字节和内存保护。此项验证签名失配处理，未在其他版本 IDE 执行加载。
- 独立备份文件测试通过：20 个 Unicode 码点摘要、非法字符、重复命名、保留数量、写失败不清理、数量为 1 时所选备份在安全快照生成后仍存在、清理后优先保留安全快照。

证据见 [正向恢复](test-results/backup-restore-20261009/positive.json)、[失败回滚](test-results/backup-restore-20261009/rollback.json)、[版本不匹配](test-results/backup-restore-20261009/unsupported.json)、[路径及修改状态](test-results/backup-restore-20261009/state.json)。完整临时记录位于 .tmp/backup-restore-live，不属于发布文件。

未覆盖密码工程、超长文件路径、恢复后程序编译运行、所有异常退出场景。当前仅支持 .e 工程；.ec 备份可显示但不能恢复。源码采用 UTF-8 BOM 与 CRLF。



## WebView2 改版审查与测试

设置页现改为共享 WebView2 宿主和嵌入式 HTML。页面展示时绑定工程路径、备份列表和 revision；点击恢复不会重新按最新目录序号选文件。旧 revision 请求被拒绝，工程切换由恢复事务再次校验。恢复期间阻止重入并禁用设置窗口，使用作用域守卫确保异常后恢复可用状态。文件名使用 textContent 渲染。目录读取失败显示错误，非 .e 工程和签名不匹配时禁用恢复。

本节描述首次 WebView2 改版的验证；最终调用链见下节。

本轮通过真实 Edge 无界面浏览器验证 7 项：Unicode 文件名、恢复 index/revision 消息、忙碌禁用、结果返回后恢复控件、IDE 不支持、非 .e 工程、空列表。脚本：scripts/test_backup_restore_webview.py。结果和截图：test-results/backup-restore-webview/。复跑现有备份文件测试通过。

上述浏览器测试使用模拟宿主消息；最初“实测”一节属于原生页面版本。新版 WebView2 在真实 IDE 中的验证见下节。


## 最终修复与真实 WebView2 回归

用户报告完整打开入口替换后恢复失败、路径改变。此前“调用约定匹配”不足以证明该入口满足内容恢复语义。真实 IDE 复现了设计页快照失败；IDA 和源码确认，完整打开入口已在 post-load 更新标题后，旧实现才改回一个 CString。GetSourceFilePath 从标题取路径，内部字段、标题和事务校验因此不一致。

最终实现使用二进制分支的分阶段原生调用，保留 IDE 对象和容器 ABI：

- 捕获原工程三个身份 CString（文档 +0x3FC/+0x400/+0x404），恢复和回滚共用同一份，不从失败现场重新取身份。
- 0x46B010 关闭旧视图；0x46B4D0 按完整二进制格式加载，内部转到 0x457490。参数依照 0x46B310 的真实汇编调用点验证。
- 使用 IDE 的 0x5BC5AA 恢复三个 CString，再执行 0x46D0C0 收尾；此时内部路径通知和标题更新使用原工程身份。
- 设置未保存状态，执行 0x46D3C0 刷新。不会把备份作为新的当前项目，也不覆盖原磁盘文件。不是文本替换或外部重打包。
- CPtrArray 使用 IDE 的 0x5BB2C6 / 0x5BB2F9 构造析构，SEH 保护；签名不匹配拒绝调用。
- 完整快照在 e5.95 可通过已验证全局文档 +0xC0 取序列化对象，不依赖活动代码页。
- 恢复请求通过主线程 PostMessage 执行，先退出 WebView2 回调再运行可能进入模态循环的原生加载器。
- 恢复期间暂缓发布中间空工程状态，事务结束后刷新实际状态；避免错误地切换 AI 会话。完成校验仍读取真实 IDE 标题，并非用缓存路径掩盖变化。

最终构建在独立复制 IDE PID 50464 中，通过真实 WebView2 DOM 点击恢复按钮：

1. 从窗口设计页直接恢复有效快照，安全备份成功，代码成为 RestoreFixtureA。
2. 写入未保存的 UnsavedBinaryRestoreTest，再恢复目标快照；随后恢复自动生成的安全备份，内容 hash 恢复为未保存状态的 hash。
3. 损坏快照触发 IDE 数据段校验错误；测试脚本确认了仅该测试 PID 的原生错误弹窗（按钮 ID 为 2），随后事务自动回滚，内容 hash 与恢复前相同。
4. 三个核心用例均断言：三个身份字段不变、标题含原 baseline.e、未保存标记为 1、原始磁盘文件逐字节不变。有效恢复的窗口事件索引与选中快照一致。
5. 旧 revision 恢复请求被拒绝，无新备份、无内容改变。
6. 在空闲测试进程中临时改变一个入口签名字节，页面禁用所有恢复按钮并显示 e5.95 限制；finally 恢复字节及页面支持状态。
7. 本次日志只有启动时的一次 source_changed；成功恢复和失败回滚均未触发中间 AI 会话换绑。

证据：[核心恢复回归](test-results/binary-content-restore-20261009/results.json)、[保护逻辑](test-results/binary-content-restore-20261009/guards.json)、同目录页面截图。临时驱动位于 .tmp/backup-restore-live/{live_regression,full-regression,guard-regression}.py，使用已保存的二进制测试快照；测试安装目录与用户 IDE 分离。

最终 fne_release|Win32 构建通过，产物 bin/fne_release/AutoLinker.fne。浏览器 7 项页面测试、备份文件测试同时通过。没有替换用户正在运行的插件。未验证密码工程、所有类型资源、超长路径及恢复后编译运行；损坏快照可能弹出 IDE 原生错误提示，需要确认后继续回滚。
