# AutoLinker MCP API 参考

本文档描述 AutoLinker 固定网关 `http://127.0.0.1:19207/mcp` 当前实际公开的
Streamable HTTP 接口。方法、参数和约束以 `src/LocalMcpServer.cpp`、
`src/AIService.cpp`、`src/AIChatToolRegistry.cpp` 及各工具实现为准。

## 目录

- [连接、会话与通用约定](#连接会话与通用约定)
- [协议级方法与 HTTP 操作](#协议级方法与-http-操作)
- [工具调用的统一响应](#工具调用的统一响应)
- [多实例路由工具](#多实例路由工具)
- [工作区与读取工具](#工作区与读取工具)
- [源码编辑工具](#源码编辑工具)
- [资源与模块工具](#资源与模块工具)
- [IDE 状态与编译工具](#ide-状态与编译工具)
- [网络工具](#网络工具)
- [推荐调用流程](#推荐调用流程)
- [常见错误](#常见错误)

## 连接、会话与通用约定

- 固定入口是 `http://127.0.0.1:19207/mcp`。`19208` 及以上是 IDE 实例后端，
  外部客户端不要直接连接。
- 请求和工具参数均为 UTF-8 JSON。向工具传递中文页名、路径或易语言源码时保持 UTF-8。
- `initialize` 响应会通过 `Mcp-Session-Id` HTTP 头返回会话 ID。后续 POST 和
  DELETE 请求应原样带回该头。
- 会话保存工作区刷新状态和所选 IDE 实例；选择结果不会影响其他 MCP 会话。会话空闲
  30 分钟后可过期，服务器最多保留 256 个外部会话。
- 服务器仅绑定 `127.0.0.1`，拒绝带非空 `Origin` 的浏览器请求，不提供 CORS 或
  Bearer Token。不要用反向代理把它暴露到局域网或互联网。
- `tools/list` 当前返回 30 个工具：28 个 AutoLinker 原生公开工具和 2 个网关路由工具。
  所有工具的根参数都是对象；工具 Schema 均禁止未声明的顶层参数。

### 工程状态和刷新门禁

以下调用不要求 IDE 已打开 `.e` 工程：`list_instances`、`select_instance`、
`get_current_eide_info` 和三个网络工具。

`refresh_workspace_mirror`、`get_current_page_info`、`add_new_file`、
`compile_with_output_path`、三个资源工具及三个模块管理工具要求当前实例已打开 `.e` 工程。

下列工具同时要求已打开工程，并要求当前 MCP 会话已经成功调用
`refresh_workspace_mirror`：

- `e_packager`、`list_files`、`search_code`、`read_file`、`read_files`、`read_code_item`
- `read_real_file`、`edit_file`、`multi_edit_file`、`write_file`、`diff_file`
- `restore_file_snapshot`、`write_window`、`delete_file`

切换 IDE 实例后，应对新实例重新调用 `refresh_workspace_mirror`，再读取或编辑工程。

### 镜像路径和可写范围

- `file_path`、`path` 和 `glob` 都相对于 e-packager 工作区镜像根目录。
- `src/` 是当前工程；其中 `src/*.txt` 可映射回 IDE 真实页面。
- `src/*.xml` 是原生窗口 UI。在通过 5.6 及以后 x86 ABI 校验的 IDE 中，可用 `read_real_file` 读取实时 XML，再用 `write_file`、`edit_file`、`multi_edit_file` 修改窗口/控件属性、层级及事件。
- 新建窗口或同时提交事件处理代码使用 `write_window`；新窗口的 `expected_base_hash` 传空串，已有窗口传实时 XML 的 `code_hash`。`code_files` 中已有程序集传刷新后 `read_file` 的哈希，新程序集传空串。
- 事件使用 `<按钮.事件 名称="被单击" 处理器="程序集::子程序"/>`；参数和返回类型以 XML 的 `.事件定义` 为准。IDE 会把处理器规范化为 `_控件名_事件名`；解绑时保留的旧处理器改成普通子程序名，结果通过 `detached_event_handlers` 告知。
- 窗口事务保留工程路径及未保存内容，回包并通过内部恢复接口加载，不保存原 `.e`；安装 EideCore 时在独立进程预检。失败时仅在事务仍拥有当前工程状态时回滚；加载后有外部修改则保留安全快照并返回 `rollback_skipped`，避免覆盖新改动。视图重建可能重置撤销、光标及断点状态。
- 涉及跨程序集引用的解绑须通过 `code_files` 显式同步引用；XML 撤回使用先前 XML/代码重新提交，`restore_file_snapshot` 用于 TXT 页。完整调用和验证见 [原生窗口编辑说明](../../../docs/native-window-editing.md)。
- `ecom/`、`elib/`、`header/` 是依赖和公开接口参考，只能读取、列出和搜索。
- 不要直接修改 `%TEMP%/AutoLinker/workspace-mirror/` 中的文件；写入必须使用 MCP 编辑工具。

### 分页和 `mirror_generation`

- 初次读取或 `offset=0`、`byte_offset=0` 时省略 `mirror_generation`，不要猜测 `1`。
- 延续上一页时，把上一页返回的 `mirror_generation` 原样传回，并使用其
  `next_offset` 或 `next_source_byte_offset`。
- 镜像在分页期间发生刷新时，旧 generation 会被拒绝，以免拼接不同版本的数据。
- `read_file` / `read_files` 每个字节窗口最多读取 1 MiB；大文件使用
  `next_source_byte_offset` 继续，并把行 `offset` 重置为 `0`。

### 真实页哈希和 CAS

- 编辑已有源码前，对同一个 `file_path` 调用 `read_real_file`，使用其 `code_hash` 作为
  `expected_base_hash`。
- 外部 MCP 调用 `edit_file`、`multi_edit_file`、`write_file`、`diff_file` 时，
  `expected_base_hash` 实际为必填，即使 JSON Schema 为兼容内置 AI 没把它列入 `required`。
- `restore_file_snapshot` 同理必须传 `expected_current_hash`。
- 普通 TXT 写入和已有窗口 XML 使用 `read_real_file.code_hash`；`write_window.code_files` 中已有程序集专门要求刷新后 `read_file` 的镜像哈希，新文件传空串。不要混用这两种基线。
- `read_real_file.content` 带有 `行号<TAB>` 显示前缀。构造 `old_text` 或 `full_code`
  时要移除这些前缀；它们不是易语言源码的一部分。

## 协议级方法与 HTTP 操作

除 GET、DELETE 和 OPTIONS 外，协议方法都使用 `POST /mcp`，请求体为 JSON-RPC 2.0。

### `initialize`

建立或重置 MCP 会话，协商协议版本，并返回服务器能力、工程状态和使用指令。

| 参数 | 类型 | 必填 | 约束与说明 |
| --- | --- | --- | --- |
| `protocolVersion` | string | 协议要求 | 支持 `2025-11-25`、`2025-03-26`、`2024-11-05`；省略或传其他值都回退为 `2025-11-25` |
| `capabilities` | object | 协议要求 | 标准 MCP 客户端能力；当前服务器不读取其子字段并兼容省略 |
| `clientInfo` | object | 协议要求 | 标准客户端名称和版本；当前服务器不读取其子字段并兼容省略 |

响应 `result` 包含 `protocolVersion`、`capabilities.tools.listChanged=false`、
`serverInfo`、`instructions`，以及 AutoLinker 扩展字段 `autolinkerState`。HTTP 响应头中的
`Mcp-Session-Id` 必须用于后续请求。每次 initialize 都会清除本会话旧的刷新状态和实例选择。

```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "initialize",
  "params": {
    "protocolVersion": "2025-11-25",
    "capabilities": {},
    "clientInfo": { "name": "example-client", "version": "1.0" }
  }
}
```

### `notifications/initialized`

通知服务端客户端已完成初始化。它是无 `id` 的 JSON-RPC notification，`params` 可传空对象。
服务器返回 HTTP 202 且无响应体。

### `ping`

检查会话和 JSON-RPC 通路。`params` 传空对象，成功结果为 `{}`。带会话头但会话未知或
已过期时，服务器返回 HTTP 404 和 JSON-RPC 错误 `-32001`。

### `tools/list`

列出当前全部 30 个工具和各自 `inputSchema`。`params` 传空对象；当前没有 cursor 分页，
响应结构为 `result.tools[]`。

### `tools/call`

调用一个 `tools/list` 返回的工具。

| 参数 | 类型 | 必填 | 约束与说明 |
| --- | --- | --- | --- |
| `name` | string | 是 | 工具名，必须存在于当前 `tools/list` |
| `arguments` | object | 否 | 工具参数；省略或传 `null` 等价于 `{}` |

```json
{
  "jsonrpc": "2.0",
  "id": 3,
  "method": "tools/call",
  "params": {
    "name": "get_current_eide_info",
    "arguments": {}
  }
}
```

参数结构错误返回 JSON-RPC `-32602`。工具被正常调度但业务执行失败时，仍是 JSON-RPC
成功响应，需检查 `result.isError` 和 `result.structuredContent.ok`。

### `DELETE /mcp`

结束会话。带 `Mcp-Session-Id` 请求头；服务器清除会话的刷新、路由和允许记录，返回
HTTP 204。缺少或无效会话头时仍返回 204。

### `GET /mcp`

非 MCP 标准的只读健康检查，不需要会话。返回 `service`、`version`、`instance_id`、
`process_id`、当前后端端口/地址、固定网关地址和 `gateway_owner`。它只说明服务进程存在，
不能替代 `initialize` 或 `ping`。

### `OPTIONS /mcp`

返回 HTTP 204，仅用于方法探测；不会返回 CORS 允许头。带非空 `Origin` 的请求会在此之前
被拒绝，因此不能借此从网页调用 MCP。

## 工具调用的统一响应

成功调度的 `tools/call` 返回：

```json
{
  "content": [{ "type": "text", "text": "..." }],
  "structuredContent": { "ok": true },
  "isError": false
}
```

- 优先读取 `structuredContent`；`content[0].text` 是为仅支持文本内容的客户端提供的摘要。
- 工具业务失败通常返回 `isError=true` 和 `structuredContent.ok=false`，并带 `error`。
- 参数 Schema 不合法、工具名不可用或缺少外部 CAS 哈希时，`tools/call` 本身返回
  JSON-RPC `-32602`，此时没有上述工具结果包装。
- `source_state=no_source_open` 表示目标 IDE 没有打开工程；先让用户打开 `.e` 文件，再调用
  `get_current_eide_info` 确认。

## 多实例路由工具

### `list_instances`

列出全部正在运行并登记的 AutoLinker IDE 实例，同时显示本会话当前选中的实例。调用会探测
其他实例是否可达。

参数：无，必须传 `{}`。

关键返回字段：`gateway_endpoint`、`gateway_instance_id`、`active_instance_id` 和
`instances[]`。每个实例包含 `instance_id`、进程信息、`source_state`、工程路径字段
`source_file_path`、`page_name` / `page_type`、后端端口、`reachable`、
`gateway_owner`、`active` 等字段。实例结果没有 `project_path`；判断目标工程时必须读取
`source_file_path`。

### `select_instance`

把本 MCP 会话后续工具调用路由到指定 IDE 实例。先调用 `list_instances`，不要自行构造 ID。

| 参数 | 类型 | 必填 | 约束与说明 |
| --- | --- | --- | --- |
| `instance_id` | string | 是 | `list_instances.instances[].instance_id`；非空，只允许这一个参数 |

成功返回 `active_instance`。实例不存在或不可达时分别返回 `instance_not_found` 或
`instance_unreachable`。选中实例退出后不会自动改投其他工程，应重新列举并明确选择。

## 工作区与读取工具

### `refresh_workspace_mirror`

从 IDE 当前内存态导出临时快照，并通过 e-packager 重建文本镜像。外部 MCP 会话首次读取或
编辑源码前必须成功调用；镜像包含尚未保存的 IDE 修改，不会改写磁盘上的 `.e` 文件。

| 参数 | 类型 | 必填 | 默认/约束 | 说明 |
| --- | --- | --- | --- | --- |
| `mode` | string | 否 | `auto` | `auto`、`main_only`、`full`；`main_only` 尽量只刷新主工程，`full` 完整重建依赖镜像 |

关键返回字段：`refresh_mode`、`mirror_generation`、`source_file_path`、`file_count`、
`mirror_root`。`mirror_root` 仅供诊断，不能直接编辑。

### `e_packager`

仅当用户明确要求解析、查看、参考或复刻外部 `.e/.ec` 文件时调用。工具通过 e-packager
将外部源码解包到当前工作区镜像的只读目录
`unimported_code/<文件名>/`。解包成功后使用 `list_files` / `search_code` 定位，再用
`read_file` / `read_files` 读取；不要对该目录调用 `read_code_item`，也不要使用写工具修改。

该工具要求当前 MCP 会话已成功调用 `refresh_workspace_mirror`。相对路径按当前工程目录解析，
也可以传绝对路径。只接受已存在的普通 `.e` 或 `.ec` 文件；同名输出目录只有在本次解包成功
后才替换。

| 参数 | 类型 | 必填 | 默认/约束 | 说明 |
| --- | --- | --- | --- | --- |
| `file_path` | string | 是 | 1 到 32768 字节 | 外部 `.e/.ec` 绝对路径，或相对当前工程目录的路径 |

成功返回：

- `ok=true`
- `operation="unpack"`
- `source_file_path`
- `output_directory`
- `file_count`
- `mirror_generation`
- `read_hint`

`e_packager` 会改变当前镜像代次。解包后不要继续使用旧的分页
`mirror_generation`；重新从第一页调用 `list_files` 或 `search_code`，再使用新代次续页。

常见失败包括：`file_path is required`、参数不是对象、文件不存在、扩展名不是 `.e/.ec`、
e-packager 不可用或解包失败。

### `list_files`

列出镜像文件。未指定 `glob` 和 `path` 时，默认只显示 `src/`、`ecom/`、`elib/`、
`header/` 等文本区域。

| 参数 | 类型 | 必填 | 默认/约束 | 说明 |
| --- | --- | --- | --- | --- |
| `glob` | string | 否 | 空 | Glob，例如 `src/**/*.txt` |
| `path` | string | 否 | 空 | 相对路径前缀，例如 `ecom/` |
| `mirror_generation` | integer | 否 | - | 仅分页续传；必须是上一页返回值，最小 1 |
| `offset` | integer | 否 | `0` | 零基结果偏移，使用上一页 `next_offset` |
| `limit` | integer | 否 | `500` | 1 到 5000 |

返回 `files[]`、匹配总数 `count`、`returned`、`has_more`、`next_offset` 和
`mirror_generation`。

### `search_code`

逐文件流式搜索镜像文本。默认做区分大小写的字面子串搜索；优先用 `patterns` 一次搜索多个
相关名称，避免重复扫描同一批文件。

| 参数 | 类型 | 必填 | 默认/约束 | 说明 |
| --- | --- | --- | --- | --- |
| `pattern` | string | 条件必填 | 最大 1024 字节 | 单个模式；`pattern` 和 `patterns` 至少提供一个非空值 |
| `patterns` | string[] | 条件必填 | 最多 16 项，每项最大 1024 字节 | 批量模式，选项对全部模式生效，重复项会合并 |
| `glob` | string | 否 | 空 | 文件过滤，例如 `src/**/*.txt` |
| `output_mode` | string | 否 | `content` | `content`、`files_with_matches` 或 `count` |
| `regex` | boolean | 否 | `false` | 启用 ECMAScript 正则；复杂或潜在灾难性正则会被拒绝 |
| `case_insensitive` | boolean | 否 | `false` | 忽略 ASCII 大小写 |
| `context` | integer | 否 | `0` | 匹配前后上下文行数，0 到 20 |
| `mirror_generation` | integer | 否 | - | 仅分页续传；传上一页准确 generation |
| `offset` | integer | 否 | `0` | 每个查询的零基结果偏移 |
| `head_limit` | integer | 否 | `200` | 每个查询返回 1 到 2000 条结果 |

单模式返回 `results`、`match_count`、`files_with_matches`、分页和完整性字段；批量模式返回
`queries[]` 及每个模式的独立结果和 continuation。`results_complete=false` 表示有文件截断、
读取失败或取消，不能把当前结果当作完整否定结论。

### `read_file`

读取一个镜像文本文件，返回带行号的内容。单个字节窗口最多 1 MiB。

| 参数 | 类型 | 必填 | 默认/约束 | 说明 |
| --- | --- | --- | --- | --- |
| `file_path` | string | 是 | 镜像相对路径 | 例如 `src/Main.txt` |
| `mirror_generation` | integer | 否 | - | 分页或字节续传时传上一页 generation |
| `byte_offset` | integer | 否 | `0` | 大文件源字节游标；使用 `next_source_byte_offset` |
| `offset` | integer | 否 | `0` | 当前字节窗口内的零基行偏移 |
| `limit` | integer | 否 | `2000` | 1 到 20000 行 |

关键返回字段：`content`、`code_kind=mirror_source`、`mirror_generation`、行分页字段、
`source_byte_offset`、`next_source_byte_offset`、`code_hash` 和 `code_hash_complete`。
只有 `code_hash_complete=true` 时该镜像哈希覆盖完整文件，但它仍不是写入 CAS 基线。

### `read_files`

批量读取最多 12 个镜像文件。优先使用它替代连续多次 `read_file`；单次总返回不超过
12000 行，某个文件失败不会隐藏其他文件的成功结果。

| 参数 | 类型 | 必填 | 默认/约束 | 说明 |
| --- | --- | --- | --- | --- |
| `file_paths` | string[] | 条件必填 | 最多 12 项 | 简单路径列表，与 `files` 至少提供一种 |
| `files` | object[] | 条件必填 | 最多 12 项 | 为每个文件单独指定游标和行范围 |
| `mirror_generation` | integer | 否 | - | 批次续传使用的统一 generation |
| `byte_offset` | integer | 否 | `0` | 所有条目的默认字节游标 |
| `offset` | integer | 否 | `0` | 所有条目的默认行偏移 |
| `limit` | integer | 否 | `1200` | 每个文件 1 到 2000 行 |

`files[]` 条目参数：

| 参数 | 类型 | 必填 | 默认/约束 | 说明 |
| --- | --- | --- | --- | --- |
| `file_path` | string | 是 | 镜像相对路径 | 目标文件 |
| `mirror_generation` | integer | 否 | 顶层值 | 兼容别名；多个条目出现时必须全部一致 |
| `byte_offset` | integer | 否 | 顶层值 | 该文件的字节游标 |
| `offset` | integer | 否 | 顶层值 | 该文件的行偏移 |
| `limit` | integer | 否 | 顶层值 | 该文件最多 1 到 2000 行 |

返回 `files[]`、`status`、`all_ok`、成功/失败计数、`truncated` 和
`omitted_requests`。`ok=true` 只表示至少一个文件成功；必须结合 `all_ok` 判断整批完整性。

### `read_code_item`

按顶级声明名称读取完整易语言代码项，适合已知子程序、DLL 声明、数据类型等名称的场景。
它能够流式定位位于文件 1 MiB 之后的代码项。

| 参数 | 类型 | 必填 | 默认/约束 | 说明 |
| --- | --- | --- | --- | --- |
| `file_path` | string | 是 | 镜像相对路径 | 声明所在文件 |
| `item_name` | string | 是 | 非空 | 顶级声明名称，匹配时忽略 ASCII 大小写 |
| `occurrence` | integer | 否 | 自动 | 从 1 开始；同名声明有多个且省略时返回候选而不猜测 |
| `mirror_generation` | integer | 否 | - | 续传一致性 generation |
| `include_references` | boolean | 否 | `false` | 同时执行有界、忽略大小写的字面引用搜索 |
| `reference_glob` | string | 否 | 空 | 引用搜索文件过滤，例如 `src/**/*.txt` |
| `reference_limit` | integer | 否 | `20` | 1 到 50 条引用结果 |

返回声明 `directive`、起止行、`content`、`item_hash` 或完整文件 `code_hash`、
`item_complete`；启用引用时增加 `references`。

### `read_real_file`

按镜像 `file_path` 读取 IDE 当前真实页面或窗口 XML，取得实时 CAS 基线。`write_window.code_files` 是例外：其中已有代码文件使用刷新后镜像的 `read_file.code_hash`。

| 参数 | 类型 | 必填 | 默认/约束 | 说明 |
| --- | --- | --- | --- | --- |
| `file_path` | string | 是 | 可写 `src/*.txt` 或已有 `src/*.xml` | 映射到 IDE 程序项或实时窗口 |
| `offset` | integer | 否 | `0` | 零基行偏移，使用 `next_offset` |
| `limit` | integer | 否 | `20000` | 1 到 20000 行 |

返回 `content`、`code_hash`、`page_name`、`type_key`、`total_lines`、
`has_more`、`next_offset`。即使分页只读取一部分，`code_hash` 仍代表整个真实页。内容中的
数字行号和 Tab 是显示前缀，提交编辑参数前必须移除。

## 源码编辑工具

普通编辑针对当前工程的 `src/*.txt`；`edit_file`、`multi_edit_file`、`write_file` 也支持已有窗口 XML，走完整工程事务。已有目标先 `read_real_file`；新窗口与原子事件代码见 `write_window`。TXT 快照和 XML 事务安全快照不同，不承诺 XML 返回可用于 `restore_file_snapshot` 的 `snapshot_id`。

### `edit_file`

在一个真实页面中执行一次精确替换。`old_text` 必须恰好匹配一次；ASCII 双引号与 IDE
左右双引号视为等价，其他标点、缩进和换行仍需精确匹配。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `file_path` | string | 是 | 目标 `src/*.txt` 或已有 `src/*.xml` 镜像路径 |
| `old_text` | string | 是 | 非空、必须只匹配一次的原文 |
| `new_text` | string | 是 | 替换文本，可为空以删除原文 |
| `expected_base_hash` | string | 外部 MCP 必填 | 刚刚读取的 `read_real_file.code_hash` |

匹配 0 次或多次都不会写入；大段或整页修改应改用 `write_file`。

### `multi_edit_file`

按数组顺序在同一个真实页面应用多项替换。

| 参数 | 类型 | 必填 | 默认/约束 | 说明 |
| --- | --- | --- | --- | --- |
| `file_path` | string | 是 | - | 目标 `src/*.txt` |
| `edits` | object[] | 是 | 非空 | 顺序执行的编辑项 |
| `fail_on_unmatched` | boolean | 否 | `true` | 任一项未匹配或多重匹配时视为失败 |
| `atomic` | boolean | 否 | `true` | 有失败时放弃整批写入 |
| `expected_base_hash` | string | 外部 MCP 必填 | - | `read_real_file.code_hash` |

`edits[]` 条目：

| 参数 | 类型 | 必填 | 默认 | 说明 |
| --- | --- | --- | --- | --- |
| `old_text` | string | 是 | - | 非空精确原文 |
| `new_text` | string | 是 | - | 替换文本 |
| `replace_all` | boolean | 否 | `false` | `false` 要求恰好一次；`true` 替换全部匹配 |

只有同时设置 `fail_on_unmatched=false` 和 `atomic=false` 才允许跳过失败项并写入成功项；
默认行为是整批原子失败。

XML 编辑始终是完整工程事务，不支持以 `atomic=false` 留下部分提交；未匹配项处理仍受 `fail_on_unmatched` 控制。

### `write_file`

用完整源码覆盖一个真实页面。适合大范围重构；调用者必须先读取所有分页、移除显示行号，
再构造完整 `full_code`。不要只把当前可见分页作为整页写回。

已有窗口 XML 也使用 `full_code` 提交完整非空 XML；新增窗口用 `write_window`。下表中的 `file_path` 在此情况下为 `src/*.xml`。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `file_path` | string | 是 | 目标 `src/*.txt` |
| `full_code` | string | 是 | 完整易语言页面源码，可为空字符串 |
| `expected_base_hash` | string | 外部 MCP 必填 | `read_real_file.code_hash` |

### `diff_file`

基于 IDE 真实页预览结构化差异，不执行写入。外部调用仍要求 CAS 哈希，确保预览针对刚读取
的版本。提供以下三种候选内容方式之一：完整源码、单项替换或批量替换。

| 参数 | 类型 | 必填 | 默认/约束 | 说明 |
| --- | --- | --- | --- | --- |
| `file_path` | string | 是 | - | 目标 `src/*.txt` |
| `new_code` | string | 条件必填 | - | 完整候选源码，优先于 `full_code` |
| `full_code` | string | 条件必填 | - | `new_code` 的兼容别名 |
| `old_text` | string | 条件必填 | 非空 | 单项精确替换原文，配合 `new_text` |
| `new_text` | string | 否 | 空 | 单项替换新文 |
| `edits` | object[] | 条件必填 | - | 与 `multi_edit_file.edits[]` 相同 |
| `fail_on_unmatched` | boolean | 否 | `true` | 替换模式下遇到失败是否立即失败 |
| `expected_base_hash` | string | 外部 MCP 必填 | - | `read_real_file.code_hash` |

返回 `base_hash`、`new_hash`、`changed_lines`、`hunks[]` 和 `proposed_code`。

### `restore_file_snapshot`

恢复最近一次写入前自动保存的真实页快照，或恢复指定 `snapshot_id`。恢复本身也会先保存
当前页为新快照，因此仍可再次撤回。

| 参数 | 类型 | 必填 | 默认/约束 | 说明 |
| --- | --- | --- | --- | --- |
| `file_path` | string | 是 | - | 目标 `src/*.txt` |
| `snapshot_id` | string | 否 | - | 指定历史快照；可从写工具成功结果取得 |
| `restore_latest` | boolean | 否 | 未传 ID 时 `true`，传 ID 时 `false` | 是否忽略 ID 并恢复最新快照 |
| `expected_current_hash` | string | 外部 MCP 必填 | - | 恢复前最新 `read_real_file.code_hash` |

指定 `restore_latest=false` 时必须提供 `snapshot_id`。

### `add_new_file`

在当前易语言工程中新建普通程序集或类，校验最终对象名称和源码，并自动重建完整镜像。

| 参数 | 类型 | 必填 | 默认/约束 | 说明 |
| --- | --- | --- | --- | --- |
| `file_type` | string | 是 | `assembly` 或 `class` | 新建程序集或类模块 |
| `name` | string | 是 | 1 到 255 字节 | 最终名称；不能含逗号、Tab 或换行，且不能与现有项重名 |
| `full_code` | string | 否 | IDE 默认源码 | 完整页面源码；其中 `.程序集` 名称会规范为 `name` |

类会保留/补齐 IDE 生命周期函数。成功结果包含最终 `file_path`、`name`、`code_hash`、
`verified=true` 和新的镜像状态。它不能新建窗口或控件。

### `write_window`

新建窗口或替换完整 XML，同时可原子提交事件代码。

| 参数 | 类型 | 必填 | 约束与说明 |
| --- | --- | --- | --- |
| `file_path` | string | 是 | 当前工程 `src/*.xml`，文件名须与 XML 窗口名一致 |
| `full_xml` | string | 是 | 完整非空 XML，最多 16777216 字符 |
| `expected_base_hash` | string | 是 | 已有窗口用实时 `read_real_file.code_hash`，新窗口为空串 |
| `assembly_name` | string | 否 | 1–255 字符；默认 `窗口程序集_窗口名`，可给未绑定窗口附加程序集，不能切换已有绑定 |
| `code_files` | object[] | 否 | 最多 16 个当前工程程序集 TXT，和 XML 原子提交 |

`code_files[]` 每项必填 `file_path`、`full_code`、`expected_base_hash`，禁止其他字段。已有 TXT 使用最新刷新后 `read_file.code_hash`，新文件哈希为空串。提交完整程序集声明、处理器和参数／返回类型。

事件名称与签名以真实 XML 的 `.事件定义` 为准；绑定形式为 `<按钮.事件 名称="被单击" 处理器="程序集::子程序"/>`。处理器必须属于窗口程序集，每条绑定使用独立处理器。解绑可能重命名保留方法；跨程序集引用需要通过 `code_files` 显式同步。

成功检查 `ok`、`verified`、最终 `code_hash`、`source_path_preserved` 和 `project_saved=false`。失败检查 `rolled_back`、`rollback_skipped`、`safety_snapshot`；候选哈希不代表已提交状态。EideCore 缺失时 `source_validation.performed=false`；加载预检不等于编译验证。窗口整体删除、窗口改名和更换已有绑定程序集未开放；撤回成功事务需重交旧 XML/代码。

### `delete_file`

删除普通程序集或类及其全部子程序。必填 `file_path`（镜像路径）和 `expected_base_hash`（最新 `read_real_file.code_hash`），均为非空字符串。拒绝窗口绑定程序集、XML、固定表和导入模块。不会改写其他页引用或自动保存。成功返回 `deleted=true`、`verified=true` 和镜像刷新结果；若已删除但镜像刷新失败，先刷新，不直接重试删除。

## 资源与模块工具

### `list_resources`

可选 `resource_type` 字符串：`all`（默认）、`picture`、`sound`。返回当前工程本地资源的 `resources`、`count`，不含导入模块资源。条目包含 ID、名称、引用、字节数、内容 SHA-256 及元数据加内容的 `resource_hash`。图片和声音表均保存原始字节集，可存任意二进制内容，用 `#名称` 引用。

### `add_resource`

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `resource_type` | string | 是 | `picture` 或 `sound` |
| `name` | string | 是 | 1–255 字符，唯一资源名 |
| `file_path` | string | 是 | 1–32768 字符，UTF-8 绝对路径，现有普通文件，最大 64 MiB |
| `comment` | string | 否 | 最多 4096 字符 |

按原始字节添加，不检查或转换图像／音频格式。验证字节后使镜像失效，不自动保存；后续源码读取前先刷新镜像。

### `delete_resource`

必填 `resource_id`（整数，1–4294967295）和 `expected_hash`（64 字符字符串）。先 `list_resources`，传回其 `resource_hash`，不是纯内容 SHA-256。删除不改写 `#名称` 引用、不保存工程；成功后先刷新镜像再读取源码。

### `list_imported_modules`

仅在用户明确要求模块管理时调用，参数为空对象，列出当前工程已导入的 EC 模块。不要用于自动探索依赖，源码检索应读取 `ecom/` 镜像。

### `add_module_to_project`

仅按用户明确要求导入。参数均为可选字段，但须提供 `module_path` 或 `module_name` 定位目标。

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `module_path` | string | 绝对 `.ec` 文件路径 |
| `module_name` | string | 从易语言模块目录解析的模块名或文件名 |
| `prefer_new_method` | boolean | 默认 true；false 使用旧 IDE 导入路径，可能仅安装到系统模块目录 |
| `allow_blocking_import` | boolean | 默认 false，仅解析目标；执行真实导入必须 true |

### `remove_module_from_project`

仅按用户明确要求移除工程引用。参数为 `module_index`（整数，最小 0）、`module_path` 或 `module_name`（字符串），至少提供一种定位方式。先依据当前模块列表核对目标；不要用 `delete_file` 删除模块引用。

## IDE 状态与编译工具

### `get_current_page_info`

获取当前 IDE 页面名称、类型和解析来源。参数：无，必须传 `{}`。要求已打开工程。

返回 `source_file_path`、`page_name`、`page_type` 和 `page_name_trace`。

### `get_current_eide_info`

获取当前 IDE 实例和工程状态。参数：无，必须传 `{}`；即使没打开工程也可以调用，是确认
`source_state` 的首选工具。

关键返回字段包括进程信息、`source_state` / `source_open`、工程路径、当前页、
`project_type`、支持的编译目标/模式、默认编译目标、实例 MCP 地址和固定网关状态。

### `compile_with_output_path`

编译当前工程并抑制 IDE 保存文件对话框。只有产物存在且相对编译前被创建或更新时才返回
成功；不能仅凭 IDE 编译调用返回值判断成功。

已安装 EideCore 时，普通及静态 EXE/DLL 编译使用独立核心，返回 `backend="eidecore"`；黑月、EC 发布和核心缺失时走原流程。独立核心编译未保存内存快照，校验输出 PE 和 SHA-256 后替换目标，失败保留旧产物。错误定位使用返回的原生源码片段，不直接套用镜像 TXT 行号。

编译命令调用与后续产物确认共用 10 分钟等待预算；IDE 原生调用不能被强制中断，
因此这不是终止编译进程的硬时限。跨实例 MCP 转发独立使用 2 秒连接超时、
5 秒发送超时和 630 秒响应等待预算（普通工具为 30 秒）；外部客户端需设置更长的请求超时。
等待响应超时返回 JSON-RPC 错误 `-32003`，`data.error=tool_response_timeout`，
并标记 `execution_state=unknown`、`may_still_be_running=true`、`retry_safe=false`。
目标可能仍在编译，应确认 IDE 状态及产物后再决定下一步，不要自动重试或重新选择实例。
请求完整发送后发生其他响应故障时返回 `tool_response_unavailable`，同样不能推断编译未执行。

| 参数 | 类型 | 必填 | 默认/约束 | 说明 |
| --- | --- | --- | --- | --- |
| `output_path` | string | 是 | 非空 | 输出文件路径，建议使用绝对路径；扩展名会按目标规范化 |
| `target` | string | 否 | `auto` | `auto`、`win_exe`、`win_console_exe`、`win_dll`、`ecom` |
| `static_compile` | boolean | 否 | EXE/DLL 为 `true`，`ecom` 为 `false` | 仅用户明确要求动态编译时对 EXE/DLL 传 `false` |
| `compiler` | string | 否 | `eide` | `eide` 或 `blackmoon`；黑月须显式设置 |
| `blackmoon_mode` | string | 否 | `vcxx` | `asm` / `cxx` / `vcxx`，普通编译省略 |

黑月编译须完全省略 `static_compile`，目标仅允许 `auto` 或当前 EXE/DLL 工程类型。

模块工程的 `target=auto` 会构建 `win_console_exe` 测试程序；只有发布 `.ec` 时才显式使用
`target=ecom`，且该目标不支持静态编译。返回产物指纹验证、编译输出、等待状态和错误光标
位置；判断结果以 `ok` / `artifact_verified` 为准。

编译结果另含 `diagnostics[]`，每项包含 `severity`、`message`、`file`、`line`、
`location_verified` 和 `location_status`。已核实的位置还包含 `code`、带行号的
`code_excerpt`、`page_name`、`ide_row` / `ide_column`。这些诊断也包含在 MCP
`content[].text` 摘要中，纯文本客户端无需依赖 `structuredContent` 才能读到错误代码。

错误行文本优先通过 IDE SDK `FN_GET_PRG_TEXT` 取得，`code_source` 标明来源；
源码行及节选通过真实页直接格式化器核实，不经整页复制或系统剪贴板。
`code_excerpt` 包含前后各两行并以 `>` 标出错误行，`focus_line`、
`excerpt_start_line` / `excerpt_end_line` 明确对应位置。SDK 不可用时允许原生行范围
读取回退。原生编译命令在 IDE 中不可用时返回 `compile_command_not_enabled`。

- `line` 从 1 开始，与 `read_real_file` 的实际源码行一致；`file` 为镜像相对路径。
- `ide_row` / `ide_column` 是 IDE 原生的零基位置，不能把 `ide_row + 1` 当作源码行号。
- 跨程序集诊断的页名优先取真实源码声明，并核对编辑器对象；IDE 自动跳转后的
  页签标题可能过期，不能据此判定错误属于哪个文件。
- 仅当原生行与整页源码映射、文件对应的编辑器对象、错误消息引用的符号均核实后，
  才返回 `location_verified=true` 和可用于定位的源码行。前缀映射不能验证时，仅接受
  唯一完整行匹配；存在歧义则 `file` / `line` 为 `null`，不猜测。
- 当前支持易语言 `错误(...)`、`警告(...)`、语法错误和编译错误文本；外部链接器、
  黑月或无源码位置的诊断可能只有原始输出，不能把当前光标当作它们的错误位置。
- 兼容字段 `caret_*` 表示采集到的光标状态，消费者应以诊断的
  `location_verified` 判断是否可以作为错误定位。

输出归属由 `output_scope`、`output_cleared_before_compile` 和
`output_belongs_to_current_compile` 描述：

| `output_scope` | 含义 |
| --- | --- |
| `compile_session` | 来自本次编译期间的内部 Hook 捕获 |
| `cleared_control_compile_session` | 编译前已清空并验证控件基线，本次读取的新输出 |
| `appended_since_compile_start` | 无已验证的清空基线，仅返回追加差分 |
| `replaced_since_compile_start` | 控件全文发生替换，不能排除历史文本 |
| `unchanged` / `unavailable` | 无文本变化或无法读取输出 |

前两种情况下 `output_belongs_to_current_compile=true`，否则为 `null`（未知）。
编译开始会清空原生输出窗口；`output_window_text` 可能包含本次编译期间的
AutoLinker 日志，并不保证其中每一行都是编译器消息。


## 网络工具

### `search_web_tavily`

通过已配置的 Tavily API 搜索公网，返回规范化摘要。需要在 AutoLinker AI 接口设置中配置
Tavily API Key。

| 参数 | 类型 | 必填 | 默认/约束 | 说明 |
| --- | --- | --- | --- | --- |
| `query` | string | 是 | 非空 | 搜索词 |
| `max_results` | integer | 否 | `5`，1 到 10 | 最大结果数 |
| `topic` | string | 否 | `general` | 直接传给 Tavily 的主题值 |

### `fetch_url`

HTTP GET 获取公开 URL 的规范化文本和基础响应元数据。仅允许 `http` / `https`，阻止
回环、私有和本地网络地址；重定向目标也会重新校验，最多跟随 5 次。

| 参数 | 类型 | 必填 | 默认/约束 | 说明 |
| --- | --- | --- | --- | --- |
| `url` | string | 是 | 公开 HTTP(S) URL | 目标地址 |
| `timeout_seconds` | integer | 否 | `60`，1 到 300 | 总请求超时 |
| `max_bytes` | integer | 否 | `524288`，4096 到 2097152 | 最大响应字节数 |

返回 `final_url`、`http_status`、`content_type`、`content_length`、`body_text` 和
`body_truncated`。需要正文而非原始文本时优先用 `extract_web_document`。

### `extract_web_document`

获取公开网页或文本文档并提取标题、可读正文、摘要和少量绝对链接。网络安全限制、超时和
字节上限与 `fetch_url` 相同。

| 参数 | 类型 | 必填 | 默认/约束 | 说明 |
| --- | --- | --- | --- | --- |
| `url` | string | 是 | 公开 HTTP(S) URL | 目标网页或文档 |
| `timeout_seconds` | integer | 否 | `60`，1 到 300 | 总请求超时 |
| `max_bytes` | integer | 否 | `524288`，4096 到 2097152 | 最大源响应字节数 |

返回 `title`、`plain_text`、`excerpt`、`links[]`、HTTP 元数据和
`source_truncated`。

## 推荐调用流程

### 探索或读取工程

1. `initialize`，保存 `Mcp-Session-Id`。
2. 多开 IDE 时调用 `list_instances`，再用 `select_instance` 明确目标。
3. 调用 `get_current_eide_info`，确认 `source_open=true` 和工程路径。
4. 调用 `refresh_workspace_mirror`；需要查 EC/支持库时优先用 `full`。
5. 用 `list_files` 探索结构，用 `search_code` 定位，用 `read_files` 批量读取；已知声明名时
   用 `read_code_item`。

### 修改已有源码

1. 完成读取流程，并确定可写的 `src/*.txt` 路径。
2. 紧接着调用 `read_real_file`，必要时读取全部分页；保存 `code_hash`。
3. 小范围修改用 `edit_file`，多项修改用 `multi_edit_file`，整页重构用 `write_file`；全部
   携带 `expected_base_hash`。
4. 写工具返回 `ok=true` 后，不要为了确认而立即重复写入；需要验证行为时调用
   `compile_with_output_path`。

### 撤销写入

1. 调用 `read_real_file` 获取当前 `code_hash`。
2. 调用 `restore_file_snapshot`，传目标 `snapshot_id` 或 `restore_latest=true`，并把当前
   哈希放入 `expected_current_hash`。

## 常见错误

按错误原文查阅 [错误原因与排查参考](troubleshooting.md)，其中区分 MCP 会话/路由、镜像准备、源码 CAS、内存快照、编译和恢复失败。参数定义仍以本页及当前实例的 `tools/list` 为准。

JSON-RPC `-32003` 的 `tool_response_timeout` / `tool_response_unavailable` 表示执行结果未知；先核查目标实例状态和日志，不要自动重发写入、恢复或编译请求。
