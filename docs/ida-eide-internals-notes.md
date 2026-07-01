# 易语言 IDE 内部对象与源码导出逆向笔记

本文记录当前围绕 `e5.95.exe.i64` 的 IDA MCP 分析结果，重点覆盖项目二进制序列化、剪贴板、代码编辑器展示源码格式化，以及当前活动编辑器对象模型。

## 基本约定

- 分析目标：`C:\Users\aiqin\OneDrive\e5.6\e5.95.exe.i64`
- 架构：x86
- IDA image base：`0x400000`
- 下文地址均按 IDA 中显示的地址记录。
- AutoLinker 代码中这些地址按 `kImageBase = 0x400000` 处理，运行时地址计算为 `moduleBase + (idaAddress - 0x400000)`。

## 项目二进制序列化

日志中的：

```text
[AutoLinker][ProjectBinarySerializer] hook serialize_to_file rva=0x458020 via=pattern
```

对应 IDA 地址 `0x458020`，即 `EProjectBinarySerializer_SaveToFile`。这里的 `rva` 是项目内按 IDA 地址风格记录的值，不需要再额外加 `0x400000`。

已命名的项目序列化相关函数：

| 地址 | 名称 | 作用 |
| --- | --- | --- |
| `0x456FC0` | `EProjectBinarySerializer_GetSavePath` | 获取/准备工程保存路径。 |
| `0x4570D0` | `EProjectBinarySerializer_LoadFromFile` | 从文件加载工程。 |
| `0x457490` | `EProjectBinarySerializer_LoadFromFileEx` | 带扩展逻辑的工程加载入口。 |
| `0x458020` | `EProjectBinarySerializer_SaveToFile` | 将当前工程保存到文件。 |
| `0x458400` | `EProjectBinarySerializer_SerializeArchive` | 通过 MFC `CArchive` 序列化工程对象。 |
| `0x45B810` | `EProjectBinarySerializer_ReencodeThroughMemory` | 通过内存中转重编码/重序列化工程数据。 |
| `0x45B9A0` | `EProjectBinarySerializer_LoadFromClipboard` | 从注册剪贴板格式读取工程/工程片段二进制数据。 |
| `0x45C4A0` | `EProjectBinarySerializer_DeserializeArchive` | 通过 `CArchive` 反序列化工程对象。 |
| `0x460B90` | `EProjectBinarySerializer_UnloadSupportLibraries` | 卸载/清理工程关联支持库。 |
| `0x460C80` | `EProjectBinarySerializer_IsModified` | 判断工程是否已修改。 |
| `0x460CB0` | `EProjectBinarySerializer_MarkSaved` | 标记工程已保存。 |
| `0x460CE0` | `EProjectBinarySerializer_ResetProjectState` | 重置工程状态。 |
| `0x460FD0` | `EProjectBinarySerializer_DestroyStringsAndArrays` | 析构/释放内部字符串和数组。 |
| `0x464050` | `EProjectBinarySerializer_SaveToHGlobal` | 将工程序列化到 `HGLOBAL`，用于剪贴板或内存导出。 |
| `0x464170` | `EProjectBinarySerializer_CopyToClipboard` | 将工程/工程片段复制到注册剪贴板格式。 |

辅助文件对象：

| 地址 | 名称 | 作用 |
| --- | --- | --- |
| `0x458380` | `EProjectMirrorFile_SetPath` | 设置内部文件路径。 |
| `0x458390` | `EProjectMirrorFile_ScalarDeletingDestructor` | scalar deleting destructor。 |
| `0x4583B0` | `EProjectMirrorFile_Destructor` | 析构内部文件对象。 |

工程片段剪贴板包装函数：

| 地址 | 名称 | 作用 |
| --- | --- | --- |
| `0x51C7F0` | `EProjectClipboard_CopySelectedAsProjectPackage` | 将选中内容复制为工程片段包。 |
| `0x51CC50` | `EProjectClipboard_PasteProjectPackage` | 从工程片段包粘贴。 |
| `0x4C86C0` | `EProjectClipboard_PastePackageIntoView` | 将片段包插入到当前视图/编辑器。 |

### CArchive

`CArchive` 是 MFC 的二进制序列化流包装，通常包在 `CFile`、`CMemFile`、`CSharedFile` 等对象上。易语言 IDE 的工程保存/加载逻辑通过它把内存对象树写成 `.e` 二进制，或从 `.e` 二进制恢复对象树。

本项目看到的路径大致是：

```text
工程对象
  -> EProjectBinarySerializer_SerializeArchive
  -> CArchive
  -> CFile / CSharedFile / HGLOBAL
```

反向加载则是：

```text
文件或 HGLOBAL
  -> CArchive
  -> EProjectBinarySerializer_DeserializeArchive
  -> 工程对象
```

### HGLOBAL

`HGLOBAL` 是 Win32 全局内存句柄，常见来源是 `GlobalAlloc(GMEM_MOVEABLE, size)`。剪贴板 API 的 `SetClipboardData` 要求很多格式的数据必须用可移动全局内存块传入，因此 IDE 在复制工程片段、源码文本时都会先分配 `HGLOBAL`，写入数据，再交给剪贴板。

简化模型：

```text
GlobalAlloc -> HGLOBAL
GlobalLock  -> 得到可写指针
写入数据
GlobalUnlock
SetClipboardData(format, HGLOBAL)
```

## 工程二进制剪贴板与展示源码剪贴板的区别

`EProjectBinarySerializer_CopyToClipboard` 和 `EProjectBinarySerializer_LoadFromClipboard` 操作的是 IDE 注册的内部二进制剪贴板格式，不是普通文本源码。

已确认：

- 注册剪贴板格式 ID 保存在 `uFormat` 附近，地址为 `0x676294`。
- 该格式用于 IDE 内部复制/粘贴工程、程序集、类、子程序等结构化对象。
- 它不是 `CF_TEXT`，直接读取无法得到代码框里可复制出来的源码文本。

要拿“人能看的源码文本”，应走代码编辑器 formatter 或 IDE 自己的复制命令，见下一节。

## 代码编辑器展示源码格式化链路

要拿某个程序集、类或页面在 IDE 代码框里展示的可复制源码，核心不是工程序列化器，而是代码编辑器对象的格式化函数。

本次通过签名扫描在 `e5.95.exe` 中唯一命中：

| 功能 | 地址 |
| --- | --- |
| 编辑器命令分发 | `0x4C2B20` |
| 获取 range 数量 | `0x4CCCB0` |
| 格式化 range 为文本 | `0x49A140` |

已在 IDA 中命名：

| 地址 | 名称 | 作用 |
| --- | --- | --- |
| `0x4C2B20` | `ECodeEditor_DispatchCommand` | 代码编辑器命令分发入口。 |
| `0x4C7B30` | `ECodeEditor_HandleCopyCommand` | 处理编辑器 Copy 命令 `0x02030001`。 |
| `0x4D52E0` | `ECodeEditor_CopySelectionToClipboard` | 将当前编辑器选区写入剪贴板。 |
| `0x49A140` | `ECodeEditor_FormatRangeToTextBuffer` | 核心源码格式化器，把内部代码模型转成人可读文本。 |
| `0x4CCCB0` | `ECodeEditor_GetRangeCount` | 返回当前代码页 range 数量。 |
| `0x4C5220` | `ECodeEditor_PrintFormattedRanges` | 打印路径的格式化调用者，也会调用 formatter。 |

### 命令分发到复制

编辑器内部 copy 命令：

```text
kEditorCmdCopy = 0x02030001
```

分发链路：

```text
ECodeEditor_DispatchCommand (0x4C2B20)
  -> 命令高位组 0x02030000
  -> ECodeEditor_HandleCopyCommand (0x4C7B30)
  -> ECodeEditor_CopySelectionToClipboard (0x4D52E0)
```

UI 命令中已知 Copy 为：

```text
kEditorUiCmdCopy = 0x2043
```

AutoLinker 当前会优先走内部命令，必要时走 UI 命令兜底。

### CopySelectionToClipboard 的两种输出

`ECodeEditor_CopySelectionToClipboard (0x4D52E0)` 会同时处理两类剪贴板数据。

内部二进制对象格式：

| 地址 | 说明 |
| --- | --- |
| `0x4D5767` | 调用 `EProjectBinarySerializer_SaveToHGlobal`，生成内部二进制剪贴板数据。 |
| `0x4D57A2` | `SetClipboardData(uFormat, hMem)`，写入 IDE 注册格式。 |

普通源码文本格式：

| 地址 | 说明 |
| --- | --- |
| `0x4D57EB` | 对选区中每个 range 调用 `ECodeEditor_FormatRangeToTextBuffer`。 |
| `0x4D58DC` | 为 `CF_TEXT` 分配 `HGLOBAL`。 |
| `0x4D591F` | `SetClipboardData(1, hMem)`，写入 `CF_TEXT`。这就是用户从代码框复制出来的文本。 |

因此，IDE 复制代码页时既会准备内部结构化数据，也会准备普通文本。我们要的“展示代码”就是 `CF_TEXT` 那一路。

### FormatRangeToTextBuffer

`ECodeEditor_FormatRangeToTextBuffer (0x49A140)` 是核心源码格式化器。

当前项目中对应函数指针类型：

```cpp
using FnEditorFormatRangeText = void(__thiscall*)(void*, int, int, void*, int);
```

调用模型：

```text
FormatRangeToTextBuffer(editorObject, startRange, endRange, outputBuffer, flags)
```

全页导出时：

```text
count = ECodeEditor_GetRangeCount(editorObject)
ECodeEditor_FormatRangeToTextBuffer(editorObject, 0, count - 1, buffer, 0)
```

函数内部会根据 `editorObject + 0x3C` 的页面类型分支处理。已观察到的类型包括 `1`、`2`、`3`、`4`、`6`、`7`、`8`。它会调用若干内部追加函数，把程序集、类、子程序、变量、常量等内部结构转换成带 `\r\n` 的易语言源码文本。

### GetRangeCount

`ECodeEditor_GetRangeCount (0x4CCCB0)` 反编译结果很短：

```cpp
int __thiscall ECodeEditor_GetRangeCount(void* this)
{
    sub_4D3C60(this);
    return sub_499EE0(v2);
}
```

对外意义是返回当前编辑器页可格式化的 range 数量。全页导出时使用 `0` 到 `count - 1`。

### 打印路径

`ECodeEditor_PrintFormattedRanges (0x4C5220)` 同样调用 `ECodeEditor_FormatRangeToTextBuffer`：

- 若只打印选区，则循环选中 range 调 formatter。
- 若打印整页，则先取 `ECodeEditor_GetRangeCount`，再调用 formatter。

这进一步确认 `0x49A140` 是通用的“内部模型到展示文本”格式化函数，不只服务剪贴板。

## AutoLinker 当前源码读取实现

相关源码：

| 文件 | 作用 |
| --- | --- |
| `src/EideInternalTextBridge.h` | 暴露真实页面源码读写接口。 |
| `src/EideInternalTextBridge.cpp` | 实现 editorObject 解析、复制、formatter 调用、剪贴板兜底。 |
| `src/AIChatToolingX86.cpp` | AI 工具层按程序树项读取真实页面源码。 |

关键接口：

```cpp
bool GetRealPageCodeByEditorObject(
    std::uintptr_t editorObject,
    std::uintptr_t moduleBase,
    std::string* outCode,
    NativeRealPageAccessResult* outResult = nullptr);

bool GetRealPageCodeByProgramTreeItemData(
    unsigned int itemData,
    std::uintptr_t moduleBase,
    std::string* outCode,
    NativeRealPageAccessResult* outResult = nullptr);
```

当前稳定路径：

```text
程序树 itemData
  -> 解析对应页面 editorObject
  -> Select All
  -> Copy
  -> fake clipboard 捕获 CF_TEXT
  -> 得到与 IDE 代码框复制一致的源码文本
```

直接 formatter 路径：

```text
editorObject
  -> ECodeEditor_GetRangeCount
  -> ECodeEditor_FormatRangeToTextBuffer
  -> InternalGenericArrayBuffer
  -> outCode
```

直接 formatter 当前受版本锁定探测影响。日志中出现过：

```text
version_locked_internal_interop_unsupported|generic_array_init=0...
```

相关内部数组函数地址：

| 地址 | 名称/用途 |
| --- | --- |
| `0x486060` | generic array init |
| `0x486260` | generic array destroy |
| `0x486920` | generic array assign |
| `0x4863A0` | generic array finalize |
| `0x574900` | generic array vtable |

也就是说，直接 formatter 路径的思路是对的，但当前实际读取仍以 IDE 复制路径更稳。

## 当前活动编辑器对象模型

结论：不是全局只有一个 `ECodeEditor` 反复初始化为不同页面。

更准确的模型：

```text
全局 mainEditorHost
  + activeEditorObject 指针槽位
      -> 当前活动页的 editorObject
```

每个打开的 MDI 页面或可解析页面通常有自己的 editor wrapper/object。切换页面时，IDE 会把全局 host 的 active editor 指针替换成对应页面的 editorObject，并通知旧/新对象。

### e5.95 地址

| 项 | 地址/偏移 |
| --- | --- |
| `mainEditorHost` | `0x6756E0` |
| `EMainEditorHost_SetActiveEditorObject` | `0x47ADF0` |
| `EMdiChild_ActivateContainedEditorObject` | `0x487060` |
| active editor 偏移 | `0x464` |

`0x47ADF0 EMainEditorHost_SetActiveEditorObject` 的关键行为：

```text
old = this + 0x464
if new != old:
    this + 0x464 = 0
    notify old deactivate/change
    this + 0x464 = new
    notify new activate/change
```

反编译中 `this + 281` 即 `this + 281 * 4 = this + 0x464`。

`0x487060 EMdiChild_ActivateContainedEditorObject` 的行为：

```text
editorObject = this + 0x40
if editorObject && activateFlag == 1:
    EMainEditorHost_SetActiveEditorObject(dword_6756E0, editorObject, 1)
```

### e571 地址

项目中已记录的 e571 profile：

| 项 | 地址/偏移 |
| --- | --- |
| `mainEditorHost` | `0x5CAE70` |
| set active editor | `0x471B30` |
| active editor caller | `0x47C840` |
| active editor 偏移 | `0x438` |

### raw editor 与 inner editor

AutoLinker 中区分：

- `rawEditorObject`：从 MDI 子窗口、全局 active host 或程序树解析得到的外层对象。
- `innerEditorObject`：某些页类型中真正接收编辑器命令/格式化的内部对象。

当前解析逻辑：

```cpp
pageType = fields[15];

switch (pageType) {
case 1: return fields[21];
case 2: return fields[16];
case 3: return fields[17];
case 4: return fields[18];
case 6:
case 7:
case 8: return fields[20];
default: return 0;
}
```

因此按页面读取源码时，不能简单假设所有页都是同一字段布局。

## 按程序集/类获取展示源码的推荐路径

如果已经有 editorObject：

```text
count = ECodeEditor_GetRangeCount(editorObject)
ECodeEditor_FormatRangeToTextBuffer(editorObject, 0, count - 1, buffer, 0)
读取 buffer 文本
```

如果只有程序树节点：

```text
程序树 itemData
  -> GetRealPageCodeByProgramTreeItemData
  -> 内部解析 editorObject
  -> 使用 IDE 复制链路或直接 formatter
  -> 返回展示源码文本
```

推荐优先使用项目现有封装：

```cpp
e571::GetRealPageCodeByProgramTreeItemData(itemData, moduleBase, &code, &result);
```

原因：

- 它能处理当前页、非当前页、隐藏解析和 UI 兜底。
- 它会在需要临时切换 active editor 时保存并恢复原 active editor。
- 它拿到的是与 IDE 代码框复制一致的文本。

## 相关但未完全展开的剪贴板/文本对象函数

这些地址已在项目中作为已知内部函数使用或初步分析，但本次没有完整展开所有结构字段：

| 项 | 地址 | 备注 |
| --- | --- | --- |
| text to object wrapper | `0x4C9220` | 位于包含函数 `0x4C8F60` 内，调用者包括 `0x498DF0`。 |
| text to object direct | `0x4CB160` | 位于包含函数 `0x4CAE20` 内，调用者包括 `0x4C3230`。 |
| insert clipboard object | `0x48B720` | 位于包含函数 `0x48B6F0` 内，调用者包括 `0x489810`、`0x489E00`、`0x48B900`、`0x48BBB0`。 |
| clipboard deserialize | `0x4535F0` | 位于包含函数 `0x453570` 内。 |
| clipboard serialize | `0x45B640` | 位于包含函数 `0x45B630` 内。 |

这些更偏向“文本粘贴解析成内部代码对象”和“内部对象序列化为剪贴板包”，与“读取展示源码文本”的主链路不同。

## 关键结论

1. `.e` 工程二进制序列化和 IDE 代码框展示源码是两条不同链路。
2. `EProjectBinarySerializer_*` 负责工程对象的二进制保存、加载、剪贴板包。
3. IDE 代码框可复制出的源码文本来自 `ECodeEditor_FormatRangeToTextBuffer (0x49A140)`。
4. Copy 命令最终在 `ECodeEditor_CopySelectionToClipboard (0x4D52E0)` 中同时写内部格式和 `CF_TEXT`。
5. `CF_TEXT` 路径才是人可读源码文本。
6. 全局不是只有一个 `ECodeEditor`；全局 host 只是保存当前 active editor 指针。
7. 切换页面时主要是替换 active editor 指针，而不是把同一个 editor 对象重新初始化成另一个页面。
8. 实际工程中按 `itemData` 调 `GetRealPageCodeByProgramTreeItemData` 是当前最稳的获取展示源码方式。
