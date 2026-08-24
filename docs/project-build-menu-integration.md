# 项目配置接入编译菜单

本文记录 AutoLinker 向易语言 IDE 的“编译”菜单增加“按配置编译”时，必须遵守的菜单生命周期和接入位置。

## 原生菜单结构

`e5.95.exe` 的原生“编译”菜单运行时包含 8 项：

1. `编译` (`0x8053`，`F7`)
2. `静态编译` (`0x8090`，`Shift+F7`)
3. `独立编译` (`0x808F`)
4. `编译为易包` (`0x80D4`，`Ctrl+F10`)
5. 分隔线
6. `编译为指定类型` 子菜单，包含 4 个目标
7. `静态编译为指定类型` 子菜单，包含 3 个目标
8. `编译生成安装软件` (`0x808D`，`Ctrl+Shift+F7`)

AutoLinker 的扩展项应追加在这组原生项目之后。当前扩展项为：

- 当前源文件使用的链接器
- `将当前 .e 反编译到目录`
- `按配置编译`

不要删除或重建 IDE 的原生菜单，也不要用固定位置替换原生项目。

## 正确调用路径

菜单调用顺序如下：

```text
TrackPopupMenu / TrackPopupMenuEx
    -> AutoLinkerHooks.cpp::PrepareAutoLinkerPopupMenu
    -> 易语言 IDE 自己处理 WM_INITMENUPOPUP
       （MainWindowSubclassProc 先调用 DefSubclassProc）
    -> AutoLinker.cpp::MainWindowSubclassProc
       -> FinalizeAutoLinkerPopupMenu
          -> HandleInitMenuPopup
             -> EnsureTopLinkerSubMenuAttached
             -> RebuildTopLinkerSubMenu
             -> RebuildProjectBuildSubMenu
```

### 前置阶段：只清理 AutoLinker 自己的旧项目

`PrepareAutoLinkerPopupMenu` 运行在 IDE 初始化弹出菜单之前。对于“编译”顶层菜单，前置阶段只能调用：

```cpp
RemoveExistingTopMenuExtensions(hMenu);
```

这一步用于处理 IDE 复用同一个 `HMENU` 的情况：上一次打开菜单时追加的链接器、项目配置和反编译项目必须先摘除，使 IDE 每次都能从原生 8 项开始初始化。

前置阶段不能追加新的项目。若在 `TrackPopupMenu` 钩子中直接 `AppendMenu`，IDE 后续按照自己的固定菜单结构刷新时会把原生编译项移除。

### 后置阶段：先让 IDE 初始化，再追加扩展

`MainWindowSubclassProc` 收到 `WM_INITMENUPOPUP` 时必须先执行：

```cpp
LRESULT result = DefSubclassProc(hWnd, uMsg, wParam, lParam);
```

然后再调用 `FinalizeAutoLinkerPopupMenu`。识别到编译顶层菜单后，`HandleInitMenuPopup` 才可以调用 `EnsureTopLinkerSubMenuAttached`，该函数会追加：

1. 链接器子菜单
2. 反编译到目录命令
3. `按配置编译` 子菜单

配置子菜单本身在 `RebuildProjectBuildSubMenu` 中从当前 `.e` 旁的 `*.autolinker.json` 读取，并为实际添加到子菜单的项目建立命令映射。命令处理只认领当前子菜单中真实存在的项目，不把整个数值区间视为 AutoLinker 命令。

## 关键安全条件

### 不能用空的子菜单句柄匹配普通命令

叶子菜单项目的 `MENUITEMINFO.hSubMenu` 通常为 `NULL`。因此删除扩展项目时必须先判断扩展句柄非空：

```cpp
const bool isProjectBuildSubMenu =
    g_projectBuildSubMenu != NULL &&
    mii.hSubMenu == g_projectBuildSubMenu;
```

错误写法如下：

```cpp
mii.hSubMenu == g_projectBuildSubMenu;
```

当 `g_projectBuildSubMenu` 尚未创建时，这个条件会对所有叶子项目成立，导致 `编译`、`静态编译`、`独立编译`、`编译为易包` 和安装软件项目被全部删除。该问题就是本次“编译菜单原有项目消失”的根因。

同样的非空保护必须用于 `g_topLinkerSubMenu`。标题回退匹配（`按配置编译`、`使用的链接器`、反编译标题）只能删除确认属于 AutoLinker 的项目，并使用 `RemoveMenu` 保留仍由 AutoLinker 管理的子菜单句柄。

### WM_COMMAND 归属判断

主窗口会收到 IDE、AutoLinker 以及其他插件的全部 `WM_COMMAND`。项目配置命令不能仅通过数值区间判断归属：处理器首先检查当前 `g_projectBuildSubMenu` 中是否真实存在该命令项；菜单被摘除时同步清空子菜单项和命令映射。未满足该条件的命令必须返回 `false`，继续交给 IDE 原窗口过程。

## 验收清单

每次修改顶层菜单逻辑后，至少验证以下项目：

- 首次打开“编译”菜单：原生 8 项全部存在。
- `编译`、`静态编译`、`独立编译`、`编译为易包` 和 `编译生成安装软件` 的命令 ID 未改变。
- `编译为指定类型` 子菜单仍有 4 个目标。
- `静态编译为指定类型` 子菜单仍有 3 个目标。
- 有配置文件时，`按配置编译` 子菜单显示文件中定义的配置。
- 没有配置文件时，子菜单只显示灰色的“目前没有编译配置”，不自动生成 `Debug`、`Release`。
- 在设置页删除最后一个配置时，配置文件会同步删除，菜单恢复为灰色的无配置状态。
- 连续关闭并重新打开菜单至少 3 次，结果保持一致。
- 当前无配置文件时显示灰色提示，而不是清空或破坏顶层原生菜单。

源码对应位置：

- `src/AutoLinkerHooks.cpp`：`MyTrackPopupMenu`、`MyTrackPopupMenuEx`
- `src/AutoLinker.cpp`：`MainWindowSubclassProc` 的 `WM_INITMENUPOPUP` 分支
- `src/AutoLinkerMenu.cpp`：菜单识别、前置清理、后置挂载和配置子菜单刷新
- `src/ProjectBuildConfigManager.cpp`：项目配置文件读写
- `src/ProjectBuildPipeline.cpp`：按配置启动异步无头编译流程
