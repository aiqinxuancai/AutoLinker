#include "AutoLinkerInternal.h"
#include <Windows.h>
#include <unordered_map>
#include <string>
#include "AIService.h"
#include "EPackagerIntegration.h"
#include "Global.h"
#include "IDEFacade.h"
#include "IdeCompileOutputCapture.h"
#include "IdeLogViewer.h"

namespace {
bool g_isContextMenuRegistered = false;
HMENU g_topLinkerSubMenu = NULL;
std::unordered_map<UINT, std::string> g_topLinkerCommandMap;
constexpr ULONG_PTR kOwnedViewMenuSeparatorTag = 0x41564D53;
constexpr wchar_t kLogViewerMenuTitle[] = L"AutoLinker日志中心";
constexpr wchar_t kDebugOutputOptimizationMenuTitle[] =
	L"AutoLinker 调试输出效率优化（试验）";

// 根据当前打开的源文件路径生成链接器父菜单项的显示标题（Wide 字符串）。
// 无源文件时返回通用名；有源文件时返回"[xxxx.e]使用的链接器"。
std::wstring GetLinkerMenuTitle()
{
	if (g_nowOpenSourceFilePath.empty()) {
		return L"源文件链接器切换";
	}
	const auto lastSep = g_nowOpenSourceFilePath.find_last_of("\\/");
	const std::string filenameAnsi = (lastSep != std::string::npos)
		? g_nowOpenSourceFilePath.substr(lastSep + 1)
		: g_nowOpenSourceFilePath;
	if (filenameAnsi.empty()) {
		return L"源文件链接器切换";
	}
	const int wlen = MultiByteToWideChar(CP_ACP, 0, filenameAnsi.c_str(), -1, nullptr, 0);
	if (wlen <= 0) {
		return L"源文件链接器切换";
	}
	std::wstring filenameW(static_cast<size_t>(wlen) - 1, L'\0');
	MultiByteToWideChar(CP_ACP, 0, filenameAnsi.c_str(), -1, filenameW.data(), wlen);
	return L"[" + filenameW + L"]使用的链接器";
}

// 在 hTargetMenu 中找到链接器子菜单所在的父菜单项，更新其标题与启用状态。
// 须在 g_nowOpenSourceFilePath 已刷新后调用。
void UpdateLinkerSubMenuParentItem(HMENU hTargetMenu)
{
	if (hTargetMenu == nullptr || g_topLinkerSubMenu == nullptr) {
		return;
	}
	const int count = GetMenuItemCount(hTargetMenu);
	for (int i = 0; i < count; ++i) {
		MENUITEMINFOW mii = {};
		mii.cbSize = sizeof(mii);
		mii.fMask = MIIM_SUBMENU;
		if (GetMenuItemInfoW(hTargetMenu, static_cast<UINT>(i), TRUE, &mii) &&
			mii.hSubMenu == g_topLinkerSubMenu) {
			std::wstring title = GetLinkerMenuTitle();
			MENUITEMINFOW miiUpdate = {};
			miiUpdate.cbSize = sizeof(miiUpdate);
			miiUpdate.fMask = MIIM_STRING | MIIM_STATE | MIIM_SUBMENU;
			miiUpdate.fState = g_nowOpenSourceFilePath.empty() ? MFS_GRAYED : MFS_ENABLED;
			miiUpdate.hSubMenu = g_topLinkerSubMenu;
			miiUpdate.dwTypeData = title.data();
			miiUpdate.cch = static_cast<UINT>(title.size());
			SetMenuItemInfoW(hTargetMenu, static_cast<UINT>(i), TRUE, &miiUpdate);
			break;
		}
	}
}
}

void RegisterIDEContextMenu()
{
	if (g_isContextMenuRegistered) {
		return;
	}

	auto& ide = IDEFacade::Instance();
	ide.RegisterContextMenuItem(IDM_AUTOLINKER_CTX_COPY_FUNC, "复制当前函数代码", []() {
		TryCopyCurrentFunctionCode();
	});
	ide.RegisterContextMenuItem(IDM_AUTOLINKER_CTX_AI_TRANSLATE_TEXT, "AI翻译选中文本", []() {
		RunAiTranslateSelectedTextTask();
	});
	ide.RegisterContextMenuItem(IDM_AUTOLINKER_CTX_AI_OPTIMIZE_FUNC, "AI优化函数", []() {
		RunAiFunctionReplaceTask(AITaskKind::OptimizeFunction);
	});
	ide.RegisterContextMenuItem(IDM_AUTOLINKER_CTX_AI_COMMENT_FUNC, "AI为当前函数添加注释", []() {
		RunAiFunctionReplaceTask(AITaskKind::AddCommentsToFunction);
	});
	ide.RegisterContextMenuItem(IDM_AUTOLINKER_CTX_AI_TRANSLATE_FUNC, "AI翻译当前函数+变量名", []() {
		RunAiFunctionReplaceTask(AITaskKind::TranslateFunctionAndVariables);
	});
	ide.RegisterContextMenuItem(IDM_AUTOLINKER_CTX_AI_ADD_BY_PAGE, "AI按当前页类型添加代码", []() {
		RunAiAddByCurrentPageTypeTask();
	});

	g_isContextMenuRegistered = true;
}

std::wstring GetMenuTitleW(HMENU hMenu, UINT item, UINT flags)
{
	wchar_t title[256] = { 0 };
	int len = GetMenuStringW(hMenu, item, title, static_cast<int>(sizeof(title) / sizeof(title[0])), flags);
	if (len <= 0) {
		return L"";
	}
	return std::wstring(title, static_cast<size_t>(len));
}

bool IsCompileOrToolsTopPopup(HMENU hPopupMenu)
{
	if (g_hwnd == NULL || hPopupMenu == NULL) {
		return false;
	}
	if (hPopupMenu == g_topLinkerSubMenu) {
		return false;
	}

	// 右键菜单里会包含该命令，直接排除，避免“链接器切换”混入右键。
	UINT copyState = GetMenuState(hPopupMenu, IDM_AUTOLINKER_CTX_COPY_FUNC, MF_BYCOMMAND);
	if (copyState != 0xFFFFFFFF) {
		return false;
	}

	auto popupKeywordMatch = [hPopupMenu]() -> bool {
		int keywordHit = 0;
		int itemCount = GetMenuItemCount(hPopupMenu);
		for (int item = 0; item < itemCount; ++item) {
			std::wstring itemTitle = GetMenuTitleW(hPopupMenu, static_cast<UINT>(item), MF_BYPOSITION);
			if (itemTitle.find(L"静态编译") != std::wstring::npos){
				++keywordHit;
			}
		}
		return keywordHit >= 1;
	};

	HMENU hMainMenu = GetMenu(g_hwnd);
	if (hMainMenu == NULL) {
		return popupKeywordMatch();
	}

	int count = GetMenuItemCount(hMainMenu);
	int popupIndex = -1;
	int compileIndex = -1;
	int toolsIndex = -1;

	for (int i = 0; i < count; ++i) {
		HMENU subMenu = GetSubMenu(hMainMenu, i);
		if (subMenu == NULL) {
			continue;
		}
		if (subMenu == hPopupMenu) {
			popupIndex = i;
		}

		std::wstring title = GetMenuTitleW(hMainMenu, static_cast<UINT>(i), MF_BYPOSITION);
		if (title.find(L"编译") != std::wstring::npos || title.find(L"Build") != std::wstring::npos) {
			if (compileIndex < 0) {
				compileIndex = i;
			}
			continue;
		}
		if (title.find(L"工具") != std::wstring::npos || title.find(L"Tools") != std::wstring::npos) {
			if (toolsIndex < 0) {
				toolsIndex = i;
			}
		}
	}

	if (popupIndex < 0) {
		return popupKeywordMatch();
	}

	int targetIndex = compileIndex >= 0 ? compileIndex : toolsIndex;
	if (targetIndex >= 0) {
		return popupIndex == targetIndex;
	}

	// 只接受主菜单的直接子菜单，避免把右键菜单误判成顶部菜单。
	return false;
}

bool IsViewTopPopup(HMENU hPopupMenu)
{
	if (g_hwnd == nullptr || hPopupMenu == nullptr) {
		return false;
	}
	auto popupKeywordMatch = [hPopupMenu]() {
		static constexpr const wchar_t* kViewKeywords[] = {
			L"工具条", L"工具栏", L"状态条", L"状态栏", L"工作夹", L"输出夹",
			L"窗口组件箱", L"自定义数据类型表", L"全局变量表", L"Dll命令定义表",
			L"DLL命令定义表", L"常量数据表", L"资源表", L"书签", L"预览被设计窗口",
			L"Toolbar", L"Status Bar", L"Workspace", L"Output", L"Global Variable",
			L"DLL Command", L"Constant", L"Resource", L"Bookmark", L"Preview"
		};
		int keywordHits = 0;
		const int itemCount = GetMenuItemCount(hPopupMenu);
		for (int item = 0; item < itemCount; ++item) {
			const std::wstring itemTitle = GetMenuTitleW(
				hPopupMenu,
				static_cast<UINT>(item),
				MF_BYPOSITION);
			for (const wchar_t* keyword : kViewKeywords) {
				if (itemTitle.find(keyword) != std::wstring::npos) {
					++keywordHits;
					break;
				}
			}
		}
		return keywordHits >= 2;
	};

	HMENU mainMenu = GetMenu(g_hwnd);
	if (mainMenu == nullptr) {
		return popupKeywordMatch();
	}

	const int count = GetMenuItemCount(mainMenu);
	for (int index = 0; index < count; ++index) {
		if (GetSubMenu(mainMenu, index) != hPopupMenu) {
			continue;
		}
		const std::wstring title = GetMenuTitleW(mainMenu, static_cast<UINT>(index), MF_BYPOSITION);
		return title.find(L"查看") != std::wstring::npos || title.find(L"View") != std::wstring::npos;
	}
	return popupKeywordMatch();
}

void RemoveOwnedViewMenuSeparators(HMENU viewMenu)
{
	for (int index = GetMenuItemCount(viewMenu) - 1; index >= 0; --index) {
		MENUITEMINFOW item = {};
		item.cbSize = sizeof(item);
		item.fMask = MIIM_FTYPE | MIIM_DATA;
		if (GetMenuItemInfoW(viewMenu, static_cast<UINT>(index), TRUE, &item) &&
			(item.fType & MFT_SEPARATOR) != 0 &&
			item.dwItemData == kOwnedViewMenuSeparatorTag) {
			DeleteMenu(viewMenu, static_cast<UINT>(index), MF_BYPOSITION);
		}
	}
}

void AppendOwnedViewMenuSeparator(HMENU viewMenu)
{
	MENUITEMINFOW separator = {};
	separator.cbSize = sizeof(separator);
	separator.fMask = MIIM_FTYPE | MIIM_DATA;
	separator.fType = MFT_SEPARATOR;
	separator.dwItemData = kOwnedViewMenuSeparatorTag;
	InsertMenuItemW(
		viewMenu,
		static_cast<UINT>(GetMenuItemCount(viewMenu)),
		TRUE,
		&separator);
}

bool IsOwnedViewMenuItemAtPosition(
	HMENU viewMenu,
	int position,
	UINT commandId,
	const wchar_t* expectedTitle)
{
	return position >= 0 &&
		GetMenuItemID(viewMenu, position) == commandId &&
		GetMenuTitleW(viewMenu, static_cast<UINT>(position), MF_BYPOSITION) == expectedTitle;
}

void RemoveOwnedViewMenuItems(
	HMENU viewMenu,
	UINT commandId,
	const wchar_t* expectedTitle)
{
	for (int index = GetMenuItemCount(viewMenu) - 1; index >= 0; --index) {
		if (IsOwnedViewMenuItemAtPosition(viewMenu, index, commandId, expectedTitle)) {
			DeleteMenu(viewMenu, static_cast<UINT>(index), MF_BYPOSITION);
		}
	}
}

void RefreshOwnedViewMenuItemState(
	HMENU viewMenu,
	UINT commandId,
	const wchar_t* expectedTitle,
	bool checked)
{
	for (int index = GetMenuItemCount(viewMenu) - 1; index >= 0; --index) {
		if (!IsOwnedViewMenuItemAtPosition(viewMenu, index, commandId, expectedTitle)) {
			continue;
		}
		EnableMenuItem(viewMenu, static_cast<UINT>(index), MF_BYPOSITION | MF_ENABLED);
		CheckMenuItem(
			viewMenu,
			static_cast<UINT>(index),
			MF_BYPOSITION | (checked ? MF_CHECKED : MF_UNCHECKED));
	}
}

void RefreshAutoLinkerViewMenuItemStates(HMENU viewMenu)
{
	RefreshOwnedViewMenuItemState(
		viewMenu,
		IDM_AUTOLINKER_LOG_CENTER,
		kLogViewerMenuTitle,
		IdeLogViewer::IsOpen());
	RefreshOwnedViewMenuItemState(
		viewMenu,
		IDM_AUTOLINKER_DEBUG_OUTPUT_OPTIMIZATION,
		kDebugOutputOptimizationMenuTitle,
		IdeCompileOutputCapture::IsDebugOutputOptimizationEnabled());
}

void EnsureAutoLinkerViewMenuItems(HMENU viewMenu)
{
	if (viewMenu == nullptr) {
		return;
	}

	const int initialCount = GetMenuItemCount(viewMenu);
	if (initialCount >= 2 &&
		IsOwnedViewMenuItemAtPosition(
			viewMenu,
			initialCount - 2,
			IDM_AUTOLINKER_LOG_CENTER,
			kLogViewerMenuTitle) &&
		IsOwnedViewMenuItemAtPosition(
			viewMenu,
			initialCount - 1,
			IDM_AUTOLINKER_DEBUG_OUTPUT_OPTIMIZATION,
			kDebugOutputOptimizationMenuTitle)) {
		RefreshAutoLinkerViewMenuItemStates(viewMenu);
		return;
	}

	// 只移动 ID 和标题都匹配的 AutoLinker 项，避免碰触其他插件的菜单命令。
	RemoveOwnedViewMenuItems(viewMenu, IDM_AUTOLINKER_LOG_CENTER, kLogViewerMenuTitle);
	RemoveOwnedViewMenuItems(
		viewMenu,
		IDM_AUTOLINKER_DEBUG_OUTPUT_OPTIMIZATION,
		kDebugOutputOptimizationMenuTitle);
	RemoveOwnedViewMenuSeparators(viewMenu);

	const int count = GetMenuItemCount(viewMenu);
	if (count > 0) {
		const UINT lastState = GetMenuState(viewMenu, static_cast<UINT>(count - 1), MF_BYPOSITION);
		if (lastState != 0xFFFFFFFF && (lastState & MF_SEPARATOR) != MF_SEPARATOR) {
			AppendOwnedViewMenuSeparator(viewMenu);
		}
	}
	AppendMenuW(
		viewMenu,
		MF_STRING | MF_ENABLED | (IdeLogViewer::IsOpen() ? MF_CHECKED : MF_UNCHECKED),
		IDM_AUTOLINKER_LOG_CENTER,
		kLogViewerMenuTitle);
	AppendMenuW(
		viewMenu,
		MF_STRING | MF_ENABLED |
			(IdeCompileOutputCapture::IsDebugOutputOptimizationEnabled() ? MF_CHECKED : MF_UNCHECKED),
		IDM_AUTOLINKER_DEBUG_OUTPUT_OPTIMIZATION,
		kDebugOutputOptimizationMenuTitle);

	// MFC 会把没有原生 ON_COMMAND 映射的动态命令置灰，需在宿主更新后恢复启用态。
	RefreshAutoLinkerViewMenuItemStates(viewMenu);
}

void ClearMenuItemsByPosition(HMENU hMenu)
{
	if (hMenu == NULL) {
		return;
	}

	int count = GetMenuItemCount(hMenu);
	for (int i = count - 1; i >= 0; --i) {
		DeleteMenu(hMenu, static_cast<UINT>(i), MF_BYPOSITION);
	}
}

bool EnsureTopLinkerSubMenu()
{
	if (g_topLinkerSubMenu == NULL || !IsMenu(g_topLinkerSubMenu)) {
		g_topLinkerSubMenu = CreatePopupMenu();
	}
	if (g_topLinkerSubMenu == NULL) {
		return false;
	}

	return true;
}

bool IsMenuSeparator(HMENU hMenu, int index)
{
	if (hMenu == NULL || index < 0) {
		return false;
	}
	UINT state = GetMenuState(hMenu, static_cast<UINT>(index), MF_BYPOSITION);
	return state != 0xFFFFFFFF && (state & MF_SEPARATOR) == MF_SEPARATOR;
}

void TrimTrailingSeparators(HMENU hMenu)
{
	if (hMenu == NULL) {
		return;
	}
	for (;;) {
		const int count = GetMenuItemCount(hMenu);
		if (count <= 0 || !IsMenuSeparator(hMenu, count - 1)) {
			return;
		}
		DeleteMenu(hMenu, static_cast<UINT>(count - 1), MF_BYPOSITION);
	}
}

void RemoveExistingTopMenuExtensions(HMENU hTargetMenu)
{
	if (hTargetMenu == NULL) {
		return;
	}

	for (int i = GetMenuItemCount(hTargetMenu) - 1; i >= 0; --i) {
		MENUITEMINFOW mii = {};
		mii.cbSize = sizeof(mii);
		mii.fMask = MIIM_SUBMENU | MIIM_ID;

		bool removeThis = false;
		bool preserveSubMenu = false;
		if (GetMenuItemInfoW(hTargetMenu, static_cast<UINT>(i), TRUE, &mii)) {
			if (mii.hSubMenu == g_topLinkerSubMenu || mii.wID == IDM_AUTOLINKER_UNPACK_SOURCE) {
				removeThis = true;
				preserveSubMenu = mii.hSubMenu == g_topLinkerSubMenu;
			}
		}
		if (!removeThis) {
			std::wstring title = GetMenuTitleW(hTargetMenu, static_cast<UINT>(i), MF_BYPOSITION);
			if (title.find(L"链接器切换") != std::wstring::npos ||
				title.find(L"使用的链接器") != std::wstring::npos ||
				title.find(L"反编译到目录") != std::wstring::npos) {
				removeThis = true;
				preserveSubMenu = mii.hSubMenu == g_topLinkerSubMenu;
			}
		}
		if (removeThis) {
			if (preserveSubMenu) {
				RemoveMenu(hTargetMenu, static_cast<UINT>(i), MF_BYPOSITION);
			}
			else {
				DeleteMenu(hTargetMenu, static_cast<UINT>(i), MF_BYPOSITION);
			}
		}
	}

	TrimTrailingSeparators(hTargetMenu);
}

void EnsureTopLinkerSubMenuAttached(HMENU hTargetMenu)
{
	if (hTargetMenu == NULL) {
		return;
	}
	if (!EnsureTopLinkerSubMenu()) {
		return;
	}

	RemoveExistingTopMenuExtensions(hTargetMenu);
	if (!EnsureTopLinkerSubMenu()) {
		return;
	}

	int count = GetMenuItemCount(hTargetMenu);
	if (count > 0) {
		UINT lastState = GetMenuState(hTargetMenu, static_cast<UINT>(count - 1), MF_BYPOSITION);
		if (lastState != 0xFFFFFFFF && (lastState & MF_SEPARATOR) != MF_SEPARATOR) {
			AppendMenuW(hTargetMenu, MF_SEPARATOR, 0, NULL);
		}
	}
	AppendMenuW(hTargetMenu, MF_POPUP | MF_STRING, reinterpret_cast<UINT_PTR>(g_topLinkerSubMenu), GetLinkerMenuTitle().c_str());
	UINT unpackFlags = MF_STRING | (EPackagerIntegration::CanUnpackCurrentSource() ? MF_ENABLED : MF_GRAYED);
	AppendMenuW(hTargetMenu, unpackFlags, IDM_AUTOLINKER_UNPACK_SOURCE, EPackagerIntegration::BuildUnpackMenuTitle().c_str());
}

void RebuildTopLinkerSubMenu()
{
	if (!EnsureTopLinkerSubMenu()) {
		return;
	}

	ClearMenuItemsByPosition(g_topLinkerSubMenu);
	g_topLinkerCommandMap.clear();

	if (g_linkerManager.getCount() <= 0) {
		AppendMenuA(g_topLinkerSubMenu, MF_STRING | MF_GRAYED, 0, "（未找到Linker配置）");
		return;
	}

	UpdateCurrentOpenSourceFile();
	std::string nowLinkConfigName = g_configManager.getValue(g_nowOpenSourceFilePath);

	UINT cmd = IDM_AUTOLINKER_LINKER_BASE;
	const auto configs = g_linkerManager.getMap();
	auto appendLinkerConfig = [&](const std::string& key, const LinkConfig& value) {
		if (cmd > IDM_AUTOLINKER_LINKER_MAX) {
			return false;
		}

		UINT flags = MF_STRING | MF_ENABLED;
		if ((value.isDefault && (nowLinkConfigName.empty() || nowLinkConfigName == key)) ||
			(!value.isDefault && nowLinkConfigName == key)) {
			flags |= MF_CHECKED;
		}

		AppendMenuA(g_topLinkerSubMenu, flags, cmd, key.c_str());
		g_topLinkerCommandMap[cmd] = key;
		++cmd;
		return true;
	};

	bool appendedDefault = false;
	for (const auto& [key, value] : configs) {
		if (value.isDefault) {
			appendedDefault = appendLinkerConfig(key, value);
			break;
		}
	}
	if (appendedDefault && configs.size() > 1 && cmd <= IDM_AUTOLINKER_LINKER_MAX) {
		AppendMenuA(g_topLinkerSubMenu, MF_SEPARATOR, 0, nullptr);
	}
	for (const auto& [key, value] : configs) {
		if (value.isDefault) {
			continue;
		}
		if (!appendLinkerConfig(key, value)) {
			break;
		}
	}
}

bool HandleTopLinkerMenuCommand(UINT cmd)
{
	if (cmd == IDM_AUTOLINKER_UNPACK_SOURCE) {
		EPackagerIntegration::RunCurrentSourceUnpackToDirectory();
		return true;
	}

	auto it = g_topLinkerCommandMap.find(cmd);
	if (it == g_topLinkerCommandMap.end()) {
		return false;
	}

	UpdateCurrentOpenSourceFile();
	if (g_nowOpenSourceFilePath.empty()) {
		OutputStringToELog("当前没有打开源文件，无法切换Linker");
		return true;
	}

	const LinkConfig& selectedConfig = g_linkerManager.getConfig(it->second);
	if (selectedConfig.isDefault) {
		g_configManager.setValue(g_nowOpenSourceFilePath, "");
	}
	else {
		g_configManager.setValue(g_nowOpenSourceFilePath, it->second);
	}
	OutputCurrentSourceLinker();
	RebuildTopLinkerSubMenu();
	return true;
}

void HandleInitMenuPopup(HMENU hMenu)
{
	if (hMenu == NULL) {
		return;
	}
	if (hMenu == g_topLinkerSubMenu) {
		RebuildTopLinkerSubMenu();
		return;
	}
	if (IsViewTopPopup(hMenu)) {
		EnsureAutoLinkerViewMenuItems(hMenu);
		return;
	}

	if (IsCompileOrToolsTopPopup(hMenu)) {
		UpdateCurrentOpenSourceFile();
		EnsureTopLinkerSubMenuAttached(hMenu);
		RebuildTopLinkerSubMenu();
		UpdateLinkerSubMenuParentItem(hMenu);
		return;
	}

	UINT state = GetMenuState(hMenu, IDM_AUTOLINKER_CTX_COPY_FUNC, MF_BYCOMMAND);
	if (state != 0xFFFFFFFF) {
		EnableMenuItem(hMenu, IDM_AUTOLINKER_CTX_COPY_FUNC, MF_BYCOMMAND | MF_ENABLED);
	}

	const UINT aiCmdIds[] = {
		IDM_AUTOLINKER_CTX_AI_OPTIMIZE_FUNC,
		IDM_AUTOLINKER_CTX_AI_COMMENT_FUNC,
		IDM_AUTOLINKER_CTX_AI_TRANSLATE_FUNC,
		IDM_AUTOLINKER_CTX_AI_TRANSLATE_TEXT,
		IDM_AUTOLINKER_CTX_AI_ADD_BY_PAGE
	};
	bool hasAnyAiCommand = false;
	for (UINT cmdId : aiCmdIds) {
		if (GetMenuState(hMenu, cmdId, MF_BYCOMMAND) != 0xFFFFFFFF) {
			hasAnyAiCommand = true;
			break;
		}
	}
	if (!hasAnyAiCommand) {
		return;
	}

	const bool hasSelectedText = IDEFacade::Instance().IsFunctionEnabled(FN_EDIT_CUT);
	for (UINT cmdId : aiCmdIds) {
		UINT aiState = GetMenuState(hMenu, cmdId, MF_BYCOMMAND);
		if (aiState != 0xFFFFFFFF) {
			if (cmdId == IDM_AUTOLINKER_CTX_AI_TRANSLATE_TEXT) {
				EnableMenuItem(hMenu, cmdId, MF_BYCOMMAND | (hasSelectedText ? MF_ENABLED : MF_GRAYED));
				continue;
			}
			EnableMenuItem(hMenu, cmdId, MF_BYCOMMAND | MF_ENABLED);
		}
	}
}

void PrepareAutoLinkerPopupMenu(HMENU hMenu)
{
	if (hMenu == NULL) {
		return;
	}

	auto& ide = IDEFacade::Instance();
	if (ide.InjectContextMenuToPopup(hMenu)) {
		ide.RefreshContextMenuEnabledState(hMenu);
		return;
	}

	if (hMenu == g_topLinkerSubMenu || IsCompileOrToolsTopPopup(hMenu) || IsViewTopPopup(hMenu)) {
		HandleInitMenuPopup(hMenu);
		return;
	}

	const bool hasKnownAutoLinkerCommand =
		GetMenuState(hMenu, IDM_AUTOLINKER_CTX_COPY_FUNC, MF_BYCOMMAND) != 0xFFFFFFFF ||
		GetMenuState(hMenu, IDM_AUTOLINKER_CTX_AI_OPTIMIZE_FUNC, MF_BYCOMMAND) != 0xFFFFFFFF ||
		GetMenuState(hMenu, IDM_AUTOLINKER_CTX_AI_COMMENT_FUNC, MF_BYCOMMAND) != 0xFFFFFFFF ||
		GetMenuState(hMenu, IDM_AUTOLINKER_CTX_AI_TRANSLATE_FUNC, MF_BYCOMMAND) != 0xFFFFFFFF ||
		GetMenuState(hMenu, IDM_AUTOLINKER_CTX_AI_TRANSLATE_TEXT, MF_BYCOMMAND) != 0xFFFFFFFF ||
		GetMenuState(hMenu, IDM_AUTOLINKER_CTX_AI_ADD_BY_PAGE, MF_BYCOMMAND) != 0xFFFFFFFF;
	if (!hasKnownAutoLinkerCommand) {
		return;
	}

	ide.RefreshContextMenuEnabledState(hMenu);
}

bool IsKnownAutoLinkerPopup(HMENU hMenu)
{
	if (hMenu == NULL) {
		return false;
	}

	return
		GetMenuState(hMenu, IDM_AUTOLINKER_CTX_COPY_FUNC, MF_BYCOMMAND) != 0xFFFFFFFF ||
		GetMenuState(hMenu, IDM_AUTOLINKER_CTX_AI_OPTIMIZE_FUNC, MF_BYCOMMAND) != 0xFFFFFFFF ||
		GetMenuState(hMenu, IDM_AUTOLINKER_CTX_AI_COMMENT_FUNC, MF_BYCOMMAND) != 0xFFFFFFFF ||
		GetMenuState(hMenu, IDM_AUTOLINKER_CTX_AI_TRANSLATE_FUNC, MF_BYCOMMAND) != 0xFFFFFFFF ||
		GetMenuState(hMenu, IDM_AUTOLINKER_CTX_AI_TRANSLATE_TEXT, MF_BYCOMMAND) != 0xFFFFFFFF ||
		GetMenuState(hMenu, IDM_AUTOLINKER_CTX_AI_ADD_BY_PAGE, MF_BYCOMMAND) != 0xFFFFFFFF;
}

void FinalizeAutoLinkerPopupMenu(HMENU hMenu)
{
	if (hMenu == NULL) {
		return;
	}

	if (hMenu == g_topLinkerSubMenu || IsCompileOrToolsTopPopup(hMenu) || IsViewTopPopup(hMenu)) {
		HandleInitMenuPopup(hMenu);
		return;
	}

	if (!IsKnownAutoLinkerPopup(hMenu)) {
		return;
	}

	auto& ide = IDEFacade::Instance();
	ide.RefreshContextMenuEnabledState(hMenu);
	HandleInitMenuPopup(hMenu);
}
