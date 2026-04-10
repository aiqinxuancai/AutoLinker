#include "AutoLinkerInternal.h"
#include <Windows.h>
#include <unordered_map>
#include <string>
#include "AIService.h"
#include "Global.h"
#include "IDEFacade.h"

namespace {
bool g_isContextMenuRegistered = false;
HMENU g_topLinkerSubMenu = NULL;
std::unordered_map<UINT, std::string> g_topLinkerCommandMap;

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
			miiUpdate.fMask = MIIM_STRING | MIIM_STATE;
			miiUpdate.fState = g_nowOpenSourceFilePath.empty() ? MFS_GRAYED : MFS_ENABLED;
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
		return false;
	}

	int targetIndex = compileIndex >= 0 ? compileIndex : toolsIndex;
	if (targetIndex >= 0) {
		return popupIndex == targetIndex;
	}

	// 只接受主菜单的直接子菜单，避免把右键菜单误判成顶部菜单。
	return false;
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

void EnsureTopLinkerSubMenuAttached(HMENU hTargetMenu)
{
	if (hTargetMenu == NULL) {
		return;
	}
	if (!EnsureTopLinkerSubMenu()) {
		return;
	}

	bool alreadyAttached = false;
	int attachedIndex = -1;
	for (int i = GetMenuItemCount(hTargetMenu) - 1; i >= 0; --i) {
		MENUITEMINFOW mii = {};
		mii.cbSize = sizeof(mii);
		mii.fMask = MIIM_SUBMENU;
		bool removeThis = false;
		if (GetMenuItemInfoW(hTargetMenu, static_cast<UINT>(i), TRUE, &mii) && mii.hSubMenu == g_topLinkerSubMenu) {
			alreadyAttached = true;
			attachedIndex = i;
			continue;
		}
		if (!removeThis) {
			std::wstring title = GetMenuTitleW(hTargetMenu, static_cast<UINT>(i), MF_BYPOSITION);
			if (title.find(L"链接器切换") != std::wstring::npos) {
				removeThis = true;
			}
		}
		if (removeThis) {
			DeleteMenu(hTargetMenu, static_cast<UINT>(i), MF_BYPOSITION);
			if (i > 0) {
				UINT prevState = GetMenuState(hTargetMenu, static_cast<UINT>(i - 1), MF_BYPOSITION);
				if (prevState != 0xFFFFFFFF && (prevState & MF_SEPARATOR) == MF_SEPARATOR) {
					DeleteMenu(hTargetMenu, static_cast<UINT>(i - 1), MF_BYPOSITION);
				}
			}
		}
	}

	if (alreadyAttached) {
		if (attachedIndex > 0) {
			UINT prevState = GetMenuState(hTargetMenu, static_cast<UINT>(attachedIndex - 1), MF_BYPOSITION);
			if (prevState == 0xFFFFFFFF || (prevState & MF_SEPARATOR) != MF_SEPARATOR) {
				InsertMenuW(hTargetMenu, static_cast<UINT>(attachedIndex), MF_BYPOSITION | MF_SEPARATOR, 0, NULL);
			}
		}
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
	for (const auto& [key, value] : g_linkerManager.getMap()) {
		if (cmd > IDM_AUTOLINKER_LINKER_MAX) {
			break;
		}

		UINT flags = MF_STRING | MF_ENABLED;
		if (!nowLinkConfigName.empty() && nowLinkConfigName == key) {
			flags |= MF_CHECKED;
		}

		AppendMenuA(g_topLinkerSubMenu, flags, cmd, key.c_str());
		g_topLinkerCommandMap[cmd] = key;
		++cmd;
	}
}

bool HandleTopLinkerMenuCommand(UINT cmd)
{
	auto it = g_topLinkerCommandMap.find(cmd);
	if (it == g_topLinkerCommandMap.end()) {
		return false;
	}

	UpdateCurrentOpenSourceFile();
	if (g_nowOpenSourceFilePath.empty()) {
		OutputStringToELog("当前没有打开源文件，无法切换Linker");
		return true;
	}

	g_configManager.setValue(g_nowOpenSourceFilePath, it->second);
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

	if (hMenu == g_topLinkerSubMenu || IsCompileOrToolsTopPopup(hMenu)) {
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

	if (hMenu == g_topLinkerSubMenu || IsCompileOrToolsTopPopup(hMenu)) {
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
