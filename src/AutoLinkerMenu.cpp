#include "AutoLinkerInternal.h"
#include <Windows.h>
#include <unordered_map>
#include <string>
#include <string_view>
#include "AIService.h"
#include "EPackagerIntegration.h"
#include "Global.h"
#include "IDEFacade.h"
#include "ProjectBuildConfigManager.h"
#include "ProjectBuildPipeline.h"
#include "UnicodeTextCodec.h"

namespace {
bool g_isContextMenuRegistered = false;
HMENU g_topLinkerSubMenu = NULL;
HMENU g_projectBuildSubMenu = NULL;
std::unordered_map<UINT, std::string> g_topLinkerCommandMap;
std::unordered_map<UINT, std::string> g_projectBuildCommandMap;

std::string ProjectBuildUtf8(std::u8string_view text)
{
	return std::string(reinterpret_cast<const char*>(text.data()), text.size());
}

void OutputProjectBuildUtf8(const std::string& text)
{
	OutputStringToELog(UnicodeTextCodec::Utf8ToLocalPreservingUnicode(text));
}

std::wstring Utf8ToWideMenuText(const std::string& text)
{
	if (text.empty()) return {};
	const int length = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (length <= 0) return std::wstring(text.begin(), text.end());
	std::wstring result(static_cast<size_t>(length), L'\0');
	MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), length);
	return result;
}

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

bool HasSubMenuItem(HMENU hMenu, HMENU subMenu)
{
	if (hMenu == nullptr || subMenu == nullptr) return false;
	const int count = GetMenuItemCount(hMenu);
	for (int i = 0; i < count; ++i) {
		MENUITEMINFOW item = {};
		item.cbSize = sizeof(item);
		item.fMask = MIIM_SUBMENU;
		if (GetMenuItemInfoW(hMenu, static_cast<UINT>(i), TRUE, &item) && item.hSubMenu == subMenu) {
			return true;
		}
	}
	return false;
}

bool EnsureProjectBuildSubMenu()
{
	if (g_projectBuildSubMenu == NULL || !IsMenu(g_projectBuildSubMenu)) g_projectBuildSubMenu = CreatePopupMenu();
	return g_projectBuildSubMenu != NULL;
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
			const bool isLinkerSubMenu =
				g_topLinkerSubMenu != NULL && mii.hSubMenu == g_topLinkerSubMenu;
			const bool isProjectBuildSubMenu =
				g_projectBuildSubMenu != NULL && mii.hSubMenu == g_projectBuildSubMenu;
			if (isLinkerSubMenu || isProjectBuildSubMenu || mii.wID == IDM_AUTOLINKER_UNPACK_SOURCE) {
				removeThis = true;
				preserveSubMenu = isLinkerSubMenu || isProjectBuildSubMenu;
			}
		}
		if (!removeThis) {
			std::wstring title = GetMenuTitleW(hTargetMenu, static_cast<UINT>(i), MF_BYPOSITION);
			if (title.find(L"链接器切换") != std::wstring::npos ||
				title.find(L"使用的链接器") != std::wstring::npos ||
				title.find(L"反编译到目录") != std::wstring::npos ||
				title.find(L"按配置编译") != std::wstring::npos) {
				removeThis = true;
				// 标题回退路径用于兼容 IDE 重建菜单后的旧句柄；当前句柄仍然有效时
				// 必须只摘除菜单项本身，不能 DeleteMenu 销毁配置子菜单。
				preserveSubMenu =
					(g_topLinkerSubMenu != NULL && mii.hSubMenu == g_topLinkerSubMenu) ||
					(g_projectBuildSubMenu != NULL && mii.hSubMenu == g_projectBuildSubMenu);
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
	// WM_INITMENUPOPUP 和 TrackPopupMenu 都可能触发刷新。菜单项已经挂载时只
	// 更新子菜单内容，避免反复摘除/重挂载时误伤 IDE 自己的菜单项。
	if (HasSubMenuItem(hTargetMenu, g_topLinkerSubMenu) &&
		HasSubMenuItem(hTargetMenu, g_projectBuildSubMenu) &&
		GetMenuState(hTargetMenu, IDM_AUTOLINKER_UNPACK_SOURCE, MF_BYCOMMAND) != 0xFFFFFFFF) {
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
	if (EnsureProjectBuildSubMenu()) {
		AppendMenuW(hTargetMenu, MF_POPUP | MF_STRING, reinterpret_cast<UINT_PTR>(g_projectBuildSubMenu), L"按配置编译");
	}
}

void RebuildProjectBuildSubMenu()
{
	if (!EnsureProjectBuildSubMenu()) return;
	ClearMenuItemsByPosition(g_projectBuildSubMenu);
	g_projectBuildCommandMap.clear();
	UpdateCurrentOpenSourceFile();
	if (g_nowOpenSourceFilePath.empty()) {
		AppendMenuW(g_projectBuildSubMenu, MF_STRING | MF_GRAYED, 0, L"（未打开源文件）");
		return;
	}
	const std::filesystem::path sourcePath(g_nowOpenSourceFilePath);
	std::string error;
	ProjectBuildConfigFile file = LoadProjectBuildConfigFile(sourcePath, &error);
	if (!error.empty()) {
		AppendMenuW(g_projectBuildSubMenu, MF_STRING | MF_GRAYED, 0, L"（配置文件无效）");
		return;
	}
	if (file.configurations.empty()) {
		AppendMenuW(g_projectBuildSubMenu, MF_STRING | MF_GRAYED, 0, L"目前没有编译配置");
		return;
	}
	UINT command = IDM_AUTOLINKER_PROJECT_BUILD_BASE;
	for (const auto& config : file.configurations) {
		if (command > IDM_AUTOLINKER_PROJECT_BUILD_MAX) break;
		const std::wstring title = Utf8ToWideMenuText(config.name);
		AppendMenuW(g_projectBuildSubMenu, MF_STRING | MF_ENABLED, command, title.c_str());
		g_projectBuildCommandMap[command] = config.name;
		++command;
	}
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

bool HandleProjectBuildMenuCommand(UINT cmd)
{
	auto it = g_projectBuildCommandMap.find(cmd);
	UpdateCurrentOpenSourceFile();
	if (g_nowOpenSourceFilePath.empty()) return true;
	const std::filesystem::path sourcePath(g_nowOpenSourceFilePath);
	std::string error;
	const ProjectBuildConfigFile file = LoadProjectBuildConfigFile(sourcePath, &error);
	if (!error.empty()) {
		MessageBoxW(g_hwnd, Utf8ToWideMenuText(error).c_str(), L"项目配置", MB_OK | MB_ICONERROR);
		return true;
	}
	std::string selectedName;
	if (it != g_projectBuildCommandMap.end()) {
		selectedName = it->second;
	}
	else if (cmd >= IDM_AUTOLINKER_PROJECT_BUILD_BASE && cmd <= IDM_AUTOLINKER_PROJECT_BUILD_MAX) {
		// 菜单由 IDE 延迟创建时，WM_COMMAND 可能先于 WM_INITMENUPOPUP 到达；
		// 仍按稳定的命令区间解析配置，避免点击后落回 IDE 原命令处理器。
		const size_t index = static_cast<size_t>(cmd - IDM_AUTOLINKER_PROJECT_BUILD_BASE);
		if (index < file.configurations.size()) selectedName = file.configurations[index].name;
	}
	if (selectedName.empty()) return false;
	const auto configIt = std::find_if(file.configurations.begin(), file.configurations.end(), [&](const ProjectBuildConfig& config) { return config.name == selectedName; });
	if (configIt == file.configurations.end()) return true;
	ProjectBuildConfigFile updated = file;
	updated.activeConfiguration = configIt->name;
	SaveProjectBuildConfigFile(sourcePath, updated, nullptr);
	OutputProjectBuildUtf8(ProjectBuildUtf8(u8"[项目配置] 开始执行配置：") + configIt->name);
	const ProjectBuildPipelineResult result = StartProjectBuildPipelineAsync(sourcePath, *configIt);
	if (result.ok) {
		OutputProjectBuildUtf8(ProjectBuildUtf8(u8"[项目配置] 异步编译任务已提交，输出目标：") + result.outputPath);
	}
	else {
		OutputProjectBuildUtf8(ProjectBuildUtf8(u8"[项目配置] 执行失败：") + result.stage + " - " + result.message);
		MessageBoxW(g_hwnd, Utf8ToWideMenuText(result.message).c_str(), L"按配置编译失败", MB_OK | MB_ICONERROR);
	}
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
	if (hMenu == g_projectBuildSubMenu) {
		RebuildProjectBuildSubMenu();
		return;
	}
	if (IsCompileOrToolsTopPopup(hMenu)) {
		UpdateCurrentOpenSourceFile();
		EnsureTopLinkerSubMenuAttached(hMenu);
		RebuildTopLinkerSubMenu();
		RebuildProjectBuildSubMenu();
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

	if (hMenu == g_topLinkerSubMenu || hMenu == g_projectBuildSubMenu) {
		HandleInitMenuPopup(hMenu);
		return;
	}
	if (IsCompileOrToolsTopPopup(hMenu)) {
		// TrackPopupMenu 钩子发生在 IDE 的 WM_INITMENUPOPUP 处理之前。菜单对象
		// 会被 IDE 复用，因此这里只摘除上一次追加的 AutoLinker 项，让 IDE 始终
		// 基于原始菜单结构更新“编译”“静态编译”等命令。新的扩展项统一留给
		// 消息处理完成后的 FinalizeAutoLinkerPopupMenu。
		RemoveExistingTopMenuExtensions(hMenu);
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

	const bool isCompilePopup = IsCompileOrToolsTopPopup(hMenu);
	if (hMenu == g_topLinkerSubMenu || hMenu == g_projectBuildSubMenu || isCompilePopup) {
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
