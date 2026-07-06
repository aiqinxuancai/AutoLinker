#include "AIChatTooling.h"
#include "AIChatToolingInternal.h"

#include <Windows.h>

#include <algorithm>
#include <cctype>
#include <chrono>
#include <CommCtrl.h>
#include <cstring>
#include <ctime>
#include <filesystem>
#include <fstream>
#include <iterator>
#include <string>
#include <sys/stat.h>
#include <unordered_set>
#include <vector>

#include "..\\thirdparty\\json.hpp"
#include "AIService.h"
#include "AIChatFeature.h"
#include "ConfigManager.h"
#include "DependencyCatalogCache.h"
#include "IDEFacade.h"
#include "EideInternalTextBridge.h"
#include "Global.h"
#include "LocalMcpServer.h"
#include "Logger.h"
#include "PageCodeCacheManager.h"
#include "PathHelper.h"
#include "RealPageCodeToolSupport.h"
#include "WorkspaceFileTools.h"
#include "WorkspaceMirror.h"
#include "WindowHelper.h"
#if defined(_M_IX86)

using ToolPerfClock = std::chrono::steady_clock;

static long long ElapsedToolMs(const ToolPerfClock::time_point& start)
{
	return static_cast<long long>(
		std::chrono::duration_cast<std::chrono::milliseconds>(ToolPerfClock::now() - start).count());
}

static void LogToolStageForAI(
	const std::string& scope,
	const std::string& stage,
	const ToolPerfClock::time_point& start)
{
	Logger::Instance().Write(
		"Tool",
		"[" + scope + "] " + stage + "|elapsed_ms=" + std::to_string(ElapsedToolMs(start)));
}

std::filesystem::path GetAvailableModuleDirectoryForAI();

struct ProgramTreeItemInfo {
	int depth = 0;
	std::string name;
	unsigned int itemData = 0;
	int image = -1;
	int selectedImage = -1;
	std::string typeKey;
	std::string typeName;
};

struct SourceEditBaseForAI {
	std::string baseCode;
	std::string currentCode;
	std::string trace;
	std::string sourceMode;
	std::string codeKind;
	bool cacheRefreshed = false;
	bool mirrorMode = false;
	std::uintptr_t editorObject = 0;
};

constexpr const char* kProgramTreeConstantTablePageNameForAI = "常量表...";
constexpr const char* kProgramTreeUserDataTypePageNameForAI = "自定义数据类型";
constexpr const char* kProgramTreeDllCommandPageNameForAI = "Dll命令";
constexpr const char* kProgramTreeGlobalVariablePageNameForAI = "全局变量";

std::vector<std::string> SplitLinesCopyForAI(const std::string& text);
std::string NormalizeLineBreaksForAI(std::string text);
std::string NormalizeRealPageCodeForLooseCompareForAI(const std::string& text);
std::string NormalizeRealPageCodeLineForStructuralCompareForAI(const std::string& line);
std::vector<std::string> BuildRealPageStructuralFingerprintForAI(const std::string& text);
bool VerifyRealPageCodeMatchesForAI(
	const std::string& expectedCode,
	const std::string& actualCode,
	std::string* outMode);
void AppendFullMirrorRefreshResult(nlohmann::json& r);
void PutPageCodeCacheEntryForAI(
	const ProgramTreeItemInfo& item,
	const std::string& code,
	const PageCodeCacheEntry* existingEntry,
	PageCodeCacheEntry* outSavedEntry);

AISourceEditMode GetActiveSourceEditModeForAI()
{
	AIJsonConfig* aiJsonConfig = GetAIChatAIJsonConfigForTooling();
	ConfigManager* configManager = GetAIChatConfigManagerForTooling();
	AISettings settings = {};
	if (aiJsonConfig != nullptr && AIService::LoadSettings(*aiJsonConfig, configManager, settings)) {
		return settings.sourceEditMode;
	}
	return AISourceEditMode::RealPageFirst;
}

std::string SourceEditModeStringForAI(AISourceEditMode mode)
{
	return AIService::SourceEditModeToString(mode);
}

std::vector<std::string> SplitLinesForNumberedViewForAI(const std::string& text)
{
	std::vector<std::string> lines;
	std::string current;
	for (size_t i = 0; i < text.size(); ++i) {
		const char ch = text[i];
		if (ch == '\r') {
			if (i + 1 < text.size() && text[i + 1] == '\n') {
				++i;
			}
			lines.push_back(current);
			current.clear();
		}
		else if (ch == '\n') {
			lines.push_back(current);
			current.clear();
		}
		else {
			current.push_back(ch);
		}
	}
	if (!current.empty() || text.empty() || text.back() == '\n' || text.back() == '\r') {
		lines.push_back(current);
	}
	return lines;
}

int GetJsonIntArgumentForAI(const nlohmann::json& args, const char* key, int defaultValue)
{
	if (!args.contains(key) || !args[key].is_number_integer()) {
		return defaultValue;
	}
	return args[key].get<int>();
}

std::string BuildNumberedViewForAI(
	const std::vector<std::string>& lines,
	int offset,
	int limit,
	int& outReturned,
	bool& outTruncated)
{
	outReturned = 0;
	outTruncated = false;
	const int total = static_cast<int>(lines.size());
	const int start = (std::clamp)(offset, 0, total);
	const int maxLines = limit <= 0 ? 20000 : (std::clamp)(limit, 1, 20000);
	const int end = (std::min)(total, start + maxLines);
	outTruncated = end < total;

	std::string view;
	for (int i = start; i < end; ++i) {
		view += std::to_string(i + 1);
		view += "\t";
		view += lines[static_cast<size_t>(i)];
		view += "\n";
		++outReturned;
	}
	return view;
}

static std::string TrimAsciiCopy(const std::string& text)
{
	size_t begin = 0;
	size_t end = text.size();
	while (begin < end && std::isspace(static_cast<unsigned char>(text[begin])) != 0) {
		++begin;
	}
	while (end > begin && std::isspace(static_cast<unsigned char>(text[end - 1])) != 0) {
		--end;
	}
	return text.substr(begin, end - begin);
}

bool IsLikelyClassModulePageNameForAI(const std::string& text)
{
	const std::string trimmed = TrimAsciiCopy(text);
	return trimmed.rfind("Class_", 0) == 0 || trimmed.rfind("类", 0) == 0;
}

std::string ToLowerAsciiCopyLocal(std::string text)
{
	for (char& ch : text) {
		ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
	}
	return text;
}

std::string LocalFromWide(const wchar_t* text)
{
	if (text == nullptr || *text == L'\0') {
		return std::string();
	}
	const int size = WideCharToMultiByte(CP_ACP, 0, text, -1, nullptr, 0, nullptr, nullptr);
	if (size <= 0) {
		return std::string();
	}
	std::string out(static_cast<size_t>(size), '\0');
	if (WideCharToMultiByte(CP_ACP, 0, text, -1, out.data(), size, nullptr, nullptr) <= 0) {
		return std::string();
	}
	if (!out.empty() && out.back() == '\0') {
		out.pop_back();
	}
	return out;
}

bool IsValidUtf8Text(const std::string& text)
{
	if (text.empty()) {
		return true;
	}
	return MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0) > 0;
}

std::string ConvertCodePage(const std::string& text, UINT fromCodePage, UINT toCodePage, DWORD fromFlags = 0)
{
	if (text.empty()) {
		return std::string();
	}

	const int wideLen = MultiByteToWideChar(
		fromCodePage,
		fromFlags,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0);
	if (wideLen <= 0) {
		return text;
	}

	std::wstring wide(static_cast<size_t>(wideLen), L'\0');
	if (MultiByteToWideChar(
		fromCodePage,
		fromFlags,
		text.data(),
		static_cast<int>(text.size()),
		wide.data(),
		wideLen) <= 0) {
		return text;
	}

	const int outLen = WideCharToMultiByte(
		toCodePage,
		0,
		wide.data(),
		wideLen,
		nullptr,
		0,
		nullptr,
		nullptr);
	if (outLen <= 0) {
		return text;
	}

	std::string out(static_cast<size_t>(outLen), '\0');
	if (WideCharToMultiByte(
		toCodePage,
		0,
		wide.data(),
		wideLen,
		out.data(),
		outLen,
		nullptr,
		nullptr) <= 0) {
		return text;
	}
	return out;
}

std::string LocalToUtf8Text(const std::string& text)
{
	if (text.empty()) {
		return std::string();
	}
	if (IsValidUtf8Text(text)) {
		return text;
	}
	return ConvertCodePage(text, CP_ACP, CP_UTF8, 0);
}

std::string Utf8ToLocalText(const std::string& text)
{
	if (text.empty()) {
		return std::string();
	}
	if (!IsValidUtf8Text(text)) {
		return text;
	}
	return ConvertCodePage(text, CP_UTF8, CP_ACP, MB_ERR_INVALID_CHARS);
}

void NormalizeJsonStringsToUtf8ForAI(nlohmann::json& value)
{
	if (value.is_string()) {
		value = LocalToUtf8Text(value.get<std::string>());
		return;
	}
	if (value.is_array()) {
		for (auto& item : value) {
			NormalizeJsonStringsToUtf8ForAI(item);
		}
		return;
	}
	if (value.is_object()) {
		for (auto& item : value.items()) {
			NormalizeJsonStringsToUtf8ForAI(item.value());
		}
	}
}

std::string JsonToLocalTextForAI(nlohmann::json value)
{
	// 工具结果可能混入 IDE/MFC 返回的本地编码字符串，统一转 UTF-8 后再序列化。
	NormalizeJsonStringsToUtf8ForAI(value);
	return Utf8ToLocalText(value.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace));
}

std::string JsonUtf8ToAsciiTextForAI(const nlohmann::json& value)
{
	return value.dump(-1, ' ', true, nlohmann::json::error_handler_t::replace);
}

std::string GetCurrentProcessPathForAI()
{
	char buffer[MAX_PATH] = {};
	const DWORD len = GetModuleFileNameA(nullptr, buffer, static_cast<DWORD>(sizeof(buffer)));
	if (len == 0 || len >= sizeof(buffer)) {
		return std::string();
	}
	return std::string(buffer, buffer + len);
}

std::string GetCurrentProcessNameForAI()
{
	const std::string fullPath = GetCurrentProcessPathForAI();
	if (fullPath.empty()) {
		return std::string();
	}
	const size_t pos = fullPath.find_last_of("\\/");
	return pos == std::string::npos ? fullPath : fullPath.substr(pos + 1);
}

std::string RefreshCurrentSourceFilePathForAI()
{
	g_nowOpenSourceFilePath = GetSourceFilePath();
	return g_nowOpenSourceFilePath;
}

std::uintptr_t GetCurrentProcessImageBaseForAI()
{
	HMODULE module = GetModuleHandleW(nullptr);
	return reinterpret_cast<std::uintptr_t>(module);
}

std::vector<HWND> CollectTreeViewWindowsForAI(HWND root)
{
	std::vector<HWND> windows;
	if (root == nullptr || !IsWindow(root)) {
		return windows;
	}

	EnumChildWindows(
		root,
		[](HWND hWnd, LPARAM lParam) -> BOOL {
			auto* out = reinterpret_cast<std::vector<HWND>*>(lParam);
			if (out == nullptr) {
				return TRUE;
			}
			char className[64] = {};
			if (GetClassNameA(hWnd, className, static_cast<int>(sizeof(className))) > 0 &&
				std::strcmp(className, WC_TREEVIEWA) == 0) {
				out->push_back(hWnd);
			}
			return TRUE;
		},
		reinterpret_cast<LPARAM>(&windows));
	return windows;
}

HTREEITEM GetTreeNextItemForAI(HWND treeHwnd, HTREEITEM item, UINT code)
{
	return reinterpret_cast<HTREEITEM>(SendMessageA(treeHwnd, TVM_GETNEXTITEM, code, reinterpret_cast<LPARAM>(item)));
}

bool QueryTreeItemInfoForAI(
	HWND treeHwnd,
	HTREEITEM item,
	std::string& outText,
	LPARAM& outParam,
	int& outImage,
	int& outSelectedImage,
	int& outChildren)
{
	outText.clear();
	outParam = 0;
	outImage = -1;
	outSelectedImage = -1;
	outChildren = 0;
	if (treeHwnd == nullptr || item == nullptr) {
		return false;
	}

	char textBuf[512] = {};
	TVITEMA tvItem = {};
	tvItem.mask = TVIF_HANDLE | TVIF_TEXT | TVIF_PARAM | TVIF_CHILDREN | TVIF_IMAGE | TVIF_SELECTEDIMAGE;
	tvItem.hItem = item;
	tvItem.pszText = textBuf;
	tvItem.cchTextMax = static_cast<int>(sizeof(textBuf));
	if (SendMessageA(treeHwnd, TVM_GETITEMA, 0, reinterpret_cast<LPARAM>(&tvItem)) == FALSE) {
		return false;
	}

	outText = textBuf;
	outParam = tvItem.lParam;
	outImage = tvItem.iImage;
	outSelectedImage = tvItem.iSelectedImage;
	outChildren = tvItem.cChildren;
	return true;
}

HWND FindProgramDataTreeViewForAI()
{
	for (HWND treeHwnd : CollectTreeViewWindowsForAI(GetAIChatMainWindowForTooling())) {
		const HTREEITEM rootItem = GetTreeNextItemForAI(treeHwnd, nullptr, TVGN_ROOT);
		std::string text;
		LPARAM itemData = 0;
		int image = -1;
		int selectedImage = -1;
		int childCount = 0;
		if (!QueryTreeItemInfoForAI(treeHwnd, rootItem, text, itemData, image, selectedImage, childCount)) {
			continue;
		}
		if (text == "程序数据") {
			return treeHwnd;
		}
	}
	return nullptr;
}

std::string GetProgramTreeTypeKey(
	unsigned int itemData,
	const std::string& text,
	int image,
	int classImage)
{
	const unsigned int typeNibble = itemData >> 28;
	switch (typeNibble) {
	case 1:
		if (IsLikelyClassModulePageNameForAI(text)) {
			return "class_module";
		}
		if (classImage >= 0 && image == classImage) {
			return "class_module";
		}
		return "assembly";
	case 2:
		return "global_var";
	case 3:
		return "user_data_type";
	case 4:
		return "dll_command";
	case 5:
		return "form";
	case 6:
		return "const_resource";
	case 7:
		return ((itemData & 0x0FFFFFFFu) == 1u) ? "picture_resource" : "sound_resource";
	default:
		return "unknown";
	}
}

std::string GetProgramTreeTypeName(const std::string& typeKey)
{
	if (typeKey == "assembly") {
		return "程序集";
	}
	if (typeKey == "class_module") {
		return "类模块";
	}
	if (typeKey == "global_var") {
		return "全局变量";
	}
	if (typeKey == "user_data_type") {
		return "自定义数据类型";
	}
	if (typeKey == "dll_command") {
		return "DLL命令";
	}
	if (typeKey == "form") {
		return "窗口/表单";
	}
	if (typeKey == "const_resource") {
		return "常量资源";
	}
	if (typeKey == "picture_resource") {
		return "图片资源";
	}
	if (typeKey == "sound_resource") {
		return "声音资源";
	}
	return "未知";
}

std::string NormalizeProgramPageNameForAI(const std::string& name)
{
	return TrimAsciiCopy(name);
}

HTREEITEM FindTreeItemByExactTextRecursiveForAI(
	HWND treeHwnd,
	HTREEITEM firstItem,
	const std::string& targetText,
	int maxDepth,
	int depth)
{
	if (treeHwnd == nullptr || firstItem == nullptr || depth > maxDepth) {
		return nullptr;
	}

	for (HTREEITEM item = firstItem; item != nullptr; item = GetTreeNextItemForAI(treeHwnd, item, TVGN_NEXT)) {
		std::string text;
		LPARAM itemData = 0;
		int image = -1;
		int selectedImage = -1;
		int childCount = 0;
		if (!QueryTreeItemInfoForAI(treeHwnd, item, text, itemData, image, selectedImage, childCount)) {
			continue;
		}
		if (text == targetText) {
			return item;
		}
		if (depth < maxDepth) {
			if (HTREEITEM found = FindTreeItemByExactTextRecursiveForAI(
					treeHwnd,
					GetTreeNextItemForAI(treeHwnd, item, TVGN_CHILD),
					targetText,
					maxDepth,
					depth + 1);
				found != nullptr) {
				return found;
			}
		}
	}

	return nullptr;
}

bool TryFindProgramTreeItemByExactNameForAI(
	const std::string& targetName,
	const std::string& typeKey,
	ProgramTreeItemInfo& outItem,
	std::string* outError = nullptr)
{
	outItem = {};
	if (outError != nullptr) {
		outError->clear();
	}

	const HWND treeHwnd = FindProgramDataTreeViewForAI();
	if (treeHwnd == nullptr) {
		if (outError != nullptr) {
			*outError = "program tree not found";
		}
		return false;
	}

	const HTREEITEM rootItem = GetTreeNextItemForAI(treeHwnd, nullptr, TVGN_ROOT);
	const HTREEITEM item = FindTreeItemByExactTextRecursiveForAI(treeHwnd, rootItem, targetName, 4, 0);
	if (item == nullptr) {
		if (outError != nullptr) {
			*outError = "tree item not found";
		}
		return false;
	}

	LPARAM itemData = 0;
	int image = -1;
	int selectedImage = -1;
	int childCount = 0;
	if (!QueryTreeItemInfoForAI(
			treeHwnd,
			item,
			outItem.name,
			itemData,
			image,
			selectedImage,
			childCount)) {
		if (outError != nullptr) {
			*outError = "read tree item info failed";
		}
		return false;
	}

	outItem.itemData = static_cast<unsigned int>(itemData);
	outItem.image = image;
	outItem.selectedImage = selectedImage;
	outItem.typeKey = typeKey;
	outItem.typeName = GetProgramTreeTypeName(outItem.typeKey);
	return true;
}

void AppendSpecialProgramTreeItemsForAI(std::vector<ProgramTreeItemInfo>& outItems)
{
	std::unordered_set<unsigned int> seenItemData;
	seenItemData.reserve(outItems.size() + 4);
	for (const auto& item : outItems) {
		seenItemData.insert(item.itemData);
	}

	struct SpecialItemSpec {
		const char* name;
		const char* typeKey;
	};
	static constexpr SpecialItemSpec kSpecialItems[] = {
		{ kProgramTreeGlobalVariablePageNameForAI, "global_var" },
		{ kProgramTreeUserDataTypePageNameForAI, "user_data_type" },
		{ kProgramTreeDllCommandPageNameForAI, "dll_command" },
		{ kProgramTreeConstantTablePageNameForAI, "const_resource" },
	};

	for (const auto& spec : kSpecialItems) {
		ProgramTreeItemInfo item;
		if (TryFindProgramTreeItemByExactNameForAI(spec.name, spec.typeKey, item) &&
			seenItemData.insert(item.itemData).second) {
			outItems.push_back(std::move(item));
		}
	}
}

void CollectProgramTreeItemsRecursiveForAI(
	HWND treeHwnd,
	HTREEITEM firstItem,
	int depth,
	int maxDepth,
	std::vector<ProgramTreeItemInfo>& outItems)
{
	if (treeHwnd == nullptr || firstItem == nullptr || depth > maxDepth) {
		return;
	}

	for (HTREEITEM item = firstItem; item != nullptr; item = GetTreeNextItemForAI(treeHwnd, item, TVGN_NEXT)) {
		std::string text;
		LPARAM itemData = 0;
		int image = -1;
		int selectedImage = -1;
		int childCount = 0;
		if (!QueryTreeItemInfoForAI(treeHwnd, item, text, itemData, image, selectedImage, childCount)) {
			continue;
		}

		const unsigned int itemDataU = static_cast<unsigned int>(itemData);
		const unsigned int typeNibble = itemDataU >> 28;
		if (typeNibble != 0 && typeNibble != 15) {
			ProgramTreeItemInfo info;
			info.depth = depth;
			info.name = text;
			info.itemData = itemDataU;
			info.image = image;
			info.selectedImage = selectedImage;
			outItems.push_back(std::move(info));
			continue;
		}

		if (depth < maxDepth) {
			CollectProgramTreeItemsRecursiveForAI(
				treeHwnd,
				GetTreeNextItemForAI(treeHwnd, item, TVGN_CHILD),
				depth + 1,
				maxDepth,
				outItems);
		}
	}
}

bool TryListProgramTreeItemsForAI(std::vector<ProgramTreeItemInfo>& outItems, std::string* outError)
{
	outItems.clear();
	if (outError != nullptr) {
		outError->clear();
	}
	const HWND mainWindow = GetAIChatMainWindowForTooling();
	if (mainWindow == nullptr || !IsWindow(mainWindow)) {
		if (outError != nullptr) {
			*outError = "main window invalid";
		}
		return false;
	}

	const HWND treeHwnd = FindProgramDataTreeViewForAI();
	if (treeHwnd == nullptr) {
		if (outError != nullptr) {
			*outError = "program tree not found";
		}
		return false;
	}

	const HTREEITEM rootItem = GetTreeNextItemForAI(treeHwnd, nullptr, TVGN_ROOT);
	const HTREEITEM firstChild = GetTreeNextItemForAI(treeHwnd, rootItem, TVGN_CHILD);
	CollectProgramTreeItemsRecursiveForAI(treeHwnd, firstChild, 0, 8, outItems);

	int classImage = -1;
	for (const auto& item : outItems) {
		if (IsLikelyClassModulePageNameForAI(item.name) && item.image >= 0) {
			classImage = item.image;
			break;
		}
	}
	for (auto& item : outItems) {
		item.typeKey = GetProgramTreeTypeKey(item.itemData, item.name, item.image, classImage);
		item.typeName = GetProgramTreeTypeName(item.typeKey);
	}
	AppendSpecialProgramTreeItemsForAI(outItems);
	return true;
}

bool MatchProgramItemKind(const ProgramTreeItemInfo& item, const std::string& kindFilter)
{
	const std::string kind = ToLowerAsciiCopyLocal(TrimAsciiCopy(kindFilter));
	if (kind.empty() || kind == "all") {
		return true;
	}
	return item.typeKey == kind;
}

bool TryGetProgramItemByNameForAI(
	const std::string& name,
	const std::string& kindFilter,
	ProgramTreeItemInfo& outItem,
	std::string& outError)
{
	outItem = {};
	outError.clear();
	const std::string normalizedName = NormalizeProgramPageNameForAI(name);

	std::vector<ProgramTreeItemInfo> items;
	if (!TryListProgramTreeItemsForAI(items, &outError)) {
		return false;
	}

	std::vector<ProgramTreeItemInfo> matched;
	for (const auto& item : items) {
		if (item.name == normalizedName && MatchProgramItemKind(item, kindFilter)) {
			matched.push_back(item);
		}
	}
	if (matched.empty()) {
		outError = "program item not found";
		return false;
	}
	if (matched.size() > 1) {
		outError = "program item name is ambiguous";
		return false;
	}

	outItem = matched.front();
	return true;
}

size_t CountExactOccurrencesForAI(const std::string& text, const std::string& needle)
{
	if (needle.empty()) {
		return 0;
	}

	size_t count = 0;
	size_t pos = 0;
	while ((pos = text.find(needle, pos)) != std::string::npos) {
		++count;
		pos += needle.size();
	}
	return count;
}

bool ReplaceExactlyOnceForAI(
	const std::string& source,
	const std::string& oldText,
	const std::string& newText,
	std::string& outResult,
	size_t& outMatchCount)
{
	outResult.clear();
	const std::string normalizedOldText = NormalizeRealCodeLineBreaksToCrLf(oldText);
	const std::string normalizedNewText = NormalizeRealCodeLineBreaksToCrLf(newText);
	outMatchCount = CountExactOccurrencesForAI(source, normalizedOldText);
	if (outMatchCount != 1) {
		return false;
	}

	const size_t pos = source.find(normalizedOldText);
	outResult.reserve(source.size() - normalizedOldText.size() + normalizedNewText.size());
	outResult.append(source.substr(0, pos));
	outResult.append(normalizedNewText);
	outResult.append(source.substr(pos + normalizedOldText.size()));
	return true;
}

std::string GetJsonStringArgumentLocal(const nlohmann::json& args, const char* key)
{
	if (!args.contains(key) || !args[key].is_string()) {
		return std::string();
	}
	return Utf8ToLocalText(args[key].get<std::string>());
}

bool GetJsonBoolArgument(const nlohmann::json& args, const char* key, bool defaultValue)
{
	if (!args.contains(key) || !args[key].is_boolean()) {
		return defaultValue;
	}
	return args[key].get<bool>();
}

PageCodeCacheEntry BuildPageCodeCacheEntryForAI(
	const ProgramTreeItemInfo& item,
	const std::string& code,
	const PageCodeCacheEntry* existingEntry)
{
	PageCodeCacheEntry entry = existingEntry != nullptr ? *existingEntry : PageCodeCacheEntry{};
	entry.pageName = item.name;
	entry.kind = item.typeKey;
	entry.pageTypeKey = item.typeKey;
	entry.pageTypeName = item.typeName;
	entry.itemData = item.itemData;
	entry.code = NormalizeRealCodeLineBreaksToCrLf(code);
	entry.codeHash = BuildStableTextHashForRealCode(entry.code);
	return entry;
}

void PutPageCodeCacheEntryForAI(
	const ProgramTreeItemInfo& item,
	const std::string& code,
	const PageCodeCacheEntry* existingEntry,
	PageCodeCacheEntry* outSavedEntry = nullptr)
{
	PageCodeCacheEntry entry = BuildPageCodeCacheEntryForAI(item, code, existingEntry);
	PageCodeCacheManager::Instance().Put(entry);
	if (outSavedEntry != nullptr) {
		*outSavedEntry = entry;
	}
}

bool IsFixedTableProgramItemForAI(const ProgramTreeItemInfo& item)
{
	return item.typeKey == "global_var" ||
		item.typeKey == "user_data_type" ||
		item.typeKey == "dll_command" ||
		item.typeKey == "const_resource";
}

std::string JoinLinesWithCrLfForAI(const std::vector<std::string>& lines)
{
	std::string text;
	for (size_t i = 0; i < lines.size(); ++i) {
		if (i != 0) {
			text += "\r\n";
		}
		text += lines[i];
	}
	return text;
}

std::string ExtractConstResourceNameForAI(const std::string& line)
{
	const std::string trimmed = TrimAsciiCopy(line);
	const std::string directive = ".常量";
	if (trimmed.rfind(directive, 0) != 0) {
		return std::string();
	}

	std::string remain = TrimAsciiCopy(trimmed.substr(directive.size()));
	const size_t comma = remain.find(',');
	if (comma != std::string::npos) {
		remain = remain.substr(0, comma);
	}
	return TrimAsciiCopy(remain);
}

bool TryReadWorkspaceMirrorTextLocalForAI(
	const std::string& filePathUtf8,
	std::string& outTextLocal,
	std::string& outError)
{
	outTextLocal.clear();
	outError.clear();

	std::filesystem::path fullPath;
	std::string relativePath;
	if (!WorkspaceMirror::ResolveFilePath(filePathUtf8, fullPath, relativePath, outError)) {
		return false;
	}

	try {
		std::ifstream in(fullPath, std::ios::binary);
		if (!in.is_open()) {
			outError = "open workspace mirror file failed: " + relativePath;
			return false;
		}
		std::string bytes((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
		if (bytes.size() >= 3 &&
			static_cast<unsigned char>(bytes[0]) == 0xEF &&
			static_cast<unsigned char>(bytes[1]) == 0xBB &&
			static_cast<unsigned char>(bytes[2]) == 0xBF) {
			bytes.erase(0, 3);
		}
		outTextLocal = NormalizeRealCodeLineBreaksToCrLf(Utf8ToLocalText(bytes));
		return true;
	}
	catch (const std::exception& ex) {
		outError = std::string("read workspace mirror file exception: ") + ex.what();
		return false;
	}
}

std::string MergeConstResourceLongTextPlaceholdersForAI(
	const std::string& targetCode,
	const std::string& fullConstCode)
{
	std::unordered_map<std::string, std::string> fullLineByName;
	for (const std::string& line : SplitLinesCopyForAI(NormalizeLineBreaksForAI(fullConstCode))) {
		if (line.find("<文本长度:") != std::string::npos) {
			continue;
		}
		const std::string name = ExtractConstResourceNameForAI(line);
		if (!name.empty()) {
			fullLineByName[name] = line;
		}
	}
	if (fullLineByName.empty()) {
		return targetCode;
	}

	std::vector<std::string> targetLines = SplitLinesCopyForAI(NormalizeLineBreaksForAI(targetCode));
	bool changed = false;
	for (std::string& line : targetLines) {
		if (line.find("<文本长度:") == std::string::npos) {
			continue;
		}
		const std::string name = ExtractConstResourceNameForAI(line);
		const auto it = fullLineByName.find(name);
		if (it != fullLineByName.end()) {
			line = it->second;
			changed = true;
		}
	}
	return changed ? JoinLinesWithCrLfForAI(targetLines) : targetCode;
}

bool OpenFixedTablePageForAI(const ProgramTreeItemInfo& item, std::string& outTrace)
{
	outTrace.clear();

	IDEFacade::ViewTab tab = IDEFacade::ViewTab::ConstResource;
	if (item.typeKey == "global_var") {
		tab = IDEFacade::ViewTab::GlobalVar;
	}
	else if (item.typeKey == "user_data_type") {
		tab = IDEFacade::ViewTab::DataType;
	}
	else if (item.typeKey == "dll_command") {
		tab = IDEFacade::ViewTab::DllCommand;
	}
	else if (item.typeKey == "const_resource") {
		tab = IDEFacade::ViewTab::ConstResource;
	}
	else {
		outTrace = "not_fixed_table";
		return false;
	}

	OutputStringToELog("[AutoLinker][FixedTable] open fixed table by tab: " + item.typeKey);
	if (IDEFacade::Instance().OpenViewTab(tab)) {
		outTrace = "open_fixed_table_tab_ok|" + item.typeKey;
		return true;
	}

	outTrace = "open_fixed_table_tab_failed|" + item.typeKey;
	return false;
}

bool TryGetCurrentPageCodeLocalForAI(
	std::string& outCode,
	std::string& outTrace,
	int maxAttempts = 10,
	DWORD waitMs = 50)
{
	outCode.clear();
	outTrace.clear();
	const auto totalStart = ToolPerfClock::now();
	for (int attempt = 0; attempt < maxAttempts; ++attempt) {
		const auto attemptStart = ToolPerfClock::now();
		std::string pageCodeUtf8;
		if (IDEFacade::Instance().GetCurrentPageCode(pageCodeUtf8)) {
			outCode = NormalizeRealCodeLineBreaksToCrLf(Utf8ToLocalText(pageCodeUtf8));
			outTrace =
				"copy_current_page_ok|attempt=" + std::to_string(attempt + 1) +
				"|attempt_ms=" + std::to_string(ElapsedToolMs(attemptStart)) +
				"|total_ms=" + std::to_string(ElapsedToolMs(totalStart));
			return true;
		}
		Sleep(waitMs);
	}
	outTrace =
		"copy_current_page_failed_after_retries|attempts=" + std::to_string(maxAttempts) +
		"|total_ms=" + std::to_string(ElapsedToolMs(totalStart));
	return false;
}

bool TryReadFixedTableRealPageCodeForAI(
	const ProgramTreeItemInfo& item,
	std::string& outCode,
	std::string& outTrace,
	std::string& outError)
{
	outCode.clear();
	outTrace.clear();
	outError.clear();

	auto copyCurrentPage = [&](const std::string& openTrace) -> bool {
		std::string copyTrace;
		if (!TryGetCurrentPageCodeLocalForAI(outCode, copyTrace, 6, 40)) {
			outTrace = openTrace + "|copy_current_fixed_table_page_failed|" + copyTrace;
			outError = "copy fixed table page code failed";
			return false;
		}
		outTrace = openTrace + "|copy_current_fixed_table_page_ok|" + copyTrace;
		outError.clear();
		return true;
	};

	std::string lastTrace;
	std::string lastError;
	for (int attempt = 0; attempt < 3; ++attempt) {
		std::string openTrace;
		if (OpenFixedTablePageForAI(item, openTrace) && copyCurrentPage(openTrace)) {
			return true;
		}
		lastTrace = outTrace.empty() ? openTrace : outTrace;
		lastError = outError;
		Sleep(30);
	}

	outTrace = lastTrace + "|fixed_table_copy_failed_no_tree_fallback";
	outError = lastError.empty() ? "copy fixed table page code failed" : lastError;
	return false;
}

bool CommitCurrentFixedTablePageForAI(const std::string& pageCode, std::string& outTrace)
{
	outTrace.clear();

	const bool forceOk = IDEFacade::Instance().RunForceProcess();
	const bool setOk = IDEFacade::Instance().RunSetAndCompilePrgItemText(pageCode, false);
	outTrace =
		std::string("force_process=") + (forceOk ? "1" : "0") +
		"|set_and_compile_text=" + (setOk ? "1" : "0");
	Sleep(50);
	return forceOk || setOk;
}

bool TryWriteFixedTableRealPageCodeByPasteForAI(
	const ProgramTreeItemInfo& item,
	const std::string& baseCode,
	const std::string& newCode,
	std::string& outFinalCode,
	std::string& outTrace,
	std::string& outError,
	bool* outRollbackAttempted = nullptr,
	bool* outRollbackSucceeded = nullptr)
{
	outFinalCode.clear();
	outTrace.clear();
	outError.clear();
	if (outRollbackAttempted != nullptr) {
		*outRollbackAttempted = false;
	}
	if (outRollbackSucceeded != nullptr) {
		*outRollbackSucceeded = false;
	}

	const std::string normalizedBaseCode = NormalizeRealCodeLineBreaksToCrLf(baseCode);
	const std::string normalizedNewCode = NormalizeRealCodeLineBreaksToCrLf(newCode);
	const std::string preparedExpectedCode = NormalizeRealCodeLineBreaksToCrLf(
		PrepareAssemblyVariablesForRealPageWrite(normalizedNewCode));

	std::string openTrace;
	if (!OpenFixedTablePageForAI(item, openTrace)) {
		outTrace = openTrace;
		outError = "open fixed table page failed";
		return false;
	}

	auto pasteCurrentPage = [&](const std::string& currentOpenTrace) -> bool {
		const bool pasteReportedOk = IDEFacade::Instance().ReplaceCurrentPageCode(normalizedNewCode, false);

		std::string verifyCode;
		std::string verifyCopyTrace;
		if (TryGetCurrentPageCodeLocalForAI(verifyCode, verifyCopyTrace, 10, 50)) {
			std::string verifyMode;
			if (!VerifyRealPageCodeMatchesForAI(preparedExpectedCode, verifyCode, &verifyMode)) {
				const bool pageChanged =
					BuildStableTextHashForRealCode(verifyCode) !=
					BuildStableTextHashForRealCode(normalizedBaseCode);
				if (pageChanged) {
					outFinalCode = verifyCode;
					std::string commitTrace;
					CommitCurrentFixedTablePageForAI(normalizedNewCode, commitTrace);
					outTrace = currentOpenTrace +
						(pasteReportedOk ? "|paste_current_fixed_table_page_ok" : "|paste_current_fixed_table_page_reported_failed") +
						"|verify_mismatch_but_page_changed" +
						"|commit_current_fixed_table|" + commitTrace;
					outError.clear();
					return true;
				}
				outTrace = currentOpenTrace +
					(pasteReportedOk ? "|paste_current_fixed_table_page_ok" : "|paste_current_fixed_table_page_reported_failed") +
					"|verify_mismatch";
				outError = pasteReportedOk
					? "fixed table paste verification mismatch"
					: "paste fixed table page code failed and verification mismatch";
				outFinalCode = verifyCode;
				return false;
			}
			outFinalCode = verifyCode;
			std::string commitTrace;
			CommitCurrentFixedTablePageForAI(normalizedNewCode, commitTrace);
			outTrace = currentOpenTrace +
				(pasteReportedOk ? "|paste_current_fixed_table_page_ok" : "|paste_current_fixed_table_page_reported_failed") +
				"|verify_ok_" + verifyMode +
				"|" +
				verifyCopyTrace +
				"|commit_current_fixed_table|" + commitTrace;
			outError.clear();
			return true;
		}

		if (pasteReportedOk) {
			outFinalCode = normalizedNewCode;
			std::string commitTrace;
			CommitCurrentFixedTablePageForAI(normalizedNewCode, commitTrace);
			outTrace = currentOpenTrace + "|paste_current_fixed_table_page_ok|verify_copy_failed_final_code=write_target|" + verifyCopyTrace +
				"|commit_current_fixed_table|" + commitTrace;
			outError.clear();
			return true;
		}

		outTrace = currentOpenTrace + "|paste_current_fixed_table_page_failed|verify_copy_failed|" + verifyCopyTrace;
		outError = "paste fixed table page code failed";
		return false;
	};

	std::string lastTrace;
	std::string lastError;
	for (int attempt = 0; attempt < 3; ++attempt) {
		std::string attemptOpenTrace = openTrace;
		if (attempt > 0) {
			if (!OpenFixedTablePageForAI(item, attemptOpenTrace)) {
				lastTrace = attemptOpenTrace;
				lastError = "open fixed table page failed";
				Sleep(30);
				continue;
			}
		}
		if (pasteCurrentPage(attemptOpenTrace + "|attempt=" + std::to_string(attempt + 1))) {
			return true;
		}
		lastTrace = outTrace;
		lastError = outError;
		Sleep(30);
	}

	outTrace = lastTrace + "|fixed_table_paste_failed_after_retries";
	outError = lastError.empty() ? "paste fixed table page code failed" : lastError;
	return false;
}

bool OpenRegularProgramItemPageForRealPageCopyForAI(
	const ProgramTreeItemInfo& item,
	std::string& outTrace,
	std::string& outError)
{
	outTrace.clear();
	outError.clear();

	const auto totalStart = ToolPerfClock::now();
	std::string openTrace;
	const auto openStart = ToolPerfClock::now();
	if (!e571::OpenProgramTreeItemPageByData(item.itemData, &openTrace)) {
		outTrace =
			(openTrace.empty() ? std::string("open_program_item_page_failed") : openTrace) +
			"|open_ms=" + std::to_string(ElapsedToolMs(openStart)) +
			"|total_ms=" + std::to_string(ElapsedToolMs(totalStart));
		outError = "open program item page failed";
		return false;
	}

	std::string currentName;
	std::string currentType;
	std::string currentTrace;
	const auto nameStart = ToolPerfClock::now();
	const bool currentOk = IDEFacade::Instance().GetCurrentPageName(
		currentName,
		&currentType,
		&currentTrace);
	outTrace =
		"open_program_item_page_by_data|" +
		openTrace +
		"|open_ms=" + std::to_string(ElapsedToolMs(openStart)) +
		"|current_page_ok=" + std::to_string(currentOk ? 1 : 0) +
		"|current_page_name=" + currentName +
		"|current_page_type=" + currentType +
		"|current_page_trace=" + currentTrace +
		"|current_page_ms=" + std::to_string(ElapsedToolMs(nameStart)) +
		"|total_ms=" + std::to_string(ElapsedToolMs(totalStart));
	return true;
}

bool TryReadRealPageCodeForAI(
	const ProgramTreeItemInfo& item,
	std::string& outCode,
	e571::NativeRealPageAccessResult& outAccessResult,
	std::string& outError)
{
	outCode.clear();
	outAccessResult = {};
	outError.clear();

	if (IsFixedTableProgramItemForAI(item)) {
		std::string trace;
		if (!TryReadFixedTableRealPageCodeForAI(item, outCode, trace, outError)) {
			outAccessResult.trace = trace;
			return false;
		}
		outAccessResult.ok = true;
		outAccessResult.usedClipboardEmulation = true;
		outAccessResult.capturedCustomFormat = false;
		outAccessResult.textBytes = outCode.size();
		outAccessResult.trace = trace + "|real_page_read=whole_page_copy";
		return true;
	}

	const auto readStart = ToolPerfClock::now();
	const std::uintptr_t moduleBase = reinterpret_cast<std::uintptr_t>(::GetModuleHandleW(nullptr));
	if (moduleBase == 0) {
		outAccessResult.trace = "internal_real_page_read_failed|module_base_invalid";
		outError = "module base invalid";
		return false;
	}

	LogToolStageForAI(
		"real_page_read",
		"before_internal_get|page=" + LocalToUtf8Text(item.name) +
			"|type=" + item.typeKey +
			"|item_data=" + std::to_string(item.itemData),
		readStart);
	if (!e571::GetRealPageCodeByProgramTreeItemData(
			item.itemData,
			moduleBase,
			&outCode,
			&outAccessResult)) {
		LogToolStageForAI(
			"real_page_read",
			"after_internal_get_failed|page=" + LocalToUtf8Text(item.name) +
				"|type=" + item.typeKey +
				"|trace=" + LocalToUtf8Text(outAccessResult.trace),
			readStart);
		outAccessResult.trace =
			"internal_real_page_read_failed|read_ms=" + std::to_string(ElapsedToolMs(readStart)) +
			(outAccessResult.trace.empty() ? std::string() : ("|" + outAccessResult.trace));
		outError = "copy real page code failed";
		return false;
	}

	outCode = NormalizeRealCodeLineBreaksToCrLf(outCode);
	outAccessResult.ok = true;
	outAccessResult.textBytes = outCode.size();
	LogToolStageForAI(
		"real_page_read",
		"after_internal_get_ok|page=" + LocalToUtf8Text(item.name) +
			"|type=" + item.typeKey +
			"|bytes=" + std::to_string(outCode.size()) +
			"|trace=" + LocalToUtf8Text(outAccessResult.trace),
		readStart);
	outAccessResult.trace =
		std::string("internal_real_page_read=whole_page_copy_paste") +
		"|read_ms=" + std::to_string(ElapsedToolMs(readStart)) +
		(outAccessResult.trace.empty() ? std::string() : ("|" + outAccessResult.trace));
	return true;
}

bool TryResolveRealPageWriteBaseForAI(
	const ProgramTreeItemInfo& item,
	const std::string& expectedBaseHash,
	bool requireCache,
	std::string& outBaseCode,
	PageCodeCacheEntry& outCacheEntry,
	std::string& outLiveTrace,
	std::string& outCurrentCode,
	bool& outCacheRefreshed,
	std::string& outError,
	std::uintptr_t* outEditorObject = nullptr)
{
	outBaseCode.clear();
	outCacheEntry = {};
	outLiveTrace.clear();
	outCurrentCode.clear();
	outCacheRefreshed = false;
	outError.clear();
	if (outEditorObject != nullptr) {
		*outEditorObject = 0;
	}

	const bool isFixedTableWriteBase = IsFixedTableProgramItemForAI(item);
	const bool hasCache = PageCodeCacheManager::Instance().Get(item.name, item.typeKey, outCacheEntry);
	if (requireCache && !hasCache && !isFixedTableWriteBase) {
		outError = "page cache missing, call read_file first or retry the edit after the real page cache is refreshed";
		return false;
	}

	const std::string normalizedExpectedHash = ToLowerAsciiCopyLocal(TrimAsciiCopy(expectedBaseHash));

	if (hasCache) {
		outCacheEntry = BuildPageCodeCacheEntryForAI(item, outCacheEntry.code, &outCacheEntry);
	}

	if (isFixedTableWriteBase) {
		std::string fixedTableCode;
		std::string fixedTableTrace;
		std::string fixedTableError;
		if (TryReadFixedTableRealPageCodeForAI(item, fixedTableCode, fixedTableTrace, fixedTableError)) {
			std::string writeBaseCode = fixedTableCode;
			if (item.typeKey == "const_resource") {
				std::string mirrorConstCode;
				std::string mirrorConstError;
				if (TryReadWorkspaceMirrorTextLocalForAI("src/.常量.txt", mirrorConstCode, mirrorConstError)) {
					writeBaseCode = NormalizeRealCodeLineBreaksToCrLf(
						MergeConstResourceLongTextPlaceholdersForAI(fixedTableCode, mirrorConstCode));
					fixedTableTrace += "|const_long_text_placeholders_merged_from_mirror";
				}
				else {
					fixedTableTrace += "|const_long_text_mirror_merge_failed:" + mirrorConstError;
				}
			}

			outLiveTrace = fixedTableTrace + "|write_base=fixed_table_real_copy";
			outCurrentCode = writeBaseCode;
			const std::string fixedTableHash = BuildStableTextHashForRealCode(writeBaseCode);
			const bool cacheSeeded = !hasCache;
			const bool cacheChanged = hasCache && outCacheEntry.codeHash != fixedTableHash;
			PutPageCodeCacheEntryForAI(item, writeBaseCode, hasCache ? &outCacheEntry : nullptr, &outCacheEntry);
			outCacheRefreshed = cacheSeeded || cacheChanged;
			if (!normalizedExpectedHash.empty() && fixedTableHash != normalizedExpectedHash) {
				outError = "expected_base_hash mismatch; cache refreshed to fixed table real page code";
				return false;
			}
			outBaseCode = writeBaseCode;
			return true;
		}
		if (!fixedTableTrace.empty()) {
			outLiveTrace = fixedTableTrace;
		}
		if (!fixedTableError.empty()) {
			outLiveTrace += "|fixed_table_real_copy_failed:" + fixedTableError;
		}
		outError = fixedTableError.empty() ? "copy fixed table real page failed" : fixedTableError;
		return false;
	}

	std::string liveCode;
	e571::NativeRealPageAccessResult liveReadResult{};
	if (!TryReadRealPageCodeForAI(item, liveCode, liveReadResult, outError)) {
		outLiveTrace = liveReadResult.trace;
		return false;
	}

	outLiveTrace = liveReadResult.trace;
	outCurrentCode = liveCode;
	if (outEditorObject != nullptr) {
		*outEditorObject = liveReadResult.editorObject;
	}
	const std::string liveHash = BuildStableTextHashForRealCode(liveCode);

	if (!normalizedExpectedHash.empty() && liveHash != normalizedExpectedHash) {
		PutPageCodeCacheEntryForAI(item, liveCode, hasCache ? &outCacheEntry : nullptr, &outCacheEntry);
		outCacheRefreshed = true;
		outError = "expected_base_hash mismatch; cache refreshed to live page code";
		return false;
	}

	if (hasCache && liveHash != outCacheEntry.codeHash) {
		PutPageCodeCacheEntryForAI(item, liveCode, &outCacheEntry, &outCacheEntry);
		outCacheRefreshed = true;
		outError = "live page code changed since cache was captured; cache refreshed, retry the edit";
		return false;
	}

	if (!hasCache) {
		PutPageCodeCacheEntryForAI(item, liveCode, nullptr, &outCacheEntry);
	}
	else {
		PutPageCodeCacheEntryForAI(item, liveCode, &outCacheEntry, &outCacheEntry);
	}

	outBaseCode = liveCode;
	return true;
}

bool TryResolveSourceEditBaseForAI(
	const ProgramTreeItemInfo& item,
	const std::string& filePathUtf8,
	const std::string& expectedBaseHash,
	SourceEditBaseForAI& outBase,
	std::string& outError)
{
	outBase = {};
	outError.clear();

	const AISourceEditMode mode = GetActiveSourceEditModeForAI();
	outBase.sourceMode = SourceEditModeStringForAI(mode);
	if (mode != AISourceEditMode::MirrorSourceBase) {
		PageCodeCacheEntry cacheEntry;
		if (!TryResolveRealPageWriteBaseForAI(
				item,
				expectedBaseHash,
				false,
				outBase.baseCode,
				cacheEntry,
				outBase.trace,
				outBase.currentCode,
				outBase.cacheRefreshed,
				outError,
				&outBase.editorObject)) {
			return false;
		}
		outBase.codeKind = "real_source";
		return true;
	}

	std::string mirrorCode;
	std::string mirrorError;
	if (!TryReadWorkspaceMirrorTextLocalForAI(filePathUtf8, mirrorCode, mirrorError)) {
		outBase.trace = "source_edit_mode=mirror_source_base|mirror_read_failed:" + mirrorError;
		outError = mirrorError.empty() ? "read workspace mirror file failed" : mirrorError;
		return false;
	}

	outBase.baseCode = NormalizeRealCodeLineBreaksToCrLf(mirrorCode);
	outBase.currentCode = outBase.baseCode;
	outBase.trace =
		std::string("source_edit_mode=mirror_source_base") +
		"|write_base=workspace_mirror" +
		"|real_page_read=skipped";
	outBase.cacheRefreshed = false;
	outBase.mirrorMode = true;
	outBase.codeKind = "mirror_source";

	const std::string normalizedExpectedHash = ToLowerAsciiCopyLocal(TrimAsciiCopy(expectedBaseHash));
	const std::string mirrorHash = BuildStableTextHashForRealCode(outBase.baseCode);
	if (!normalizedExpectedHash.empty() && mirrorHash != normalizedExpectedHash) {
		outError = "expected_base_hash mismatch; mirror write base refreshed from workspace mirror";
		return false;
	}

	return true;
}

bool TryWriteRealPageCodeForAI(
	const ProgramTreeItemInfo& item,
	const std::string& baseCode,
	const std::string& newCode,
	const std::string& snapshotNote,
	PageCodeSnapshotEntry& outSnapshot,
	std::string& outFinalCode,
	std::string& outTrace,
	bool& outRollbackAttempted,
	bool& outRollbackSucceeded,
	std::string& outError,
	std::uintptr_t preferredEditorObject = 0,
	const std::string* snapshotBaseCode = nullptr)
{
	outSnapshot = {};
	outFinalCode.clear();
	outTrace.clear();
	outRollbackAttempted = false;
	outRollbackSucceeded = false;
	outError.clear();

	const std::string& rollbackBaseCode =
		(snapshotBaseCode != nullptr && !snapshotBaseCode->empty()) ? *snapshotBaseCode : baseCode;

	PageCodeCacheEntry existingEntry;
	const bool hasCache = PageCodeCacheManager::Instance().Get(item.name, item.typeKey, existingEntry);
	PutPageCodeCacheEntryForAI(item, rollbackBaseCode, hasCache ? &existingEntry : nullptr, &existingEntry);
	if (!PageCodeCacheManager::Instance().AddSnapshot(item.name, item.typeKey, rollbackBaseCode, snapshotNote, outSnapshot)) {
		outError = "add page snapshot failed";
		return false;
	}

	const std::string normalizedBaseCode = NormalizeRealCodeLineBreaksToCrLf(baseCode);
	const std::string normalizedRollbackBaseCode = NormalizeRealCodeLineBreaksToCrLf(rollbackBaseCode);
	const std::string normalizedNewCode = NormalizeRealCodeLineBreaksToCrLf(newCode);
	const std::string preparedExpectedCode = NormalizeRealCodeLineBreaksToCrLf(
		PrepareAssemblyVariablesForRealPageWrite(normalizedNewCode));

	std::string openTrace;
	std::string openError;
	const std::uintptr_t moduleBase = reinterpret_cast<std::uintptr_t>(::GetModuleHandleW(nullptr));
	if (moduleBase == 0) {
		outTrace = "internal_real_page_write_failed|module_base_invalid";
		outError = "module base invalid";
		return false;
	}

	if (IsFixedTableProgramItemForAI(item)) {
		e571::NativeRealPageAccessResult fixedTableWriteResult{};
		const auto fixedTableWriteStart = ToolPerfClock::now();
		const bool fixedTableWriteOk = e571::ReplaceRealPageCodeByProgramTreeItemData(
			item.itemData,
			moduleBase,
			normalizedNewCode,
			&normalizedRollbackBaseCode,
			&fixedTableWriteResult);
		const auto fixedTableWriteMs = ElapsedToolMs(fixedTableWriteStart);
		outRollbackAttempted = fixedTableWriteResult.rollbackAttempted;
		outRollbackSucceeded = fixedTableWriteResult.rollbackSucceeded;

		const std::string fixedTableDirectTrace =
			std::string("write_by_fixed_table_internal_editor_object") +
			"|editor_object=" + std::to_string(fixedTableWriteResult.editorObject) +
			"|write_ok=" + std::to_string(fixedTableWriteOk ? 1 : 0) +
			"|write_ms=" + std::to_string(fixedTableWriteMs) +
			(fixedTableWriteResult.trace.empty() ? std::string() : ("|" + fixedTableWriteResult.trace));
		if (fixedTableWriteOk) {
			outFinalCode = fixedTableWriteResult.pageCode.empty()
				? preparedExpectedCode
				: NormalizeRealCodeLineBreaksToCrLf(fixedTableWriteResult.pageCode);

			PageCodeCacheEntry postSnapshotEntry;
			const bool hasPostSnapshotEntry = PageCodeCacheManager::Instance().Get(item.name, item.typeKey, postSnapshotEntry);
			PutPageCodeCacheEntryForAI(item, outFinalCode, hasPostSnapshotEntry ? &postSnapshotEntry : nullptr, nullptr);

			outTrace = fixedTableDirectTrace + "|final_code=real_copy";
			return true;
		}

		std::string fallbackFinalCode;
		std::string fallbackTrace;
		std::string fallbackError;
		bool fallbackRollbackAttempted = false;
		bool fallbackRollbackSucceeded = false;
		if (!TryWriteFixedTableRealPageCodeByPasteForAI(
				item,
				normalizedRollbackBaseCode,
				normalizedNewCode,
				fallbackFinalCode,
				fallbackTrace,
				fallbackError,
				&fallbackRollbackAttempted,
				&fallbackRollbackSucceeded)) {
			outRollbackAttempted = outRollbackAttempted || fallbackRollbackAttempted;
			outRollbackSucceeded = outRollbackSucceeded || fallbackRollbackSucceeded;
			outTrace = fixedTableDirectTrace + "|fallback_page_paste_failed|" + fallbackTrace;
			outError = fallbackError.empty() ? "paste fixed table page code failed" : fallbackError;
			return false;
		}

		outRollbackAttempted = outRollbackAttempted || fallbackRollbackAttempted;
		outRollbackSucceeded = outRollbackSucceeded || fallbackRollbackSucceeded;
		PageCodeCacheEntry postSnapshotEntry;
		const bool hasPostSnapshotEntry = PageCodeCacheManager::Instance().Get(item.name, item.typeKey, postSnapshotEntry);
		PutPageCodeCacheEntryForAI(item, fallbackFinalCode, hasPostSnapshotEntry ? &postSnapshotEntry : nullptr, nullptr);

		outFinalCode = fallbackFinalCode;
		outTrace = fixedTableDirectTrace + "|fallback_page_paste_ok|" + fallbackTrace + "|final_code=real_copy";
		return true;
	}

	e571::NativeRealPageAccessResult writeResult{};
	const auto writeStart = ToolPerfClock::now();
	const bool reusedEditorObject = preferredEditorObject != 0;
	const bool writeOk = reusedEditorObject
		? e571::ReplaceRealPageCodeByEditorObject(
			preferredEditorObject,
			moduleBase,
			normalizedNewCode,
			&normalizedRollbackBaseCode,
			&writeResult)
		: e571::ReplaceRealPageCodeByProgramTreeItemData(
			item.itemData,
			moduleBase,
			normalizedNewCode,
			&normalizedRollbackBaseCode,
			&writeResult);
	const auto writeMs = ElapsedToolMs(writeStart);
	outRollbackAttempted = writeResult.rollbackAttempted;
	outRollbackSucceeded = writeResult.rollbackSucceeded;

	if (writeOk) {
		outFinalCode = writeResult.pageCode.empty()
			? preparedExpectedCode
			: NormalizeRealCodeLineBreaksToCrLf(writeResult.pageCode);
	}

	const std::string writeTrace =
		std::string("write_by_internal_editor_object_whole_page") +
		"|editor_object_source=" + std::string(reusedEditorObject ? "read_base" : "program_tree") +
		"|editor_object=" + std::to_string(reusedEditorObject ? preferredEditorObject : writeResult.editorObject) +
		"|write_ok=" + std::to_string(writeOk ? 1 : 0) +
		"|write_ms=" + std::to_string(writeMs) +
		(writeResult.trace.empty() ? std::string() : ("|" + writeResult.trace));

	if (!writeOk) {
		outTrace = writeTrace;
		outError = "whole page paste failed";
		return false;
	}

	PageCodeCacheEntry postSnapshotEntry;
	const bool hasPostSnapshotEntry = PageCodeCacheManager::Instance().Get(item.name, item.typeKey, postSnapshotEntry);
	PutPageCodeCacheEntryForAI(item, outFinalCode, hasPostSnapshotEntry ? &postSnapshotEntry : nullptr, nullptr);

	outTrace = writeTrace;
	return true;
}

nlohmann::json BuildRealPagePatchHunkJsonForAI(const RealPageStructuredPatchHunk& hunk)
{
	nlohmann::json lines = nlohmann::json::array();
	for (const auto& line : hunk.lines) {
		lines.push_back(LocalToUtf8Text(line));
	}

	nlohmann::json row;
	row["old_start"] = hunk.oldStart;
	row["old_lines"] = hunk.oldLines;
	row["new_start"] = hunk.newStart;
	row["new_lines"] = hunk.newLines;
	row["lines"] = std::move(lines);
	return row;
}

nlohmann::json BuildRealPageTextEditResultJsonForAI(const RealPageTextEditApplyResult& result)
{
	nlohmann::json row;
	row["match_count"] = result.matchCount;
	row["applied"] = result.applied;
	if (!result.error.empty()) {
		row["error"] = result.error;
	}
	return row;
}

bool TryParseRealPageTextEditsFromJson(
	const nlohmann::json& args,
	const char* key,
	std::vector<RealPageTextEditRequest>& outEdits,
	std::string& outError)
{
	outEdits.clear();
	outError.clear();

	if (!args.contains(key) || !args[key].is_array()) {
		outError = std::string(key) + " must be an array";
		return false;
	}

	for (const auto& item : args[key]) {
		if (!item.is_object()) {
			outError = std::string(key) + " must contain objects";
			return false;
		}

		RealPageTextEditRequest edit;
		edit.oldText = item.contains("old_text") && item["old_text"].is_string()
			? Utf8ToLocalText(item["old_text"].get<std::string>())
			: std::string();
		edit.newText = item.contains("new_text") && item["new_text"].is_string()
			? Utf8ToLocalText(item["new_text"].get<std::string>())
			: std::string();
		edit.replaceAll = item.contains("replace_all") && item["replace_all"].is_boolean()
			? item["replace_all"].get<bool>()
			: false;
		outEdits.push_back(std::move(edit));
	}

	return true;
}

bool TryListImportedModulePathsForAI(std::vector<std::string>& outPaths, std::string* outError)
{
	outPaths.clear();
	if (outError != nullptr) {
		outError->clear();
	}

	int count = 0;
	if (!IDEFacade::Instance().GetImportedECOMCount(count)) {
		if (outError != nullptr) {
			*outError = "GetImportedECOMCount failed";
		}
		return false;
	}

	for (int i = 0; i < count; ++i) {
		std::string path;
		if (IDEFacade::Instance().GetImportedECOMPath(i, path) && !TrimAsciiCopy(path).empty()) {
			outPaths.push_back(path);
		}
	}
	return true;
}

std::string GetFileNameOnlyForAI(const std::string& path)
{
	const size_t pos = path.find_last_of("\\/");
	return pos == std::string::npos ? path : path.substr(pos + 1);
}

std::string GetFileStemForAI(const std::string& path)
{
	const std::string fileName = GetFileNameOnlyForAI(path);
	const size_t pos = fileName.find_last_of('.');
	return pos == std::string::npos ? fileName : fileName.substr(0, pos);
}

bool EqualsInsensitiveForAI(const std::string& left, const std::string& right)
{
	return ToLowerAsciiCopyLocal(left) == ToLowerAsciiCopyLocal(right);
}

bool ResolveImportedModulePathForAI(
	const std::string& moduleName,
	const std::string& modulePath,
	std::string& outResolvedPath,
	std::string& outError)
{
	outResolvedPath.clear();
	outError.clear();

	const std::string trimmedPath = TrimAsciiCopy(modulePath);
	if (!trimmedPath.empty()) {
		outResolvedPath = trimmedPath;
		return true;
	}

	const std::string trimmedName = TrimAsciiCopy(moduleName);
	if (trimmedName.empty()) {
		outError = "module_name or module_path is required";
		return false;
	}

	std::vector<std::string> paths;
	if (!TryListImportedModulePathsForAI(paths, &outError)) {
		return false;
	}

	std::vector<std::string> matched;
	for (const auto& path : paths) {
		const std::string fileName = GetFileNameOnlyForAI(path);
		const std::string stem = GetFileStemForAI(path);
		if (EqualsInsensitiveForAI(fileName, trimmedName) ||
			EqualsInsensitiveForAI(stem, trimmedName) ||
			EqualsInsensitiveForAI(path, trimmedName)) {
			matched.push_back(path);
		}
	}

	if (matched.empty()) {
		outError = "module not found";
		return false;
	}
	if (matched.size() > 1) {
		outError = "module name is ambiguous";
		return false;
	}

	outResolvedPath = matched.front();
	return true;
}

std::string NormalizePathForAI(const std::string& pathText)
{
	if (TrimAsciiCopy(pathText).empty()) {
		return std::string();
	}

	try {
		std::filesystem::path path(pathText);
		path = path.lexically_normal();
		if (path.is_relative()) {
			path = std::filesystem::absolute(path);
		}
		path = path.lexically_normal();
		return path.string();
	}
	catch (...) {
		return pathText;
	}
}

bool TryListAvailableModulePathsForAI(
	std::vector<std::string>& outPaths,
	std::string* outError,
	std::string* outRootPath = nullptr)
{
	outPaths.clear();
	if (outError != nullptr) {
		outError->clear();
	}
	if (outRootPath != nullptr) {
		outRootPath->clear();
	}

	const std::filesystem::path root = GetAvailableModuleDirectoryForAI();
	if (root.empty()) {
		if (outError != nullptr) {
			*outError = "available module root is empty";
		}
		return false;
	}

	const std::string rootPath = NormalizePathForAI(root.string());
	if (outRootPath != nullptr) {
		*outRootPath = rootPath;
	}

	std::error_code ec;
	if (!std::filesystem::exists(root, ec) || !std::filesystem::is_directory(root, ec)) {
		if (outError != nullptr) {
			*outError = "available module root not found";
		}
		return false;
	}

	const auto options = std::filesystem::directory_options::skip_permission_denied;
	for (std::filesystem::recursive_directory_iterator it(root, options, ec), end; it != end; it.increment(ec)) {
		if (ec) {
			ec.clear();
			continue;
		}

		const auto& entry = *it;
		if (!entry.is_regular_file(ec)) {
			if (ec) {
				ec.clear();
			}
			continue;
		}

		if (ToLowerAsciiCopyLocal(entry.path().extension().string()) != ".ec") {
			continue;
		}

		const std::string normalizedPath = NormalizePathForAI(entry.path().string());
		if (!TrimAsciiCopy(normalizedPath).empty()) {
			outPaths.push_back(normalizedPath);
		}
	}

	std::sort(outPaths.begin(), outPaths.end(), [](const std::string& left, const std::string& right) {
		return ToLowerAsciiCopyLocal(left) < ToLowerAsciiCopyLocal(right);
	});
	outPaths.erase(std::unique(outPaths.begin(), outPaths.end(), [](const std::string& left, const std::string& right) {
		return EqualsInsensitiveForAI(left, right);
	}), outPaths.end());
	return true;
}

bool ResolveAvailableModulePathForAI(
	const std::string& moduleName,
	const std::string& modulePath,
	std::string& outResolvedPath,
	std::string& outError,
	std::vector<std::string>* outCandidates = nullptr,
	std::string* outRootPath = nullptr)
{
	outResolvedPath.clear();
	outError.clear();
	if (outCandidates != nullptr) {
		outCandidates->clear();
	}
	if (outRootPath != nullptr) {
		outRootPath->clear();
	}

	const std::filesystem::path root = GetAvailableModuleDirectoryForAI();
	const std::string normalizedRoot = NormalizePathForAI(root.string());
	if (outRootPath != nullptr) {
		*outRootPath = normalizedRoot;
	}

	const std::string trimmedPath = TrimAsciiCopy(modulePath);
	if (!trimmedPath.empty()) {
		std::filesystem::path candidatePath(trimmedPath);
		if (candidatePath.is_relative() && !root.empty()) {
			candidatePath = root / candidatePath;
		}
		const std::string normalizedPath = NormalizePathForAI(candidatePath.string());
		std::error_code ec;
		if (!normalizedPath.empty() &&
			std::filesystem::exists(candidatePath, ec) &&
			std::filesystem::is_regular_file(candidatePath, ec) &&
			ToLowerAsciiCopyLocal(candidatePath.extension().string()) == ".ec") {
			outResolvedPath = normalizedPath;
			return true;
		}

		outError = "module file not found";
		return false;
	}

	const std::string trimmedName = TrimAsciiCopy(moduleName);
	if (trimmedName.empty()) {
		outError = "module_name or module_path is required";
		return false;
	}

	DependencyCatalogCache::Instance().StartAsyncRefreshIfNeeded(false);
	std::string catalogPath;
	std::vector<DependencyCatalogCache::ModuleEntry> catalogCandidates;
	std::string catalogError;
	if (DependencyCatalogCache::Instance().ResolveModulePath(
			trimmedName,
			std::string(),
			catalogPath,
			&catalogCandidates,
			catalogError)) {
		outResolvedPath = catalogPath;
		return true;
	}
	if (catalogError == "module name is ambiguous" && outCandidates != nullptr) {
		for (const auto& candidate : catalogCandidates) {
			outCandidates->push_back(candidate.pathLocal);
		}
	}

	std::vector<std::string> paths;
	if (!TryListAvailableModulePathsForAI(paths, &outError, outRootPath)) {
		return false;
	}

	const std::string loweredName = ToLowerAsciiCopyLocal(trimmedName);
	std::vector<std::string> matched;
	for (const auto& path : paths) {
		const std::string fileName = GetFileNameOnlyForAI(path);
		const std::string stem = GetFileStemForAI(path);
		std::string relativePath;
		try {
			if (!normalizedRoot.empty()) {
				relativePath = NormalizePathForAI(
					std::filesystem::path(path).lexically_relative(std::filesystem::path(normalizedRoot)).string());
			}
		}
		catch (...) {
			relativePath.clear();
		}

		if (EqualsInsensitiveForAI(fileName, trimmedName) ||
			EqualsInsensitiveForAI(stem, trimmedName) ||
			EqualsInsensitiveForAI(path, trimmedName) ||
			(!relativePath.empty() && EqualsInsensitiveForAI(relativePath, trimmedName)) ||
			ToLowerAsciiCopyLocal(fileName).find(loweredName) != std::string::npos ||
			ToLowerAsciiCopyLocal(stem).find(loweredName) != std::string::npos ||
			(!relativePath.empty() && ToLowerAsciiCopyLocal(relativePath).find(loweredName) != std::string::npos)) {
			matched.push_back(path);
		}
	}

	if (outCandidates != nullptr) {
		*outCandidates = matched;
	}
	if (matched.empty()) {
		outError = "module not found in ecom directory";
		return false;
	}
	if (matched.size() > 1) {
		outError = "module name is ambiguous";
		return false;
	}

	outResolvedPath = matched.front();
	return true;
}

std::filesystem::path GetAvailableModuleDirectoryForAI()
{
	try {
		return std::filesystem::path(GetBasePath()) / "ecom";
	}
	catch (...) {
		return std::filesystem::path();
	}
}

std::string DependencyRefreshStateToStringForAI(DependencyCatalogCache::RefreshState state)
{
	switch (state) {
	case DependencyCatalogCache::RefreshState::Idle:
		return "idle";
	case DependencyCatalogCache::RefreshState::Running:
		return "running";
	case DependencyCatalogCache::RefreshState::Ready:
		return "ready";
	case DependencyCatalogCache::RefreshState::Failed:
		return "failed";
	default:
		return "unknown";
	}
}

void AppendDependencyCatalogStatusForAI(nlohmann::json& r)
{
	const DependencyCatalogCache::Status status = DependencyCatalogCache::Instance().GetStatus();
	nlohmann::json row;
	row["state"] = DependencyRefreshStateToStringForAI(status.state);
	row["has_snapshot"] = status.hasSnapshot;
	row["running"] = status.running;
	row["refresh_generation"] = status.refreshGeneration;
	row["last_duration_ms"] = status.lastDurationMs;
	row["module_count"] = status.moduleCount;
	row["library_count"] = status.libraryCount;
	row["decoded_module_count"] = status.decodedModuleCount;
	row["decoded_library_count"] = status.decodedLibraryCount;
	row["cache_root"] = LocalToUtf8Text(status.cacheRootLocal);
	row["ecom_root"] = LocalToUtf8Text(status.ecomRootLocal);
	row["lib_root"] = LocalToUtf8Text(status.libRootLocal);
	if (!status.lastErrorLocal.empty()) {
		row["last_error"] = LocalToUtf8Text(status.lastErrorLocal);
	}
	r["catalog_status"] = std::move(row);
}

bool IsPathSameForAI(const std::string& left, const std::string& right)
{
	if (TrimAsciiCopy(left).empty() || TrimAsciiCopy(right).empty()) {
		return false;
	}
	try {
		return EqualsInsensitiveForAI(
			std::filesystem::path(left).lexically_normal().string(),
			std::filesystem::path(right).lexically_normal().string());
	}
	catch (...) {
		return EqualsInsensitiveForAI(left, right);
	}
}

bool IsModuleImportedForAI(const std::string& modulePathLocal)
{
	std::vector<std::string> paths;
	std::string error;
	if (!TryListImportedModulePathsForAI(paths, &error)) {
		return false;
	}
	for (const auto& path : paths) {
		if (IsPathSameForAI(path, modulePathLocal)) {
			return true;
		}
	}
	return false;
}

std::string NormalizeSupportLibraryLookupNameForAI(const std::string& text)
{
	std::string value = TrimAsciiCopy(text);
	if (value.size() >= 2 &&
		((value.front() == '"' && value.back() == '"') ||
			(value.front() == '\'' && value.back() == '\''))) {
		value = value.substr(1, value.size() - 2);
	}
	try {
		std::filesystem::path path(value);
		if (path.has_extension()) {
			const std::string ext = ToLowerAsciiCopyLocal(path.extension().string());
			if (ext == ".fne" || ext == ".fnr") {
				value = path.stem().string();
			}
		}
	}
	catch (...) {
	}
	return TrimAsciiCopy(value);
}

bool TryListLoadedSupportLibraryInfoForAI(std::vector<std::string>& outRows, std::string& outError)
{
	outRows.clear();
	outError.clear();

	int count = 0;
	if (!IDEFacade::Instance().RunGetNumLib(count)) {
		outError = "RunGetNumLib failed";
		return false;
	}
	for (int i = 0; i < count; ++i) {
		std::string text;
		if (IDEFacade::Instance().RunGetLibInfoText(i, text) && !TrimAsciiCopy(text).empty()) {
			outRows.push_back(text);
		}
	}
	return true;
}

bool IsSupportLibraryLoadedForAI(
	const std::string& libraryNameLocal,
	const std::string& fileNameLocal,
	int* outIndex = nullptr,
	std::string* outInfoText = nullptr)
{
	if (outIndex != nullptr) {
		*outIndex = -1;
	}
	if (outInfoText != nullptr) {
		outInfoText->clear();
	}

	std::vector<std::string> rows;
	std::string error;
	if (!TryListLoadedSupportLibraryInfoForAI(rows, error)) {
		return false;
	}

	const std::string libraryName = NormalizeSupportLibraryLookupNameForAI(libraryNameLocal);
	const std::string fileStem = NormalizeSupportLibraryLookupNameForAI(fileNameLocal);
	for (size_t i = 0; i < rows.size(); ++i) {
		const std::string& row = rows[i];
		const bool nameMatched = !libraryName.empty() && row.find(libraryName) != std::string::npos;
		const bool fileMatched = !fileStem.empty() && row.find(fileStem) != std::string::npos;
		if (!nameMatched && !fileMatched) {
			continue;
		}
		if (outIndex != nullptr) {
			*outIndex = static_cast<int>(i);
		}
		if (outInfoText != nullptr) {
			*outInfoText = row;
		}
		return true;
	}
	return false;
}

bool IsSupportLibraryDirectiveTargetForAI(const ProgramTreeItemInfo& item)
{
	return item.typeKey == "assembly" || item.typeKey == "class_module";
}

bool TryResolveSupportLibraryDirectiveTargetForAI(
	ProgramTreeItemInfo& outItem,
	std::string& outTrace,
	std::string& outError)
{
	outItem = {};
	outTrace.clear();
	outError.clear();

	std::string currentName;
	std::string currentType;
	std::string currentTrace;
	if (IDEFacade::Instance().GetCurrentPageName(currentName, &currentType, &currentTrace)) {
		ProgramTreeItemInfo currentItem;
		std::string currentError;
		if (TryGetProgramItemByNameForAI(currentName, "all", currentItem, currentError) &&
			IsSupportLibraryDirectiveTargetForAI(currentItem)) {
			outItem = currentItem;
			outTrace = "target=current_page|page=" + currentName + "|type=" + currentItem.typeKey + "|" + currentTrace;
			return true;
		}
		outTrace = "current_page_not_usable|page=" + currentName + "|type=" + currentType + "|error=" + currentError;
	}
	else {
		outTrace = "current_page_name_failed|" + currentTrace;
	}

	std::vector<ProgramTreeItemInfo> items;
	if (!TryListProgramTreeItemsForAI(items, &outError)) {
		return false;
	}
	for (const auto& item : items) {
		if (item.typeKey == "assembly") {
			outItem = item;
			outTrace += "|target=first_assembly|page=" + item.name;
			return true;
		}
	}
	for (const auto& item : items) {
		if (item.typeKey == "class_module") {
			outItem = item;
			outTrace += "|target=first_class_module|page=" + item.name;
			return true;
		}
	}

	outError = "no assembly or class module page found for support library directive";
	return false;
}

bool CodeHasSupportLibraryDirectiveForAI(const std::string& code, const std::string& libraryNameLocal)
{
	const std::string target = NormalizeSupportLibraryLookupNameForAI(libraryNameLocal);
	for (const auto& line : SplitRealCodeLines(NormalizeRealCodeLineBreaksToCrLf(code))) {
		const std::string trimmed = TrimAsciiCopy(line);
		if (trimmed.rfind(".支持库", 0) != 0) {
			continue;
		}
		const std::string name = NormalizeSupportLibraryLookupNameForAI(trimmed.substr(std::string(".支持库").size()));
		if (EqualsInsensitiveForAI(name, target)) {
			return true;
		}
	}
	return false;
}

std::vector<std::string> ExtractSupportLibraryDirectiveNamesForAI(const std::string& code)
{
	std::vector<std::string> names;
	for (const auto& line : SplitRealCodeLines(NormalizeRealCodeLineBreaksToCrLf(code))) {
		const std::string trimmed = TrimAsciiCopy(line);
		if (trimmed.rfind(".支持库", 0) != 0) {
			continue;
		}
		const std::string name = NormalizeSupportLibraryLookupNameForAI(trimmed.substr(std::string(".支持库").size()));
		if (name.empty()) {
			continue;
		}
		bool exists = false;
		for (const std::string& existing : names) {
			if (EqualsInsensitiveForAI(existing, name)) {
				exists = true;
				break;
			}
		}
		if (!exists) {
			names.push_back(name);
		}
	}
	return names;
}

std::string InsertSupportLibraryDirectiveForAI(const std::string& code, const std::string& libraryNameLocal)
{
	const std::string normalizedCode = NormalizeRealCodeLineBreaksToCrLf(code);
	if (CodeHasSupportLibraryDirectiveForAI(normalizedCode, libraryNameLocal)) {
		return normalizedCode;
	}

	std::vector<std::string> lines = SplitRealCodeLines(normalizedCode);
	const std::string directive = ".支持库 " + NormalizeSupportLibraryLookupNameForAI(libraryNameLocal);
	size_t insertIndex = 0;
	for (size_t i = 0; i < lines.size(); ++i) {
		const std::string trimmed = TrimAsciiCopy(lines[i]);
		if (trimmed.rfind(".支持库", 0) == 0) {
			insertIndex = i + 1;
		}
	}
	if (insertIndex == 0) {
		for (size_t i = 0; i < lines.size(); ++i) {
			const std::string trimmed = TrimAsciiCopy(lines[i]);
			if (trimmed.rfind(".版本", 0) == 0) {
				insertIndex = i + 1;
				break;
			}
		}
	}
	lines.insert(lines.begin() + static_cast<std::ptrdiff_t>((std::min)(insertIndex, lines.size())), directive);
	return JoinRealCodeLines(lines);
}

std::string Utf8FileStemFromMirrorRelativePathForAI(std::string relativePathUtf8)
{
	std::replace(relativePathUtf8.begin(), relativePathUtf8.end(), '\\', '/');
	const size_t slash = relativePathUtf8.find_last_of('/');
	std::string fileName = slash == std::string::npos ? relativePathUtf8 : relativePathUtf8.substr(slash + 1);
	const size_t dot = fileName.find_last_of('.');
	if (dot != std::string::npos) {
		fileName.erase(dot);
	}
	return fileName;
}

bool IsMirrorTextSourceFileForAI(std::string relativePathUtf8)
{
	std::replace(relativePathUtf8.begin(), relativePathUtf8.end(), '\\', '/');
	const std::string lowered = ToLowerAsciiCopyLocal(relativePathUtf8);
	return lowered.rfind("src/", 0) == 0 &&
		lowered.size() > 4 &&
		lowered.substr(lowered.size() - 4) == ".txt";
}

bool TryReadSupportLibraryDirectiveMirrorCodeForAI(
	const ProgramTreeItemInfo& item,
	std::string& outCode,
	std::string& outRelativePathUtf8,
	std::string& outTrace)
{
	outCode.clear();
	outRelativePathUtf8.clear();
	outTrace.clear();

	std::vector<std::string> files;
	std::string listError;
	if (!WorkspaceMirror::ListMirrorFiles(files, listError)) {
		outTrace = "mirror_list_failed:" + listError;
		return false;
	}

	const std::string itemNameUtf8 = LocalToUtf8Text(item.name);
	for (const std::string& relativePath : files) {
		if (!IsMirrorTextSourceFileForAI(relativePath)) {
			continue;
		}
		const std::string stemUtf8 = Utf8FileStemFromMirrorRelativePathForAI(relativePath);
		if (stemUtf8 != itemNameUtf8 && !EqualsInsensitiveForAI(Utf8ToLocalText(stemUtf8), item.name)) {
			continue;
		}

		std::string code;
		std::string readError;
		if (!TryReadWorkspaceMirrorTextLocalForAI(relativePath, code, readError)) {
			outTrace = "mirror_read_failed:" + readError + "|path=" + Utf8ToLocalText(relativePath);
			return false;
		}
		outCode = NormalizeRealCodeLineBreaksToCrLf(code);
		outRelativePathUtf8 = relativePath;
		outTrace = "mirror_read_ok|path=" + Utf8ToLocalText(relativePath);
		return true;
	}

	outTrace = "mirror_source_page_not_found|page=" + item.name;
	return false;
}

std::string MergeSupportLibraryDirectivesFromMirrorForAI(
	const std::string& baseCode,
	const std::string& mirrorCode,
	int& outMergedCount)
{
	outMergedCount = 0;
	std::string merged = NormalizeRealCodeLineBreaksToCrLf(baseCode);
	for (const std::string& name : ExtractSupportLibraryDirectiveNamesForAI(mirrorCode)) {
		if (CodeHasSupportLibraryDirectiveForAI(merged, name)) {
			continue;
		}
		merged = InsertSupportLibraryDirectiveForAI(merged, name);
		++outMergedCount;
	}
	return merged;
}

std::string BuildRefreshDependencyCatalogJsonForAI(const std::string& argumentsJson, bool& outOk)
{
	outOk = false;
	nlohmann::json args;
	try {
		args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = std::string("invalid arguments json: ") + ex.what();
		return JsonToLocalTextForAI(r);
	}

	const bool force = args.contains("force") && args["force"].is_boolean() ? args["force"].get<bool>() : false;
	const bool wait = args.contains("wait") && args["wait"].is_boolean() ? args["wait"].get<bool>() : true;
	const int timeoutMs = args.contains("timeout_ms") && args["timeout_ms"].is_number_integer()
		? (std::clamp)(args["timeout_ms"].get<int>(), 0, 600000)
		: 0;

	if (!wait) {
		DependencyCatalogCache::Instance().StartAsyncRefreshIfNeeded(force);
		nlohmann::json r;
		r["ok"] = true;
		r["started"] = true;
		r["waited"] = false;
		AppendDependencyCatalogStatusForAI(r);
		outOk = true;
		return JsonUtf8ToAsciiTextForAI(r);
	}

	DependencyCatalogCache::Instance().StartAsyncRefreshIfNeeded(force);
	const bool idle = DependencyCatalogCache::Instance().WaitForIdle(static_cast<unsigned int>(timeoutMs));
	const DependencyCatalogCache::Status status = DependencyCatalogCache::Instance().GetStatus();
	nlohmann::json r;
	r["ok"] = idle && status.state == DependencyCatalogCache::RefreshState::Ready;
	r["started"] = true;
	r["waited"] = idle;
	if (!idle) {
		r["error"] = "dependency catalog refresh timeout";
	}
	else if (status.state != DependencyCatalogCache::RefreshState::Ready) {
		r["error"] = status.lastErrorLocal.empty() ? "dependency catalog refresh failed" : status.lastErrorLocal;
	}
	AppendDependencyCatalogStatusForAI(r);
	outOk = r["ok"].get<bool>();
	return JsonUtf8ToAsciiTextForAI(r);
}

std::string BuildSearchAvailableModulesJsonForAI(const std::string& argumentsJson, bool& outOk)
{
	outOk = false;
	nlohmann::json args;
	try {
		args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = std::string("invalid arguments json: ") + ex.what();
		return JsonToLocalTextForAI(r);
	}

	DependencyCatalogCache::SearchOptions options;
	options.queryUtf8 = args.contains("query") && args["query"].is_string() ? args["query"].get<std::string>() : std::string();
	options.limit = args.contains("limit") && args["limit"].is_number_integer() ? args["limit"].get<int>() : 50;
	options.includeSnippets = !args.contains("include_snippets") || !args["include_snippets"].is_boolean() || args["include_snippets"].get<bool>();
	const auto results = DependencyCatalogCache::Instance().SearchModules(options);

	nlohmann::json rows = nlohmann::json::array();
	for (const auto& result : results) {
		nlohmann::json row;
		row["module_name"] = LocalToUtf8Text(result.entry.moduleNameLocal);
		row["file_name"] = LocalToUtf8Text(result.entry.fileNameLocal);
		row["path"] = LocalToUtf8Text(result.entry.pathLocal);
		row["cache_path"] = LocalToUtf8Text(result.entry.cachePathLocal);
		row["decoded"] = result.entry.decoded;
		row["imported"] = IsModuleImportedForAI(result.entry.pathLocal);
		row["score"] = result.score;
		if (!result.entry.decodeErrorLocal.empty()) {
			row["decode_error"] = LocalToUtf8Text(result.entry.decodeErrorLocal);
		}
		if (!result.snippetUtf8.empty()) {
			row["snippet"] = result.snippetUtf8;
		}
		rows.push_back(std::move(row));
	}

	nlohmann::json r;
	r["ok"] = true;
	r["query"] = options.queryUtf8;
	r["count"] = rows.size();
	r["results"] = std::move(rows);
	AppendDependencyCatalogStatusForAI(r);
	outOk = true;
	return JsonUtf8ToAsciiTextForAI(r);
}

std::string BuildSearchAvailableSupportLibrariesJsonForAI(const std::string& argumentsJson, bool& outOk)
{
	outOk = false;
	nlohmann::json args;
	try {
		args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = std::string("invalid arguments json: ") + ex.what();
		return JsonToLocalTextForAI(r);
	}

	DependencyCatalogCache::SearchOptions options;
	options.queryUtf8 = args.contains("query") && args["query"].is_string() ? args["query"].get<std::string>() : std::string();
	options.limit = args.contains("limit") && args["limit"].is_number_integer() ? args["limit"].get<int>() : 50;
	options.includeSnippets = !args.contains("include_snippets") || !args["include_snippets"].is_boolean() || args["include_snippets"].get<bool>();
	const auto results = DependencyCatalogCache::Instance().SearchLibraries(options);

	nlohmann::json rows = nlohmann::json::array();
	for (const auto& result : results) {
		int loadedIndex = -1;
		const bool loaded = IsSupportLibraryLoadedForAI(
			result.entry.libraryNameLocal,
			result.entry.fileNameLocal,
			&loadedIndex,
			nullptr);
		nlohmann::json row;
		row["library_name"] = LocalToUtf8Text(result.entry.libraryNameLocal);
		row["file_name"] = LocalToUtf8Text(result.entry.fileNameLocal);
		row["path"] = LocalToUtf8Text(result.entry.pathLocal);
		row["cache_path"] = LocalToUtf8Text(result.entry.cachePathLocal);
		row["decoded"] = result.entry.decoded;
		row["loaded"] = loaded;
		if (loadedIndex >= 0) {
			row["loaded_index"] = loadedIndex;
		}
		row["score"] = result.score;
		if (!result.entry.decodeErrorLocal.empty()) {
			row["decode_error"] = LocalToUtf8Text(result.entry.decodeErrorLocal);
		}
		if (!result.snippetUtf8.empty()) {
			row["snippet"] = result.snippetUtf8;
		}
		rows.push_back(std::move(row));
	}

	nlohmann::json r;
	r["ok"] = true;
	r["query"] = options.queryUtf8;
	r["count"] = rows.size();
	r["results"] = std::move(rows);
	AppendDependencyCatalogStatusForAI(r);
	outOk = true;
	return JsonUtf8ToAsciiTextForAI(r);
}

std::string BuildAddSupportLibraryToProjectJsonOnMainThread(const std::string& argumentsJson, bool& outOk)
{
	outOk = false;
	nlohmann::json args;
	try {
		args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = std::string("invalid arguments json: ") + ex.what();
		return JsonToLocalTextForAI(r);
	}

	DependencyCatalogCache::Instance().StartAsyncRefreshIfNeeded(false);
	DependencyCatalogCache::Instance().WaitForIdle(30000);

	const std::string libraryName = args.contains("library_name") && args["library_name"].is_string()
		? Utf8ToLocalText(args["library_name"].get<std::string>())
		: std::string();
	const std::string libraryPath = args.contains("library_path") && args["library_path"].is_string()
		? Utf8ToLocalText(args["library_path"].get<std::string>())
		: std::string();

	DependencyCatalogCache::LibraryEntry entry;
	std::vector<DependencyCatalogCache::LibraryEntry> candidates;
	std::string resolveError;
	if (!DependencyCatalogCache::Instance().ResolveLibrary(
			libraryName,
			libraryPath,
			entry,
			&candidates,
			resolveError)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = resolveError.empty() ? "support library not found" : resolveError;
		if (!candidates.empty()) {
			nlohmann::json rows = nlohmann::json::array();
			for (const auto& candidate : candidates) {
				rows.push_back({
					{"library_name", LocalToUtf8Text(candidate.libraryNameLocal)},
					{"file_name", LocalToUtf8Text(candidate.fileNameLocal)},
					{"path", LocalToUtf8Text(candidate.pathLocal)},
					{"decoded", candidate.decoded}
				});
			}
			r["candidates"] = std::move(rows);
		}
		AppendDependencyCatalogStatusForAI(r);
		return JsonUtf8ToAsciiTextForAI(r);
	}

	const std::string normalizedLibraryName = NormalizeSupportLibraryLookupNameForAI(entry.libraryNameLocal);
	if (normalizedLibraryName.empty()) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = "resolved support library name is empty";
		r["library_path"] = LocalToUtf8Text(entry.pathLocal);
		return JsonUtf8ToAsciiTextForAI(r);
	}

	ProgramTreeItemInfo targetItem;
	std::string targetTrace;
	std::string targetError;
	if (!TryResolveSupportLibraryDirectiveTargetForAI(targetItem, targetTrace, targetError)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = targetError.empty() ? "resolve support library directive target failed" : targetError;
		r["trace"] = LocalToUtf8Text(targetTrace);
		r["library_name"] = LocalToUtf8Text(normalizedLibraryName);
		r["library_path"] = LocalToUtf8Text(entry.pathLocal);
		return JsonUtf8ToAsciiTextForAI(r);
	}

	std::string baseCode;
	e571::NativeRealPageAccessResult readResult{};
	std::string readError;
	if (!TryReadRealPageCodeForAI(targetItem, baseCode, readResult, readError)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = readError.empty() ? "read directive target page failed" : readError;
		r["trace"] = LocalToUtf8Text(targetTrace + "|" + readResult.trace);
		r["library_name"] = LocalToUtf8Text(normalizedLibraryName);
		r["target_page"] = LocalToUtf8Text(targetItem.name);
		r["target_type"] = targetItem.typeKey;
		return JsonUtf8ToAsciiTextForAI(r);
	}

	std::string writeBaseCode = baseCode;
	std::string mirrorDirectiveTrace;
	std::string mirrorDirectivePathUtf8;
	std::string mirrorCode;
	int mergedDirectiveCount = 0;
	if (TryReadSupportLibraryDirectiveMirrorCodeForAI(targetItem, mirrorCode, mirrorDirectivePathUtf8, mirrorDirectiveTrace)) {
		writeBaseCode = MergeSupportLibraryDirectivesFromMirrorForAI(baseCode, mirrorCode, mergedDirectiveCount);
	}

	const bool directiveAlreadyExists = CodeHasSupportLibraryDirectiveForAI(writeBaseCode, normalizedLibraryName);
	std::string finalCode = writeBaseCode;
	bool writeSucceeded = false;
	bool rollbackAttempted = false;
	bool rollbackSucceeded = false;
	std::string writeTrace;
	PageCodeSnapshotEntry snapshot;
	if (!directiveAlreadyExists) {
		const std::string newCode = InsertSupportLibraryDirectiveForAI(writeBaseCode, normalizedLibraryName);
		if (BuildStableTextHashForRealCode(newCode) != BuildStableTextHashForRealCode(writeBaseCode)) {
			std::string writeError;
			if (!TryWriteRealPageCodeForAI(
					targetItem,
					writeBaseCode,
					newCode,
					"add_support_library_to_project",
					snapshot,
					finalCode,
					writeTrace,
					rollbackAttempted,
					rollbackSucceeded,
					writeError)) {
				nlohmann::json r;
				r["ok"] = false;
				r["error"] = writeError.empty() ? "write support library directive failed" : writeError;
				r["trace"] = LocalToUtf8Text(targetTrace + "|" + readResult.trace + "|" + mirrorDirectiveTrace + "|" + writeTrace);
				r["library_name"] = LocalToUtf8Text(normalizedLibraryName);
				r["library_path"] = LocalToUtf8Text(entry.pathLocal);
				r["target_page"] = LocalToUtf8Text(targetItem.name);
				r["target_type"] = targetItem.typeKey;
				r["merged_support_directive_count"] = mergedDirectiveCount;
				r["support_directive_mirror_trace"] = LocalToUtf8Text(mirrorDirectiveTrace);
				r["rollback_attempted"] = rollbackAttempted;
				r["rollback_succeeded"] = rollbackSucceeded;
				return JsonUtf8ToAsciiTextForAI(r);
			}
			writeSucceeded = true;
		}
	}

	const bool refreshCalled = IDEFacade::Instance().RunRefrushLib();
	Sleep(200);
	int loadedAfterIndex = -1;
	std::string loadedAfterInfo;
	const bool verified = IsSupportLibraryLoadedForAI(
		normalizedLibraryName,
		entry.fileNameLocal,
		&loadedAfterIndex,
		&loadedAfterInfo);

	nlohmann::json r;
	r["ok"] = verified;
	r["library_name"] = LocalToUtf8Text(normalizedLibraryName);
	r["file_name"] = LocalToUtf8Text(entry.fileNameLocal);
	r["library_path"] = LocalToUtf8Text(entry.pathLocal);
	r["cache_path"] = LocalToUtf8Text(entry.cachePathLocal);
	r["method"] = "source_directive_internal";
	r["already_loaded"] = directiveAlreadyExists && verified;
	r["directive_already_exists"] = directiveAlreadyExists;
	r["added"] = writeSucceeded || directiveAlreadyExists;
	r["write_succeeded"] = writeSucceeded;
	r["verified"] = verified;
	r["refresh_called"] = refreshCalled;
	r["target_page"] = LocalToUtf8Text(targetItem.name);
	r["target_type"] = targetItem.typeKey;
	r["merged_support_directive_count"] = mergedDirectiveCount;
	r["support_directive_mirror_trace"] = LocalToUtf8Text(mirrorDirectiveTrace);
	if (!mirrorDirectivePathUtf8.empty()) {
		r["support_directive_mirror_path"] = mirrorDirectivePathUtf8;
	}
	r["trace"] = LocalToUtf8Text(targetTrace + "|" + readResult.trace + "|" + mirrorDirectiveTrace + "|" + writeTrace);
	r["rollback_attempted"] = rollbackAttempted;
	r["rollback_succeeded"] = rollbackSucceeded;
	if (!snapshot.snapshotId.empty()) {
		r["snapshot_id"] = snapshot.snapshotId;
	}
	if (loadedAfterIndex >= 0) {
		r["loaded_index"] = loadedAfterIndex;
	}
	if (!verified) {
		r["error"] = "support library directive was written but library load verification failed";
		r["warning"] = LocalToUtf8Text("当前实现通过源码页 .支持库 指令加入支持库；如果 IDE 未立即载入，请保存/重开工程或在支持库设置中确认。");
		return JsonUtf8ToAsciiTextForAI(r);
	}

	AppendFullMirrorRefreshResult(r);
	outOk = true;
	return JsonUtf8ToAsciiTextForAI(r);
}

std::vector<std::string> SplitLinesCopyForAI(const std::string& text)
{
	std::vector<std::string> lines;
	size_t begin = 0;
	while (begin <= text.size()) {
		const size_t end = text.find('\n', begin);
		if (end == std::string::npos) {
			lines.push_back(text.substr(begin));
			break;
		}
		lines.push_back(text.substr(begin, end - begin));
		begin = end + 1;
	}
	return lines;
}

std::string NormalizeLineBreaksForAI(std::string text)
{
	text.erase(std::remove(text.begin(), text.end(), '\r'), text.end());
	return text;
}

std::string NormalizeRealPageCodeForLooseCompareForAI(const std::string& text)
{
	const std::string normalized = NormalizeRealCodeLineBreaksToCrLf(
		NormalizeRealPageAssemblyVariableAliasesForCompare(text));
	std::string result;
	result.reserve(normalized.size());
	bool lastKeptLineBlank = false;

	size_t start = 0;
	while (start <= normalized.size()) {
		size_t end = normalized.find("\r\n", start);
		if (end == std::string::npos) {
			end = normalized.size();
		}

		const std::string line = normalized.substr(start, end - start);
		const std::string trimmedLine = TrimAsciiCopy(line);
		if (trimmedLine.rfind(".支持库", 0) != 0) {
			const bool isBlankLine = trimmedLine.empty();
			if (!(isBlankLine && lastKeptLineBlank)) {
				result.append(line);
				if (end != normalized.size()) {
					result.append("\r\n");
				}
				lastKeptLineBlank = isBlankLine;
			}
		}

		if (end == normalized.size()) {
			break;
		}
		start = end + 2;
	}

	while (result.size() >= 2 && result.compare(result.size() - 2, 2, "\r\n") == 0) {
		result.resize(result.size() - 2);
	}
	return result;
}

std::string NormalizeRealPageCodeLineForStructuralCompareForAI(const std::string& line)
{
	const std::string trimmedLine = TrimAsciiCopy(line);
	if (trimmedLine.empty()) {
		return std::string();
	}

	std::string normalized;
	normalized.reserve(trimmedLine.size());
	bool inString = false;
	for (size_t i = 0; i < trimmedLine.size(); ++i) {
		const char ch = trimmedLine[i];
		if (!inString && std::isspace(static_cast<unsigned char>(ch)) != 0) {
			continue;
		}

		normalized.push_back(ch);
		if (ch != '"') {
			continue;
		}

		if (inString && i + 1 < trimmedLine.size() && trimmedLine[i + 1] == '"') {
			normalized.push_back(trimmedLine[i + 1]);
			++i;
			continue;
		}

		inString = !inString;
	}
	return normalized;
}

std::vector<std::string> BuildRealPageStructuralFingerprintForAI(const std::string& text)
{
	const std::vector<std::string> lines = SplitRealCodeLines(
		NormalizeRealCodeLineBreaksToCrLf(
			NormalizeRealPageAssemblyVariableAliasesForCompare(text)));
	std::vector<std::string> fingerprint;
	fingerprint.reserve(lines.size());
	for (const auto& line : lines) {
		const std::string trimmedLine = TrimAsciiCopy(line);
		if (trimmedLine.empty() || trimmedLine.rfind(".支持库", 0) == 0) {
			continue;
		}

		const std::string normalizedLine = NormalizeRealPageCodeLineForStructuralCompareForAI(trimmedLine);
		if (!normalizedLine.empty()) {
			fingerprint.push_back(normalizedLine);
		}
	}
	return fingerprint;
}

bool VerifyRealPageCodeMatchesForAI(
	const std::string& expectedCode,
	const std::string& actualCode,
	std::string* outMode)
{
	if (outMode != nullptr) {
		outMode->clear();
	}

	const std::string normalizedExpected = NormalizeRealPageCodeForLooseCompareForAI(expectedCode);
	const std::string normalizedActual = NormalizeRealPageCodeForLooseCompareForAI(actualCode);
	if (normalizedExpected == normalizedActual) {
		if (outMode != nullptr) {
			*outMode = "loose_text";
		}
		return true;
	}

	const std::vector<std::string> expectedFingerprint = BuildRealPageStructuralFingerprintForAI(expectedCode);
	const std::vector<std::string> actualFingerprint = BuildRealPageStructuralFingerprintForAI(actualCode);
	if (expectedFingerprint == actualFingerprint) {
		if (outMode != nullptr) {
			*outMode = "structural";
		}
		return true;
	}

	return false;
}

std::string BuildListImportedModulesJsonOnMainThread(bool& outOk)
{
	std::vector<std::string> paths;
	std::string error;
	if (!TryListImportedModulePathsForAI(paths, &error)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = error.empty() ? "list imported modules failed" : error;
		return JsonToLocalTextForAI(r);
	}

	nlohmann::json modules = nlohmann::json::array();
	for (size_t i = 0; i < paths.size(); ++i) {
		nlohmann::json row;
		row["index"] = static_cast<int>(i);
		row["path"] = LocalToUtf8Text(paths[i]);
		row["file_name"] = LocalToUtf8Text(GetFileNameOnlyForAI(paths[i]));
		row["module_name"] = LocalToUtf8Text(GetFileStemForAI(paths[i]));
		modules.push_back(std::move(row));
	}

	nlohmann::json r;
	r["ok"] = true;
	r["count"] = modules.size();
	r["warning"] = LocalToUtf8Text("这里列出的是项目当前导入的易模块路径。");
	r["modules"] = std::move(modules);
	outOk = true;
	return JsonToLocalTextForAI(r);
}

void AppendFullMirrorRefreshResult(nlohmann::json& r);

std::string BuildAddModuleToProjectJsonOnMainThread(const std::string& argumentsJson, bool& outOk)
{
	const auto totalStart = ToolPerfClock::now();
	nlohmann::json args;
	try {
		args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = std::string("invalid arguments json: ") + ex.what();
		return JsonToLocalTextForAI(r);
	}

	const std::string moduleName = args.contains("module_name") && args["module_name"].is_string()
		? Utf8ToLocalText(args["module_name"].get<std::string>())
		: std::string();
	const std::string modulePath = args.contains("module_path") && args["module_path"].is_string()
		? Utf8ToLocalText(args["module_path"].get<std::string>())
		: std::string();
	const bool preferNewMethod = args.contains("prefer_new_method") && args["prefer_new_method"].is_boolean()
		? args["prefer_new_method"].get<bool>()
		: false;
	const bool allowBlockingImport = args.contains("allow_blocking_import") && args["allow_blocking_import"].is_boolean()
		? args["allow_blocking_import"].get<bool>()
		: false;

	std::string resolvedPath;
	std::string searchRoot;
	std::vector<std::string> candidates;
	std::string resolveError;
	if (!ResolveAvailableModulePathForAI(
			moduleName,
			modulePath,
			resolvedPath,
			resolveError,
			&candidates,
			&searchRoot)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = resolveError;
		if (!searchRoot.empty()) {
			r["search_root"] = LocalToUtf8Text(searchRoot);
		}
		if (!candidates.empty()) {
			nlohmann::json rows = nlohmann::json::array();
			for (const auto& path : candidates) {
				rows.push_back(LocalToUtf8Text(path));
			}
			r["candidates"] = std::move(rows);
		}
		return JsonToLocalTextForAI(r);
	}
	LogToolStageForAI(
		"add_module_to_project",
		"resolved|path=" + LocalToUtf8Text(resolvedPath) +
			"|prefer_new_method=" + std::to_string(preferNewMethod ? 1 : 0),
		totalStart);

	IDEFacade& ide = IDEFacade::Instance();
	const int existingIndex = ide.FindECOMIndex(resolvedPath);
	if (existingIndex >= 0) {
		nlohmann::json r;
		r["ok"] = true;
		r["module_path"] = LocalToUtf8Text(resolvedPath);
		r["module_name"] = LocalToUtf8Text(GetFileStemForAI(resolvedPath));
		r["file_name"] = LocalToUtf8Text(GetFileNameOnlyForAI(resolvedPath));
		r["already_imported"] = true;
		r["added"] = false;
		r["module_index"] = existingIndex;
		r["warning"] = LocalToUtf8Text("目标模块已在当前工程中，无需重复加入。");
		outOk = true;
		return JsonToLocalTextForAI(r);
	}

	if (!allowBlockingImport) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = "module import requires allow_blocking_import=true; E IDE module import APIs can block the MCP main-thread tool call";
		r["resolved"] = true;
		r["add_blocked"] = true;
		r["allow_blocking_import_required"] = true;
		r["module_path"] = LocalToUtf8Text(resolvedPath);
		r["module_name"] = LocalToUtf8Text(GetFileStemForAI(resolvedPath));
		r["file_name"] = LocalToUtf8Text(GetFileNameOnlyForAI(resolvedPath));
		r["warning"] = LocalToUtf8Text("已解析到可用模块，但默认不执行可能卡死 IDE 的导入调用；确认要真实导入时传 allow_blocking_import=true。");
		return JsonToLocalTextForAI(r);
	}

	bool addedByNewMethod = false;
	bool added = false;
	std::string addTrace;
	if (preferNewMethod) {
		LogToolStageForAI("add_module_to_project", "before_add_ecom2", totalStart);
		addedByNewMethod = ide.AddECOM2(resolvedPath);
		addTrace += std::string("AddECOM2=") + (addedByNewMethod ? "1" : "0");
		LogToolStageForAI("add_module_to_project", "after_add_ecom2|ok=" + std::to_string(addedByNewMethod ? 1 : 0), totalStart);
		if (!addedByNewMethod) {
			LogToolStageForAI("add_module_to_project", "before_add_ecom_legacy_after_new_failed", totalStart);
			added = ide.AddECOM(resolvedPath);
			addTrace += std::string("|AddECOM=") + (added ? "1" : "0");
			LogToolStageForAI("add_module_to_project", "after_add_ecom_legacy|ok=" + std::to_string(added ? 1 : 0), totalStart);
		}
		else {
			added = true;
		}
	}
	else {
		LogToolStageForAI("add_module_to_project", "before_add_ecom_legacy", totalStart);
		added = ide.AddECOM(resolvedPath);
		addTrace += std::string("AddECOM=") + (added ? "1" : "0");
		LogToolStageForAI("add_module_to_project", "after_add_ecom_legacy|ok=" + std::to_string(added ? 1 : 0), totalStart);
	}
	if (!added) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = "add_ecom_failed";
		r["module_path"] = LocalToUtf8Text(resolvedPath);
		r["module_name"] = LocalToUtf8Text(GetFileStemForAI(resolvedPath));
		r["file_name"] = LocalToUtf8Text(GetFileNameOnlyForAI(resolvedPath));
		r["add_trace"] = addTrace;
		return JsonToLocalTextForAI(r);
	}

	const int moduleIndex = ide.FindECOMIndex(resolvedPath);
	if (moduleIndex < 0) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = "add_ecom_succeeded_but_index_not_found";
		r["module_path"] = LocalToUtf8Text(resolvedPath);
		r["module_name"] = LocalToUtf8Text(GetFileStemForAI(resolvedPath));
		r["file_name"] = LocalToUtf8Text(GetFileNameOnlyForAI(resolvedPath));
		r["add_method"] = addedByNewMethod ? "AddECOM2" : "AddECOM";
		r["add_trace"] = addTrace;
		AppendFullMirrorRefreshResult(r);
		return JsonToLocalTextForAI(r);
	}

	nlohmann::json r;
	r["ok"] = true;
	r["module_path"] = LocalToUtf8Text(resolvedPath);
	r["module_name"] = LocalToUtf8Text(GetFileStemForAI(resolvedPath));
	r["file_name"] = LocalToUtf8Text(GetFileNameOnlyForAI(resolvedPath));
	r["already_imported"] = false;
	r["added"] = true;
	r["module_index"] = moduleIndex;
	r["add_method"] = addedByNewMethod ? "AddECOM2" : "AddECOM";
	r["add_trace"] = addTrace;
	r["warning"] = LocalToUtf8Text("模块已加入当前工程。后续可用 list_imported_modules 确认导入状态，或通过 refresh_workspace_mirror 刷新后使用 list_files/search_code/read_file 读取镜像内容。");
	AppendFullMirrorRefreshResult(r);
	outOk = true;
	return JsonToLocalTextForAI(r);
}

void AppendFullMirrorRefreshResult(nlohmann::json& r)
{
	std::string refreshError;
	std::string refreshMode;
	if (WorkspaceMirror::RefreshMirror(refreshError, &refreshMode, WorkspaceMirror::RefreshMode::Full)) {
		r["workspace_mirror_refreshed"] = true;
		r["refresh_mode"] = refreshMode;
		return;
	}
	r["workspace_mirror_refreshed"] = false;
	r["refresh_error"] = refreshError.empty() ? "refresh workspace mirror failed" : refreshError;
}

std::string BuildRemoveModuleFromProjectJsonOnMainThread(const std::string& argumentsJson, bool& outOk)
{
	nlohmann::json args;
	try {
		args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = std::string("invalid arguments json: ") + ex.what();
		return JsonToLocalTextForAI(r);
	}

	IDEFacade& ide = IDEFacade::Instance();
	int moduleIndex = -1;
	std::string modulePath;
	std::string moduleName;

	if (args.contains("module_index") && args["module_index"].is_number_integer()) {
		moduleIndex = args["module_index"].get<int>();
		int count = 0;
		if (!ide.GetImportedECOMCount(count)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = "GetImportedECOMCount failed";
			return JsonToLocalTextForAI(r);
		}
		if (moduleIndex < 0 || moduleIndex >= count) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = "module_index out of range";
			r["module_count"] = count;
			return JsonToLocalTextForAI(r);
		}
		ide.GetImportedECOMPath(moduleIndex, modulePath);
		moduleName = GetFileStemForAI(modulePath);
	}
	else {
		modulePath = args.contains("module_path") && args["module_path"].is_string()
			? Utf8ToLocalText(args["module_path"].get<std::string>())
			: std::string();
		moduleName = args.contains("module_name") && args["module_name"].is_string()
			? Utf8ToLocalText(args["module_name"].get<std::string>())
			: std::string();
		if (!TrimAsciiCopy(modulePath).empty()) {
			moduleIndex = ide.FindECOMIndex(TrimAsciiCopy(modulePath));
			if (moduleIndex >= 0) {
				ide.GetImportedECOMPath(moduleIndex, modulePath);
				moduleName = GetFileStemForAI(modulePath);
			}
		}
		else if (!TrimAsciiCopy(moduleName).empty()) {
			moduleIndex = ide.FindECOMNameIndex(TrimAsciiCopy(moduleName));
			if (moduleIndex >= 0) {
				ide.GetImportedECOMPath(moduleIndex, modulePath);
				moduleName = GetFileStemForAI(modulePath);
			}
		}
	}

	if (moduleIndex < 0) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = "module not found";
		if (!TrimAsciiCopy(modulePath).empty()) {
			r["module_path"] = LocalToUtf8Text(modulePath);
		}
		if (!TrimAsciiCopy(moduleName).empty()) {
			r["module_name"] = LocalToUtf8Text(moduleName);
		}
		std::vector<std::string> paths;
		std::string listError;
		if (TryListImportedModulePathsForAI(paths, &listError)) {
			nlohmann::json rows = nlohmann::json::array();
			for (size_t i = 0; i < paths.size(); ++i) {
				nlohmann::json row;
				row["module_index"] = i;
				row["module_path"] = LocalToUtf8Text(paths[i]);
				row["module_name"] = LocalToUtf8Text(GetFileStemForAI(paths[i]));
				row["file_name"] = LocalToUtf8Text(GetFileNameOnlyForAI(paths[i]));
				rows.push_back(std::move(row));
			}
			r["imported_modules"] = std::move(rows);
		}
		return JsonToLocalTextForAI(r);
	}

	if (TrimAsciiCopy(modulePath).empty()) {
		ide.GetImportedECOMPath(moduleIndex, modulePath);
	}
	if (TrimAsciiCopy(moduleName).empty()) {
		moduleName = GetFileStemForAI(modulePath);
	}

	if (!ide.RemoveECOM(moduleIndex)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = "remove_ecom_failed";
		r["module_index"] = moduleIndex;
		r["module_path"] = LocalToUtf8Text(modulePath);
		r["module_name"] = LocalToUtf8Text(moduleName);
		r["file_name"] = LocalToUtf8Text(GetFileNameOnlyForAI(modulePath));
		return JsonToLocalTextForAI(r);
	}

	nlohmann::json r;
	r["ok"] = true;
	r["removed"] = true;
	r["module_index"] = moduleIndex;
	r["module_path"] = LocalToUtf8Text(modulePath);
	r["module_name"] = LocalToUtf8Text(moduleName);
	r["file_name"] = LocalToUtf8Text(GetFileNameOnlyForAI(modulePath));
	AppendFullMirrorRefreshResult(r);
	outOk = true;
	return JsonToLocalTextForAI(r);
}

bool TryReadMappedRealPageCodeForAI(
	const WorkspaceMirror::ProgramItemRef& item,
	std::string& outCode,
	std::string& outCodeHash,
	std::string& outTrace,
	std::string& outError);

std::string ExecuteMappedEditFileToolForAI(
	const nlohmann::json& args,
	const std::string& filePathUtf8,
	const ProgramTreeItemInfo& item,
	bool& outOk)
{
	outOk = false;
	const auto totalStart = ToolPerfClock::now();
	LogToolStageForAI(
		"edit_file",
		"begin|page=" + LocalToUtf8Text(item.name) +
			"|type=" + item.typeKey +
			"|item_data=" + std::to_string(item.itemData),
		totalStart);

	const std::string oldText = GetJsonStringArgumentLocal(args, "old_text");
	const std::string newText = GetJsonStringArgumentLocal(args, "new_text");
	if (oldText.empty()) {
		LogToolStageForAI("edit_file", "invalid_arguments|old_text_empty", totalStart);
		return R"({"ok":false,"error":"old_text is required"})";
	}

	std::string error;
	SourceEditBaseForAI editBase;
	LogToolStageForAI("edit_file", "before_read_base", totalStart);
	if (!TryResolveSourceEditBaseForAI(item, filePathUtf8, std::string(), editBase, error)) {
		LogToolStageForAI(
			"edit_file",
			"after_read_base_failed|error=" + LocalToUtf8Text(error) +
				"|trace=" + LocalToUtf8Text(editBase.trace),
			totalStart);
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = error;
		r["page_name"] = LocalToUtf8Text(item.name);
		r["type_key"] = item.typeKey;
		r["trace"] = LocalToUtf8Text(editBase.trace);
		r["source_edit_mode"] = editBase.sourceMode.empty() ? SourceEditModeStringForAI(GetActiveSourceEditModeForAI()) : editBase.sourceMode;
		r["base_code_kind"] = editBase.codeKind;
		r["cache_refreshed"] = editBase.cacheRefreshed;
		if (!editBase.currentCode.empty()) {
			r["code"] = LocalToUtf8Text(editBase.currentCode);
			r["code_hash"] = BuildStableTextHashForRealCode(editBase.currentCode);
		}
		return JsonToLocalTextForAI(r);
	}
	const std::string& baseCode = editBase.baseCode;
	LogToolStageForAI(
		"edit_file",
		"after_read_base_ok|bytes=" + std::to_string(baseCode.size()) +
			"|editor_object=" + std::to_string(editBase.editorObject) +
			"|source_edit_mode=" + editBase.sourceMode +
			"|trace=" + LocalToUtf8Text(editBase.trace),
		totalStart);

	size_t matchCount = 0;
	std::string replacedCode;
	if (!ReplaceExactlyOnceForAI(baseCode, oldText, newText, replacedCode, matchCount)) {
		LogToolStageForAI(
			"edit_file",
			"replace_match_failed|match_count=" + std::to_string(matchCount),
			totalStart);
		const std::string baseLabel = editBase.mirrorMode ? "workspace mirror source" : "real IDE page";
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = matchCount == 0 ? "old_text not found in active edit base" : "old_text matched multiple times in active edit base";
		r["hint"] = "old_text must be copied exactly from the active edit base; preserve line breaks, indentation and full-width punctuation. Use write_file when replacing a large block.";
		r["active_edit_base"] = baseLabel;
		r["match_count"] = matchCount;
		r["page_name"] = LocalToUtf8Text(item.name);
		r["type_key"] = item.typeKey;
		r["trace"] = LocalToUtf8Text(editBase.trace);
		r["source_edit_mode"] = editBase.sourceMode;
		r["base_code_kind"] = editBase.codeKind;
		r["code_hash"] = BuildStableTextHashForRealCode(baseCode);
		r["code"] = LocalToUtf8Text(baseCode);
		return JsonToLocalTextForAI(r);
	}

	if (BuildStableTextHashForRealCode(replacedCode) == BuildStableTextHashForRealCode(baseCode)) {
		LogToolStageForAI("edit_file", "no_changes", totalStart);
		nlohmann::json r;
		r["ok"] = true;
		r["no_changes"] = true;
		r["page_name"] = LocalToUtf8Text(item.name);
		r["type_key"] = item.typeKey;
		r["type_name"] = LocalToUtf8Text(item.typeName);
		r["item_data"] = item.itemData;
		r["match_count"] = matchCount;
		r["source_edit_mode"] = editBase.sourceMode;
		r["base_code_kind"] = editBase.codeKind;
		r["code_kind"] = editBase.codeKind.empty() ? "real_source" : editBase.codeKind;
		r["code_hash"] = BuildStableTextHashForRealCode(baseCode);
		r["code"] = LocalToUtf8Text(baseCode);
		outOk = true;
		return JsonToLocalTextForAI(r);
	}

	PageCodeSnapshotEntry snapshot;
	std::string finalCode;
	bool rollbackAttempted = false;
	bool rollbackSucceeded = false;
	std::string writeTrace;
	LogToolStageForAI(
		"edit_file",
		"before_write|bytes=" + std::to_string(replacedCode.size()) +
			"|editor_object=" + std::to_string(editBase.editorObject),
		totalStart);
	if (!TryWriteRealPageCodeForAI(
			item,
			baseCode,
			replacedCode,
			"edit_file",
			snapshot,
			finalCode,
			writeTrace,
			rollbackAttempted,
			rollbackSucceeded,
			error,
			editBase.editorObject)) {
		LogToolStageForAI(
			"edit_file",
			"after_write_failed|error=" + LocalToUtf8Text(error) +
				"|trace=" + LocalToUtf8Text(writeTrace),
			totalStart);
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = error;
		r["page_name"] = LocalToUtf8Text(item.name);
		r["type_key"] = item.typeKey;
		r["trace"] = LocalToUtf8Text(editBase.trace + "|" + writeTrace);
		r["source_edit_mode"] = editBase.sourceMode;
		r["base_code_kind"] = editBase.codeKind;
		r["rollback_attempted"] = rollbackAttempted;
		r["rollback_succeeded"] = rollbackSucceeded;
		return JsonToLocalTextForAI(r);
	}
	LogToolStageForAI(
		"edit_file",
		"after_write_ok|bytes=" + std::to_string(finalCode.size()) +
			"|trace=" + LocalToUtf8Text(writeTrace),
		totalStart);

	const auto hunks = BuildRealPageStructuredPatch(baseCode, finalCode);
	nlohmann::json r;
	r["ok"] = true;
	r["page_name"] = LocalToUtf8Text(item.name);
	r["type_key"] = item.typeKey;
	r["type_name"] = LocalToUtf8Text(item.typeName);
	r["item_data"] = item.itemData;
	r["match_count"] = matchCount;
	r["trace"] = LocalToUtf8Text(editBase.trace + "|" + writeTrace);
	r["source_edit_mode"] = editBase.sourceMode;
	r["base_code_kind"] = editBase.codeKind;
	r["rollback_attempted"] = rollbackAttempted;
	r["rollback_succeeded"] = rollbackSucceeded;
	r["snapshot_id"] = snapshot.snapshotId;
	r["snapshot_note"] = LocalToUtf8Text(snapshot.note);
	r["base_hash"] = BuildStableTextHashForRealCode(baseCode);
	r["code_hash"] = BuildStableTextHashForRealCode(finalCode);
	r["changed_lines"] = CountStructuredPatchChangedLines(hunks);
	r["code_kind"] = "real_source";
	r["code"] = LocalToUtf8Text(finalCode);
	outOk = true;
	return JsonToLocalTextForAI(r);
}

std::string ExecuteMappedMultiEditFileToolForAI(
	const nlohmann::json& args,
	const std::string& filePathUtf8,
	const ProgramTreeItemInfo& item,
	bool& outOk)
{
	outOk = false;

	const bool failOnUnmatched = GetJsonBoolArgument(args, "fail_on_unmatched", true);
	const bool atomic = GetJsonBoolArgument(args, "atomic", true);

	std::vector<RealPageTextEditRequest> edits;
	std::string error;
	if (!TryParseRealPageTextEditsFromJson(args, "edits", edits, error)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = error;
		return JsonToLocalTextForAI(r);
	}
	if (edits.empty()) {
		return R"({"ok":false,"error":"edits is required"})";
	}

	SourceEditBaseForAI editBase;
	if (!TryResolveSourceEditBaseForAI(item, filePathUtf8, std::string(), editBase, error)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = error;
		r["page_name"] = LocalToUtf8Text(item.name);
		r["type_key"] = item.typeKey;
		r["trace"] = LocalToUtf8Text(editBase.trace);
		r["source_edit_mode"] = editBase.sourceMode.empty() ? SourceEditModeStringForAI(GetActiveSourceEditModeForAI()) : editBase.sourceMode;
		r["base_code_kind"] = editBase.codeKind;
		r["cache_refreshed"] = editBase.cacheRefreshed;
		if (!editBase.currentCode.empty()) {
			r["code"] = LocalToUtf8Text(editBase.currentCode);
			r["code_hash"] = BuildStableTextHashForRealCode(editBase.currentCode);
		}
		return JsonToLocalTextForAI(r);
	}
	const std::string& baseCode = editBase.baseCode;

	std::string candidateCode;
	std::vector<RealPageTextEditApplyResult> applyResults;
	ApplyRealPageTextEdits(baseCode, edits, false, candidateCode, applyResults, error);

	size_t appliedCount = 0;
	bool hasFailures = false;
	nlohmann::json resultRows = nlohmann::json::array();
	for (const auto& result : applyResults) {
		if (result.applied) {
			++appliedCount;
		}
		if (!result.error.empty()) {
			hasFailures = true;
		}
		resultRows.push_back(BuildRealPageTextEditResultJsonForAI(result));
	}

	if (hasFailures && (failOnUnmatched || atomic)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = failOnUnmatched ? "one or more edits did not match" : "atomic edit aborted because one or more edits failed";
		r["results"] = std::move(resultRows);
		r["page_name"] = LocalToUtf8Text(item.name);
		r["type_key"] = item.typeKey;
		r["trace"] = LocalToUtf8Text(editBase.trace);
		r["source_edit_mode"] = editBase.sourceMode;
		r["base_code_kind"] = editBase.codeKind;
		r["code_hash"] = BuildStableTextHashForRealCode(baseCode);
		r["code"] = LocalToUtf8Text(baseCode);
		return JsonToLocalTextForAI(r);
	}

	if (appliedCount == 0) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = "no edits applied";
		r["results"] = std::move(resultRows);
		r["page_name"] = LocalToUtf8Text(item.name);
		r["type_key"] = item.typeKey;
		r["trace"] = LocalToUtf8Text(editBase.trace);
		r["source_edit_mode"] = editBase.sourceMode;
		r["base_code_kind"] = editBase.codeKind;
		r["code_hash"] = BuildStableTextHashForRealCode(baseCode);
		r["code"] = LocalToUtf8Text(baseCode);
		return JsonToLocalTextForAI(r);
	}

	if (BuildStableTextHashForRealCode(candidateCode) == BuildStableTextHashForRealCode(baseCode)) {
		nlohmann::json r;
		r["ok"] = true;
		r["no_changes"] = true;
		r["results"] = std::move(resultRows);
		r["page_name"] = LocalToUtf8Text(item.name);
		r["type_key"] = item.typeKey;
		r["type_name"] = LocalToUtf8Text(item.typeName);
		r["item_data"] = item.itemData;
		r["source_edit_mode"] = editBase.sourceMode;
		r["base_code_kind"] = editBase.codeKind;
		r["code_hash"] = BuildStableTextHashForRealCode(baseCode);
		r["code"] = LocalToUtf8Text(baseCode);
		outOk = true;
		return JsonToLocalTextForAI(r);
	}

	PageCodeSnapshotEntry snapshot;
	std::string finalCode;
	bool rollbackAttempted = false;
	bool rollbackSucceeded = false;
	std::string writeTrace;
	if (!TryWriteRealPageCodeForAI(
			item,
			baseCode,
			candidateCode,
			"multi_edit_file",
			snapshot,
			finalCode,
			writeTrace,
			rollbackAttempted,
			rollbackSucceeded,
			error,
			editBase.editorObject)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = error;
		r["results"] = std::move(resultRows);
		r["trace"] = LocalToUtf8Text(editBase.trace + "|" + writeTrace);
		r["source_edit_mode"] = editBase.sourceMode;
		r["base_code_kind"] = editBase.codeKind;
		r["rollback_attempted"] = rollbackAttempted;
		r["rollback_succeeded"] = rollbackSucceeded;
		return JsonToLocalTextForAI(r);
	}

	const auto hunks = BuildRealPageStructuredPatch(baseCode, finalCode);
	nlohmann::json r;
	r["ok"] = true;
	r["page_name"] = LocalToUtf8Text(item.name);
	r["type_key"] = item.typeKey;
	r["type_name"] = LocalToUtf8Text(item.typeName);
	r["item_data"] = item.itemData;
	r["results"] = std::move(resultRows);
	r["applied_count"] = appliedCount;
	r["trace"] = LocalToUtf8Text(editBase.trace + "|" + writeTrace);
	r["source_edit_mode"] = editBase.sourceMode;
	r["base_code_kind"] = editBase.codeKind;
	r["rollback_attempted"] = rollbackAttempted;
	r["rollback_succeeded"] = rollbackSucceeded;
	r["snapshot_id"] = snapshot.snapshotId;
	r["snapshot_note"] = LocalToUtf8Text(snapshot.note);
	r["base_hash"] = BuildStableTextHashForRealCode(baseCode);
	r["code_hash"] = BuildStableTextHashForRealCode(finalCode);
	r["changed_lines"] = CountStructuredPatchChangedLines(hunks);
	r["code_kind"] = "real_source";
	r["code"] = LocalToUtf8Text(finalCode);
	outOk = true;
	return JsonToLocalTextForAI(r);
}

std::string ExecuteMappedWriteFileToolForAI(
	const nlohmann::json& args,
	const std::string& filePathUtf8,
	const ProgramTreeItemInfo& item,
	bool& outOk)
{
	outOk = false;

	const std::string fullCode = GetJsonStringArgumentLocal(args, "full_code");
	const std::string expectedBaseHash = GetJsonStringArgumentLocal(args, "expected_base_hash");
	if (fullCode.empty()) {
		return R"({"ok":false,"error":"full_code is required"})";
	}

	std::string error;
	SourceEditBaseForAI editBase;
	if (!TryResolveSourceEditBaseForAI(item, filePathUtf8, expectedBaseHash, editBase, error)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = error;
		r["page_name"] = LocalToUtf8Text(item.name);
		r["type_key"] = item.typeKey;
		r["trace"] = LocalToUtf8Text(editBase.trace);
		r["source_edit_mode"] = editBase.sourceMode.empty() ? SourceEditModeStringForAI(GetActiveSourceEditModeForAI()) : editBase.sourceMode;
		r["base_code_kind"] = editBase.codeKind;
		r["cache_refreshed"] = editBase.cacheRefreshed;
		if (!editBase.currentCode.empty()) {
			r["code"] = LocalToUtf8Text(editBase.currentCode);
			r["code_hash"] = BuildStableTextHashForRealCode(editBase.currentCode);
		}
		return JsonToLocalTextForAI(r);
	}
	const std::string& baseCode = editBase.baseCode;

	std::string preparedFullCode = fullCode;
	if (item.typeKey == "const_resource") {
		preparedFullCode = MergeConstResourceLongTextPlaceholdersForAI(preparedFullCode, baseCode);
	}
	const std::string normalizedFullCode = NormalizeRealCodeLineBreaksToCrLf(preparedFullCode);
	if (BuildStableTextHashForRealCode(normalizedFullCode) == BuildStableTextHashForRealCode(baseCode)) {
		nlohmann::json r;
		r["ok"] = true;
		r["no_changes"] = true;
		r["page_name"] = LocalToUtf8Text(item.name);
		r["type_key"] = item.typeKey;
		r["type_name"] = LocalToUtf8Text(item.typeName);
		r["item_data"] = item.itemData;
		r["source_edit_mode"] = editBase.sourceMode;
		r["base_code_kind"] = editBase.codeKind;
		r["code_hash"] = BuildStableTextHashForRealCode(baseCode);
		r["code"] = LocalToUtf8Text(baseCode);
		outOk = true;
		return JsonToLocalTextForAI(r);
	}

	PageCodeSnapshotEntry snapshot;
	std::string finalCode;
	bool rollbackAttempted = false;
	bool rollbackSucceeded = false;
	std::string writeTrace;
	if (!TryWriteRealPageCodeForAI(
			item,
			baseCode,
			normalizedFullCode,
			"write_file",
			snapshot,
			finalCode,
			writeTrace,
			rollbackAttempted,
			rollbackSucceeded,
			error,
			editBase.editorObject)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = error;
		r["trace"] = LocalToUtf8Text(editBase.trace + "|" + writeTrace);
		r["source_edit_mode"] = editBase.sourceMode;
		r["base_code_kind"] = editBase.codeKind;
		r["rollback_attempted"] = rollbackAttempted;
		r["rollback_succeeded"] = rollbackSucceeded;
		return JsonToLocalTextForAI(r);
	}

	const auto hunks = BuildRealPageStructuredPatch(baseCode, finalCode);
	nlohmann::json r;
	r["ok"] = true;
	r["page_name"] = LocalToUtf8Text(item.name);
	r["type_key"] = item.typeKey;
	r["type_name"] = LocalToUtf8Text(item.typeName);
	r["item_data"] = item.itemData;
	r["trace"] = LocalToUtf8Text(editBase.trace + "|" + writeTrace);
	r["source_edit_mode"] = editBase.sourceMode;
	r["base_code_kind"] = editBase.codeKind;
	r["rollback_attempted"] = rollbackAttempted;
	r["rollback_succeeded"] = rollbackSucceeded;
	r["snapshot_id"] = snapshot.snapshotId;
	r["snapshot_note"] = LocalToUtf8Text(snapshot.note);
	r["base_hash"] = BuildStableTextHashForRealCode(baseCode);
	r["code_hash"] = BuildStableTextHashForRealCode(finalCode);
	r["changed_lines"] = CountStructuredPatchChangedLines(hunks);
	r["code_kind"] = "real_source";
	r["code"] = LocalToUtf8Text(finalCode);
	outOk = true;
	return JsonToLocalTextForAI(r);
}

std::string ExecuteMappedDiffFileToolForAI(
	const nlohmann::json& args,
	const std::string& filePathUtf8,
	const ProgramTreeItemInfo& item,
	bool& outOk)
{
	outOk = false;

	const std::string newCodeInput = [&]() {
		const std::string newCode = GetJsonStringArgumentLocal(args, "new_code");
		return newCode.empty() ? GetJsonStringArgumentLocal(args, "full_code") : newCode;
	}();
	const std::string oldText = GetJsonStringArgumentLocal(args, "old_text");
	const std::string newText = GetJsonStringArgumentLocal(args, "new_text");
	const bool failOnUnmatched = GetJsonBoolArgument(args, "fail_on_unmatched", true);

	std::string error;
	SourceEditBaseForAI editBase;
	if (!TryResolveSourceEditBaseForAI(item, filePathUtf8, std::string(), editBase, error)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = error.empty() ? "read real page code failed" : error;
		r["trace"] = LocalToUtf8Text(editBase.trace);
		r["source_edit_mode"] = editBase.sourceMode.empty() ? SourceEditModeStringForAI(GetActiveSourceEditModeForAI()) : editBase.sourceMode;
		r["base_code_kind"] = editBase.codeKind;
		r["cache_refreshed"] = editBase.cacheRefreshed;
		if (!editBase.currentCode.empty()) {
			r["code"] = LocalToUtf8Text(editBase.currentCode);
			r["code_hash"] = BuildStableTextHashForRealCode(editBase.currentCode);
		}
		return JsonToLocalTextForAI(r);
	}
	const std::string& baseCode = editBase.baseCode;

	std::string candidateCode;
	nlohmann::json editResults = nlohmann::json::array();
	if (!newCodeInput.empty()) {
		candidateCode = NormalizeRealCodeLineBreaksToCrLf(
			item.typeKey == "const_resource"
				? MergeConstResourceLongTextPlaceholdersForAI(newCodeInput, baseCode)
				: newCodeInput);
	}
	else if (!oldText.empty()) {
		std::vector<RealPageTextEditRequest> edits = {
			RealPageTextEditRequest{oldText, newText, false}
		};
		std::vector<RealPageTextEditApplyResult> applyResults;
		if (!ApplyRealPageTextEdits(baseCode, edits, failOnUnmatched, candidateCode, applyResults, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error;
			for (const auto& result : applyResults) {
				editResults.push_back(BuildRealPageTextEditResultJsonForAI(result));
			}
			r["results"] = std::move(editResults);
			return JsonToLocalTextForAI(r);
		}
		for (const auto& result : applyResults) {
			editResults.push_back(BuildRealPageTextEditResultJsonForAI(result));
		}
	}
	else if (args.contains("edits")) {
		std::vector<RealPageTextEditRequest> edits;
		if (!TryParseRealPageTextEditsFromJson(args, "edits", edits, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error;
			return JsonToLocalTextForAI(r);
		}
		std::vector<RealPageTextEditApplyResult> applyResults;
		if (!ApplyRealPageTextEdits(baseCode, edits, failOnUnmatched, candidateCode, applyResults, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error;
			for (const auto& result : applyResults) {
				editResults.push_back(BuildRealPageTextEditResultJsonForAI(result));
			}
			r["results"] = std::move(editResults);
			return JsonToLocalTextForAI(r);
		}
		for (const auto& result : applyResults) {
			editResults.push_back(BuildRealPageTextEditResultJsonForAI(result));
		}
	}
	else {
		return R"({"ok":false,"error":"new_code, full_code, or edit parameters are required"})";
	}

	const auto hunks = BuildRealPageStructuredPatch(baseCode, candidateCode);
	nlohmann::json hunkRows = nlohmann::json::array();
	for (const auto& hunk : hunks) {
		hunkRows.push_back(BuildRealPagePatchHunkJsonForAI(hunk));
	}

	nlohmann::json r;
	r["ok"] = true;
	r["page_name"] = LocalToUtf8Text(item.name);
	r["type_key"] = item.typeKey;
	r["type_name"] = LocalToUtf8Text(item.typeName);
	r["item_data"] = item.itemData;
	r["trace"] = LocalToUtf8Text(editBase.trace);
	r["source_edit_mode"] = editBase.sourceMode;
	r["base_code_kind"] = editBase.codeKind;
	r["from_cache"] = false;
	r["base_hash"] = BuildStableTextHashForRealCode(baseCode);
	r["new_hash"] = BuildStableTextHashForRealCode(candidateCode);
	r["changed_lines"] = CountStructuredPatchChangedLines(hunks);
	r["hunk_count"] = hunks.size();
	r["hunks"] = std::move(hunkRows);
	if (!editResults.empty()) {
		r["results"] = std::move(editResults);
	}
	r["proposed_code"] = LocalToUtf8Text(candidateCode);
	outOk = true;
	return JsonToLocalTextForAI(r);
}

std::string ExecuteMappedRestoreFileSnapshotToolForAI(
	const nlohmann::json& args,
	const ProgramTreeItemInfo& item,
	bool& outOk)
{
	outOk = false;

	const std::string snapshotId = GetJsonStringArgumentLocal(args, "snapshot_id");
	const bool restoreLatest = GetJsonBoolArgument(args, "restore_latest", snapshotId.empty());
	if (!restoreLatest && snapshotId.empty()) {
		return R"({"ok":false,"error":"snapshot_id is required when restore_latest is false"})";
	}

	PageCodeSnapshotEntry snapshot;
	const bool gotSnapshot = restoreLatest
		? PageCodeCacheManager::Instance().GetLatestSnapshot(item.name, item.typeKey, snapshot)
		: PageCodeCacheManager::Instance().GetSnapshot(item.name, item.typeKey, snapshotId, snapshot);
	if (!gotSnapshot) {
		return R"({"ok":false,"error":"snapshot not found"})";
	}

	std::string error;
	std::string baseCode;
	PageCodeCacheEntry cacheEntry;
	std::string trace;
	std::string currentCode;
	bool cacheRefreshed = false;
	std::uintptr_t baseEditorObject = 0;
	if (!TryResolveRealPageWriteBaseForAI(item, std::string(), false, baseCode, cacheEntry, trace, currentCode, cacheRefreshed, error, &baseEditorObject)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = error;
		r["trace"] = LocalToUtf8Text(trace);
		r["cache_refreshed"] = cacheRefreshed;
		if (!currentCode.empty()) {
			r["code"] = LocalToUtf8Text(currentCode);
			r["code_hash"] = BuildStableTextHashForRealCode(currentCode);
		}
		return JsonToLocalTextForAI(r);
	}

	if (BuildStableTextHashForRealCode(snapshot.code) == BuildStableTextHashForRealCode(baseCode)) {
		nlohmann::json r;
		r["ok"] = true;
		r["no_changes"] = true;
		r["snapshot_id"] = snapshot.snapshotId;
		r["snapshot_note"] = LocalToUtf8Text(snapshot.note);
		r["code_hash"] = BuildStableTextHashForRealCode(baseCode);
		r["code"] = LocalToUtf8Text(baseCode);
		outOk = true;
		return JsonToLocalTextForAI(r);
	}

	PageCodeSnapshotEntry createdSnapshot;
	std::string finalCode;
	bool rollbackAttempted = false;
	bool rollbackSucceeded = false;
	if (!TryWriteRealPageCodeForAI(
			item,
			baseCode,
			snapshot.code,
			std::string("restore_file_snapshot:") + snapshot.snapshotId,
			createdSnapshot,
			finalCode,
			trace,
			rollbackAttempted,
			rollbackSucceeded,
			error,
			baseEditorObject)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = error;
		r["trace"] = LocalToUtf8Text(trace);
		r["rollback_attempted"] = rollbackAttempted;
		r["rollback_succeeded"] = rollbackSucceeded;
		return JsonToLocalTextForAI(r);
	}

	const auto hunks = BuildRealPageStructuredPatch(baseCode, finalCode);
	nlohmann::json r;
	r["ok"] = true;
	r["page_name"] = LocalToUtf8Text(item.name);
	r["type_key"] = item.typeKey;
	r["type_name"] = LocalToUtf8Text(item.typeName);
	r["item_data"] = item.itemData;
	r["restored_snapshot_id"] = snapshot.snapshotId;
	r["restored_snapshot_note"] = LocalToUtf8Text(snapshot.note);
	r["snapshot_id"] = createdSnapshot.snapshotId;
	r["snapshot_note"] = LocalToUtf8Text(createdSnapshot.note);
	r["trace"] = LocalToUtf8Text(trace);
	r["rollback_attempted"] = rollbackAttempted;
	r["rollback_succeeded"] = rollbackSucceeded;
	r["base_hash"] = BuildStableTextHashForRealCode(baseCode);
	r["code_hash"] = BuildStableTextHashForRealCode(finalCode);
	r["changed_lines"] = CountStructuredPatchChangedLines(hunks);
	r["code_kind"] = "real_source";
	r["code"] = LocalToUtf8Text(finalCode);
	outOk = true;
	return JsonToLocalTextForAI(r);
}

std::string ExecuteFileMappedRealPageToolForAI(
	const std::string& publicToolName,
	const std::string& argumentsJson,
	bool invalidateOnWrite,
	bool& outOk)
{
	outOk = false;
	const auto totalStart = ToolPerfClock::now();
	LogToolStageForAI(
		"file_mapped_tool",
		"begin|tool=" + publicToolName,
		totalStart);

	nlohmann::json args;
	try {
		args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		LogToolStageForAI(
			"file_mapped_tool",
			"parse_arguments_failed|tool=" + publicToolName +
				"|error=" + ex.what(),
			totalStart);
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = std::string("invalid arguments json: ") + ex.what();
		return JsonToLocalTextForAI(r);
	}

	if (!args.contains("file_path") || !args["file_path"].is_string()) {
		LogToolStageForAI(
			"file_mapped_tool",
			"invalid_arguments|tool=" + publicToolName +
				"|file_path_missing",
			totalStart);
		return R"({"ok":false,"error":"file_path is required"})";
	}

	const std::string filePathUtf8 = args["file_path"].get<std::string>();
	WorkspaceMirror::ProgramItemRef item;
	std::string error;
	LogToolStageForAI(
		"file_mapped_tool",
		"before_resolve_file|tool=" + publicToolName +
			"|file_path=" + filePathUtf8,
		totalStart);
	if (!WorkspaceMirror::ResolveFileToProgramItem(filePathUtf8, item, error)) {
		LogToolStageForAI(
			"file_mapped_tool",
			"after_resolve_file_failed|tool=" + publicToolName +
				"|file_path=" + filePathUtf8 +
				"|error=" + error,
			totalStart);
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = error.empty() ? "resolve file_path failed" : error;
		r["file_path"] = filePathUtf8;
		return JsonToLocalTextForAI(r);
	}
	LogToolStageForAI(
		"file_mapped_tool",
		"after_resolve_file_ok|tool=" + publicToolName +
			"|file_path=" + filePathUtf8 +
			"|page=" + LocalToUtf8Text(item.pageNameLocal) +
			"|kind=" + item.kind,
		totalStart);

	ProgramTreeItemInfo programItem;
	LogToolStageForAI(
		"file_mapped_tool",
		"before_program_item_lookup|tool=" + publicToolName +
			"|page=" + LocalToUtf8Text(item.pageNameLocal) +
			"|kind=" + item.kind,
		totalStart);
	if (!TryGetProgramItemByNameForAI(item.pageNameLocal, item.kind, programItem, error)) {
		LogToolStageForAI(
			"file_mapped_tool",
			"after_program_item_lookup_failed|tool=" + publicToolName +
				"|page=" + LocalToUtf8Text(item.pageNameLocal) +
				"|kind=" + item.kind +
				"|error=" + LocalToUtf8Text(error),
			totalStart);
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = error.empty() ? "program item lookup failed" : error;
		r["file_path"] = filePathUtf8;
		r["mapped_page_name"] = LocalToUtf8Text(item.pageNameLocal);
		r["mapped_kind"] = item.kind;
		return JsonToLocalTextForAI(r);
	}
	LogToolStageForAI(
		"file_mapped_tool",
		"after_program_item_lookup_ok|tool=" + publicToolName +
			"|page=" + LocalToUtf8Text(programItem.name) +
			"|type=" + programItem.typeKey +
			"|item_data=" + std::to_string(programItem.itemData),
		totalStart);

	std::string resultLocal;
	LogToolStageForAI(
		"file_mapped_tool",
		"before_execute_mapped_tool|tool=" + publicToolName,
		totalStart);
	if (publicToolName == "read_real_file") {
		std::string realCode;
		e571::NativeRealPageAccessResult accessResult{};
		if (!TryReadRealPageCodeForAI(programItem, realCode, accessResult, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "copy real page code failed" : error;
			r["trace"] = LocalToUtf8Text(accessResult.trace);
			resultLocal = JsonToLocalTextForAI(r);
			outOk = false;
		}
		else {
			PageCodeCacheEntry cacheEntry;
			PutPageCodeCacheEntryForAI(programItem, realCode, nullptr, &cacheEntry);
			const int offset = (std::max)(0, GetJsonIntArgumentForAI(args, "offset", 0));
			const int limit = GetJsonIntArgumentForAI(args, "limit", 0);
			const std::vector<std::string> lines = SplitLinesForNumberedViewForAI(realCode);
			int returnedLines = 0;
			bool truncated = false;
			const std::string content = BuildNumberedViewForAI(lines, offset, limit, returnedLines, truncated);
			nlohmann::json r;
			r["ok"] = true;
			r["page_name"] = LocalToUtf8Text(programItem.name);
			r["type_key"] = programItem.typeKey;
			r["type_name"] = LocalToUtf8Text(programItem.typeName);
			r["item_data"] = programItem.itemData;
			r["source_edit_mode"] = SourceEditModeStringForAI(GetActiveSourceEditModeForAI());
			r["code_kind"] = "real_source";
			r["code_hash"] = cacheEntry.codeHash;
			r["total_lines"] = lines.size();
			r["offset"] = offset;
			r["returned_lines"] = returnedLines;
			r["truncated"] = truncated;
			r["content"] = LocalToUtf8Text(content);
			r["code"] = LocalToUtf8Text(realCode);
			r["trace"] = LocalToUtf8Text(accessResult.trace);
			resultLocal = JsonToLocalTextForAI(r);
			outOk = true;
		}
	}
	else if (publicToolName == "edit_file") {
		resultLocal = ExecuteMappedEditFileToolForAI(args, filePathUtf8, programItem, outOk);
	}
	else if (publicToolName == "multi_edit_file") {
		resultLocal = ExecuteMappedMultiEditFileToolForAI(args, filePathUtf8, programItem, outOk);
	}
	else if (publicToolName == "write_file") {
		resultLocal = ExecuteMappedWriteFileToolForAI(args, filePathUtf8, programItem, outOk);
	}
	else if (publicToolName == "diff_file") {
		resultLocal = ExecuteMappedDiffFileToolForAI(args, filePathUtf8, programItem, outOk);
	}
	else if (publicToolName == "restore_file_snapshot") {
		resultLocal = ExecuteMappedRestoreFileSnapshotToolForAI(args, programItem, outOk);
	}
	else {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = "unknown file mapped tool: " + publicToolName;
		LogToolStageForAI(
			"file_mapped_tool",
			"unknown_tool|tool=" + publicToolName,
			totalStart);
		return JsonToLocalTextForAI(r);
	}
	LogToolStageForAI(
		"file_mapped_tool",
		"after_execute_mapped_tool|tool=" + publicToolName +
			"|ok=" + std::to_string(outOk ? 1 : 0),
		totalStart);

	nlohmann::json result;
	bool parsed = false;
	try {
		result = nlohmann::json::parse(LocalToUtf8Text(resultLocal));
		parsed = result.is_object();
	}
	catch (...) {
		parsed = false;
	}
	if (!parsed) {
		if (outOk && invalidateOnWrite) {
			WorkspaceMirror::InvalidateMirror();
		}
		return resultLocal;
	}

	result["tool"] = publicToolName;
	result["file_path"] = filePathUtf8;
	result["mapped_page_name"] = LocalToUtf8Text(item.pageNameLocal);
	result["mapped_kind"] = item.kind;

	const bool ok = result.value("ok", false);
	const bool noChanges = result.value("no_changes", false);
	if (item.fixedTable && ok && result.contains("code_kind")) {
		result["code_kind"] = "fixed_table_real_page";
	}
	if (ok && invalidateOnWrite && !noChanges) {
		bool mirrorUpdated = false;
		std::string mirrorUpdateError;
		if (!item.fixedTable &&
			GetActiveSourceEditModeForAI() == AISourceEditMode::MirrorSourceBase &&
			result.contains("code") &&
			result["code"].is_string()) {
			const std::string finalCodeLocal = Utf8ToLocalText(result["code"].get<std::string>());
			mirrorUpdated = WorkspaceMirror::UpdateMirrorTextFile(filePathUtf8, finalCodeLocal, mirrorUpdateError);
			if (mirrorUpdated) {
				result["workspace_mirror_updated"] = true;
				result["workspace_mirror_invalidated"] = false;
			}
		}

		if (!mirrorUpdated) {
			WorkspaceMirror::InvalidateMirror();
			result["workspace_mirror_invalidated"] = true;
			if (!mirrorUpdateError.empty()) {
				result["workspace_mirror_update_error"] = LocalToUtf8Text(mirrorUpdateError);
			}
		}

		if (item.fixedTable) {
			result["code_kind"] = "fixed_table_real_page";
			result["post_write_real_page_refreshed"] = true;
		}
	}

	if (!ok &&
		!item.fixedTable &&
		(publicToolName == "edit_file" || publicToolName == "multi_edit_file") &&
		!result.contains("code")) {
		std::string realCode;
		std::string realCodeHash;
		std::string readTrace;
		std::string readError;
		if (TryReadMappedRealPageCodeForAI(item, realCode, realCodeHash, readTrace, readError)) {
			result["real_code"] = LocalToUtf8Text(realCode);
			result["real_code_hash"] = realCodeHash;
			result["real_code_trace"] = LocalToUtf8Text(readTrace);
			result["real_code_note"] = "old_text matching is performed against this real IDE page text";
		}
	}

	return JsonToLocalTextForAI(result);
}

bool TryReadMappedRealPageCodeForAI(
	const WorkspaceMirror::ProgramItemRef& item,
	std::string& outCode,
	std::string& outCodeHash,
	std::string& outTrace,
	std::string& outError)
{
	outCode.clear();
	outCodeHash.clear();
	outTrace.clear();
	outError.clear();

	ProgramTreeItemInfo programItem;
	if (!TryGetProgramItemByNameForAI(item.pageNameLocal, item.kind, programItem, outError)) {
		if (outError.empty()) {
			outError = "program item lookup failed";
		}
		return false;
	}

	e571::NativeRealPageAccessResult accessResult{};
	if (!TryReadRealPageCodeForAI(programItem, outCode, accessResult, outError)) {
		outTrace = accessResult.trace;
		if (outError.empty()) {
			outError = "copy real page code failed";
		}
		return false;
	}

	PageCodeCacheEntry cacheEntry;
	PutPageCodeCacheEntryForAI(programItem, outCode, nullptr, &cacheEntry);
	outCodeHash = cacheEntry.codeHash;
	outTrace = accessResult.trace;
	return true;
}

std::string ExecuteToolCallOnMainThreadImpl(const std::string& toolName, const std::string& argumentsJson, bool& outOk)
{
	outOk = false;

	if (toolName == "update_plan") {
		std::string resultJsonLocal;
		if (AIChatFeature::UpdatePlanFromTool(argumentsJson, resultJsonLocal, outOk)) {
			return resultJsonLocal;
		}

		nlohmann::json r;
		r["ok"] = false;
		r["error"] = "update_plan failed";
		return JsonToLocalTextForAI(r);
	}

	if (toolName == "refresh_workspace_mirror") {
		nlohmann::json args = nlohmann::json::object();
		try {
			args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return JsonToLocalTextForAI(r);
		}
		std::string requestedMode = "auto";
		WorkspaceMirror::RefreshMode refreshMode = WorkspaceMirror::RefreshMode::Auto;
		if (args.contains("mode") && args["mode"].is_string()) {
			requestedMode = ToLowerAsciiCopyLocal(TrimAsciiCopy(args["mode"].get<std::string>()));
			if (requestedMode == "full") {
				refreshMode = WorkspaceMirror::RefreshMode::Full;
			}
			else if (requestedMode == "main_only" || requestedMode == "source_only") {
				requestedMode = "main_only";
				refreshMode = WorkspaceMirror::RefreshMode::MainOnly;
			}
			else if (requestedMode.empty() || requestedMode == "auto") {
				requestedMode = "auto";
			}
			else {
				nlohmann::json r;
				r["ok"] = false;
				r["error"] = "mode must be auto, main_only, or full";
				return JsonToLocalTextForAI(r);
			}
		}
		std::string error;
		std::string mode;
		if (!WorkspaceMirror::RefreshMirror(error, &mode, refreshMode)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "refresh workspace mirror failed" : error;
			r["requested_mode"] = requestedMode;
			return JsonToLocalTextForAI(r);
		}

		std::filesystem::path mirrorRoot;
		if (!WorkspaceMirror::GetMirrorRoot(mirrorRoot, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "query refreshed workspace mirror failed" : error;
			r["requested_mode"] = requestedMode;
			r["refresh_mode"] = mode;
			return JsonToLocalTextForAI(r);
		}

		std::vector<std::string> files;
		if (!WorkspaceMirror::ListMirrorFiles(files, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "list refreshed workspace mirror failed" : error;
			r["requested_mode"] = requestedMode;
			r["refresh_mode"] = mode;
			r["mirror_root"] = LocalToUtf8Text(mirrorRoot.string());
			return JsonToLocalTextForAI(r);
		}

		nlohmann::json r;
		r["ok"] = true;
		r["requested_mode"] = requestedMode;
		r["refresh_mode"] = mode;
		r["mirror_root"] = LocalToUtf8Text(mirrorRoot.string());
		r["file_count"] = files.size();
		r["note"] = "workspace mirror refreshed from current IDE memory; call read_file, search_code, or list_files after this";
		outOk = true;
		return JsonToLocalTextForAI(r);
	}

	if (WorkspaceFileTools::CanHandleTool(toolName)) {
		return WorkspaceFileTools::ExecuteTool(toolName, argumentsJson, outOk);
	}

	if (toolName == "edit_file") {
		return ExecuteFileMappedRealPageToolForAI(toolName, argumentsJson, true, outOk);
	}
	if (toolName == "read_real_file") {
		return ExecuteFileMappedRealPageToolForAI(toolName, argumentsJson, false, outOk);
	}
	if (toolName == "multi_edit_file") {
		return ExecuteFileMappedRealPageToolForAI(toolName, argumentsJson, true, outOk);
	}
	if (toolName == "write_file") {
		return ExecuteFileMappedRealPageToolForAI(toolName, argumentsJson, true, outOk);
	}
	if (toolName == "diff_file") {
		return ExecuteFileMappedRealPageToolForAI(toolName, argumentsJson, false, outOk);
	}
	if (toolName == "restore_file_snapshot") {
		return ExecuteFileMappedRealPageToolForAI(toolName, argumentsJson, true, outOk);
	}

	if (toolName == "get_current_page_info") {
		const std::string sourceFilePath = RefreshCurrentSourceFilePathForAI();
		std::string pageName;
		std::string pageType;
		std::string pageNameTrace;
		if (!IDEFacade::Instance().GetCurrentPageName(pageName, &pageType, &pageNameTrace)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = "GetCurrentPageName failed";
			r["page_name_trace"] = pageNameTrace;
			return JsonToLocalTextForAI(r);
		}

		nlohmann::json r;
		r["ok"] = true;
		r["source_file_path"] = LocalToUtf8Text(sourceFilePath);
		r["page_name"] = LocalToUtf8Text(pageName);
		r["page_type"] = LocalToUtf8Text(pageType);
		r["page_name_trace"] = pageNameTrace;
		LocalMcpServer::UpdateInstanceHints(sourceFilePath, pageName, pageType);
		outOk = true;
		return JsonToLocalTextForAI(r);
	}

	if (toolName == "get_current_eide_info") {
		const std::string sourceFilePath = RefreshCurrentSourceFilePathForAI();
		std::string pageName;
		std::string pageType;
		std::string pageNameTrace;
		const bool pageNameOk = IDEFacade::Instance().GetCurrentPageName(pageName, &pageType, &pageNameTrace);

		std::string mainWindowTitle;
		if (HWND mainWindow = IDEFacade::Instance().GetMainWindow();
			mainWindow != nullptr && IsWindow(mainWindow)) {
			char title[512] = {};
			GetWindowTextA(mainWindow, title, static_cast<int>(sizeof(title)));
			mainWindowTitle = title;
		}

		// 从主窗口标题解析项目（源文件）类型。
		// 标题格式示例：
		//   易语言5.71 - [xxx.e] - Windows窗口程序 - [程序集: ...]
		//   KanColleEx - C:\...\KanCore56.e [无编译条件] - Windows易语言模块 - [程序集: ...]
		struct ProjectKindInfo {
			const char* titleKeyword;   // 标题中出现的关键字
			const char* type;           // compile_with_output_path target 枚举值
			const char* label;          // 中文名称
			bool supportsStaticCompile; // 是否支持静态编译
		};
		static const ProjectKindInfo kProjectKinds[] = {
			{" Windows易语言模块 ",  "ecom",            "Windows易语言模块",  false},
			{" Windows动态链接库 ",  "win_dll",         "Windows动态链接库",  true},
			{" Windows控制台程序 ",  "win_console_exe", "Windows控制台程序",  true},
			{" Windows窗口程序 ",    "win_exe",         "Windows窗口程序",    true},
		};

		std::string projectType    = "unknown";
		std::string projectTypeLabel = "\xe6\x9c\xaa\xe7\x9f\xa5"; // UTF-8: 未知
		bool        projectSupportsStaticCompile = false;
		for (const auto& k : kProjectKinds) {
			if (mainWindowTitle.find(k.titleKeyword) != std::string::npos) {
				projectType                  = k.type;
				projectTypeLabel             = k.label;
				projectSupportsStaticCompile = k.supportsStaticCompile;
				break;
			}
		}

		// 支持的编译目标及模式（供 AI 决策 compile_with_output_path 参数）。
		nlohmann::json compileModes = nlohmann::json::array();
		compileModes.push_back("compile");
		if (projectSupportsStaticCompile) {
			compileModes.push_back("static_compile");
		}

		LocalMcpServer::UpdateInstanceHints(sourceFilePath, pageName, pageType);

		nlohmann::json r;
		r["ok"] = true;
		r["process_id"] = GetCurrentProcessId();
		r["process_path"] = LocalToUtf8Text(GetCurrentProcessPathForAI());
		r["process_name"] = LocalToUtf8Text(GetCurrentProcessNameForAI());
		r["source_file_path"] = LocalToUtf8Text(sourceFilePath);
		r["source_file_name"] = LocalToUtf8Text(sourceFilePath.empty() ? std::string() : std::filesystem::path(sourceFilePath).filename().string());
		r["source_directory"] = LocalToUtf8Text(sourceFilePath.empty() ? std::string() : std::filesystem::path(sourceFilePath).parent_path().string());
		r["source_file_exists"] = !sourceFilePath.empty() && std::filesystem::exists(std::filesystem::path(sourceFilePath));
		r["page_name_ok"] = pageNameOk;
		r["current_page_name"] = LocalToUtf8Text(pageName);
		r["current_page_type"] = LocalToUtf8Text(pageType);
		r["page_name_trace"] = pageNameTrace;
		r["main_window_title"] = LocalToUtf8Text(mainWindowTitle);
		r["project_type"] = projectType;
		r["project_type_label"] = LocalToUtf8Text(projectTypeLabel);
		r["project_supported_compile_modes"] = compileModes;
		r["mcp_running"] = LocalMcpServer::IsRunning();
		r["mcp_instance_id"] = LocalMcpServer::GetInstanceId();
		r["mcp_port"] = LocalMcpServer::GetBoundPort();
		r["mcp_endpoint"] = LocalMcpServer::GetEndpoint();
		outOk = true;
		return JsonToLocalTextForAI(r);
	}

	if (toolName == "refresh_dependency_catalog") {
		return BuildRefreshDependencyCatalogJsonForAI(argumentsJson, outOk);
	}

	if (toolName == "search_available_modules") {
		return BuildSearchAvailableModulesJsonForAI(argumentsJson, outOk);
	}

	if (toolName == "search_available_support_libraries") {
		return BuildSearchAvailableSupportLibrariesJsonForAI(argumentsJson, outOk);
	}

	if (toolName == "list_imported_modules") {
		return BuildListImportedModulesJsonOnMainThread(outOk);
	}

	if (toolName == "add_module_to_project") {
		return BuildAddModuleToProjectJsonOnMainThread(argumentsJson, outOk);
	}

	if (toolName == "remove_module_from_project") {
		return BuildRemoveModuleFromProjectJsonOnMainThread(argumentsJson, outOk);
	}

	if (toolName == "add_support_library_to_project") {
		return BuildAddSupportLibraryToProjectJsonOnMainThread(argumentsJson, outOk);
	}

	if (toolName == "compile_with_output_path") {
		nlohmann::json args;
		try {
			args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return JsonToLocalTextForAI(r);
		}

		const std::string target = args.contains("target") && args["target"].is_string()
			? ToLowerAsciiCopyLocal(args["target"].get<std::string>())
			: std::string();
		const std::string outputPath = args.contains("output_path") && args["output_path"].is_string()
			? Utf8ToLocalText(args["output_path"].get<std::string>())
			: std::string();
		const bool staticCompile = args.contains("static_compile") && args["static_compile"].is_boolean()
			? args["static_compile"].get<bool>()
			: false;

		if (TrimAsciiCopy(target).empty()) {
			return R"({"ok":false,"error":"target is required"})";
		}
		if (TrimAsciiCopy(outputPath).empty()) {
			return R"({"ok":false,"error":"output_path is required"})";
		}

		IDEFacade::CompileOutputKind kind = IDEFacade::CompileOutputKind::WinExe;
		if (target == "win_exe") {
			kind = IDEFacade::CompileOutputKind::WinExe;
		}
		else if (target == "win_console_exe") {
			kind = IDEFacade::CompileOutputKind::WinConsoleExe;
		}
		else if (target == "win_dll") {
			kind = IDEFacade::CompileOutputKind::WinDll;
		}
		else if (target == "ecom") {
			kind = IDEFacade::CompileOutputKind::Ecom;
		}
		else {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = "unsupported target";
			return JsonToLocalTextForAI(r);
		}

		// 编译前：快照输出窗口文本，记录开始时间戳（用于判断产物是否刷新）。
		std::string preOutputText;
		IDEFacade::Instance().GetOutputWindowText(preOutputText);
		const __time64_t compileStartTimestamp = _time64(nullptr);

		std::string normalizedPath;
		std::string diagnostics;
		const bool compileOk = IDEFacade::Instance().CompileWithOutputPath(
			kind,
			outputPath,
			staticCompile,
			&normalizedPath,
			&diagnostics);

		// 编译完成后，读取光标所在行内容及当前程序集。
		// 编译失败时 IDE 通常会跳转到出错行，此处内容可辅助 AI 定位问题。
		int caretRow = -1;
		int caretCol = -1;
		std::string caretLineText;
		std::string caretPageName;
		std::string caretPageType;
		IDEFacade::Instance().GetCaretPosition(caretRow, caretCol);
		if (caretRow >= 0) {
			caretLineText = IDEFacade::Instance().GetRowFullText(caretRow);
		}
		IDEFacade::Instance().GetCurrentPageName(caretPageName, &caretPageType);

		if (!compileOk) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = diagnostics.empty() ? "compile_with_output_path_failed" : diagnostics;
			r["target"] = target;
			r["static_compile"] = staticCompile;
			r["caret_row"] = caretRow;
			r["caret_line_text"] = LocalToUtf8Text(caretLineText);
			r["caret_page_name"] = LocalToUtf8Text(caretPageName);
			r["caret_page_type"] = LocalToUtf8Text(caretPageType);
			return JsonToLocalTextForAI(r);
		}

		if (diagnostics == "compile_invoked_dialog_pending" ||
			diagnostics == "compile_invoked_dialog_suppressed") {
			const auto waitDeadline = std::chrono::steady_clock::now() + std::chrono::seconds(30);
			bool outputFileReadySeen = false;
			std::chrono::steady_clock::time_point outputFileReadySince = {};
			while (std::chrono::steady_clock::now() < waitDeadline) {
				MSG msg = {};
				while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE)) {
					TranslateMessage(&msg);
					DispatchMessageW(&msg);
				}

				const bool requestConsumed = WasSilentCompileOutputPathRequestConsumed();
				if (requestConsumed) {
					diagnostics = "compile_invoked_dialog_suppressed";
				}
				if (!IsSilentCompileOutputPathRequestActive() && !requestConsumed) {
					break;
				}

				bool outputFileReady = false;
				if (!normalizedPath.empty()) {
					struct __stat64 pendingFileStat = {};
					if (_stat64(normalizedPath.c_str(), &pendingFileStat) == 0 &&
						pendingFileStat.st_mtime >= compileStartTimestamp - 1) {
						outputFileReady = true;
					}
				}
				if (outputFileReady) {
					if (!outputFileReadySeen) {
						outputFileReadySeen = true;
						outputFileReadySince = std::chrono::steady_clock::now();
					}
					if (std::chrono::steady_clock::now() - outputFileReadySince >= std::chrono::milliseconds(1200)) {
						break;
					}
				}
				else {
					outputFileReadySeen = false;
				}
				Sleep(20);
			}
			CancelSilentCompileOutputPathRequest();
		}

		// 编译后：读取输出窗口，提取本次编译产生的新内容。
		// 易语言编译为同步操作，函数返回时编译输出已写入完毕。
		std::string postOutputText;
		IDEFacade::Instance().GetOutputWindowText(postOutputText);

		std::string newOutput;
		if (postOutputText.size() > preOutputText.size()) {
			// IDE 追加输出模式：只取新增部分
			newOutput = postOutputText.substr(preOutputText.size());
		}
		else {
			// IDE 清空后重写，或编译为异步（输出尚未刷新）：返回当前全部文本
			newOutput = postOutputText;
		}

		// 检查产物文件是否在编译启动后被创建/更新，用于确认编译是否成功。
		bool outputFileExists = false;
		bool outputFileModifiedAfterCompile = false;
		if (!normalizedPath.empty()) {
			struct __stat64 fileStat = {};
			if (_stat64(normalizedPath.c_str(), &fileStat) == 0) {
				outputFileExists = true;
				outputFileModifiedAfterCompile = (fileStat.st_mtime >= compileStartTimestamp - 1);
			}
		}

		nlohmann::json r;
		r["ok"] = true;
		r["target"] = target;
		r["static_compile"] = staticCompile;
		r["output_path"] = LocalToUtf8Text(normalizedPath);
		r["output_window_text"] = LocalToUtf8Text(newOutput);
		r["output_file_exists"] = outputFileExists;
		r["output_file_modified_after_compile"] = outputFileModifiedAfterCompile;
		r["trace"] = diagnostics;
		r["caret_row"] = caretRow;
		r["caret_line_text"] = LocalToUtf8Text(caretLineText);
		r["caret_page_name"] = LocalToUtf8Text(caretPageName);
		r["caret_page_type"] = LocalToUtf8Text(caretPageType);
		outOk = true;
		return JsonToLocalTextForAI(r);
	}

	nlohmann::json r;
	r["ok"] = false;
	r["error"] = "unknown tool: " + toolName;
	return JsonToLocalTextForAI(r);
}

// 顶层异常防线：工具调用经 SendMessage 在主线程的 WndProc 中执行
// （HandleToolExecRequest）。若任何工具内部抛出未捕获异常，异常会逃出窗口
// 过程、穿过 Win32 消息派发边界，触发 MSVC 运行时 abort（“abnormal program
// termination”），整个 IDE 进程崩溃。这里统一兜底，把异常转成失败结果返回，
// 保证无论工具内部发生什么都不会崩进程。
std::string ExecuteToolCallOnMainThread(const std::string& toolName, const std::string& argumentsJson, bool& outOk)
{
	try {
		return ExecuteToolCallOnMainThreadImpl(toolName, argumentsJson, outOk);
	}
	catch (const std::exception& ex) {
		outOk = false;
		OutputStringToELog(std::string("[Tool] tool_execution_exception tool=") + toolName + " what=" + ex.what());
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = std::string("tool execution failed: ") + ex.what();
		r["exception"] = true;
		return JsonToLocalTextForAI(r);
	}
	catch (...) {
		outOk = false;
		OutputStringToELog(std::string("[Tool] tool_execution_exception tool=") + toolName + " what=<unknown>");
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = "tool execution failed: unknown exception";
		r["exception"] = true;
		return JsonToLocalTextForAI(r);
	}
}

#endif


