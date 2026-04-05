#include "AIChatTooling.h"
#include "AIChatToolingInternal.h"

#include <Windows.h>

#include <algorithm>
#include <array>
#include <cctype>
#include <chrono>
#include <CommCtrl.h>
#include <cstring>
#include <filesystem>
#include <fstream>
#include <format>
#include <limits>
#include <mutex>
#include <regex>
#include <string>
#include <tlhelp32.h>
#include <unordered_map>
#include <unordered_set>
#include <vector>
#include <wincrypt.h>

#include "..\\thirdparty\\json.hpp"
#include "..\\elib\\lib2.h"

#include "IDEFacade.h"
#include "EideInternalTextBridge.h"
#include "Global.h"
#include "LocalMcpServer.h"
#include "PageCodeCacheManager.h"
#include "PathHelper.h"
#include "ProjectSourceCacheManager.h"
#include "RealPageCodeToolSupport.h"
#include "WindowHelper.h"
#if defined(_M_IX86)
#include "direct_global_search_debug.hpp"
#include "native_module_public_info.hpp"

struct ModulePublicInfoCacheEntry {
	std::string md5;
	e571::ModulePublicInfoDump dump;
	std::string error;
	bool ok = false;
	std::string normalizedText;
	std::vector<std::string> lines;
};

struct SupportLibraryDumpCacheEntry {
	std::string md5;
	nlohmann::json dumpJson = nlohmann::json::object();
	std::string error;
	bool ok = false;
	std::string sourceKind;
	std::string normalizedText;
	std::vector<std::string> lines;
};

std::mutex g_modulePublicInfoCacheMutex;
std::unordered_map<std::string, ModulePublicInfoCacheEntry> g_modulePublicInfoCache;
std::mutex g_supportLibraryDumpCacheMutex;
std::unordered_map<std::string, SupportLibraryDumpCacheEntry> g_supportLibraryDumpCache;

struct ProgramTreeItemInfo {
	int depth = 0;
	std::string name;
	unsigned int itemData = 0;
	int image = -1;
	int selectedImage = -1;
	std::string typeKey;
	std::string typeName;
};

struct KeywordSearchResultInfo {
	std::string pageName;
	std::string pageTypeKey;
	std::string pageTypeName;
	int lineNumber = -1;
	std::string text;
	int hitIndexInPage = 0;
	int hitTotalInPage = 0;
	int sameTextOccurrenceIndex = 0;
	int sameTextOccurrenceTotal = 0;
};

constexpr const char* kProgramTreeConstantTablePageNameForAI = "常量表...";
constexpr const char* kSearchResultConstantDefinitionTableNameForAI = "常量定义表";
constexpr const char* kSearchResultUserDataTypeTableNameForAI = "自定义数据类型表";
constexpr const char* kProgramTreeUserDataTypePageNameForAI = "自定义数据类型";
constexpr const char* kSearchResultDllCommandTableNameForAI = "DLL命令定义表";
constexpr const char* kProgramTreeDllCommandPageNameForAI = "Dll命令";
constexpr const char* kSearchResultGlobalVariableTableNameForAI = "全局变量表";
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
	const std::string trimmed = TrimAsciiCopy(name);
	if (trimmed == kSearchResultConstantDefinitionTableNameForAI) {
		return kProgramTreeConstantTablePageNameForAI;
	}
	if (trimmed == kSearchResultUserDataTypeTableNameForAI) {
		return kProgramTreeUserDataTypePageNameForAI;
	}
	if (trimmed == kSearchResultDllCommandTableNameForAI) {
		return kProgramTreeDllCommandPageNameForAI;
	}
	if (trimmed == kSearchResultGlobalVariableTableNameForAI) {
		return kProgramTreeGlobalVariablePageNameForAI;
	}
	return trimmed;
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
	outItem.typeKey = "const_resource";
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

	ProgramTreeItemInfo constantTableItem;
	if (TryFindProgramTreeItemByExactNameForAI(kProgramTreeConstantTablePageNameForAI, constantTableItem) &&
		seenItemData.insert(constantTableItem.itemData).second) {
		outItems.push_back(std::move(constantTableItem));
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
	outMatchCount = CountExactOccurrencesForAI(source, oldText);
	if (outMatchCount != 1) {
		return false;
	}

	const size_t pos = source.find(oldText);
	outResult.reserve(source.size() - oldText.size() + newText.size());
	outResult.append(source.substr(0, pos));
	outResult.append(newText);
	outResult.append(source.substr(pos + oldText.size()));
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

int GetJsonIntArgument(const nlohmann::json& args, const char* key, int defaultValue)
{
	if (!args.contains(key) || !args[key].is_number_integer()) {
		return defaultValue;
	}
	return args[key].get<int>();
}

size_t GetJsonSizeArgument(const nlohmann::json& args, const char* key, size_t defaultValue)
{
	if (!args.contains(key) || !args[key].is_number_integer()) {
		return defaultValue;
	}
	const int value = args[key].get<int>();
	return value <= 0 ? 0u : static_cast<size_t>(value);
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

bool TryReadRealPageCodeForAI(
	const ProgramTreeItemInfo& item,
	std::string& outCode,
	e571::NativeRealPageAccessResult& outAccessResult,
	std::string& outError)
{
	outCode.clear();
	outAccessResult = {};
	outError.clear();

	if (!e571::GetRealPageCodeByProgramTreeItemData(
			item.itemData,
			GetCurrentProcessImageBaseForAI(),
			&outCode,
			&outAccessResult)) {
		outError = "get real page code failed";
		return false;
	}

	outCode = NormalizeRealCodeLineBreaksToCrLf(outCode);
	return true;
}

bool TryLoadRealPageCodeForReadForAI(
	const ProgramTreeItemInfo& item,
	bool refreshCache,
	std::string& outCode,
	PageCodeCacheEntry& outCacheEntry,
	bool& outFromCache,
	std::string& outTrace,
	std::string& outError)
{
	outCode.clear();
	outCacheEntry = {};
	outFromCache = false;
	outTrace.clear();
	outError.clear();

	if (!refreshCache && PageCodeCacheManager::Instance().Get(item.name, item.typeKey, outCacheEntry)) {
		if (outCacheEntry.codeHash.empty()) {
			PutPageCodeCacheEntryForAI(item, outCacheEntry.code, &outCacheEntry, &outCacheEntry);
		}
		outCode = outCacheEntry.code;
		outFromCache = true;
		outTrace = "cache";
		return true;
	}

	e571::NativeRealPageAccessResult accessResult{};
	if (!TryReadRealPageCodeForAI(item, outCode, accessResult, outError)) {
		outTrace = accessResult.trace;
		return false;
	}

	PutPageCodeCacheEntryForAI(item, outCode, nullptr, &outCacheEntry);
	outFromCache = false;
	outTrace = accessResult.trace;
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
	std::string& outError)
{
	outBaseCode.clear();
	outCacheEntry = {};
	outLiveTrace.clear();
	outCurrentCode.clear();
	outCacheRefreshed = false;
	outError.clear();

	const bool hasCache = PageCodeCacheManager::Instance().Get(item.name, item.typeKey, outCacheEntry);
	if (requireCache && !hasCache) {
		outError = "page cache missing, call read_program_item_real_code or get_program_item_real_code first";
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
	const std::string liveHash = BuildStableTextHashForRealCode(liveCode);
	const std::string normalizedExpectedHash = ToLowerAsciiCopyLocal(TrimAsciiCopy(expectedBaseHash));

	if (hasCache) {
		outCacheEntry = BuildPageCodeCacheEntryForAI(item, outCacheEntry.code, &outCacheEntry);
	}

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
	std::string& outError)
{
	outSnapshot = {};
	outFinalCode.clear();
	outTrace.clear();
	outRollbackAttempted = false;
	outRollbackSucceeded = false;
	outError.clear();

	PageCodeCacheEntry existingEntry;
	const bool hasCache = PageCodeCacheManager::Instance().Get(item.name, item.typeKey, existingEntry);
	PutPageCodeCacheEntryForAI(item, baseCode, hasCache ? &existingEntry : nullptr, &existingEntry);
	if (!PageCodeCacheManager::Instance().AddSnapshot(item.name, item.typeKey, baseCode, snapshotNote, outSnapshot)) {
		outError = "add page snapshot failed";
		return false;
	}

	const std::string normalizedBaseCode = NormalizeRealCodeLineBreaksToCrLf(baseCode);
	const std::string normalizedNewCode = NormalizeRealCodeLineBreaksToCrLf(newCode);

	e571::NativeRealPageAccessResult writeResult{};
	if (!e571::ReplaceRealPageCodeByProgramTreeItemData(
			item.itemData,
			GetCurrentProcessImageBaseForAI(),
			normalizedNewCode,
			&normalizedBaseCode,
			&writeResult)) {
		outRollbackAttempted = writeResult.rollbackAttempted;
		outRollbackSucceeded = writeResult.rollbackSucceeded;
		outTrace = writeResult.trace;

		std::string recoveredCode;
		e571::NativeRealPageAccessResult recoveredReadResult{};
		std::string recoveredReadError;
		if (TryReadRealPageCodeForAI(item, recoveredCode, recoveredReadResult, recoveredReadError)) {
			std::string recoveredVerifyMode;
			if (VerifyRealPageCodeMatchesForAI(normalizedNewCode, recoveredCode, &recoveredVerifyMode)) {
				PageCodeCacheEntry postSnapshotEntry;
				const bool hasPostSnapshotEntry = PageCodeCacheManager::Instance().Get(item.name, item.typeKey, postSnapshotEntry);
				PutPageCodeCacheEntryForAI(item, recoveredCode, hasPostSnapshotEntry ? &postSnapshotEntry : nullptr, nullptr);

				outFinalCode = recoveredCode;
				outTrace =
					writeResult.trace +
					"|post_failure_verify_ok_" +
					recoveredVerifyMode +
					"|" +
					recoveredReadResult.trace;
				outError.clear();
				return true;
			}

			outTrace =
				writeResult.trace +
				"|post_failure_verify_mismatch|" +
				recoveredReadResult.trace;
		}
		else if (!recoveredReadError.empty()) {
			outTrace =
				writeResult.trace +
				"|post_failure_verify_read_failed|" +
				recoveredReadResult.trace;
		}

		outError = "replace real page code failed";
		return false;
	}

	std::string finalCode;
	e571::NativeRealPageAccessResult finalReadResult{};
	if (!TryReadRealPageCodeForAI(item, finalCode, finalReadResult, outError)) {
		outRollbackAttempted = writeResult.rollbackAttempted;
		outRollbackSucceeded = writeResult.rollbackSucceeded;
		outTrace = writeResult.trace + "|" + finalReadResult.trace;
		outError = "final real page verification failed";
		return false;
	}

	PageCodeCacheEntry postSnapshotEntry;
	const bool hasPostSnapshotEntry = PageCodeCacheManager::Instance().Get(item.name, item.typeKey, postSnapshotEntry);
	PutPageCodeCacheEntryForAI(item, finalCode, hasPostSnapshotEntry ? &postSnapshotEntry : nullptr, nullptr);

	outFinalCode = finalCode;
	outRollbackAttempted = writeResult.rollbackAttempted;
	outRollbackSucceeded = writeResult.rollbackSucceeded;
	outTrace = writeResult.trace + "|" + finalReadResult.trace;
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

nlohmann::json BuildRealPageSymbolJsonForAI(const RealPageSymbolInfo& symbol)
{
	nlohmann::json row;
	row["name"] = LocalToUtf8Text(symbol.name);
	row["kind"] = symbol.kind;
	row["display_name"] = LocalToUtf8Text(symbol.displayName);
	if (!symbol.parentName.empty()) {
		row["parent_name"] = LocalToUtf8Text(symbol.parentName);
	}
	row["declaration_line"] = symbol.declarationLine;
	row["start_line"] = symbol.startLine;
	row["end_line"] = symbol.endLine;
	row["is_event_handler"] = symbol.isEventHandler;
	return row;
}

nlohmann::json BuildRealPageSearchMatchJsonForAI(const RealPageSearchMatch& match)
{
	nlohmann::json before = nlohmann::json::array();
	for (const auto& line : match.beforeContextLines) {
		before.push_back(LocalToUtf8Text(line));
	}
	nlohmann::json after = nlohmann::json::array();
	for (const auto& line : match.afterContextLines) {
		after.push_back(LocalToUtf8Text(line));
	}

	nlohmann::json row;
	row["line_number"] = match.lineNumber;
	row["match_column"] = match.matchColumn;
	row["line_text"] = LocalToUtf8Text(match.lineText);
	row["before_context_lines"] = std::move(before);
	row["after_context_lines"] = std::move(after);
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

bool ParsePageNameFromSearchDisplayText(const std::string& displayText, std::string& outPageName)
{
	outPageName.clear();
	const size_t arrowPos = displayText.find(" -> ");
	if (arrowPos != std::string::npos && arrowPos > 0) {
		outPageName = NormalizeProgramPageNameForAI(displayText.substr(0, arrowPos));
		return !outPageName.empty();
	}

	const size_t colonPos = displayText.find(": ");
	const size_t parenPos = displayText.find(" (");
	if (parenPos != std::string::npos && parenPos > 0 && colonPos != std::string::npos && colonPos > parenPos) {
		outPageName = NormalizeProgramPageNameForAI(displayText.substr(0, parenPos));
		return !outPageName.empty();
	}

	if (parenPos != std::string::npos && parenPos > 0) {
		outPageName = NormalizeProgramPageNameForAI(displayText.substr(0, parenPos));
		return !outPageName.empty();
	}

	return false;
}

std::string BuildProjectSearchGroupingKeyForAI(const std::string& pageName, const std::string& text)
{
	return ToLowerAsciiCopyLocal(TrimAsciiCopy(pageName)) + "\n" + ToLowerAsciiCopyLocal(TrimAsciiCopy(text));
}

std::vector<std::string> BuildSearchTextCandidatesForAI(const std::string& searchText, const std::string& pageName)
{
	std::vector<std::string> candidates;
	const auto pushUnique = [&](std::string text) {
		text = TrimAsciiCopy(text);
		if (text.empty()) {
			return;
		}
		for (const auto& existing : candidates) {
			if (existing == text) {
				return;
			}
		}
		candidates.push_back(std::move(text));
	};

	pushUnique(searchText);

	const size_t arrowPos = searchText.find(" -> ");
	if (arrowPos != std::string::npos) {
		pushUnique(searchText.substr(arrowPos + 4));
	}

	const size_t colonPos = searchText.find(": ");
	const size_t parenPos = searchText.find(" (");
	if (colonPos != std::string::npos && parenPos != std::string::npos && colonPos > parenPos) {
		pushUnique(searchText.substr(colonPos + 2));
	}

	const std::string trimmedPageName = TrimAsciiCopy(pageName);
	if (!trimmedPageName.empty() && searchText.rfind(trimmedPageName, 0) == 0) {
		std::string suffix = searchText.substr(trimmedPageName.size());
		suffix = TrimAsciiCopy(suffix);
		if (suffix.rfind("->", 0) == 0) {
			suffix = TrimAsciiCopy(suffix.substr(2));
		}
		pushUnique(suffix);
	}

	return candidates;
}

bool TryResolveProjectSearchLineNumberForAI(
	const std::vector<std::string>& pageLines,
	const std::vector<std::string>& candidates,
	int sameTextOccurrenceIndex,
	int hintLineNumber,
	int& outResolvedLineNumber,
	std::string& outResolvedBy,
	std::string& outMatchedCandidate,
	int& outMatchedCandidateTotal)
{
	outResolvedLineNumber = -1;
	outResolvedBy.clear();
	outMatchedCandidate.clear();
	outMatchedCandidateTotal = 0;

	if (pageLines.empty()) {
		outResolvedBy = "page_empty";
		return false;
	}

	std::vector<std::pair<int, std::string>> matches;
	for (const auto& candidate : candidates) {
		std::vector<int> candidateMatches;
		const std::string candidateLower = ToLowerAsciiCopyLocal(candidate);
		const bool caseSensitive = false;
		for (size_t i = 0; i < pageLines.size(); ++i) {
			const bool matched = caseSensitive
				? (pageLines[i].find(candidate) != std::string::npos)
				: (pageLines[i].find(candidate) != std::string::npos ||
				   ToLowerAsciiCopyLocal(pageLines[i]).find(candidateLower) != std::string::npos);
			if (matched) {
				candidateMatches.push_back(static_cast<int>(i) + 1);
			}
		}

		if (!candidateMatches.empty()) {
			if (matches.empty()) {
				for (const int lineNumber : candidateMatches) {
					matches.emplace_back(lineNumber, candidate);
				}
				outMatchedCandidateTotal = static_cast<int>(candidateMatches.size());
			}

			if (sameTextOccurrenceIndex > 0 &&
				sameTextOccurrenceIndex <= static_cast<int>(candidateMatches.size())) {
				outResolvedLineNumber = candidateMatches[static_cast<size_t>(sameTextOccurrenceIndex - 1)];
				outMatchedCandidate = candidate;
				outMatchedCandidateTotal = static_cast<int>(candidateMatches.size());
				outResolvedBy = "matched_text_occurrence_index";
				return true;
			}
		}
	}

	if (!matches.empty()) {
		size_t bestIndex = 0;
		if (hintLineNumber > 0) {
			int bestDistance = (std::numeric_limits<int>::max)();
			for (size_t i = 0; i < matches.size(); ++i) {
				const int distance = std::abs(matches[i].first - hintLineNumber);
				if (distance < bestDistance) {
					bestDistance = distance;
					bestIndex = i;
				}
			}
		}

		outResolvedLineNumber = matches[bestIndex].first;
		outMatchedCandidate = matches[bestIndex].second;
		if (sameTextOccurrenceIndex > 0) {
			outResolvedBy = matches.size() == 1
				? "same_text_occurrence_fallback_unique"
				: "same_text_occurrence_fallback_nearest_hint";
		}
		else {
			outResolvedBy = matches.size() == 1 ? "matched_text_unique" : "matched_text_nearest_hint";
		}
		return true;
	}

	if (hintLineNumber > 0 && hintLineNumber <= static_cast<int>(pageLines.size())) {
		outResolvedLineNumber = hintLineNumber;
		outResolvedBy = "search_hint_fallback";
		return true;
	}

	outResolvedBy = "not_found";
	return false;
}

std::string BuildSearchJumpToken(const e571::DirectGlobalSearchDebugHit& hit)
{
	return std::format(
		"v1:{}:{}:{}:{}:{}",
		hit.type,
		hit.extra,
		hit.outerIndex,
		hit.innerIndex,
		hit.matchOffset);
}

bool ParseSearchJumpToken(const std::string& token, e571::DirectGlobalSearchDebugHit& outHit)
{
	outHit = {};
	const std::string text = TrimAsciiCopy(token);
	static constexpr char kPrefix[] = "v1:";
	if (text.empty() || text.rfind(kPrefix, 0) != 0) {
		return false;
	}

	std::vector<int> values;
	size_t begin = sizeof(kPrefix) - 1;
	while (begin < text.size()) {
		const size_t end = text.find(':', begin);
		const std::string part = (end == std::string::npos)
			? text.substr(begin)
			: text.substr(begin, end - begin);
		if (part.empty()) {
			return false;
		}
		try {
			values.push_back(std::stoi(part));
		}
		catch (...) {
			return false;
		}
		if (end == std::string::npos) {
			break;
		}
		begin = end + 1;
	}

	if (values.size() != 5) {
		return false;
	}

	outHit.type = values[0];
	outHit.extra = values[1];
	outHit.outerIndex = values[2];
	outHit.innerIndex = values[3];
	outHit.matchOffset = values[4];
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

bool TryGetFileMd5HexForAI(const std::string& path, std::string& outHex)
{
	outHex.clear();
	HANDLE file = CreateFileA(
		path.c_str(),
		GENERIC_READ,
		FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
		nullptr,
		OPEN_EXISTING,
		FILE_ATTRIBUTE_NORMAL,
		nullptr);
	if (file == INVALID_HANDLE_VALUE) {
		return false;
	}

	HCRYPTPROV provider = 0;
	HCRYPTHASH hash = 0;
	bool ok = false;
	if (CryptAcquireContextA(&provider, nullptr, nullptr, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) &&
		CryptCreateHash(provider, CALG_MD5, 0, 0, &hash)) {
		std::array<BYTE, 8192> buffer{};
		DWORD readBytes = 0;
		while (ReadFile(file, buffer.data(), static_cast<DWORD>(buffer.size()), &readBytes, nullptr) && readBytes > 0) {
			if (!CryptHashData(hash, buffer.data(), readBytes, 0)) {
				break;
			}
		}

		BYTE digest[16] = {};
		DWORD digestSize = sizeof(digest);
		if (CryptGetHashParam(hash, HP_HASHVAL, digest, &digestSize, 0) != FALSE) {
			static constexpr char kHex[] = "0123456789abcdef";
			outHex.reserve(digestSize * 2);
			for (DWORD i = 0; i < digestSize; ++i) {
				outHex.push_back(kHex[(digest[i] >> 4) & 0x0F]);
				outHex.push_back(kHex[digest[i] & 0x0F]);
			}
			ok = true;
		}
	}

	if (hash != 0) {
		CryptDestroyHash(hash);
	}
	if (provider != 0) {
		CryptReleaseContext(provider, 0);
	}
	CloseHandle(file);
	return ok;
}

std::string BuildFileCacheKeyForAI(const std::string& path)
{
	return ToLowerAsciiCopyLocal(TrimAsciiCopy(path));
}

std::filesystem::path GetAICacheDirectoryForAI()
{
	try {
		return std::filesystem::path(GetBasePath()) / "AutoLinker" / "cache";
	}
	catch (...) {
		return std::filesystem::path();
	}
}

std::filesystem::path GetModulePublicInfoDirectoryForAI()
{
	try {
		return std::filesystem::path(GetBasePath()) / "AutoLinker" / "ecomInfo";
	}
	catch (...) {
		return std::filesystem::path();
	}
}

std::filesystem::path GetModulePublicInfoCachePathForAI(const std::string& md5)
{
	if (TrimAsciiCopy(md5).empty()) {
		return std::filesystem::path();
	}
	return GetModulePublicInfoDirectoryForAI() / (md5 + ".json");
}

std::filesystem::path GetLegacyModulePublicInfoCachePathForAI(const std::string& md5)
{
	if (TrimAsciiCopy(md5).empty()) {
		return std::filesystem::path();
	}
	return GetAICacheDirectoryForAI() / "module_public_info" / (md5 + ".json");
}

std::filesystem::path GetSupportLibraryInfoCachePathForAI(const std::string& md5)
{
	if (TrimAsciiCopy(md5).empty()) {
		return std::filesystem::path();
	}
	return GetAICacheDirectoryForAI() / "support_library_info" / (md5 + ".json");
}

bool TryReadJsonCacheFileForAI(const std::filesystem::path& path, nlohmann::json& outJson)
{
	outJson = nlohmann::json::object();
	if (path.empty()) {
		return false;
	}

	try {
		std::ifstream in(path, std::ios::in | std::ios::binary);
		if (!in.is_open()) {
			return false;
		}
		in.seekg(0, std::ios::end);
		const std::streamoff size = in.tellg();
		if (size <= 0) {
			return false;
		}
		in.seekg(0, std::ios::beg);
		std::string text(static_cast<size_t>(size), '\0');
		in.read(text.data(), size);
		if (!in.good() && static_cast<std::streamoff>(in.gcount()) != size) {
			return false;
		}
		outJson = nlohmann::json::parse(text);
		return true;
	}
	catch (...) {
		return false;
	}
}

bool TryWriteJsonCacheFileForAI(const std::filesystem::path& path, const nlohmann::json& jsonValue)
{
	if (path.empty()) {
		return false;
	}

	try {
		std::filesystem::create_directories(path.parent_path());
		std::ofstream out(path, std::ios::out | std::ios::binary | std::ios::trunc);
		if (!out.is_open()) {
			return false;
		}
		const std::string text = jsonValue.dump();
		out.write(text.data(), static_cast<std::streamsize>(text.size()));
		return out.good();
	}
	catch (...) {
		return false;
	}
}

constexpr const char* kModulePublicInfoCacheSchemaForAI = "module_public_info_v2";

nlohmann::json SerializeModulePublicInfoParamForAI(const e571::ModulePublicInfoParam& param)
{
	nlohmann::json row = nlohmann::json::object();
	row["name"] = LocalToUtf8Text(param.name);
	row["type_text"] = LocalToUtf8Text(param.typeText);
	row["flags_text"] = LocalToUtf8Text(param.flagsText);
	row["comment"] = LocalToUtf8Text(param.comment);
	return row;
}

bool DeserializeModulePublicInfoParamForAI(const nlohmann::json& row, e571::ModulePublicInfoParam& outParam)
{
	if (!row.is_object()) {
		return false;
	}
	outParam = {};
	if (row.contains("name") && row["name"].is_string()) {
		outParam.name = Utf8ToLocalText(row["name"].get<std::string>());
	}
	if (row.contains("type_text") && row["type_text"].is_string()) {
		outParam.typeText = Utf8ToLocalText(row["type_text"].get<std::string>());
	}
	if (row.contains("flags_text") && row["flags_text"].is_string()) {
		outParam.flagsText = Utf8ToLocalText(row["flags_text"].get<std::string>());
	}
	if (row.contains("comment") && row["comment"].is_string()) {
		outParam.comment = Utf8ToLocalText(row["comment"].get<std::string>());
	}
	return true;
}

nlohmann::json SerializeModulePublicInfoRecordForAI(const e571::ModulePublicInfoRecord& record)
{
	nlohmann::json row = nlohmann::json::object();
	row["tag"] = record.tag;
	row["body_size"] = record.bodySize;
	row["header_ints"] = record.headerInts;
	row["payload_offset"] = record.payloadOffset;
	row["raw_bytes"] = record.rawBytes;
	row["extracted_strings"] = nlohmann::json::array();
	for (const auto& text : record.extractedStrings) {
		row["extracted_strings"].push_back(LocalToUtf8Text(text));
	}
	row["kind"] = LocalToUtf8Text(record.kind);
	row["name"] = LocalToUtf8Text(record.name);
	row["type_text"] = LocalToUtf8Text(record.typeText);
	row["flags_text"] = LocalToUtf8Text(record.flagsText);
	row["comment"] = LocalToUtf8Text(record.comment);
	row["signature_text"] = LocalToUtf8Text(record.signatureText);
	row["params"] = nlohmann::json::array();
	for (const auto& param : record.params) {
		row["params"].push_back(SerializeModulePublicInfoParamForAI(param));
	}
	return row;
}

bool DeserializeModulePublicInfoRecordForAI(const nlohmann::json& row, e571::ModulePublicInfoRecord& outRecord)
{
	if (!row.is_object()) {
		return false;
	}
	outRecord = {};
	if (row.contains("tag") && row["tag"].is_number_integer()) {
		outRecord.tag = row["tag"].get<int>();
	}
	if (row.contains("body_size") && row["body_size"].is_number_unsigned()) {
		outRecord.bodySize = row["body_size"].get<unsigned int>();
	}
	if (row.contains("header_ints") && row["header_ints"].is_array()) {
		for (const auto& item : row["header_ints"]) {
			if (item.is_number_integer()) {
				outRecord.headerInts.push_back(item.get<int>());
			}
		}
	}
	if (row.contains("payload_offset") && row["payload_offset"].is_number_integer()) {
		outRecord.payloadOffset = row["payload_offset"].get<int>();
	}
	if (row.contains("raw_bytes") && row["raw_bytes"].is_array()) {
		for (const auto& item : row["raw_bytes"]) {
			if (item.is_number_unsigned()) {
				outRecord.rawBytes.push_back(static_cast<unsigned char>(item.get<unsigned int>()));
			}
		}
	}
	if (row.contains("extracted_strings") && row["extracted_strings"].is_array()) {
		for (const auto& item : row["extracted_strings"]) {
			if (item.is_string()) {
				outRecord.extractedStrings.push_back(Utf8ToLocalText(item.get<std::string>()));
			}
		}
	}
	if (row.contains("kind") && row["kind"].is_string()) {
		outRecord.kind = Utf8ToLocalText(row["kind"].get<std::string>());
	}
	if (row.contains("name") && row["name"].is_string()) {
		outRecord.name = Utf8ToLocalText(row["name"].get<std::string>());
	}
	if (row.contains("type_text") && row["type_text"].is_string()) {
		outRecord.typeText = Utf8ToLocalText(row["type_text"].get<std::string>());
	}
	if (row.contains("flags_text") && row["flags_text"].is_string()) {
		outRecord.flagsText = Utf8ToLocalText(row["flags_text"].get<std::string>());
	}
	if (row.contains("comment") && row["comment"].is_string()) {
		outRecord.comment = Utf8ToLocalText(row["comment"].get<std::string>());
	}
	if (row.contains("signature_text") && row["signature_text"].is_string()) {
		outRecord.signatureText = Utf8ToLocalText(row["signature_text"].get<std::string>());
	}
	if (row.contains("params") && row["params"].is_array()) {
		for (const auto& item : row["params"]) {
			e571::ModulePublicInfoParam param;
			if (DeserializeModulePublicInfoParamForAI(item, param)) {
				outRecord.params.push_back(std::move(param));
			}
		}
	}
	return true;
}

nlohmann::json SerializeModulePublicInfoDumpForAI(const e571::ModulePublicInfoDump& dump)
{
	nlohmann::json row = nlohmann::json::object();
	row["module_path"] = LocalToUtf8Text(dump.modulePath);
	row["loader_error"] = LocalToUtf8Text(dump.loaderError);
	row["trace"] = LocalToUtf8Text(dump.trace);
	row["source_kind"] = dump.sourceKind;
	row["version_text"] = LocalToUtf8Text(dump.versionText);
	row["module_name"] = LocalToUtf8Text(dump.moduleName);
	row["assembly_name"] = LocalToUtf8Text(dump.assemblyName);
	row["assembly_comment"] = LocalToUtf8Text(dump.assemblyComment);
	row["formatted_text"] = LocalToUtf8Text(dump.formattedText);
	row["native_result"] = dump.nativeResult;
	row["records"] = nlohmann::json::array();
	for (const auto& record : dump.records) {
		row["records"].push_back(SerializeModulePublicInfoRecordForAI(record));
	}
	return row;
}

bool DeserializeModulePublicInfoDumpForAI(const nlohmann::json& row, e571::ModulePublicInfoDump& outDump)
{
	if (!row.is_object()) {
		return false;
	}

	outDump = {};
	if (row.contains("module_path") && row["module_path"].is_string()) {
		outDump.modulePath = Utf8ToLocalText(row["module_path"].get<std::string>());
	}
	if (row.contains("loader_error") && row["loader_error"].is_string()) {
		outDump.loaderError = Utf8ToLocalText(row["loader_error"].get<std::string>());
	}
	if (row.contains("trace") && row["trace"].is_string()) {
		outDump.trace = Utf8ToLocalText(row["trace"].get<std::string>());
	}
	if (row.contains("source_kind") && row["source_kind"].is_string()) {
		outDump.sourceKind = row["source_kind"].get<std::string>();
	}
	if (row.contains("version_text") && row["version_text"].is_string()) {
		outDump.versionText = Utf8ToLocalText(row["version_text"].get<std::string>());
	}
	if (row.contains("module_name") && row["module_name"].is_string()) {
		outDump.moduleName = Utf8ToLocalText(row["module_name"].get<std::string>());
	}
	if (row.contains("assembly_name") && row["assembly_name"].is_string()) {
		outDump.assemblyName = Utf8ToLocalText(row["assembly_name"].get<std::string>());
	}
	if (row.contains("assembly_comment") && row["assembly_comment"].is_string()) {
		outDump.assemblyComment = Utf8ToLocalText(row["assembly_comment"].get<std::string>());
	}
	if (row.contains("formatted_text") && row["formatted_text"].is_string()) {
		outDump.formattedText = Utf8ToLocalText(row["formatted_text"].get<std::string>());
	}
	if (row.contains("native_result") && row["native_result"].is_number_integer()) {
		outDump.nativeResult = row["native_result"].get<int>();
	}
	if (row.contains("records") && row["records"].is_array()) {
		for (const auto& item : row["records"]) {
			e571::ModulePublicInfoRecord record;
			if (DeserializeModulePublicInfoRecordForAI(item, record)) {
				outDump.records.push_back(std::move(record));
			}
		}
	}
	return true;
}

void BuildModulePublicInfoLineCacheForAI(ModulePublicInfoCacheEntry& entry)
{
	entry.normalizedText = NormalizeLineBreaksForAI(entry.dump.formattedText);
	entry.lines = SplitLinesCopyForAI(entry.normalizedText);
}

std::string JoinLinesForAI(const std::vector<std::string>& lines)
{
	std::string text;
	for (size_t i = 0; i < lines.size(); ++i) {
		text += lines[i];
		if (i + 1 < lines.size()) {
			text += "\n";
		}
	}
	return text;
}

std::string GetJsonStringFieldLocalForAI(const nlohmann::json& row, const char* key)
{
	if (!row.is_object() || key == nullptr) {
		return std::string();
	}
	if (!row.contains(key) || !row[key].is_string()) {
		return std::string();
	}
	return Utf8ToLocalText(row[key].get<std::string>());
}

int GetJsonIntFieldForAI(const nlohmann::json& row, const char* key, int defaultValue = 0)
{
	if (!row.is_object() || key == nullptr) {
		return defaultValue;
	}
	if (!row.contains(key) || !row[key].is_number_integer()) {
		return defaultValue;
	}
	return row[key].get<int>();
}

std::string DecodeSupportLibraryDataTypeForAI(int typeValue)
{
	const DATA_TYPE type = static_cast<DATA_TYPE>(typeValue);
	const bool isArray = (type & DT_IS_ARY) != 0;
	const bool isVar = (type & DT_IS_VAR) != 0;
	const DATA_TYPE baseType = static_cast<DATA_TYPE>(type & ~DT_IS_ARY);

	std::string text;
	switch (baseType) {
	case _SDT_NULL: text = "空类型"; break;
	case _SDT_ALL: text = "通用型"; break;
	case SDT_BYTE: text = "字节型"; break;
	case SDT_SHORT: text = "短整数型"; break;
	case SDT_INT: text = "整数型"; break;
	case SDT_INT64: text = "长整数型"; break;
	case SDT_FLOAT: text = "小数型"; break;
	case SDT_DOUBLE: text = "双精度小数型"; break;
	case SDT_BOOL: text = "逻辑型"; break;
	case SDT_DATE_TIME: text = "日期时间型"; break;
	case SDT_TEXT: text = "文本型"; break;
	case SDT_BIN: text = "字节集"; break;
	case SDT_SUB_PTR: text = "子程序指针"; break;
	case SDT_STATMENT: text = "子语句"; break;
	default:
		if ((baseType & DTM_USER_DATA_TYPE_MASK) != 0) {
			text = std::format("用户自定义类型(0x{:08X})", static_cast<unsigned int>(baseType));
		}
		else if ((baseType & DTM_SYS_DATA_TYPE_MASK) != 0) {
			text = std::format("系统类型(0x{:08X})", static_cast<unsigned int>(baseType));
		}
		else {
			text = std::format("库类型(0x{:08X})", static_cast<unsigned int>(baseType));
		}
		break;
	}

	if (isArray) {
		text += "[]";
	}
	if (isVar) {
		text += "&";
	}
	return text;
}

std::string DecodeSupportLibraryConstTypeForAI(int typeValue)
{
	switch (static_cast<SHORT>(typeValue)) {
	case CT_NULL: return "空";
	case CT_NUM: return "数值";
	case CT_BOOL: return "逻辑";
	case CT_TEXT: return "文本";
	default: return std::format("未知({})", typeValue);
	}
}

std::string BuildSupportLibraryPublicTextFromDumpForAI(const nlohmann::json& dumpJson)
{
	if (!dumpJson.is_object()) {
		return std::string();
	}

	std::vector<std::string> lines;
	const std::string supportLibraryName = GetJsonStringFieldLocalForAI(dumpJson, "support_library_name");
	const std::string versionText = GetJsonStringFieldLocalForAI(dumpJson, "version");
	const std::string explainText = GetJsonStringFieldLocalForAI(dumpJson, "explain");
	const std::string authorText = GetJsonStringFieldLocalForAI(dumpJson, "author");
	const std::string filePath = GetJsonStringFieldLocalForAI(dumpJson, "file_path");

	if (!supportLibraryName.empty()) {
		lines.push_back("支持库名称：" + supportLibraryName);
	}
	if (!versionText.empty()) {
		lines.push_back("版本：" + versionText);
	}
	if (!authorText.empty()) {
		lines.push_back("作者：" + authorText);
	}
	if (!filePath.empty()) {
		lines.push_back("文件路径：" + filePath);
	}
	if (!explainText.empty()) {
		lines.push_back("说明：" + explainText);
	}

	if (dumpJson.contains("commands") && dumpJson["commands"].is_array()) {
		lines.push_back("");
		lines.push_back("[命令]");
		for (const auto& cmd : dumpJson["commands"]) {
			const std::string nameText = GetJsonStringFieldLocalForAI(cmd, "name");
			const std::string categoryText = std::to_string(GetJsonIntFieldForAI(cmd, "category", 0));
			const std::string returnTypeText = DecodeSupportLibraryDataTypeForAI(
				GetJsonIntFieldForAI(cmd, "return_type", 0));
			const std::string explain = GetJsonStringFieldLocalForAI(cmd, "explain");

			lines.push_back(std::format(
				".命令 {}, {}, 分类={}",
				nameText.empty() ? std::string("<未命名>") : nameText,
				returnTypeText,
				categoryText));
			if (!explain.empty()) {
				lines.push_back("  说明：" + explain);
			}

			if (cmd.contains("args") && cmd["args"].is_array()) {
				for (const auto& arg : cmd["args"]) {
					const std::string argName = GetJsonStringFieldLocalForAI(arg, "name");
					const std::string argExplain = GetJsonStringFieldLocalForAI(arg, "explain");
					const std::string argTypeText = DecodeSupportLibraryDataTypeForAI(
						GetJsonIntFieldForAI(arg, "data_type", 0));
					std::string argLine = std::format(
						"  .参数 {}, {}",
						argName.empty() ? std::string("<未命名>") : argName,
						argTypeText);
					if (!argExplain.empty()) {
						argLine += ", " + argExplain;
					}
					lines.push_back(std::move(argLine));
				}
			}
		}
	}

	if (dumpJson.contains("data_types") && dumpJson["data_types"].is_array()) {
		lines.push_back("");
		lines.push_back("[数据类型]");
		for (const auto& item : dumpJson["data_types"]) {
			const std::string nameText = GetJsonStringFieldLocalForAI(item, "name");
			const std::string explain = GetJsonStringFieldLocalForAI(item, "explain");
			lines.push_back(std::format(
				".数据类型 {}",
				nameText.empty() ? std::string("<未命名>") : nameText));
			if (!explain.empty()) {
				lines.push_back("  说明：" + explain);
			}

			if (item.contains("members") && item["members"].is_array()) {
				for (const auto& member : item["members"]) {
					const std::string memberName = GetJsonStringFieldLocalForAI(member, "name");
					const std::string memberExplain = GetJsonStringFieldLocalForAI(member, "explain");
					const std::string memberTypeText = DecodeSupportLibraryDataTypeForAI(
						GetJsonIntFieldForAI(member, "data_type", 0));
					std::string memberLine = std::format(
						"  .成员 {}, {}",
						memberName.empty() ? std::string("<未命名>") : memberName,
						memberTypeText);
					if (!memberExplain.empty()) {
						memberLine += ", " + memberExplain;
					}
					lines.push_back(std::move(memberLine));
				}
			}

			if (item.contains("member_commands") && item["member_commands"].is_array()) {
				for (const auto& memberCmd : item["member_commands"]) {
					const std::string memberCmdName = GetJsonStringFieldLocalForAI(memberCmd, "name");
					if (!memberCmdName.empty()) {
						lines.push_back("  .成员命令 " + memberCmdName);
					}
				}
			}
		}
	}

	if (dumpJson.contains("constants") && dumpJson["constants"].is_array()) {
		lines.push_back("");
		lines.push_back("[常量]");
		for (const auto& item : dumpJson["constants"]) {
			const std::string nameText = GetJsonStringFieldLocalForAI(item, "name");
			const std::string explain = GetJsonStringFieldLocalForAI(item, "explain");
			const std::string textValue = GetJsonStringFieldLocalForAI(item, "text_value");
			const std::string constTypeText = DecodeSupportLibraryConstTypeForAI(
				GetJsonIntFieldForAI(item, "type", 0));
			std::string line = std::format(
				".常量 {}, {}",
				nameText.empty() ? std::string("<未命名>") : nameText,
				constTypeText);
			if (!textValue.empty()) {
				line += ", " + textValue;
			}
			else if (item.contains("numeric_value") && item["numeric_value"].is_number()) {
				line += ", " + item["numeric_value"].dump();
			}
			if (!explain.empty()) {
				line += ", " + explain;
			}
			lines.push_back(std::move(line));
		}
	}

	return JoinLinesForAI(lines);
}

void BuildSupportLibraryLineCacheForAI(SupportLibraryDumpCacheEntry& entry)
{
	std::string text = GetJsonStringFieldLocalForAI(entry.dumpJson, "formatted_text");
	if (text.empty()) {
		text = BuildSupportLibraryPublicTextFromDumpForAI(entry.dumpJson);
		if (!text.empty()) {
			entry.dumpJson["formatted_text"] = LocalToUtf8Text(text);
		}
	}

	entry.normalizedText = NormalizeLineBreaksForAI(text);
	entry.lines = SplitLinesCopyForAI(entry.normalizedText);
}

bool TryLoadModulePublicInfoCacheEntryFromJsonForAI(
	const nlohmann::json& cacheJson,
	const std::string& fallbackModulePath,
	const std::string& md5,
	ModulePublicInfoCacheEntry& outEntry)
{
	outEntry = {};
	if (!cacheJson.is_object() ||
		cacheJson.value("schema", std::string()) != kModulePublicInfoCacheSchemaForAI ||
		cacheJson.value("md5", std::string()) != md5 ||
		!cacheJson.contains("dump")) {
		return false;
	}

	e571::ModulePublicInfoDump cachedDump;
	if (!DeserializeModulePublicInfoDumpForAI(cacheJson["dump"], cachedDump)) {
		return false;
	}
	if (!TrimAsciiCopy(fallbackModulePath).empty()) {
		cachedDump.modulePath = fallbackModulePath;
	}

	outEntry.md5 = md5;
	outEntry.dump = std::move(cachedDump);
	outEntry.ok = true;
	BuildModulePublicInfoLineCacheForAI(outEntry);
	return true;
}

bool TryLoadModulePublicInfoCacheEntryByMd5ForAI(
	const std::string& md5,
	ModulePublicInfoCacheEntry& outEntry,
	std::string& outError)
{
	outEntry = {};
	outError.clear();
	const std::string trimmedMd5 = TrimAsciiCopy(md5);
	if (trimmedMd5.empty()) {
		outError = "md5 is required";
		return false;
	}

	{
		std::lock_guard<std::mutex> guard(g_modulePublicInfoCacheMutex);
		for (const auto& [cacheKey, entry] : g_modulePublicInfoCache) {
			(void)cacheKey;
			if (entry.md5 == trimmedMd5) {
				outEntry = entry;
				outError = entry.error;
				return entry.ok;
			}
		}
	}

	const auto tryLoadFromPath = [&](const std::filesystem::path& cachePath, bool migrateToNewPath) -> bool {
		nlohmann::json cacheJson = nlohmann::json::object();
		if (!TryReadJsonCacheFileForAI(cachePath, cacheJson)) {
			return false;
		}

		ModulePublicInfoCacheEntry entry;
		if (!TryLoadModulePublicInfoCacheEntryFromJsonForAI(cacheJson, std::string(), trimmedMd5, entry)) {
			return false;
		}

		if (migrateToNewPath) {
			TryWriteJsonCacheFileForAI(GetModulePublicInfoCachePathForAI(trimmedMd5), cacheJson);
		}

		const std::string cacheKey = BuildFileCacheKeyForAI(entry.dump.modulePath);
		if (!cacheKey.empty()) {
			std::lock_guard<std::mutex> guard(g_modulePublicInfoCacheMutex);
			g_modulePublicInfoCache[cacheKey] = entry;
		}

		outEntry = std::move(entry);
		outError.clear();
		return true;
	};

	if (tryLoadFromPath(GetModulePublicInfoCachePathForAI(trimmedMd5), false)) {
		return true;
	}
	if (tryLoadFromPath(GetLegacyModulePublicInfoCachePathForAI(trimmedMd5), true)) {
		return true;
	}

	outError = "module public info cache not found by md5";
	return false;
}

bool TryLoadModulePublicInfoCacheEntryForAI(
	const std::string& modulePath,
	ModulePublicInfoCacheEntry& outEntry,
	std::string& outError,
	bool* outCacheHit = nullptr)
{
	outEntry = {};
	outError.clear();
	if (outCacheHit != nullptr) {
		*outCacheHit = false;
	}

	std::string md5;
	if (!TryGetFileMd5HexForAI(modulePath, md5) || md5.empty()) {
		e571::ModulePublicInfoDump dump;
		const bool ok = e571::LoadModulePublicInfoDump(
			modulePath,
			GetCurrentProcessImageBaseForAI(),
			&dump,
			&outError);
		outEntry.md5.clear();
		outEntry.dump = std::move(dump);
		outEntry.error = outError;
		outEntry.ok = ok;
		BuildModulePublicInfoLineCacheForAI(outEntry);
		return ok;
	}

	const std::string cacheKey = BuildFileCacheKeyForAI(modulePath);
	{
		std::lock_guard<std::mutex> guard(g_modulePublicInfoCacheMutex);
		const auto it = g_modulePublicInfoCache.find(cacheKey);
		if (it != g_modulePublicInfoCache.end() && it->second.md5 == md5) {
			outEntry = it->second;
			outError = it->second.error;
			if (outCacheHit != nullptr) {
				*outCacheHit = true;
			}
			return it->second.ok;
		}
	}

	const auto tryLoadCachePath = [&](const std::filesystem::path& cachePath, bool migrateToNewPath) -> bool {
		nlohmann::json cacheJson;
		if (!TryReadJsonCacheFileForAI(cachePath, cacheJson)) {
			return false;
		}

		ModulePublicInfoCacheEntry entry;
		if (!TryLoadModulePublicInfoCacheEntryFromJsonForAI(cacheJson, modulePath, md5, entry)) {
			return false;
		}
		if (migrateToNewPath) {
			TryWriteJsonCacheFileForAI(GetModulePublicInfoCachePathForAI(md5), cacheJson);
		}
		{
			std::lock_guard<std::mutex> guard(g_modulePublicInfoCacheMutex);
			g_modulePublicInfoCache[cacheKey] = entry;
		}
		outEntry = std::move(entry);
		if (outCacheHit != nullptr) {
			*outCacheHit = true;
		}
		return true;
	};

	if (tryLoadCachePath(GetModulePublicInfoCachePathForAI(md5), false)) {
		return true;
	}
	if (tryLoadCachePath(GetLegacyModulePublicInfoCachePathForAI(md5), true)) {
		return true;
	}

	e571::ModulePublicInfoDump loadedDump;
	std::string loadError;
	const bool ok = e571::LoadModulePublicInfoDump(
		modulePath,
		GetCurrentProcessImageBaseForAI(),
		&loadedDump,
		&loadError);

	ModulePublicInfoCacheEntry entry;
	entry.md5 = md5;
	entry.dump = loadedDump;
	entry.error = loadError;
	entry.ok = ok;
	BuildModulePublicInfoLineCacheForAI(entry);
	{
		std::lock_guard<std::mutex> guard(g_modulePublicInfoCacheMutex);
		g_modulePublicInfoCache[cacheKey] = entry;
	}

	if (ok) {
		nlohmann::json cacheJson = nlohmann::json::object();
		cacheJson["schema"] = kModulePublicInfoCacheSchemaForAI;
		cacheJson["md5"] = md5;
		cacheJson["dump"] = SerializeModulePublicInfoDumpForAI(loadedDump);
		TryWriteJsonCacheFileForAI(GetModulePublicInfoCachePathForAI(md5), cacheJson);
	}

	outEntry = std::move(entry);
	outError = loadError;
	return ok;
}

bool TryLoadModulePublicInfoCachedForAI(
	const std::string& modulePath,
	e571::ModulePublicInfoDump& outDump,
	std::string& outError,
	bool* outCacheHit = nullptr)
{
	ModulePublicInfoCacheEntry entry;
	const bool ok = TryLoadModulePublicInfoCacheEntryForAI(modulePath, entry, outError, outCacheHit);
	outDump = std::move(entry.dump);
	return ok;
}

struct SupportLibraryInfoHeaderForAI {
	int index = -1;
	std::string rawName;
	std::string name;
	std::string versionText;
	std::string fileName;
	std::string fileStem;
	std::string filePath;
	std::string rawText;
	std::string resolveTrace;
};

struct LoadedSupportLibraryModuleForAI {
	std::string filePath;
	std::string fileName;
	std::string fileStem;
};

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
	const std::string normalized = NormalizeRealCodeLineBreaksToCrLf(text);
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
	const std::vector<std::string> lines = SplitRealCodeLines(NormalizeRealCodeLineBreaksToCrLf(text));
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

bool ContainsKeywordInsensitiveForAI(
	const std::string& text,
	const std::string& keyword,
	const std::string& keywordLower)
{
	if (keyword.empty()) {
		return true;
	}
	if (text.find(keyword) != std::string::npos) {
		return true;
	}
	return ToLowerAsciiCopyLocal(text).find(keywordLower) != std::string::npos;
}

struct PublicCodeSearchSpecForAI {
	std::vector<std::string> keywords;
	std::vector<std::string> keywordsLower;
	std::string keywordMode = "all";
	bool caseSensitive = false;
	bool useRegex = false;
	std::string regexText;
	std::string regexFlags;
	std::regex regex;
	bool searchModules = true;
	bool searchSupportLibraries = true;
	bool searchProject = true;
	bool searchProjectCache = true;
};

bool TryBuildPublicCodeSearchSpecForAI(
	const nlohmann::json& args,
	PublicCodeSearchSpecForAI& outSpec,
	std::string& outError)
{
	outSpec = {};
	outError.clear();

	if (args.contains("keyword") && args["keyword"].is_string()) {
		const std::string keyword = TrimAsciiCopy(Utf8ToLocalText(args["keyword"].get<std::string>()));
		if (!keyword.empty()) {
			outSpec.keywords.push_back(keyword);
		}
	}

	if (args.contains("keywords") && args["keywords"].is_array()) {
		for (const auto& item : args["keywords"]) {
			if (!item.is_string()) {
				continue;
			}
			const std::string keyword = TrimAsciiCopy(Utf8ToLocalText(item.get<std::string>()));
			if (!keyword.empty()) {
				outSpec.keywords.push_back(keyword);
			}
		}
	}

	if (args.contains("keyword_mode") && args["keyword_mode"].is_string()) {
		outSpec.keywordMode = ToLowerAsciiCopyLocal(TrimAsciiCopy(args["keyword_mode"].get<std::string>()));
	}
	if (outSpec.keywordMode != "all" && outSpec.keywordMode != "any") {
		outError = "keyword_mode must be all or any";
		return false;
	}

	outSpec.caseSensitive = args.contains("case_sensitive") && args["case_sensitive"].is_boolean()
		? args["case_sensitive"].get<bool>()
		: false;

	for (const auto& keyword : outSpec.keywords) {
		outSpec.keywordsLower.push_back(outSpec.caseSensitive ? keyword : ToLowerAsciiCopyLocal(keyword));
	}

	if (args.contains("regex") && args["regex"].is_string()) {
		outSpec.regexText = TrimAsciiCopy(Utf8ToLocalText(args["regex"].get<std::string>()));
	}
	if (args.contains("regex_flags") && args["regex_flags"].is_string()) {
		outSpec.regexFlags = ToLowerAsciiCopyLocal(TrimAsciiCopy(args["regex_flags"].get<std::string>()));
	}
	if (!outSpec.regexText.empty()) {
		std::regex_constants::syntax_option_type options = std::regex_constants::ECMAScript;
		if (outSpec.regexFlags.find('i') != std::string::npos) {
			options |= std::regex_constants::icase;
		}
		try {
			outSpec.regex = std::regex(outSpec.regexText, options);
			outSpec.useRegex = true;
		}
		catch (const std::regex_error& ex) {
			outError = std::string("invalid regex: ") + ex.what();
			return false;
		}
	}

	if (args.contains("target_types") && args["target_types"].is_array()) {
		outSpec.searchModules = false;
		outSpec.searchSupportLibraries = false;
		outSpec.searchProject = false;
		outSpec.searchProjectCache = false;
		for (const auto& item : args["target_types"]) {
			if (!item.is_string()) {
				continue;
			}
			const std::string type = ToLowerAsciiCopyLocal(TrimAsciiCopy(item.get<std::string>()));
			if (type == "module" || type == "modules") {
				outSpec.searchModules = true;
			}
			else if (type == "support_library" || type == "support_libraries" || type == "library") {
				outSpec.searchSupportLibraries = true;
			}
			else if (type == "project" || type == "ide" || type == "project_page" || type == "ide_project") {
				outSpec.searchProject = true;
			}
			else if (type == "project_cache" || type == "source_cache" || type == "parsed_project") {
				outSpec.searchProjectCache = true;
			}
		}
		if (!outSpec.searchModules &&
			!outSpec.searchSupportLibraries &&
			!outSpec.searchProject &&
			!outSpec.searchProjectCache) {
			outError = "target_types must contain project and/or project_cache and/or module and/or support_library";
			return false;
		}
	}

	if (outSpec.keywords.empty() && !outSpec.useRegex) {
		outError = "keyword/keywords/regex is required";
		return false;
	}

	return true;
}

bool MatchPublicCodeSearchLineForAI(
	const std::string& line,
	const PublicCodeSearchSpecForAI& spec,
	std::vector<std::string>& outMatchedKeywords)
{
	outMatchedKeywords.clear();

	if (spec.useRegex && !std::regex_search(line, spec.regex)) {
		return false;
	}

	if (spec.keywords.empty()) {
		return true;
	}

	if (spec.keywordMode == "all") {
		for (size_t i = 0; i < spec.keywords.size(); ++i) {
			const bool matched = spec.caseSensitive
				? (line.find(spec.keywords[i]) != std::string::npos)
				: ContainsKeywordInsensitiveForAI(line, spec.keywords[i], spec.keywordsLower[i]);
			if (!matched) {
				outMatchedKeywords.clear();
				return false;
			}
			outMatchedKeywords.push_back(spec.keywords[i]);
		}
		return true;
	}

	for (size_t i = 0; i < spec.keywords.size(); ++i) {
		const bool matched = spec.caseSensitive
			? (line.find(spec.keywords[i]) != std::string::npos)
			: ContainsKeywordInsensitiveForAI(line, spec.keywords[i], spec.keywordsLower[i]);
		if (matched) {
			outMatchedKeywords.push_back(spec.keywords[i]);
		}
	}
	return !outMatchedKeywords.empty();
}

std::vector<KeywordSearchResultInfo> BuildKeywordSearchResultInfosForAI(
	const std::vector<e571::DirectGlobalSearchDebugHit>& hits,
	const std::vector<ProgramTreeItemInfo>& items)
{
	std::vector<KeywordSearchResultInfo> infos;
	infos.reserve(hits.size());

	std::unordered_map<std::string, int> pageTotals;
	std::unordered_map<std::string, int> textTotals;
	for (size_t i = 0; i < hits.size(); ++i) {
		KeywordSearchResultInfo info;
		info.text = hits[i].displayText;
		info.lineNumber = hits[i].outerIndex + 1;
		ParsePageNameFromSearchDisplayText(hits[i].displayText, info.pageName);
		for (const auto& item : items) {
			if (!info.pageName.empty() && item.name == info.pageName) {
				info.pageTypeKey = item.typeKey;
				info.pageTypeName = item.typeName;
				break;
			}
		}
		++pageTotals[ToLowerAsciiCopyLocal(TrimAsciiCopy(info.pageName))];
		++textTotals[BuildProjectSearchGroupingKeyForAI(info.pageName, info.text)];
		infos.push_back(std::move(info));
	}

	std::unordered_map<std::string, int> pageSeen;
	std::unordered_map<std::string, int> textSeen;
	for (auto& info : infos) {
		const std::string pageKey = ToLowerAsciiCopyLocal(TrimAsciiCopy(info.pageName));
		const std::string textKey = BuildProjectSearchGroupingKeyForAI(info.pageName, info.text);
		info.hitIndexInPage = ++pageSeen[pageKey];
		info.hitTotalInPage = pageTotals[pageKey];
		info.sameTextOccurrenceIndex = ++textSeen[textKey];
		info.sameTextOccurrenceTotal = textTotals[textKey];
	}

	return infos;
}

struct ProjectCacheSearchHitForAI {
	KeywordSearchResultInfo info;
	std::vector<std::string> matchedKeywords;
	std::string jumpToken;
	int pageIndex = -1;
};

std::string BuildProjectCacheSourceKindForAI(const project_source_cache::Snapshot& snapshot)
{
	(void)snapshot;
	return "e2txt_project_cache";
}

bool AreProjectCachePageTypesCompatibleForAI(
	const std::string& pageTypeKey,
	const std::string& itemTypeKey)
{
	if (pageTypeKey == itemTypeKey) {
		return true;
	}

	const auto isAssemblyLike = [](const std::string& value) {
		return value == "assembly" || value == "class_module";
	};
	return isAssemblyLike(pageTypeKey) && isAssemblyLike(itemTypeKey);
}

const project_source_cache::Page* FindProjectCachePageByProgramItemForAI(
	const project_source_cache::Snapshot& snapshot,
	const ProgramTreeItemInfo& item,
	std::string& outTrace,
	std::string& outError)
{
	outTrace.clear();
	outError.clear();

	const std::string normalizedName = NormalizeProgramPageNameForAI(item.name);
	if (normalizedName.empty()) {
		outError = "program item name is empty";
		return nullptr;
	}

	const project_source_cache::Page* exactPage = nullptr;
	int exactCount = 0;
	const project_source_cache::Page* loosePage = nullptr;
	int looseCount = 0;

	for (const auto& page : snapshot.pages) {
		if (page.name != normalizedName) {
			continue;
		}

		if (loosePage == nullptr) {
			loosePage = &page;
		}
		++looseCount;

		if (!AreProjectCachePageTypesCompatibleForAI(page.typeKey, item.typeKey)) {
			continue;
		}

		if (exactPage == nullptr) {
			exactPage = &page;
		}
		++exactCount;
	}

	if (exactCount == 1 && exactPage != nullptr) {
		outTrace =
			"page_lookup_exact"
			"|name=" + normalizedName +
			"|type_key=" + item.typeKey +
			"|page_index=" + std::to_string(exactPage->pageIndex);
		return exactPage;
	}
	if (exactCount > 1) {
		outError = "project cache page name is ambiguous";
		outTrace =
			"page_lookup_ambiguous_exact"
			"|name=" + normalizedName +
			"|type_key=" + item.typeKey +
			"|count=" + std::to_string(exactCount);
		return nullptr;
	}
	if (looseCount == 1 && loosePage != nullptr) {
		outTrace =
			"page_lookup_name_only"
			"|name=" + normalizedName +
			"|item_type_key=" + item.typeKey +
			"|page_type_key=" + loosePage->typeKey +
			"|page_index=" + std::to_string(loosePage->pageIndex);
		return loosePage;
	}
	if (looseCount > 1) {
		outError = "project cache page name is ambiguous";
		outTrace =
			"page_lookup_ambiguous_name"
			"|name=" + normalizedName +
			"|item_type_key=" + item.typeKey +
			"|count=" + std::to_string(looseCount);
		return nullptr;
	}

	outError = "program item page not found in project cache";
	outTrace =
		"page_lookup_not_found"
		"|name=" + normalizedName +
		"|type_key=" + item.typeKey;
	return nullptr;
}

bool TryGetProgramItemProjectCacheCodeForAI(
	const ProgramTreeItemInfo& item,
	bool refreshCache,
	std::string& outCode,
	project_source_cache::Snapshot& outSnapshot,
	const project_source_cache::Page*& outPage,
	bool& outRefreshed,
	std::string& outTrace,
	std::string& outError)
{
	outCode.clear();
	outSnapshot = {};
	outPage = nullptr;
	outRefreshed = false;
	outTrace.clear();
	outError.clear();

	std::string snapshotTrace;
	if (!project_source_cache::ProjectSourceCacheManager::Instance().EnsureCurrentSourceLatest(
			outSnapshot,
			refreshCache,
			&outRefreshed,
			&outError,
			&snapshotTrace)) {
		outTrace = snapshotTrace;
		return false;
	}

	std::string pageTrace;
	const project_source_cache::Page* page = FindProjectCachePageByProgramItemForAI(
		outSnapshot,
		item,
		pageTrace,
		outError);
	if (page == nullptr) {
		outTrace = snapshotTrace.empty() ? pageTrace : (snapshotTrace + "|" + pageTrace);
		return false;
	}

	outCode = NormalizeRealCodeLineBreaksToCrLf(JoinRealCodeLines(page->lines));
	if (outCode.empty()) {
		outError = "project cache page code is empty";
		outTrace =
			(snapshotTrace.empty() ? std::string() : (snapshotTrace + "|")) +
			pageTrace +
			"|page_empty";
		return false;
	}

	outPage = page;
	outTrace =
		(snapshotTrace.empty() ? std::string() : (snapshotTrace + "|")) +
		pageTrace +
		"|no_switch=1"
		"|line_count=" + std::to_string(page->lines.size()) +
		"|code_bytes=" + std::to_string(outCode.size());
	return true;
}

template <typename MatchFn>
std::vector<ProjectCacheSearchHitForAI> CollectProjectCacheSearchHitsForAI(
	const project_source_cache::Snapshot& snapshot,
	MatchFn&& matchLine)
{
	std::vector<ProjectCacheSearchHitForAI> hits;
	std::unordered_map<std::string, int> pageTotals;
	std::unordered_map<std::string, int> textTotals;

	for (const auto& page : snapshot.pages) {
		for (size_t lineIndex = 0; lineIndex < page.lines.size(); ++lineIndex) {
			std::vector<std::string> matchedKeywords;
			const std::string& line = page.lines[lineIndex];
			if (!matchLine(line, matchedKeywords)) {
				continue;
			}

			ProjectCacheSearchHitForAI hit;
			hit.info.pageName = page.name;
			hit.info.pageTypeKey = page.typeKey;
			hit.info.pageTypeName = page.typeName;
			hit.info.lineNumber = static_cast<int>(lineIndex) + 1;
			hit.info.text = line;
			hit.matchedKeywords = std::move(matchedKeywords);
			hit.pageIndex = page.pageIndex;
			++pageTotals[ToLowerAsciiCopyLocal(TrimAsciiCopy(hit.info.pageName))];
			++textTotals[BuildProjectSearchGroupingKeyForAI(hit.info.pageName, hit.info.text)];
			hits.push_back(std::move(hit));
		}
	}

	std::unordered_map<std::string, int> pageSeen;
	std::unordered_map<std::string, int> textSeen;
	auto& manager = project_source_cache::ProjectSourceCacheManager::Instance();
	for (auto& hit : hits) {
		const std::string pageKey = ToLowerAsciiCopyLocal(TrimAsciiCopy(hit.info.pageName));
		const std::string textKey = BuildProjectSearchGroupingKeyForAI(hit.info.pageName, hit.info.text);
		hit.info.hitIndexInPage = ++pageSeen[pageKey];
		hit.info.hitTotalInPage = pageTotals[pageKey];
		hit.info.sameTextOccurrenceIndex = ++textSeen[textKey];
		hit.info.sameTextOccurrenceTotal = textTotals[textKey];

		project_source_cache::HitToken token;
		token.sourcePath = snapshot.sourcePath;
		token.revision = snapshot.revision;
		token.pageIndex = hit.pageIndex;
		token.lineNumber = hit.info.lineNumber;
		token.pageName = hit.info.pageName;
		token.pageTypeKey = hit.info.pageTypeKey;
		token.pageTypeName = hit.info.pageTypeName;
		hit.jumpToken = manager.BuildHitToken(token);
	}

	return hits;
}

const project_source_cache::Page* FindProjectCachePageForAI(
	const project_source_cache::Snapshot& snapshot,
	const project_source_cache::HitToken& token)
{
	if (token.pageIndex >= 0 && token.pageIndex < static_cast<int>(snapshot.pages.size())) {
		const auto& page = snapshot.pages[static_cast<size_t>(token.pageIndex)];
		if ((token.pageName.empty() || page.name == token.pageName) &&
			(token.pageTypeKey.empty() || page.typeKey == token.pageTypeKey)) {
			return &page;
		}
	}

	for (const auto& page : snapshot.pages) {
		if (!token.pageName.empty() && page.name != token.pageName) {
			continue;
		}
		if (!token.pageTypeKey.empty() && page.typeKey != token.pageTypeKey) {
			continue;
		}
		return &page;
	}

	return nullptr;
}

bool TryGetAnsiPtrTextLengthForAI(LPCSTR textPtr, size_t& outLength)
{
	outLength = 0;
	if (textPtr == nullptr) {
		return false;
	}

	__try {
		while (outLength < 1024 * 1024) {
			const unsigned char ch = static_cast<unsigned char>(textPtr[outLength]);
			if (ch == 0) {
				return true;
			}
			++outLength;
		}
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
		return false;
	}

	return false;
}

bool TryReadAnsiPtrTextForAI(LPCSTR textPtr, std::string& outText)
{
	outText.clear();
	size_t length = 0;
	if (!TryGetAnsiPtrTextLengthForAI(textPtr, length)) {
		return false;
	}
	outText.assign(textPtr, length);
	return true;
}

std::string ReadAnsiPtrTextOrEmptyForAI(LPCSTR textPtr)
{
	std::string text;
	TryReadAnsiPtrTextForAI(textPtr, text);
	return text;
}

std::string FindPossibleFileTokenForAI(const std::string& text)
{
	const std::array<const char*, 4> exts = { ".fne", ".fnr", ".dll", ".FNX" };
	for (const char* ext : exts) {
		size_t pos = text.find(ext);
		if (pos == std::string::npos) {
			pos = ToLowerAsciiCopyLocal(text).find(ToLowerAsciiCopyLocal(ext));
		}
		if (pos == std::string::npos) {
			continue;
		}

		size_t begin = pos;
		while (begin > 0) {
			const char ch = text[begin - 1];
			if (std::isalnum(static_cast<unsigned char>(ch)) != 0 ||
				ch == '_' || ch == '-' || ch == '.' || ch == '\\' || ch == '/' || ch == ':') {
				--begin;
				continue;
			}
			break;
		}

		size_t end = pos + std::strlen(ext);
		while (end < text.size()) {
			const char ch = text[end];
			if (std::isalnum(static_cast<unsigned char>(ch)) != 0 ||
				ch == '_' || ch == '-' || ch == '.' || ch == '\\' || ch == '/' || ch == ':') {
				++end;
				continue;
			}
			break;
		}
		return TrimAsciiCopy(text.substr(begin, end - begin));
	}
	return std::string();
}

void ParseSupportLibraryHeaderTextForAI(
	int index,
	const std::string& rawText,
	SupportLibraryInfoHeaderForAI& outInfo)
{
	outInfo = {};
	outInfo.index = index;
	outInfo.rawText = rawText;

	const auto lines = SplitLinesCopyForAI(NormalizeLineBreaksForAI(rawText));
	for (const auto& rawLine : lines) {
		const std::string line = TrimAsciiCopy(rawLine);
		if (line.empty()) {
			continue;
		}

		if (outInfo.name.empty()) {
			if (line.find("支持库") != std::string::npos ||
				line.find("库名") != std::string::npos ||
				line.find("名称") != std::string::npos) {
				const size_t sep = line.find_first_of(":：");
				if (sep != std::string::npos) {
					outInfo.name = TrimAsciiCopy(line.substr(sep + 1));
				}
			}
			if (outInfo.name.empty()) {
				outInfo.name = line;
			}
			outInfo.rawName = outInfo.name;
		}

		if (outInfo.versionText.empty() &&
			(line.find("版本") != std::string::npos || line.find("Version") != std::string::npos)) {
			outInfo.versionText = line;
		}

		if (outInfo.fileName.empty()) {
			std::string fileToken = FindPossibleFileTokenForAI(line);
			if (!fileToken.empty()) {
				outInfo.fileName = GetFileNameOnlyForAI(fileToken);
				outInfo.fileStem = GetFileStemForAI(fileToken);
				if (fileToken.find('\\') != std::string::npos || fileToken.find('/') != std::string::npos) {
					outInfo.filePath = fileToken;
				}
			}
		}
	}

	if (outInfo.name.empty()) {
		outInfo.name = std::format("support_library_{}", index);
		outInfo.rawName = outInfo.name;
	}
}

bool TryLoadSupportLibraryBasicInfoForAI(
	const std::string& filePath,
	std::string& outName,
	std::string& outVersionText,
	std::string* outGuid = nullptr)
{
	outName.clear();
	outVersionText.clear();
	if (outGuid != nullptr) {
		outGuid->clear();
	}

	HMODULE module = LoadLibraryExA(filePath.c_str(), nullptr, DONT_RESOLVE_DLL_REFERENCES);
	if (module == nullptr) {
		return false;
	}

	auto closeModule = [&]() {
		if (module != nullptr) {
			FreeLibrary(module);
			module = nullptr;
		}
	};

	auto* getInfoProc = reinterpret_cast<PFN_GET_LIB_INFO>(GetProcAddress(module, FUNCNAME_GET_LIB_INFO));
	if (getInfoProc == nullptr) {
		closeModule();
		return false;
	}

	const auto* libInfo = getInfoProc();
	if (libInfo == nullptr) {
		closeModule();
		return false;
	}

	outName = ReadAnsiPtrTextOrEmptyForAI(libInfo->m_szName);
	outVersionText = std::format(
		"{}.{}.{}",
		libInfo->m_nMajorVersion,
		libInfo->m_nMinorVersion,
		libInfo->m_nBuildNumber);
	if (outGuid != nullptr) {
		*outGuid = ReadAnsiPtrTextOrEmptyForAI(libInfo->m_szGuid);
	}

	closeModule();
	return !outName.empty();
}

std::string GetSupportLibraryDirectoryForAI()
{
	try {
		return (std::filesystem::path(GetBasePath()) / "lib").string();
	}
	catch (...) {
		return std::string();
	}
}

std::string ExtractLeadingAsciiStemForSupportLibraryAI(const std::string& text)
{
	std::string stem;
	for (char ch : text) {
		const unsigned char uch = static_cast<unsigned char>(ch);
		if (std::isalnum(uch) != 0 || ch == '_' || ch == '-') {
			stem.push_back(ch);
			continue;
		}
		break;
	}
	if (stem.size() < 2) {
		return std::string();
	}
	return stem;
}

std::vector<LoadedSupportLibraryModuleForAI> EnumerateLoadedSupportLibraryModulesForAI()
{
	std::vector<LoadedSupportLibraryModuleForAI> modules;
	const std::string libDir = GetSupportLibraryDirectoryForAI();
	if (libDir.empty() || !std::filesystem::exists(std::filesystem::path(libDir))) {
		return modules;
	}

	HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE | TH32CS_SNAPMODULE32, GetCurrentProcessId());
	if (snapshot == INVALID_HANDLE_VALUE) {
		return modules;
	}

	const std::string libDirLower = ToLowerAsciiCopyLocal(libDir);
	MODULEENTRY32 entry = {};
	entry.dwSize = sizeof(entry);
	if (Module32First(snapshot, &entry) == FALSE) {
		CloseHandle(snapshot);
		return modules;
	}

	do {
		const std::filesystem::path modulePath(entry.szExePath);
		const std::string path = modulePath.string();
		const std::string pathLower = ToLowerAsciiCopyLocal(path);
		if (pathLower.rfind(libDirLower, 0) != 0) {
			continue;
		}

		const std::string ext = ToLowerAsciiCopyLocal(modulePath.extension().string());
		if (ext != ".fne" && ext != ".fnr" && ext != ".dll") {
			continue;
		}

		LoadedSupportLibraryModuleForAI module;
		module.filePath = path;
		module.fileName = modulePath.filename().string();
		module.fileStem = modulePath.stem().string();
		modules.push_back(std::move(module));
	} while (Module32Next(snapshot, &entry) != FALSE);

	CloseHandle(snapshot);
	return modules;
}

bool TryAssignSupportLibraryPathByLoadedModulesForAI(
	SupportLibraryInfoHeaderForAI& info,
	const std::vector<LoadedSupportLibraryModuleForAI>& modules,
	std::unordered_set<size_t>& usedModuleIndexes)
{
	auto tryMatch = [&](const std::string& candidate, const char* trace) -> bool {
		const std::string needle = ToLowerAsciiCopyLocal(TrimAsciiCopy(candidate));
		if (needle.empty()) {
			return false;
		}

		int matchedIndex = -1;
		for (size_t i = 0; i < modules.size(); ++i) {
			if (usedModuleIndexes.find(i) != usedModuleIndexes.end()) {
				continue;
			}
			if (EqualsInsensitiveForAI(modules[i].fileName, needle) ||
				EqualsInsensitiveForAI(modules[i].fileStem, needle) ||
				EqualsInsensitiveForAI(modules[i].filePath, needle)) {
				if (matchedIndex >= 0) {
					return false;
				}
				matchedIndex = static_cast<int>(i);
			}
		}

		if (matchedIndex < 0) {
			return false;
		}

		const auto& module = modules[static_cast<size_t>(matchedIndex)];
		info.filePath = module.filePath;
		info.fileName = module.fileName;
		info.fileStem = module.fileStem;
		info.resolveTrace = trace;
		usedModuleIndexes.insert(static_cast<size_t>(matchedIndex));
		return true;
	};

	if (tryMatch(info.fileName, "loaded_module_file_name")) {
		return true;
	}
	if (tryMatch(info.fileStem, "loaded_module_file_stem")) {
		return true;
	}
	if (tryMatch(info.name, "loaded_module_name")) {
		return true;
	}
	if (tryMatch(ExtractLeadingAsciiStemForSupportLibraryAI(info.name), "loaded_module_ascii_prefix_name")) {
		return true;
	}
	if (tryMatch(ExtractLeadingAsciiStemForSupportLibraryAI(info.rawText), "loaded_module_ascii_prefix_raw")) {
		return true;
	}
	return false;
}

bool TryListSupportLibrariesForAI(std::vector<SupportLibraryInfoHeaderForAI>& outInfos, std::string* outError)
{
	outInfos.clear();
	if (outError != nullptr) {
		outError->clear();
	}

	int count = 0;
	if (!IDEFacade::Instance().RunGetNumLib(count)) {
		if (outError != nullptr) {
			*outError = "RunGetNumLib failed";
		}
		return false;
	}

	for (int i = 0; i < count; ++i) {
		std::string text;
		if (!IDEFacade::Instance().RunGetLibInfoText(i, text)) {
			continue;
		}

		SupportLibraryInfoHeaderForAI info;
		ParseSupportLibraryHeaderTextForAI(i, text, info);
		outInfos.push_back(std::move(info));
	}

	const auto loadedModules = EnumerateLoadedSupportLibraryModulesForAI();
	std::unordered_set<size_t> usedModuleIndexes;
	for (auto& info : outInfos) {
		if (!info.filePath.empty()) {
			info.resolveTrace = "header_text_path";
		}
		else {
			TryAssignSupportLibraryPathByLoadedModulesForAI(info, loadedModules, usedModuleIndexes);
		}

		if (!info.filePath.empty()) {
			std::string resolvedName;
			std::string resolvedVersion;
			if (TryLoadSupportLibraryBasicInfoForAI(info.filePath, resolvedName, resolvedVersion, nullptr)) {
				info.name = resolvedName;
				info.versionText = resolvedVersion;
			}
		}
	}

	if (loadedModules.size() == outInfos.size()) {
		for (size_t i = 0; i < outInfos.size(); ++i) {
			auto& info = outInfos[i];
			if (!info.filePath.empty()) {
				continue;
			}
			if (usedModuleIndexes.find(i) != usedModuleIndexes.end()) {
				continue;
			}
			info.filePath = loadedModules[i].filePath;
			info.fileName = loadedModules[i].fileName;
			info.fileStem = loadedModules[i].fileStem;
			info.resolveTrace = "loaded_module_order_fallback";
			usedModuleIndexes.insert(i);
			std::string resolvedName;
			std::string resolvedVersion;
			if (TryLoadSupportLibraryBasicInfoForAI(info.filePath, resolvedName, resolvedVersion, nullptr)) {
				info.name = resolvedName;
				info.versionText = resolvedVersion;
			}
		}
	}

	return true;
}

bool ResolveSupportLibraryHeaderForAI(
	const nlohmann::json& args,
	SupportLibraryInfoHeaderForAI& outInfo,
	std::string& outError)
{
	outInfo = {};
	outError.clear();

	const std::string filePath = args.contains("file_path") && args["file_path"].is_string()
		? Utf8ToLocalText(args["file_path"].get<std::string>())
		: std::string();
	const std::string name = args.contains("name") && args["name"].is_string()
		? Utf8ToLocalText(args["name"].get<std::string>())
		: std::string();
	const int index = args.contains("index") && args["index"].is_number_integer()
		? args["index"].get<int>()
		: -1;

	if (!TrimAsciiCopy(filePath).empty()) {
		outInfo.filePath = TrimAsciiCopy(filePath);
		outInfo.fileName = GetFileNameOnlyForAI(outInfo.filePath);
		outInfo.name = GetFileStemForAI(outInfo.filePath);
		outInfo.index = index;
		return true;
	}

	std::vector<SupportLibraryInfoHeaderForAI> libs;
	if (!TryListSupportLibrariesForAI(libs, &outError)) {
		return false;
	}

	std::vector<SupportLibraryInfoHeaderForAI> matched;
	for (const auto& lib : libs) {
		if (index >= 0) {
			if (lib.index == index) {
				matched.push_back(lib);
			}
			continue;
		}

		const std::string trimmedName = TrimAsciiCopy(name);
		if (trimmedName.empty()) {
			continue;
		}

		if (EqualsInsensitiveForAI(lib.name, trimmedName) ||
			EqualsInsensitiveForAI(lib.fileName, trimmedName) ||
			(!lib.filePath.empty() && EqualsInsensitiveForAI(lib.filePath, trimmedName)) ||
			(!lib.filePath.empty() && EqualsInsensitiveForAI(GetFileStemForAI(lib.filePath), trimmedName))) {
			matched.push_back(lib);
		}
	}

	if (matched.empty()) {
		outError = index >= 0 ? "support library index not found" : "support library not found";
		return false;
	}
	if (matched.size() > 1) {
		outError = "support library is ambiguous";
		return false;
	}

	outInfo = matched.front();
	return true;
}

bool LoadSupportLibraryDumpFromFileForAI(
	const std::string& filePath,
	nlohmann::json& outJson,
	std::string& outError)
{
	outJson = nlohmann::json::object();
	outError.clear();

	HMODULE module = LoadLibraryExA(filePath.c_str(), nullptr, DONT_RESOLVE_DLL_REFERENCES);
	if (module == nullptr) {
		outError = "LoadLibraryEx failed";
		return false;
	}

	const auto closeModule = [&]() {
		if (module != nullptr) {
			FreeLibrary(module);
			module = nullptr;
		}
	};

	auto* getInfoProc = reinterpret_cast<PFN_GET_LIB_INFO>(GetProcAddress(module, FUNCNAME_GET_LIB_INFO));
	if (getInfoProc == nullptr) {
		closeModule();
		outError = "GetNewInf not found";
		return false;
	}

	const auto* libInfo = getInfoProc();
	if (libInfo == nullptr) {
		closeModule();
		outError = "GetNewInf returned null";
		return false;
	}

	nlohmann::json categories = nlohmann::json::array();
	if (libInfo->m_nCategoryCount > 0 && libInfo->m_szzCategory != nullptr) {
		const char* cursor = libInfo->m_szzCategory;
		for (int i = 0; i < libInfo->m_nCategoryCount; ++i) {
			const int bitmapIndex = *reinterpret_cast<const int*>(cursor);
			cursor += sizeof(int);
			std::string nameText = ReadAnsiPtrTextOrEmptyForAI(cursor);
			cursor += nameText.size() + 1;
			categories.push_back({
				{"index", i + 1},
				{"bitmap_index", bitmapIndex},
				{"name", LocalToUtf8Text(nameText)}
			});
		}
	}

	nlohmann::json commands = nlohmann::json::array();
	if (libInfo->m_nCmdCount > 0 && libInfo->m_pBeginCmdInfo != nullptr) {
		for (int i = 0; i < libInfo->m_nCmdCount; ++i) {
			const CMD_INFO& cmd = libInfo->m_pBeginCmdInfo[i];
			nlohmann::json row;
			row["index"] = i;
			row["name"] = LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(cmd.m_szName));
			row["eg_name"] = LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(cmd.m_szEgName));
			row["explain"] = LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(cmd.m_szExplain));
			row["category"] = cmd.m_shtCategory;
			row["state"] = cmd.m_wState;
			row["return_type"] = cmd.m_dtRetValType;
			row["user_level"] = cmd.m_shtUserLevel;
			row["bitmap_index"] = cmd.m_shtBitmapIndex;
			row["bitmap_count"] = cmd.m_shtBitmapCount;
			row["arg_count"] = cmd.m_nArgCount;
			row["is_object_member"] = (cmd.m_shtCategory == -1);

			nlohmann::json args = nlohmann::json::array();
			if (cmd.m_nArgCount > 0 && cmd.m_pBeginArgInfo != nullptr) {
				for (int argIndex = 0; argIndex < cmd.m_nArgCount; ++argIndex) {
					const ARG_INFO& arg = cmd.m_pBeginArgInfo[argIndex];
					args.push_back({
						{"index", argIndex},
						{"name", LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(arg.m_szName))},
						{"explain", LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(arg.m_szExplain))},
						{"bitmap_index", arg.m_shtBitmapIndex},
						{"bitmap_count", arg.m_shtBitmapCount},
						{"data_type", arg.m_dtType},
						{"default_value", arg.m_nDefault},
						{"state", arg.m_dwState}
					});
				}
			}
			row["args"] = std::move(args);
			commands.push_back(std::move(row));
		}
	}

	nlohmann::json dataTypes = nlohmann::json::array();
	if (libInfo->m_nDataTypeCount > 0 && libInfo->m_pDataType != nullptr) {
		for (int i = 0; i < libInfo->m_nDataTypeCount; ++i) {
			const LIB_DATA_TYPE_INFO& dataType = libInfo->m_pDataType[i];
			nlohmann::json row;
			row["index"] = i;
			row["name"] = LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(dataType.m_szName));
			row["eg_name"] = LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(dataType.m_szEgName));
			row["explain"] = LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(dataType.m_szExplain));
			row["cmd_count"] = dataType.m_nCmdCount;
			row["state"] = dataType.m_dwState;
			row["element_count"] = dataType.m_nElementCount;
			row["event_count"] = dataType.m_nEventCount;
			row["property_count"] = dataType.m_nPropertyCount;

			nlohmann::json members = nlohmann::json::array();
			if (dataType.m_nElementCount > 0 && dataType.m_pElementBegin != nullptr) {
				for (int memberIndex = 0; memberIndex < dataType.m_nElementCount; ++memberIndex) {
					const auto& member = dataType.m_pElementBegin[memberIndex];
					members.push_back({
						{"index", memberIndex},
						{"name", LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(member.m_szName))},
						{"eg_name", LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(member.m_szEgName))},
						{"explain", LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(member.m_szExplain))},
						{"data_type", member.m_dtType},
						{"state", member.m_dwState},
						{"default_value", member.m_nDefault}
					});
				}
			}
			row["members"] = std::move(members);

			nlohmann::json memberCmds = nlohmann::json::array();
			if (dataType.m_nCmdCount > 0 && dataType.m_pnCmdsIndex != nullptr) {
				for (int cmdIndex = 0; cmdIndex < dataType.m_nCmdCount; ++cmdIndex) {
					const int globalCmdIndex = dataType.m_pnCmdsIndex[cmdIndex];
					if (globalCmdIndex >= 0 && globalCmdIndex < libInfo->m_nCmdCount) {
						memberCmds.push_back({
							{"cmd_index", globalCmdIndex},
							{"name", LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(libInfo->m_pBeginCmdInfo[globalCmdIndex].m_szName))}
						});
					}
				}
			}
			row["member_commands"] = std::move(memberCmds);
			dataTypes.push_back(std::move(row));
		}
	}

	nlohmann::json constants = nlohmann::json::array();
	if (libInfo->m_nLibConstCount > 0 && libInfo->m_pLibConst != nullptr) {
		for (int i = 0; i < libInfo->m_nLibConstCount; ++i) {
			const LIB_CONST_INFO& item = libInfo->m_pLibConst[i];
			const std::string textValue = ReadAnsiPtrTextOrEmptyForAI(item.m_szText);
			const std::string nameText = ReadAnsiPtrTextOrEmptyForAI(item.m_szName);
			const std::string explainText = ReadAnsiPtrTextOrEmptyForAI(item.m_szExplain);
			nlohmann::json row;
			row["index"] = i;
			row["name"] = LocalToUtf8Text(nameText);
			row["eg_name"] = LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(item.m_szEgName));
			row["explain"] = LocalToUtf8Text(explainText);
			row["layout"] = item.m_shtLayout;
			row["type"] = item.m_shtType;
			row["text_value"] = LocalToUtf8Text(textValue);
			row["numeric_value"] = item.m_dbValue;
			constants.push_back(std::move(row));
		}
	}

	outJson["ok"] = true;
	outJson["file_path"] = LocalToUtf8Text(filePath);
	outJson["file_name"] = LocalToUtf8Text(GetFileNameOnlyForAI(filePath));
	outJson["support_library_name"] = LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(libInfo->m_szName));
	outJson["version"] = std::format("{}.{}.{}", libInfo->m_nMajorVersion, libInfo->m_nMinorVersion, libInfo->m_nBuildNumber);
	outJson["major_version"] = libInfo->m_nMajorVersion;
	outJson["minor_version"] = libInfo->m_nMinorVersion;
	outJson["build_number"] = libInfo->m_nBuildNumber;
	outJson["guid"] = LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(libInfo->m_szGuid));
	outJson["language"] = libInfo->m_nLanguage;
	outJson["state"] = libInfo->m_dwState;
	outJson["explain"] = LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(libInfo->m_szExplain));
	outJson["author"] = LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(libInfo->m_szAuthor));
	outJson["zip_code"] = LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(libInfo->m_szZipCode));
	outJson["address"] = LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(libInfo->m_szAddress));
	outJson["phone"] = LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(libInfo->m_szPhoto));
	outJson["fax"] = LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(libInfo->m_szFax));
	outJson["email"] = LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(libInfo->m_szEmail));
	outJson["home_page"] = LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(libInfo->m_szHomePage));
	outJson["other"] = LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(libInfo->m_szOther));
	outJson["depend_files"] = LocalToUtf8Text(ReadAnsiPtrTextOrEmptyForAI(libInfo->m_szzDependFiles));
	outJson["category_count"] = libInfo->m_nCategoryCount;
	outJson["command_count"] = libInfo->m_nCmdCount;
	outJson["data_type_count"] = libInfo->m_nDataTypeCount;
	outJson["constant_count"] = libInfo->m_nLibConstCount;
	outJson["categories"] = std::move(categories);
	outJson["commands"] = std::move(commands);
	outJson["data_types"] = std::move(dataTypes);
	outJson["constants"] = std::move(constants);

	closeModule();
	return true;
}

bool TryLoadSupportLibraryDumpCachedForAI(
	const std::string& filePath,
	nlohmann::json& outJson,
	std::string& outError,
	bool* outCacheHit = nullptr)
{
	outJson = nlohmann::json::object();
	outError.clear();
	if (outCacheHit != nullptr) {
		*outCacheHit = false;
	}

	std::string md5;
	if (!TryGetFileMd5HexForAI(filePath, md5) || md5.empty()) {
		return LoadSupportLibraryDumpFromFileForAI(filePath, outJson, outError);
	}

	const std::string cacheKey = BuildFileCacheKeyForAI(filePath);
	const std::filesystem::path cachePath = GetSupportLibraryInfoCachePathForAI(md5);
	{
		std::lock_guard<std::mutex> guard(g_supportLibraryDumpCacheMutex);
		const auto it = g_supportLibraryDumpCache.find(cacheKey);
		if (it != g_supportLibraryDumpCache.end() && it->second.md5 == md5) {
			outJson = it->second.dumpJson;
			outError = it->second.error;
			if (outCacheHit != nullptr) {
				*outCacheHit = true;
			}
			return it->second.ok;
		}
	}

	{
		nlohmann::json cacheJson = nlohmann::json::object();
		if (TryReadJsonCacheFileForAI(cachePath, cacheJson) &&
			cacheJson.is_object() &&
			cacheJson.value("schema", "") == "support_library_info_v1" &&
			cacheJson.value("md5", "") == md5 &&
			cacheJson.contains("dump") &&
			cacheJson["dump"].is_object()) {
			nlohmann::json cachedDump = cacheJson["dump"];
			cachedDump["file_path"] = LocalToUtf8Text(filePath);
			cachedDump["file_name"] = LocalToUtf8Text(GetFileNameOnlyForAI(filePath));

			SupportLibraryDumpCacheEntry entry;
			entry.md5 = md5;
			entry.dumpJson = cachedDump;
			entry.error.clear();
			entry.ok = true;
			entry.sourceKind = "getnewinf";
			BuildSupportLibraryLineCacheForAI(entry);
			{
				std::lock_guard<std::mutex> guard(g_supportLibraryDumpCacheMutex);
				g_supportLibraryDumpCache[cacheKey] = entry;
			}

			outJson = std::move(cachedDump);
			if (outCacheHit != nullptr) {
				*outCacheHit = true;
			}
			return true;
		}
	}

	nlohmann::json loadedJson = nlohmann::json::object();
	std::string loadError;
	const bool ok = LoadSupportLibraryDumpFromFileForAI(filePath, loadedJson, loadError);

	SupportLibraryDumpCacheEntry entry;
	entry.md5 = md5;
	entry.dumpJson = loadedJson;
	entry.error = loadError;
	entry.ok = ok;
	entry.sourceKind = "getnewinf";
	BuildSupportLibraryLineCacheForAI(entry);
	{
		std::lock_guard<std::mutex> guard(g_supportLibraryDumpCacheMutex);
		g_supportLibraryDumpCache[cacheKey] = std::move(entry);
	}

	if (ok) {
		nlohmann::json cacheJson = nlohmann::json::object();
		cacheJson["schema"] = "support_library_info_v1";
		cacheJson["md5"] = md5;
		cacheJson["dump"] = loadedJson;
		TryWriteJsonCacheFileForAI(cachePath, cacheJson);
	}

	outJson = std::move(loadedJson);
	outError = loadError;
	return ok;
}

bool TryLoadSupportLibraryDumpCacheEntryByMd5ForAI(
	const std::string& md5,
	SupportLibraryDumpCacheEntry& outEntry,
	std::string& outError)
{
	outEntry = {};
	outError.clear();

	const std::string trimmedMd5 = TrimAsciiCopy(md5);
	if (trimmedMd5.empty()) {
		outError = "md5 is required";
		return false;
	}

	{
		std::lock_guard<std::mutex> guard(g_supportLibraryDumpCacheMutex);
		for (const auto& [cacheKey, entry] : g_supportLibraryDumpCache) {
			(void)cacheKey;
			if (entry.md5 == trimmedMd5) {
				outEntry = entry;
				outError = entry.error;
				return entry.ok;
			}
		}
	}

	nlohmann::json cacheJson = nlohmann::json::object();
	if (!TryReadJsonCacheFileForAI(GetSupportLibraryInfoCachePathForAI(trimmedMd5), cacheJson) ||
		!cacheJson.is_object() ||
		cacheJson.value("schema", "") != "support_library_info_v1" ||
		cacheJson.value("md5", "") != trimmedMd5 ||
		!cacheJson.contains("dump") ||
		!cacheJson["dump"].is_object()) {
		outError = "support library cache not found by md5";
		return false;
	}

	outEntry.md5 = trimmedMd5;
	outEntry.dumpJson = cacheJson["dump"];
	outEntry.error.clear();
	outEntry.ok = true;
	outEntry.sourceKind = "getnewinf";
	BuildSupportLibraryLineCacheForAI(outEntry);
	return true;
}

const char* GetModulePublicInfoTagKeyForAI(int tag)
{
	switch (tag) {
	case 250: return "tag_250";
	case 251: return "tag_251";
	case 252: return "tag_252";
	case 253: return "tag_253";
	case 301: return "tag_301";
	case 302: return "tag_302";
	case 303: return "tag_303";
	case 305: return "tag_305";
	case 306: return "tag_306";
	case 307: return "tag_307";
	case 308: return "tag_308";
	case 309: return "tag_309";
	case 311: return "tag_311";
	default: return "tag_unknown";
	}
}

nlohmann::json BuildModulePublicRecordJsonForAI(
	const e571::ModulePublicInfoRecord& record,
	int index,
	int maxStringsPerRecord)
{
	nlohmann::json row;
	row["index"] = index;
	row["tag"] = record.tag;
	row["tag_key"] = GetModulePublicInfoTagKeyForAI(record.tag);
	row["body_size"] = record.bodySize;
	row["payload_offset"] = record.payloadOffset;
	row["header_ints"] = record.headerInts;
	if (!record.kind.empty()) {
		row["kind"] = record.kind;
	}
	if (!record.name.empty()) {
		row["name"] = LocalToUtf8Text(record.name);
	}
	if (!record.typeText.empty()) {
		row["type_text"] = LocalToUtf8Text(record.typeText);
	}
	if (!record.flagsText.empty()) {
		row["flags_text"] = LocalToUtf8Text(record.flagsText);
	}
	if (!record.comment.empty()) {
		row["comment"] = LocalToUtf8Text(record.comment);
	}
	if (!record.signatureText.empty()) {
		row["signature_text"] = LocalToUtf8Text(record.signatureText);
	}
	if (!record.params.empty()) {
		nlohmann::json params = nlohmann::json::array();
		for (const auto& param : record.params) {
			nlohmann::json paramRow;
			paramRow["name"] = LocalToUtf8Text(param.name);
			if (!param.typeText.empty()) {
				paramRow["type_text"] = LocalToUtf8Text(param.typeText);
			}
			if (!param.flagsText.empty()) {
				paramRow["flags_text"] = LocalToUtf8Text(param.flagsText);
			}
			if (!param.comment.empty()) {
				paramRow["comment"] = LocalToUtf8Text(param.comment);
			}
			params.push_back(std::move(paramRow));
		}
		row["params"] = std::move(params);
	}
	if (!record.extractedStrings.empty()) {
		if (!row.contains("name")) {
			row["name"] = LocalToUtf8Text(record.extractedStrings.front());
		}
		nlohmann::json strings = nlohmann::json::array();
		for (int i = 0; i < static_cast<int>(record.extractedStrings.size()) && i < maxStringsPerRecord; ++i) {
			strings.push_back(LocalToUtf8Text(record.extractedStrings[i]));
		}
		row["strings"] = std::move(strings);
		row["string_count"] = record.extractedStrings.size();
	}
	return row;
}


std::string BuildListImportedModulesJsonOnMainThread(bool& outOk)
{
	std::vector<std::string> paths;
	std::string error;
	if (!TryListImportedModulePathsForAI(paths, &error)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = error.empty() ? "list imported modules failed" : error;
		return Utf8ToLocalText(r.dump());
	}

	nlohmann::json modules = nlohmann::json::array();
	for (size_t i = 0; i < paths.size(); ++i) {
		nlohmann::json row;
		row["index"] = static_cast<int>(i);
		row["path"] = LocalToUtf8Text(paths[i]);
		row["file_name"] = LocalToUtf8Text(GetFileNameOnlyForAI(paths[i]));
		row["module_name"] = LocalToUtf8Text(GetFileStemForAI(paths[i]));
		std::string md5;
		if (TryGetFileMd5HexForAI(paths[i], md5)) {
			row["md5"] = md5;
		}
		modules.push_back(std::move(row));
	}

	nlohmann::json r;
	r["ok"] = true;
	r["count"] = modules.size();
	r["warning"] = LocalToUtf8Text("这里列出的是项目当前导入的易模块路径；模块公开信息优先来自 IDE 模块公开信息窗口的隐藏抓取，必要时才退回 .ec 离线解析，且仅可作为公开接口/伪代码参考。");
	r["modules"] = std::move(modules);
	outOk = true;
	return Utf8ToLocalText(r.dump());
}

std::string BuildListSupportLibrariesJsonOnMainThread(bool& outOk)
{
	std::vector<SupportLibraryInfoHeaderForAI> libs;
	std::string error;
	if (!TryListSupportLibrariesForAI(libs, &error)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = error.empty() ? "list support libraries failed" : error;
		return Utf8ToLocalText(r.dump());
	}

	nlohmann::json rows = nlohmann::json::array();
	for (const auto& lib : libs) {
		nlohmann::json row;
		row["index"] = lib.index;
		row["name"] = LocalToUtf8Text(lib.name);
		row["raw_name"] = LocalToUtf8Text(lib.rawName);
		row["version_text"] = LocalToUtf8Text(lib.versionText);
		row["file_name"] = LocalToUtf8Text(lib.fileName);
		row["file_path"] = LocalToUtf8Text(lib.filePath);
		row["resolve_trace"] = lib.resolveTrace;
		row["info_text"] = LocalToUtf8Text(lib.rawText);
		rows.push_back(std::move(row));
	}

	nlohmann::json r;
	r["ok"] = true;
	r["count"] = rows.size();
	r["warning"] = LocalToUtf8Text("这里列出的是 IDE 当前已选支持库。若能解析到支持库文件路径，则可进一步通过 GetNewInf/lib2.h 读取其命令、常量、数据类型等公开定义。");
	r["libraries"] = std::move(rows);
	outOk = true;
	return Utf8ToLocalText(r.dump());
}

std::string BuildGetSupportLibraryInfoJsonOnMainThread(const std::string& argumentsJson, bool& outOk)
{
	nlohmann::json args;
	try {
		args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = std::string("invalid arguments json: ") + ex.what();
		return Utf8ToLocalText(r.dump());
	}

	SupportLibraryInfoHeaderForAI header;
	std::string resolveError;
	if (!ResolveSupportLibraryHeaderForAI(args, header, resolveError)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = resolveError;
		return Utf8ToLocalText(r.dump());
	}

	nlohmann::json r;
	if (!header.filePath.empty()) {
		std::string loadError;
		if (TryLoadSupportLibraryDumpCachedForAI(header.filePath, r, loadError, nullptr)) {
			r["index"] = header.index;
			r["resolved_header_name"] = LocalToUtf8Text(header.name);
			r["raw_name_from_ide_text"] = LocalToUtf8Text(header.rawName);
			r["resolved_header_version_text"] = LocalToUtf8Text(header.versionText);
			r["info_text"] = LocalToUtf8Text(header.rawText);
			r["resolve_trace"] = header.resolveTrace;
			r["source_kind"] = "getnewinf";
			r["warning"] = LocalToUtf8Text("支持库公开信息来自支持库文件 GetNewInf/lib2.h 结构解析，可作为公开接口参考。");
			outOk = true;
			return Utf8ToLocalText(r.dump());
		}
	}

	r["ok"] = true;
	r["index"] = header.index;
	r["name"] = LocalToUtf8Text(header.name);
	r["raw_name"] = LocalToUtf8Text(header.rawName);
	r["version_text"] = LocalToUtf8Text(header.versionText);
	r["file_name"] = LocalToUtf8Text(header.fileName);
	r["file_path"] = LocalToUtf8Text(header.filePath);
	r["resolve_trace"] = header.resolveTrace;
	r["info_text"] = LocalToUtf8Text(header.rawText);
	r["source_kind"] = "ide_text";
	r["warning"] = LocalToUtf8Text("当前未解析到支持库文件路径或 GetNewInf 失败，以下内容来自 IDE 返回的支持库信息文本。");
	outOk = true;
	return Utf8ToLocalText(r.dump());
}

std::string BuildSearchSupportLibraryInfoJsonOnMainThread(const std::string& argumentsJson, bool& outOk)
{
	nlohmann::json args;
	try {
		args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = std::string("invalid arguments json: ") + ex.what();
		return Utf8ToLocalText(r.dump());
	}

	const std::string keyword = args.contains("keyword") && args["keyword"].is_string()
		? Utf8ToLocalText(args["keyword"].get<std::string>())
		: std::string();
	const int limit = args.contains("limit") && args["limit"].is_number_integer()
		? (std::clamp)(args["limit"].get<int>(), 1, 200)
		: 50;

	const std::string keywordLocal = TrimAsciiCopy(keyword);
	if (keywordLocal.empty()) {
		return R"({"ok":false,"error":"keyword is required"})";
	}

	std::vector<SupportLibraryInfoHeaderForAI> libs;
	std::string error;
	if (args.contains("index") || args.contains("name") || args.contains("file_path")) {
		SupportLibraryInfoHeaderForAI header;
		if (!ResolveSupportLibraryHeaderForAI(args, header, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error;
			return Utf8ToLocalText(r.dump());
		}
		libs.push_back(std::move(header));
	}
	else if (!TryListSupportLibrariesForAI(libs, &error)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = error.empty() ? "list support libraries failed" : error;
		return Utf8ToLocalText(r.dump());
	}

	const auto matchesKeyword = [&](const std::string& text) {
		return
			text.find(keywordLocal) != std::string::npos ||
			ToLowerAsciiCopyLocal(text).find(ToLowerAsciiCopyLocal(keywordLocal)) != std::string::npos;
	};

	nlohmann::json matches = nlohmann::json::array();
	for (const auto& lib : libs) {
		bool usedStructured = false;
		if (!lib.filePath.empty()) {
			nlohmann::json dump;
			std::string loadError;
			if (TryLoadSupportLibraryDumpCachedForAI(lib.filePath, dump, loadError, nullptr)) {
				usedStructured = true;

				if (dump.contains("commands") && dump["commands"].is_array()) {
					for (const auto& cmd : dump["commands"]) {
						if (static_cast<int>(matches.size()) >= limit) {
							break;
						}
						std::vector<std::string> hitTexts;
						const std::string nameText = Utf8ToLocalText(cmd.value("name", ""));
						const std::string explainText = Utf8ToLocalText(cmd.value("explain", ""));
						if (matchesKeyword(nameText)) {
							hitTexts.push_back(nameText);
						}
						if (matchesKeyword(explainText)) {
							hitTexts.push_back(explainText);
						}
						if (cmd.contains("args") && cmd["args"].is_array()) {
							for (const auto& arg : cmd["args"]) {
								const std::string argName = Utf8ToLocalText(arg.value("name", ""));
								const std::string argExplain = Utf8ToLocalText(arg.value("explain", ""));
								if (matchesKeyword(argName)) {
									hitTexts.push_back(argName);
								}
								if (matchesKeyword(argExplain)) {
									hitTexts.push_back(argExplain);
								}
							}
						}
						if (hitTexts.empty()) {
							continue;
						}
						nlohmann::json row;
						row["support_library_name"] = dump.value("support_library_name", "");
						row["file_path"] = dump.value("file_path", "");
						row["source"] = "command";
						row["name"] = cmd.value("name", "");
						row["index"] = cmd.value("index", -1);
						row["matched_strings"] = hitTexts;
						matches.push_back(std::move(row));
					}
				}

				if (dump.contains("constants") && dump["constants"].is_array()) {
					for (const auto& item : dump["constants"]) {
						if (static_cast<int>(matches.size()) >= limit) {
							break;
						}
						std::vector<std::string> hitTexts;
						const std::string nameText = Utf8ToLocalText(item.value("name", ""));
						const std::string explainText = Utf8ToLocalText(item.value("explain", ""));
						const std::string textValue = Utf8ToLocalText(item.value("text_value", ""));
						if (matchesKeyword(nameText)) {
							hitTexts.push_back(nameText);
						}
						if (matchesKeyword(explainText)) {
							hitTexts.push_back(explainText);
						}
						if (matchesKeyword(textValue)) {
							hitTexts.push_back(textValue);
						}
						if (hitTexts.empty()) {
							continue;
						}
						nlohmann::json row;
						row["support_library_name"] = dump.value("support_library_name", "");
						row["file_path"] = dump.value("file_path", "");
						row["source"] = "constant";
						row["name"] = item.value("name", "");
						row["index"] = item.value("index", -1);
						row["matched_strings"] = hitTexts;
						matches.push_back(std::move(row));
					}
				}

				if (dump.contains("data_types") && dump["data_types"].is_array()) {
					for (const auto& item : dump["data_types"]) {
						if (static_cast<int>(matches.size()) >= limit) {
							break;
						}
						std::vector<std::string> hitTexts;
						const std::string nameText = Utf8ToLocalText(item.value("name", ""));
						const std::string explainText = Utf8ToLocalText(item.value("explain", ""));
						if (matchesKeyword(nameText)) {
							hitTexts.push_back(nameText);
						}
						if (matchesKeyword(explainText)) {
							hitTexts.push_back(explainText);
						}
						if (item.contains("members") && item["members"].is_array()) {
							for (const auto& member : item["members"]) {
								const std::string memberName = Utf8ToLocalText(member.value("name", ""));
								const std::string memberExplain = Utf8ToLocalText(member.value("explain", ""));
								if (matchesKeyword(memberName)) {
									hitTexts.push_back(memberName);
								}
								if (matchesKeyword(memberExplain)) {
									hitTexts.push_back(memberExplain);
								}
							}
						}
						if (hitTexts.empty()) {
							continue;
						}
						nlohmann::json row;
						row["support_library_name"] = dump.value("support_library_name", "");
						row["file_path"] = dump.value("file_path", "");
						row["source"] = "data_type";
						row["name"] = item.value("name", "");
						row["index"] = item.value("index", -1);
						row["matched_strings"] = hitTexts;
						matches.push_back(std::move(row));
					}
				}
			}
		}

		if (!usedStructured && !lib.rawText.empty()) {
			const auto lines = SplitLinesCopyForAI(NormalizeLineBreaksForAI(lib.rawText));
			for (size_t i = 0; i < lines.size() && static_cast<int>(matches.size()) < limit; ++i) {
				const std::string line = TrimAsciiCopy(lines[i]);
				if (line.empty() || !matchesKeyword(line)) {
					continue;
				}
				nlohmann::json row;
				row["support_library_name"] = LocalToUtf8Text(lib.name);
				row["file_path"] = LocalToUtf8Text(lib.filePath);
				row["source"] = "ide_text";
				row["line_index"] = static_cast<int>(i);
				row["matched_strings"] = nlohmann::json::array({ LocalToUtf8Text(line) });
				matches.push_back(std::move(row));
			}
		}

		if (static_cast<int>(matches.size()) >= limit) {
			break;
		}
	}

	nlohmann::json r;
	r["ok"] = true;
	r["keyword"] = LocalToUtf8Text(keywordLocal);
	r["match_count"] = matches.size();
	r["warning"] = LocalToUtf8Text("支持库检索优先来自支持库文件 GetNewInf/lib2.h 结构解析；无法解析文件时退回 IDE 返回的支持库信息文本。结果属于公开接口参考，不是项目源码页。");
	r["matches"] = std::move(matches);
	outOk = true;
	return Utf8ToLocalText(r.dump());
}

bool TryResolveSupportLibraryEntryForReadForAI(
	const nlohmann::json& args,
	SupportLibraryDumpCacheEntry& outEntry,
	SupportLibraryInfoHeaderForAI* outHeader,
	std::string& outError)
{
	outEntry = {};
	outError.clear();
	if (outHeader != nullptr) {
		*outHeader = {};
	}

	const std::string md5 = args.contains("md5") && args["md5"].is_string()
		? TrimAsciiCopy(args["md5"].get<std::string>())
		: std::string();
	if (!md5.empty()) {
		return TryLoadSupportLibraryDumpCacheEntryByMd5ForAI(md5, outEntry, outError);
	}

	SupportLibraryInfoHeaderForAI header;
	if (!ResolveSupportLibraryHeaderForAI(args, header, outError)) {
		return false;
	}
	if (outHeader != nullptr) {
		*outHeader = header;
	}

	if (!header.filePath.empty()) {
		std::string loadError;
		nlohmann::json dumpJson = nlohmann::json::object();
		if (TryLoadSupportLibraryDumpCachedForAI(header.filePath, dumpJson, loadError, nullptr)) {
			outEntry.dumpJson = std::move(dumpJson);
			outEntry.ok = true;
			outEntry.error.clear();
			outEntry.sourceKind = "getnewinf";
			TryGetFileMd5HexForAI(header.filePath, outEntry.md5);
			BuildSupportLibraryLineCacheForAI(outEntry);
			return true;
		}
	}

	outEntry.dumpJson = nlohmann::json::object();
	outEntry.dumpJson["ok"] = true;
	outEntry.dumpJson["support_library_name"] = LocalToUtf8Text(header.name);
	outEntry.dumpJson["file_path"] = LocalToUtf8Text(header.filePath);
	outEntry.dumpJson["file_name"] = LocalToUtf8Text(header.fileName);
	outEntry.dumpJson["version"] = LocalToUtf8Text(header.versionText);
	outEntry.dumpJson["info_text"] = LocalToUtf8Text(header.rawText);
	outEntry.dumpJson["formatted_text"] = LocalToUtf8Text(header.rawText);
	outEntry.ok = true;
	outEntry.sourceKind = "ide_text";
	outEntry.normalizedText = NormalizeLineBreaksForAI(header.rawText);
	outEntry.lines = SplitLinesCopyForAI(outEntry.normalizedText);
	return true;
}

std::string BuildSearchSupportLibraryPublicCodeJsonOnMainThread(const std::string& argumentsJson, bool& outOk)
{
	nlohmann::json args;
	try {
		args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = std::string("invalid arguments json: ") + ex.what();
		return Utf8ToLocalText(r.dump());
	}

	const std::string keyword = args.contains("keyword") && args["keyword"].is_string()
		? Utf8ToLocalText(args["keyword"].get<std::string>())
		: std::string();
	const int limit = args.contains("limit") && args["limit"].is_number_integer()
		? (std::clamp)(args["limit"].get<int>(), 1, 500)
		: 100;

	const std::string keywordLocal = TrimAsciiCopy(keyword);
	if (keywordLocal.empty()) {
		return R"({"ok":false,"error":"keyword is required"})";
	}
	const std::string keywordLower = ToLowerAsciiCopyLocal(keywordLocal);

	std::vector<SupportLibraryInfoHeaderForAI> libs;
	std::string error;
	if (args.contains("index") || args.contains("name") || args.contains("file_path")) {
		SupportLibraryInfoHeaderForAI header;
		if (!ResolveSupportLibraryHeaderForAI(args, header, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error;
			return Utf8ToLocalText(r.dump());
		}
		libs.push_back(std::move(header));
	}
	else if (!TryListSupportLibrariesForAI(libs, &error)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = error.empty() ? "list support libraries failed" : error;
		return Utf8ToLocalText(r.dump());
	}

	nlohmann::json matches = nlohmann::json::array();
	for (const auto& lib : libs) {
		nlohmann::json resolveArgs = nlohmann::json::object();
		if (lib.index >= 0) {
			resolveArgs["index"] = lib.index;
		}
		if (!lib.filePath.empty()) {
			resolveArgs["file_path"] = LocalToUtf8Text(lib.filePath);
		}
		else if (!lib.name.empty()) {
			resolveArgs["name"] = LocalToUtf8Text(lib.name);
		}

		SupportLibraryDumpCacheEntry entry;
		SupportLibraryInfoHeaderForAI resolvedHeader;
		std::string loadError;
		if (!TryResolveSupportLibraryEntryForReadForAI(resolveArgs, entry, &resolvedHeader, loadError)) {
			continue;
		}

		for (size_t lineIndex = 0;
			lineIndex < entry.lines.size() && static_cast<int>(matches.size()) < limit;
			++lineIndex) {
			const std::string& line = entry.lines[lineIndex];
			if (!ContainsKeywordInsensitiveForAI(line, keywordLocal, keywordLower)) {
				continue;
			}

			nlohmann::json row;
			row["support_library_name"] = LocalToUtf8Text(
				GetJsonStringFieldLocalForAI(entry.dumpJson, "support_library_name"));
			row["file_path"] = LocalToUtf8Text(
				GetJsonStringFieldLocalForAI(entry.dumpJson, "file_path"));
			row["file_name"] = LocalToUtf8Text(
				GetJsonStringFieldLocalForAI(entry.dumpJson, "file_name"));
			row["md5"] = entry.md5;
			row["source_kind"] = entry.sourceKind;
			row["line_number"] = static_cast<int>(lineIndex) + 1;
			row["text"] = LocalToUtf8Text(line);
			if (!entry.md5.empty()) {
				row["cache_path"] = LocalToUtf8Text(GetSupportLibraryInfoCachePathForAI(entry.md5).string());
			}
			matches.push_back(std::move(row));
		}

		if (static_cast<int>(matches.size()) >= limit) {
			break;
		}
	}

	nlohmann::json r;
	r["ok"] = true;
	r["keyword"] = LocalToUtf8Text(keywordLocal);
	r["match_count"] = matches.size();
	r["warning"] = LocalToUtf8Text("这里搜索的是支持库公开信息的按行文本。优先来自支持库文件 GetNewInf/lib2.h 结构解析；无法定位文件时退回 IDE 支持库信息文本。结果属于公开接口参考，不是项目源码页。");
	r["matches"] = std::move(matches);
	outOk = true;
	return Utf8ToLocalText(r.dump());
}

std::string BuildReadSupportLibraryPublicCodeJsonOnMainThread(const std::string& argumentsJson, bool& outOk)
{
	nlohmann::json args;
	try {
		args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = std::string("invalid arguments json: ") + ex.what();
		return Utf8ToLocalText(r.dump());
	}

	const int startLine = args.contains("start_line") && args["start_line"].is_number_integer()
		? args["start_line"].get<int>()
		: -1;
	const int endLine = args.contains("end_line") && args["end_line"].is_number_integer()
		? args["end_line"].get<int>()
		: startLine;
	if (startLine <= 0 || endLine <= 0 || endLine < startLine) {
		return R"({"ok":false,"error":"start_line/end_line is invalid"})";
	}

	SupportLibraryDumpCacheEntry entry;
	SupportLibraryInfoHeaderForAI header;
	std::string loadError;
	if (!TryResolveSupportLibraryEntryForReadForAI(args, entry, &header, loadError)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = loadError.empty() ? "load support library public code failed" : loadError;
		return Utf8ToLocalText(r.dump());
	}

	const int totalLines = static_cast<int>(entry.lines.size());
	if (totalLines <= 0) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = "support library public code text is empty";
		r["support_library_name"] = LocalToUtf8Text(
			GetJsonStringFieldLocalForAI(entry.dumpJson, "support_library_name"));
		r["file_path"] = LocalToUtf8Text(
			GetJsonStringFieldLocalForAI(entry.dumpJson, "file_path"));
		r["md5"] = entry.md5;
		return Utf8ToLocalText(r.dump());
	}
	if (startLine > totalLines) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = "start_line exceeds total_lines";
		r["support_library_name"] = LocalToUtf8Text(
			GetJsonStringFieldLocalForAI(entry.dumpJson, "support_library_name"));
		r["file_path"] = LocalToUtf8Text(
			GetJsonStringFieldLocalForAI(entry.dumpJson, "file_path"));
		r["md5"] = entry.md5;
		r["total_lines"] = totalLines;
		return Utf8ToLocalText(r.dump());
	}

	const int clampedEndLine = (std::min)(endLine, totalLines);
	nlohmann::json lines = nlohmann::json::array();
	std::string text;
	for (int lineNo = startLine; lineNo <= clampedEndLine; ++lineNo) {
		const std::string& line = entry.lines[static_cast<size_t>(lineNo - 1)];
		if (!text.empty()) {
			text += "\n";
		}
		text += line;
		lines.push_back({
			{"line_number", lineNo},
			{"text", LocalToUtf8Text(line)}
		});
	}

	nlohmann::json r;
	r["ok"] = true;
	r["support_library_name"] = LocalToUtf8Text(
		GetJsonStringFieldLocalForAI(entry.dumpJson, "support_library_name"));
	r["file_path"] = LocalToUtf8Text(
		GetJsonStringFieldLocalForAI(entry.dumpJson, "file_path"));
	r["file_name"] = LocalToUtf8Text(
		GetJsonStringFieldLocalForAI(entry.dumpJson, "file_name"));
	r["md5"] = entry.md5;
	r["source_kind"] = entry.sourceKind;
	r["total_lines"] = totalLines;
	r["start_line"] = startLine;
	r["end_line"] = clampedEndLine;
	r["returned_line_count"] = clampedEndLine - startLine + 1;
	r["code_kind"] = "pseudo_reference";
	r["warning"] = LocalToUtf8Text("这里返回的是支持库公开信息文本的指定行范围。优先来自支持库文件 GetNewInf/lib2.h 结构解析；无法定位文件时退回 IDE 支持库信息文本。结果属于公开接口参考，不是项目源码页。");
	if (!entry.md5.empty()) {
		r["cache_path"] = LocalToUtf8Text(GetSupportLibraryInfoCachePathForAI(entry.md5).string());
	}
	r["text"] = LocalToUtf8Text(text);
	r["lines"] = std::move(lines);
	outOk = true;
	return Utf8ToLocalText(r.dump());
}

std::string BuildGetModulePublicInfoJsonOnMainThread(const std::string& argumentsJson, bool& outOk)
{
	nlohmann::json args;
	try {
		args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = std::string("invalid arguments json: ") + ex.what();
		return Utf8ToLocalText(r.dump());
	}

	const std::string moduleName = args.contains("module_name") && args["module_name"].is_string()
		? Utf8ToLocalText(args["module_name"].get<std::string>())
		: std::string();
	const std::string modulePath = args.contains("module_path") && args["module_path"].is_string()
		? Utf8ToLocalText(args["module_path"].get<std::string>())
		: std::string();
	const int maxRecords = args.contains("max_records") && args["max_records"].is_number_integer()
		? (std::clamp)(args["max_records"].get<int>(), 1, 500)
		: 120;
	const int maxStringsPerRecord = args.contains("max_strings_per_record") && args["max_strings_per_record"].is_number_integer()
		? (std::clamp)(args["max_strings_per_record"].get<int>(), 1, 20)
		: 8;

	std::string resolvedPath;
	std::string resolveError;
	if (!ResolveImportedModulePathForAI(moduleName, modulePath, resolvedPath, resolveError)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = resolveError;
		return Utf8ToLocalText(r.dump());
	}

	e571::ModulePublicInfoDump dump;
	std::string loadError;
	if (!TryLoadModulePublicInfoCachedForAI(
			resolvedPath,
			dump,
			loadError,
			nullptr)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = loadError.empty() ? "load module public info failed" : loadError;
		r["module_path"] = LocalToUtf8Text(resolvedPath);
		r["trace"] = dump.trace;
		r["loader_error"] = LocalToUtf8Text(dump.loaderError);
		return Utf8ToLocalText(r.dump());
	}

	nlohmann::json records = nlohmann::json::array();
	nlohmann::json tagCounts = nlohmann::json::object();
	for (size_t i = 0; i < dump.records.size(); ++i) {
		const auto& record = dump.records[i];
		const std::string tagKey = GetModulePublicInfoTagKeyForAI(record.tag);
		tagCounts[tagKey] = tagCounts.value(tagKey, 0) + 1;
		if (static_cast<int>(records.size()) < maxRecords) {
			records.push_back(BuildModulePublicRecordJsonForAI(
				record,
				static_cast<int>(i),
				maxStringsPerRecord));
		}
	}

	std::string md5;
	TryGetFileMd5HexForAI(resolvedPath, md5);

	nlohmann::json r;
	r["ok"] = true;
	r["module_path"] = LocalToUtf8Text(resolvedPath);
	r["file_name"] = LocalToUtf8Text(GetFileNameOnlyForAI(resolvedPath));
	r["module_name"] = LocalToUtf8Text(GetFileStemForAI(resolvedPath));
	r["md5"] = md5;
	r["native_result"] = dump.nativeResult;
	r["source_kind"] = dump.sourceKind;
	r["version_text"] = LocalToUtf8Text(dump.versionText);
	r["assembly_name"] = LocalToUtf8Text(dump.assemblyName);
	r["assembly_comment"] = LocalToUtf8Text(dump.assemblyComment);
	r["formatted_text"] = LocalToUtf8Text(dump.formattedText);
	r["record_count"] = dump.records.size();
	r["records_returned"] = records.size();
	r["records_truncated"] = dump.records.size() > records.size();
	r["trace"] = dump.trace;
	r["loader_error"] = LocalToUtf8Text(dump.loaderError);
	r["tag_counts"] = std::move(tagCounts);
	r["warning"] = LocalToUtf8Text("模块公开信息优先来自 IDE 模块公开信息窗口的隐藏抓取；必要时会退回 .ec 离线解析。它仍不是 IDE 正常编辑页，也不是模块完整源码，只能作为公开接口/伪代码参考。");
	r["records"] = std::move(records);
	outOk = true;
	return Utf8ToLocalText(r.dump());
}

std::string BuildSearchModulePublicInfoJsonOnMainThread(const std::string& argumentsJson, bool& outOk)
{
	nlohmann::json args;
	try {
		args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = std::string("invalid arguments json: ") + ex.what();
		return Utf8ToLocalText(r.dump());
	}

	const std::string keyword = args.contains("keyword") && args["keyword"].is_string()
		? Utf8ToLocalText(args["keyword"].get<std::string>())
		: std::string();
	const std::string moduleName = args.contains("module_name") && args["module_name"].is_string()
		? Utf8ToLocalText(args["module_name"].get<std::string>())
		: std::string();
	const std::string modulePath = args.contains("module_path") && args["module_path"].is_string()
		? Utf8ToLocalText(args["module_path"].get<std::string>())
		: std::string();
	const int limit = args.contains("limit") && args["limit"].is_number_integer()
		? (std::clamp)(args["limit"].get<int>(), 1, 200)
		: 50;

	const std::string keywordLocal = TrimAsciiCopy(keyword);
	if (keywordLocal.empty()) {
		return R"({"ok":false,"error":"keyword is required"})";
	}

	std::vector<std::string> paths;
	std::string resolveError;
	if (!TrimAsciiCopy(moduleName).empty() || !TrimAsciiCopy(modulePath).empty()) {
		std::string resolvedPath;
		if (!ResolveImportedModulePathForAI(moduleName, modulePath, resolvedPath, resolveError)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = resolveError;
			return Utf8ToLocalText(r.dump());
		}
		paths.push_back(resolvedPath);
	}
	else if (!TryListImportedModulePathsForAI(paths, &resolveError)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = resolveError.empty() ? "list imported modules failed" : resolveError;
		return Utf8ToLocalText(r.dump());
	}

	nlohmann::json matches = nlohmann::json::array();
	for (const auto& path : paths) {
		e571::ModulePublicInfoDump dump;
		std::string loadError;
		if (!TryLoadModulePublicInfoCachedForAI(
				path,
				dump,
				loadError,
				nullptr)) {
			continue;
		}

		for (size_t i = 0; i < dump.records.size() && static_cast<int>(matches.size()) < limit; ++i) {
			const auto& record = dump.records[i];
			std::vector<std::string> matchedStrings;
			const auto matchesKeyword = [&](const std::string& text) {
				return
					text.find(keywordLocal) != std::string::npos ||
					ToLowerAsciiCopyLocal(text).find(ToLowerAsciiCopyLocal(keywordLocal)) != std::string::npos;
			};
			if (!record.name.empty() && matchesKeyword(record.name)) {
				matchedStrings.push_back(record.name);
			}
			if (!record.comment.empty() && matchesKeyword(record.comment)) {
				matchedStrings.push_back(record.comment);
			}
			if (!record.signatureText.empty() && matchesKeyword(record.signatureText)) {
				matchedStrings.push_back(record.signatureText);
			}
			for (const auto& param : record.params) {
				if (matchesKeyword(param.name)) {
					matchedStrings.push_back(param.name);
				}
				if (!param.comment.empty() && matchesKeyword(param.comment)) {
					matchedStrings.push_back(param.comment);
				}
			}
			for (const auto& text : record.extractedStrings) {
				if (matchesKeyword(text)) {
					matchedStrings.push_back(text);
				}
			}
			if (matchedStrings.empty()) {
				continue;
			}

			nlohmann::json row;
			row["module_path"] = LocalToUtf8Text(path);
			row["module_name"] = LocalToUtf8Text(GetFileStemForAI(path));
			row["record_index"] = static_cast<int>(i);
			row["tag"] = record.tag;
			row["tag_key"] = GetModulePublicInfoTagKeyForAI(record.tag);
			row["kind"] = record.kind;
			row["name"] = !record.name.empty()
				? LocalToUtf8Text(record.name)
				: (record.extractedStrings.empty() ? "" : LocalToUtf8Text(record.extractedStrings.front()));
			if (!record.signatureText.empty()) {
				row["signature_text"] = LocalToUtf8Text(record.signatureText);
			}
			nlohmann::json strings = nlohmann::json::array();
			for (const auto& matched : matchedStrings) {
				strings.push_back(LocalToUtf8Text(matched));
			}
			row["matched_strings"] = std::move(strings);
			matches.push_back(std::move(row));
		}

		if (static_cast<int>(matches.size()) >= limit) {
			break;
		}
	}

	nlohmann::json r;
	r["ok"] = true;
	r["keyword"] = LocalToUtf8Text(keywordLocal);
	r["match_count"] = matches.size();
	r["warning"] = LocalToUtf8Text("这里搜索的是模块公开信息窗口抓取到的公开接口文本；必要时会退回 .ec 离线解析。它不是模块完整源码，只能作为公开接口/伪代码参考。");
	r["matches"] = std::move(matches);
	outOk = true;
	return Utf8ToLocalText(r.dump());
}

bool TryResolveModulePublicInfoEntryForReadForAI(
	const nlohmann::json& args,
	ModulePublicInfoCacheEntry& outEntry,
	std::string& outError)
{
	outEntry = {};
	outError.clear();

	const std::string md5 = args.contains("md5") && args["md5"].is_string()
		? TrimAsciiCopy(args["md5"].get<std::string>())
		: std::string();
	if (!md5.empty()) {
		return TryLoadModulePublicInfoCacheEntryByMd5ForAI(md5, outEntry, outError);
	}

	const std::string moduleName = args.contains("module_name") && args["module_name"].is_string()
		? Utf8ToLocalText(args["module_name"].get<std::string>())
		: std::string();
	const std::string modulePath = args.contains("module_path") && args["module_path"].is_string()
		? Utf8ToLocalText(args["module_path"].get<std::string>())
		: std::string();

	std::string resolvedPath;
	if (!ResolveImportedModulePathForAI(moduleName, modulePath, resolvedPath, outError)) {
		return false;
	}

	return TryLoadModulePublicInfoCacheEntryForAI(resolvedPath, outEntry, outError, nullptr);
}

std::string BuildSearchModulePublicCodeJsonOnMainThread(const std::string& argumentsJson, bool& outOk)
{
	nlohmann::json args;
	try {
		args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = std::string("invalid arguments json: ") + ex.what();
		return Utf8ToLocalText(r.dump());
	}

	const std::string keyword = args.contains("keyword") && args["keyword"].is_string()
		? Utf8ToLocalText(args["keyword"].get<std::string>())
		: std::string();
	const std::string moduleName = args.contains("module_name") && args["module_name"].is_string()
		? Utf8ToLocalText(args["module_name"].get<std::string>())
		: std::string();
	const std::string modulePath = args.contains("module_path") && args["module_path"].is_string()
		? Utf8ToLocalText(args["module_path"].get<std::string>())
		: std::string();
	const int limit = args.contains("limit") && args["limit"].is_number_integer()
		? (std::clamp)(args["limit"].get<int>(), 1, 500)
		: 100;

	const std::string keywordLocal = TrimAsciiCopy(keyword);
	if (keywordLocal.empty()) {
		return R"({"ok":false,"error":"keyword is required"})";
	}
	const std::string keywordLower = ToLowerAsciiCopyLocal(keywordLocal);

	std::vector<std::string> paths;
	std::string resolveError;
	if (!TrimAsciiCopy(moduleName).empty() || !TrimAsciiCopy(modulePath).empty()) {
		std::string resolvedPath;
		if (!ResolveImportedModulePathForAI(moduleName, modulePath, resolvedPath, resolveError)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = resolveError;
			return Utf8ToLocalText(r.dump());
		}
		paths.push_back(resolvedPath);
	}
	else if (!TryListImportedModulePathsForAI(paths, &resolveError)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = resolveError.empty() ? "list imported modules failed" : resolveError;
		return Utf8ToLocalText(r.dump());
	}

	nlohmann::json matches = nlohmann::json::array();
	for (const auto& path : paths) {
		ModulePublicInfoCacheEntry entry;
		std::string loadError;
		if (!TryLoadModulePublicInfoCacheEntryForAI(path, entry, loadError, nullptr)) {
			continue;
		}

		for (size_t lineIndex = 0;
			lineIndex < entry.lines.size() && static_cast<int>(matches.size()) < limit;
			++lineIndex) {
			const std::string& line = entry.lines[lineIndex];
			if (!ContainsKeywordInsensitiveForAI(line, keywordLocal, keywordLower)) {
				continue;
			}

			nlohmann::json row;
			row["module_path"] = LocalToUtf8Text(entry.dump.modulePath);
			row["module_name"] = LocalToUtf8Text(GetFileStemForAI(entry.dump.modulePath));
			row["md5"] = entry.md5;
			row["source_kind"] = entry.dump.sourceKind;
			row["line_number"] = static_cast<int>(lineIndex) + 1;
			row["text"] = LocalToUtf8Text(line);
			if (!entry.md5.empty()) {
				row["cache_path"] = LocalToUtf8Text(GetModulePublicInfoCachePathForAI(entry.md5).string());
			}
			matches.push_back(std::move(row));
		}

		if (static_cast<int>(matches.size()) >= limit) {
			break;
		}
	}

	nlohmann::json r;
	r["ok"] = true;
	r["keyword"] = LocalToUtf8Text(keywordLocal);
	r["match_count"] = matches.size();
	r["warning"] = LocalToUtf8Text("这里搜索的是模块公开声明文本的按行结果。它来自模块公开信息窗口抓取或 .ec 离线解析，只能作为公开接口/伪代码参考，不是模块完整源码。");
	r["matches"] = std::move(matches);
	outOk = true;
	return Utf8ToLocalText(r.dump());
}

std::string BuildReadModulePublicCodeJsonOnMainThread(const std::string& argumentsJson, bool& outOk)
{
	nlohmann::json args;
	try {
		args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = std::string("invalid arguments json: ") + ex.what();
		return Utf8ToLocalText(r.dump());
	}

	const int startLine = args.contains("start_line") && args["start_line"].is_number_integer()
		? args["start_line"].get<int>()
		: -1;
	const int endLine = args.contains("end_line") && args["end_line"].is_number_integer()
		? args["end_line"].get<int>()
		: startLine;
	if (startLine <= 0 || endLine <= 0 || endLine < startLine) {
		return R"({"ok":false,"error":"start_line/end_line is invalid"})";
	}

	ModulePublicInfoCacheEntry entry;
	std::string loadError;
	if (!TryResolveModulePublicInfoEntryForReadForAI(args, entry, loadError)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = loadError.empty() ? "load module public code failed" : loadError;
		return Utf8ToLocalText(r.dump());
	}

	const int totalLines = static_cast<int>(entry.lines.size());
	if (totalLines <= 0) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = "module public code text is empty";
		r["module_path"] = LocalToUtf8Text(entry.dump.modulePath);
		r["md5"] = entry.md5;
		return Utf8ToLocalText(r.dump());
	}
	if (startLine > totalLines) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = "start_line exceeds total_lines";
		r["module_path"] = LocalToUtf8Text(entry.dump.modulePath);
		r["md5"] = entry.md5;
		r["total_lines"] = totalLines;
		return Utf8ToLocalText(r.dump());
	}

	const int clampedEndLine = (std::min)(endLine, totalLines);
	nlohmann::json lines = nlohmann::json::array();
	std::string text;
	for (int lineNo = startLine; lineNo <= clampedEndLine; ++lineNo) {
		const std::string& line = entry.lines[static_cast<size_t>(lineNo - 1)];
		if (!text.empty()) {
			text += "\n";
		}
		text += line;
		lines.push_back({
			{"line_number", lineNo},
			{"text", LocalToUtf8Text(line)}
		});
	}

	nlohmann::json r;
	r["ok"] = true;
	r["module_path"] = LocalToUtf8Text(entry.dump.modulePath);
	r["module_name"] = LocalToUtf8Text(GetFileStemForAI(entry.dump.modulePath));
	r["file_name"] = LocalToUtf8Text(GetFileNameOnlyForAI(entry.dump.modulePath));
	r["md5"] = entry.md5;
	r["source_kind"] = entry.dump.sourceKind;
	r["total_lines"] = totalLines;
	r["start_line"] = startLine;
	r["end_line"] = clampedEndLine;
	r["returned_line_count"] = clampedEndLine - startLine + 1;
	r["code_kind"] = "pseudo_reference";
	r["warning"] = LocalToUtf8Text("这里返回的是模块公开声明文本的指定行范围。它来自模块公开信息窗口抓取或 .ec 离线解析，只能作为公开接口/伪代码参考，不是模块完整源码。");
	if (!entry.md5.empty()) {
		r["cache_path"] = LocalToUtf8Text(GetModulePublicInfoCachePathForAI(entry.md5).string());
	}
	r["text"] = LocalToUtf8Text(text);
	r["lines"] = std::move(lines);
	outOk = true;
	return Utf8ToLocalText(r.dump());
}

std::string BuildRefreshProjectSourceCacheJsonOnMainThread(const std::string& argumentsJson, bool& outOk)
{
	(void)argumentsJson;

	project_source_cache::Snapshot snapshot;
	bool refreshed = false;
	std::string error;
	std::string trace;
	if (!project_source_cache::ProjectSourceCacheManager::Instance().EnsureCurrentSourceLatest(
			snapshot,
			true,
			&refreshed,
			&error,
			&trace)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = error.empty() ? "refresh project source cache failed" : error;
		r["trace"] = LocalToUtf8Text(trace);
		return Utf8ToLocalText(r.dump());
	}

	nlohmann::json pages = nlohmann::json::array();
	const size_t previewCount = (std::min)(snapshot.pages.size(), static_cast<size_t>(20));
	for (size_t i = 0; i < previewCount; ++i) {
		nlohmann::json row;
		row["page_index"] = snapshot.pages[i].pageIndex;
		row["page_name"] = LocalToUtf8Text(snapshot.pages[i].name);
		row["page_type_key"] = snapshot.pages[i].typeKey;
		row["page_type_name"] = LocalToUtf8Text(snapshot.pages[i].typeName);
		row["line_count"] = snapshot.pages[i].lines.size();
		pages.push_back(std::move(row));
	}

	nlohmann::json r;
	r["ok"] = true;
	r["refreshed"] = refreshed;
	r["source_kind"] = BuildProjectCacheSourceKindForAI(snapshot);
	r["parsed_input_kind"] = snapshot.parsedInputKind;
	r["source_file_path"] = LocalToUtf8Text(snapshot.sourcePath);
	r["cache_revision"] = snapshot.revision;
	r["project_name"] = LocalToUtf8Text(snapshot.projectName);
	r["version_text"] = LocalToUtf8Text(snapshot.versionText);
	r["page_count"] = snapshot.pages.size();
	if (!snapshot.snapshotPath.empty()) {
		r["snapshot_path"] = LocalToUtf8Text(snapshot.snapshotPath);
	}
	r["trace"] = LocalToUtf8Text(trace);
	r["pages_preview"] = std::move(pages);
	r["warning"] = LocalToUtf8Text("该工具仅使用内存直序列化：会把当前工程从 IDE 内存直接序列化为二进制字节，并直接交给 e2txt 重新解析刷新内存缓存。若当前会话尚未捕获可用的序列化上下文，它会直接报错，不会再走保存重定向或磁盘兜底。");
	outOk = true;
	return Utf8ToLocalText(r.dump());
}

std::string BuildSearchProjectSourceCacheJsonOnMainThread(const std::string& argumentsJson, bool& outOk)
{
	nlohmann::json args;
	try {
		args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = std::string("invalid arguments json: ") + ex.what();
		return Utf8ToLocalText(r.dump());
	}

	PublicCodeSearchSpecForAI spec;
	std::string specError;
	if (!TryBuildPublicCodeSearchSpecForAI(args, spec, specError)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = specError;
		return Utf8ToLocalText(r.dump());
	}

	const int limit = args.contains("limit") && args["limit"].is_number_integer()
		? (std::clamp)(args["limit"].get<int>(), 1, 500)
		: 100;

	project_source_cache::Snapshot snapshot;
	bool refreshed = false;
	std::string error;
	std::string trace;
	if (!project_source_cache::ProjectSourceCacheManager::Instance().EnsureCurrentSourceLatest(
			snapshot,
			true,
			&refreshed,
			&error,
			&trace)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = error.empty() ? "refresh project source cache failed" : error;
		r["trace"] = LocalToUtf8Text(trace);
		return Utf8ToLocalText(r.dump());
	}

	const auto hits = CollectProjectCacheSearchHitsForAI(snapshot, [&](const std::string& line, std::vector<std::string>& matchedKeywords) {
		return MatchPublicCodeSearchLineForAI(line, spec, matchedKeywords);
	});

	nlohmann::json results = nlohmann::json::array();
	for (size_t i = 0; i < hits.size() && static_cast<int>(results.size()) < limit; ++i) {
		nlohmann::json row;
		row["index"] = static_cast<int>(i);
		row["target_type"] = "project_cache";
		row["read_tool"] = "read_project_source_cache_code";
		row["jump_tool"] = "jump_to_search_result";
		row["page_name"] = LocalToUtf8Text(hits[i].info.pageName);
		row["page_type_key"] = hits[i].info.pageTypeKey;
		row["page_type_name"] = LocalToUtf8Text(hits[i].info.pageTypeName);
		row["line_number"] = hits[i].info.lineNumber;
		row["text"] = LocalToUtf8Text(hits[i].info.text);
		row["jump_token"] = hits[i].jumpToken;
		row["source_kind"] = BuildProjectCacheSourceKindForAI(snapshot);
		row["parsed_input_kind"] = snapshot.parsedInputKind;
		row["source_file_path"] = LocalToUtf8Text(snapshot.sourcePath);
		row["cache_revision"] = snapshot.revision;
		row["hit_index_in_page"] = hits[i].info.hitIndexInPage;
		row["hit_total_in_page"] = hits[i].info.hitTotalInPage;
		row["same_text_occurrence_index"] = hits[i].info.sameTextOccurrenceIndex;
		row["same_text_occurrence_total"] = hits[i].info.sameTextOccurrenceTotal;
		if (!snapshot.snapshotPath.empty()) {
			row["snapshot_path"] = LocalToUtf8Text(snapshot.snapshotPath);
		}
		if (!hits[i].matchedKeywords.empty()) {
			nlohmann::json keywords = nlohmann::json::array();
			for (const auto& keyword : hits[i].matchedKeywords) {
				keywords.push_back(LocalToUtf8Text(keyword));
			}
			row["matched_keywords"] = std::move(keywords);
		}
		if (spec.useRegex) {
			row["matched_regex"] = LocalToUtf8Text(spec.regexText);
		}
		results.push_back(std::move(row));
	}

	nlohmann::json r;
	r["ok"] = true;
	r["match_count"] = hits.size();
	r["keyword_mode"] = spec.keywordMode;
	r["case_sensitive"] = spec.caseSensitive;
	r["source_kind"] = BuildProjectCacheSourceKindForAI(snapshot);
	r["parsed_input_kind"] = snapshot.parsedInputKind;
	r["source_file_path"] = LocalToUtf8Text(snapshot.sourcePath);
	r["cache_revision"] = snapshot.revision;
	if (!snapshot.snapshotPath.empty()) {
		r["snapshot_path"] = LocalToUtf8Text(snapshot.snapshotPath);
	}
	if (!spec.keywords.empty()) {
		nlohmann::json keywords = nlohmann::json::array();
		for (const auto& keyword : spec.keywords) {
			keywords.push_back(LocalToUtf8Text(keyword));
		}
		r["keywords"] = std::move(keywords);
	}
	if (spec.useRegex) {
		r["regex"] = LocalToUtf8Text(spec.regexText);
		if (!spec.regexFlags.empty()) {
			r["regex_flags"] = spec.regexFlags;
		}
	}
	r["trace"] = LocalToUtf8Text(trace);
	r["warning"] = LocalToUtf8Text("这里搜索的是当前工程源码缓存：仅依赖内存直序列化，把当前工程从 IDE 内存导出到临时快照并重新解析，再在解析后的页代码里逐行搜索。若当前会话尚未捕获可用的序列化上下文，它会直接报错，不会再走保存重定向或磁盘兜底。若主要查当前工程源码，优先用这个工具而不是 IDE 隐藏搜索。");
	r["results"] = std::move(results);
	outOk = true;
	return Utf8ToLocalText(r.dump());
}

std::string BuildReadProjectSourceCacheCodeJsonOnMainThread(const std::string& argumentsJson, bool& outOk)
{
	nlohmann::json args;
	try {
		args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = std::string("invalid arguments json: ") + ex.what();
		return Utf8ToLocalText(r.dump());
	}

	const std::string jumpToken = (args.contains("jump_token") && args["jump_token"].is_string())
		? args["jump_token"].get<std::string>()
		: std::string();
	const std::string searchText = GetJsonStringArgumentLocal(args, "search_text");
	const int sameTextOccurrenceIndex = (std::max)(0, GetJsonIntArgument(args, "same_text_occurrence_index", 0));
	const int contextBefore = (std::clamp)(GetJsonIntArgument(args, "context_before", 3), 0, 50);
	const int contextAfter = (std::clamp)(GetJsonIntArgument(args, "context_after", 8), 0, 50);
	const bool refreshCache = GetJsonBoolArgument(args, "refresh_cache", true);
	if (TrimAsciiCopy(jumpToken).empty()) {
		return R"({"ok":false,"error":"jump_token is required"})";
	}

	project_source_cache::HitToken projectToken;
	std::string projectTokenError;
	if (!project_source_cache::ProjectSourceCacheManager::Instance().ParseHitToken(
			jumpToken,
			projectToken,
			&projectTokenError)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = projectTokenError.empty() ? "invalid project cache jump_token" : projectTokenError;
		return Utf8ToLocalText(r.dump());
	}

	project_source_cache::Snapshot snapshot;
	bool refreshed = false;
	std::string resolveError;
	std::string resolveTrace;
	if (!project_source_cache::ProjectSourceCacheManager::Instance().ResolveSnapshotForHit(
			projectToken,
			refreshCache,
			snapshot,
			&refreshed,
			&resolveError,
			&resolveTrace)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = resolveError.empty() ? "resolve project cache snapshot failed" : resolveError;
		r["trace"] = LocalToUtf8Text(resolveTrace);
		r["jump_token"] = jumpToken;
		return Utf8ToLocalText(r.dump());
	}

	const auto* page = FindProjectCachePageForAI(snapshot, projectToken);
	if (page == nullptr) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = "page not found in project cache snapshot";
		r["trace"] = LocalToUtf8Text(resolveTrace);
		r["jump_token"] = jumpToken;
		r["cache_revision"] = snapshot.revision;
		return Utf8ToLocalText(r.dump());
	}

	const auto& pageLines = page->lines;
	const int totalLines = static_cast<int>(pageLines.size());
	if (totalLines <= 0) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = "project cache page is empty";
		r["trace"] = LocalToUtf8Text(resolveTrace);
		r["page_name"] = LocalToUtf8Text(page->name);
		r["type_key"] = page->typeKey;
		r["type_name"] = LocalToUtf8Text(page->typeName);
		return Utf8ToLocalText(r.dump());
	}

	const int hintLineNumber = projectToken.lineNumber;
	int resolvedLineNumber = -1;
	std::string resolvedBy;
	std::string matchedCandidate;
	int matchedCandidateTotal = 0;
	if (snapshot.revision == projectToken.revision &&
		hintLineNumber >= 1 &&
		hintLineNumber <= totalLines) {
		resolvedLineNumber = hintLineNumber;
		resolvedBy = "token_line_number";
	}
	else if (!TrimAsciiCopy(searchText).empty()) {
		const std::vector<std::string> candidates = BuildSearchTextCandidatesForAI(searchText, page->name);
		if (!TryResolveProjectSearchLineNumberForAI(
				pageLines,
				candidates,
				sameTextOccurrenceIndex,
				hintLineNumber,
				resolvedLineNumber,
				resolvedBy,
				matchedCandidate,
				matchedCandidateTotal)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = "resolve matched line failed";
			r["trace"] = LocalToUtf8Text(resolveTrace);
			r["page_name"] = LocalToUtf8Text(page->name);
			r["type_key"] = page->typeKey;
			r["type_name"] = LocalToUtf8Text(page->typeName);
			r["search_line_hint"] = hintLineNumber;
			r["resolve_method"] = resolvedBy;
			r["search_text"] = LocalToUtf8Text(searchText);
			r["cache_revision"] = snapshot.revision;
			return Utf8ToLocalText(r.dump());
		}
	}
	else {
		resolvedLineNumber = (std::clamp)(hintLineNumber, 1, totalLines);
		resolvedBy = "token_line_hint_fallback";
	}

	const int startLine = (std::max)(1, resolvedLineNumber - contextBefore);
	const int endLine = (std::min)(totalLines, resolvedLineNumber + contextAfter);
	nlohmann::json lines = nlohmann::json::array();
	std::string text;
	for (int lineNo = startLine; lineNo <= endLine; ++lineNo) {
		const std::string& line = pageLines[static_cast<size_t>(lineNo - 1)];
		if (!text.empty()) {
			text += "\n";
		}
		text += line;
		nlohmann::json row;
		row["line_number"] = lineNo;
		row["text"] = LocalToUtf8Text(line);
		row["is_match_line"] = (lineNo == resolvedLineNumber);
		lines.push_back(std::move(row));
	}

	nlohmann::json r;
	r["ok"] = true;
	r["target_type"] = "project_cache";
	r["page_name"] = LocalToUtf8Text(page->name);
	r["type_key"] = page->typeKey;
	r["type_name"] = LocalToUtf8Text(page->typeName);
	r["jump_token"] = jumpToken;
	r["search_text"] = LocalToUtf8Text(searchText);
	r["same_text_occurrence_index"] = sameTextOccurrenceIndex;
	r["search_line_hint"] = hintLineNumber;
	r["resolved_line_number"] = resolvedLineNumber;
	r["resolve_method"] = resolvedBy;
	r["matched_candidate"] = LocalToUtf8Text(matchedCandidate);
	r["matched_candidate_total"] = matchedCandidateTotal;
	r["from_cache"] = !refreshed;
	r["code_hash"] = std::format("project-cache:{}:{}", snapshot.revision, page->pageIndex);
	r["source_kind"] = BuildProjectCacheSourceKindForAI(snapshot);
	r["parsed_input_kind"] = snapshot.parsedInputKind;
	r["source_file_path"] = LocalToUtf8Text(snapshot.sourcePath);
	r["cache_revision"] = snapshot.revision;
	if (!snapshot.snapshotPath.empty()) {
		r["snapshot_path"] = LocalToUtf8Text(snapshot.snapshotPath);
	}
	r["total_lines"] = totalLines;
	r["start_line"] = startLine;
	r["end_line"] = endLine;
	r["returned_line_count"] = endLine - startLine + 1;
	r["trace"] = LocalToUtf8Text(resolveTrace);
	r["warning"] = LocalToUtf8Text("该结果来自当前工程源码缓存，不会切换 IDE 页面。若缓存版本已变化且 refresh_cache=true，系统会先刷新缓存，再优先按 search_text 与 same_text_occurrence_index 重新定位匹配行。");
	r["text"] = LocalToUtf8Text(text);
	r["lines"] = std::move(lines);
	outOk = true;
	return Utf8ToLocalText(r.dump());
}

std::string BuildSearchPublicCodeJsonOnMainThread(const std::string& argumentsJson, bool& outOk)
{
	nlohmann::json args;
	try {
		args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = std::string("invalid arguments json: ") + ex.what();
		return Utf8ToLocalText(r.dump());
	}

	PublicCodeSearchSpecForAI spec;
	std::string specError;
	if (!TryBuildPublicCodeSearchSpecForAI(args, spec, specError)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = specError;
		return Utf8ToLocalText(r.dump());
	}

	const int limit = args.contains("limit") && args["limit"].is_number_integer()
		? (std::clamp)(args["limit"].get<int>(), 1, 500)
		: 100;

	nlohmann::json matches = nlohmann::json::array();
	std::vector<std::string> warnings;

	if (spec.searchProjectCache && static_cast<int>(matches.size()) < limit) {
		project_source_cache::Snapshot snapshot;
		bool refreshed = false;
		std::string projectError;
		std::string projectTrace;
		if (project_source_cache::ProjectSourceCacheManager::Instance().EnsureCurrentSourceLatest(
				snapshot,
				true,
				&refreshed,
				&projectError,
				&projectTrace)) {
			const auto projectHits = CollectProjectCacheSearchHitsForAI(snapshot, [&](const std::string& line, std::vector<std::string>& matchedKeywords) {
				return MatchPublicCodeSearchLineForAI(line, spec, matchedKeywords);
			});

			for (size_t i = 0; i < projectHits.size() && static_cast<int>(matches.size()) < limit; ++i) {
				nlohmann::json row;
				row["target_type"] = "project_cache";
				row["read_tool"] = "read_project_source_cache_code";
				row["jump_tool"] = "jump_to_search_result";
				row["source_kind"] = BuildProjectCacheSourceKindForAI(snapshot);
				row["parsed_input_kind"] = snapshot.parsedInputKind;
				row["source_file_path"] = LocalToUtf8Text(snapshot.sourcePath);
				row["cache_revision"] = snapshot.revision;
				row["page_name"] = LocalToUtf8Text(projectHits[i].info.pageName);
				row["page_type_key"] = projectHits[i].info.pageTypeKey;
				row["page_type_name"] = LocalToUtf8Text(projectHits[i].info.pageTypeName);
				row["line_number"] = projectHits[i].info.lineNumber;
				row["text"] = LocalToUtf8Text(projectHits[i].info.text);
				row["jump_token"] = projectHits[i].jumpToken;
				row["hit_index_in_page"] = projectHits[i].info.hitIndexInPage;
				row["hit_total_in_page"] = projectHits[i].info.hitTotalInPage;
				row["same_text_occurrence_index"] = projectHits[i].info.sameTextOccurrenceIndex;
				row["same_text_occurrence_total"] = projectHits[i].info.sameTextOccurrenceTotal;
				if (!snapshot.snapshotPath.empty()) {
					row["snapshot_path"] = LocalToUtf8Text(snapshot.snapshotPath);
				}
				if (!projectHits[i].matchedKeywords.empty()) {
					nlohmann::json keywords = nlohmann::json::array();
					for (const auto& keyword : projectHits[i].matchedKeywords) {
						keywords.push_back(LocalToUtf8Text(keyword));
					}
					row["matched_keywords"] = std::move(keywords);
				}
				if (spec.useRegex) {
					row["matched_regex"] = LocalToUtf8Text(spec.regexText);
				}
				matches.push_back(std::move(row));
			}
		}
		else {
			warnings.push_back(std::string("工程源码缓存搜索失败: ") +
				(projectError.empty() ? "project_cache_refresh_failed" : projectError));
			if (!projectTrace.empty()) {
				warnings.push_back(std::string("工程源码缓存 trace: ") + projectTrace);
			}
		}
	}

	if (spec.searchProject && static_cast<int>(matches.size()) < limit) {
		if (spec.keywords.empty()) {
			warnings.push_back("当前请求包含 project 搜索，但未提供 keyword/keywords。IDE 工程搜索只能以关键字作为种子，因此本次未包含 project 命中。");
		}
		else {
			const std::string ideSeedKeyword = spec.keywords.front();
			bool dialogHandled = false;
			const auto hits = e571::DebugSearchDirectGlobalKeywordHiddenDetailed(
				ideSeedKeyword.c_str(),
				GetCurrentProcessImageBaseForAI(),
				&dialogHandled);

			std::vector<ProgramTreeItemInfo> items;
			std::string listError;
			TryListProgramTreeItemsForAI(items, &listError);

			if (!listError.empty()) {
				warnings.push_back(std::string("工程页类型解析存在异常: ") + listError);
			}

			const auto infos = BuildKeywordSearchResultInfosForAI(hits, items);
			for (size_t i = 0; i < infos.size() && static_cast<int>(matches.size()) < limit; ++i) {
				std::vector<std::string> matchedKeywords;
				if (!MatchPublicCodeSearchLineForAI(infos[i].text, spec, matchedKeywords)) {
					continue;
				}

				nlohmann::json row;
				row["target_type"] = "project";
				row["read_tool"] = "read_project_search_result_code";
				row["jump_tool"] = "jump_to_search_result";
				row["source_kind"] = "ide_hidden_search";
				row["page_name"] = LocalToUtf8Text(infos[i].pageName);
				row["page_type_key"] = infos[i].pageTypeKey;
				row["page_type_name"] = LocalToUtf8Text(infos[i].pageTypeName);
				row["line_number"] = infos[i].lineNumber;
				row["text"] = LocalToUtf8Text(infos[i].text);
				row["jump_token"] = BuildSearchJumpToken(hits[i]);
				row["hit_index_in_page"] = infos[i].hitIndexInPage;
				row["hit_total_in_page"] = infos[i].hitTotalInPage;
				row["same_text_occurrence_index"] = infos[i].sameTextOccurrenceIndex;
				row["same_text_occurrence_total"] = infos[i].sameTextOccurrenceTotal;
				if (!matchedKeywords.empty()) {
					nlohmann::json keywords = nlohmann::json::array();
					for (const auto& keyword : matchedKeywords) {
						keywords.push_back(LocalToUtf8Text(keyword));
					}
					row["matched_keywords"] = std::move(keywords);
				}
				if (spec.useRegex) {
					row["matched_regex"] = LocalToUtf8Text(spec.regexText);
				}
				matches.push_back(std::move(row));
			}

			if (dialogHandled) {
				warnings.push_back("IDE 搜索过程中出现过对话框处理。");
			}
		}
	}

	if (spec.searchModules) {
		std::vector<std::string> modulePaths;
		std::string resolveError;

		const std::string moduleName = args.contains("module_name") && args["module_name"].is_string()
			? Utf8ToLocalText(args["module_name"].get<std::string>())
			: std::string();
		const std::string modulePath = args.contains("module_path") && args["module_path"].is_string()
			? Utf8ToLocalText(args["module_path"].get<std::string>())
			: std::string();

		if (!TrimAsciiCopy(moduleName).empty() || !TrimAsciiCopy(modulePath).empty()) {
			std::string resolvedPath;
			if (!ResolveImportedModulePathForAI(moduleName, modulePath, resolvedPath, resolveError)) {
				nlohmann::json r;
				r["ok"] = false;
				r["error"] = resolveError;
				return Utf8ToLocalText(r.dump());
			}
			modulePaths.push_back(resolvedPath);
		}
		else if (!TryListImportedModulePathsForAI(modulePaths, &resolveError)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = resolveError.empty() ? "list imported modules failed" : resolveError;
			return Utf8ToLocalText(r.dump());
		}

		for (const auto& path : modulePaths) {
			ModulePublicInfoCacheEntry entry;
			std::string loadError;
			if (!TryLoadModulePublicInfoCacheEntryForAI(path, entry, loadError, nullptr)) {
				continue;
			}

			for (size_t lineIndex = 0;
				lineIndex < entry.lines.size() && static_cast<int>(matches.size()) < limit;
				++lineIndex) {
				std::vector<std::string> matchedKeywords;
				const std::string& line = entry.lines[lineIndex];
				if (!MatchPublicCodeSearchLineForAI(line, spec, matchedKeywords)) {
					continue;
				}

				nlohmann::json row;
				row["target_type"] = "module";
				row["read_tool"] = "read_module_public_code";
				row["module_path"] = LocalToUtf8Text(entry.dump.modulePath);
				row["module_name"] = LocalToUtf8Text(GetFileStemForAI(entry.dump.modulePath));
				row["file_name"] = LocalToUtf8Text(GetFileNameOnlyForAI(entry.dump.modulePath));
				row["md5"] = entry.md5;
				row["source_kind"] = entry.dump.sourceKind;
				row["line_number"] = static_cast<int>(lineIndex) + 1;
				row["text"] = LocalToUtf8Text(line);
				if (!matchedKeywords.empty()) {
					nlohmann::json keywords = nlohmann::json::array();
					for (const auto& keyword : matchedKeywords) {
						keywords.push_back(LocalToUtf8Text(keyword));
					}
					row["matched_keywords"] = std::move(keywords);
				}
				if (spec.useRegex) {
					row["matched_regex"] = LocalToUtf8Text(spec.regexText);
				}
				if (!entry.md5.empty()) {
					row["cache_path"] = LocalToUtf8Text(GetModulePublicInfoCachePathForAI(entry.md5).string());
				}
				matches.push_back(std::move(row));
			}

			if (static_cast<int>(matches.size()) >= limit) {
				break;
			}
		}
	}

	if (spec.searchSupportLibraries && static_cast<int>(matches.size()) < limit) {
		std::vector<SupportLibraryInfoHeaderForAI> libs;
		std::string error;

		const bool hasSupportLibraryFilter =
			(args.contains("support_library_index") && args["support_library_index"].is_number_integer()) ||
			(args.contains("support_library_name") && args["support_library_name"].is_string()) ||
			(args.contains("support_library_path") && args["support_library_path"].is_string());

		if (hasSupportLibraryFilter) {
			nlohmann::json libArgs = nlohmann::json::object();
			if (args.contains("support_library_index")) {
				libArgs["index"] = args["support_library_index"];
			}
			if (args.contains("support_library_name")) {
				libArgs["name"] = args["support_library_name"];
			}
			if (args.contains("support_library_path")) {
				libArgs["file_path"] = args["support_library_path"];
			}

			SupportLibraryInfoHeaderForAI header;
			if (!ResolveSupportLibraryHeaderForAI(libArgs, header, error)) {
				nlohmann::json r;
				r["ok"] = false;
				r["error"] = error;
				return Utf8ToLocalText(r.dump());
			}
			libs.push_back(std::move(header));
		}
		else if (!TryListSupportLibrariesForAI(libs, &error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "list support libraries failed" : error;
			return Utf8ToLocalText(r.dump());
		}

		for (const auto& lib : libs) {
			nlohmann::json resolveArgs = nlohmann::json::object();
			if (lib.index >= 0) {
				resolveArgs["index"] = lib.index;
			}
			if (!lib.filePath.empty()) {
				resolveArgs["file_path"] = LocalToUtf8Text(lib.filePath);
			}
			else if (!lib.name.empty()) {
				resolveArgs["name"] = LocalToUtf8Text(lib.name);
			}

			SupportLibraryDumpCacheEntry entry;
			SupportLibraryInfoHeaderForAI resolvedHeader;
			std::string loadError;
			if (!TryResolveSupportLibraryEntryForReadForAI(resolveArgs, entry, &resolvedHeader, loadError)) {
				continue;
			}

			for (size_t lineIndex = 0;
				lineIndex < entry.lines.size() && static_cast<int>(matches.size()) < limit;
				++lineIndex) {
				std::vector<std::string> matchedKeywords;
				const std::string& line = entry.lines[lineIndex];
				if (!MatchPublicCodeSearchLineForAI(line, spec, matchedKeywords)) {
					continue;
				}

				nlohmann::json row;
				row["target_type"] = "support_library";
				row["read_tool"] = "read_support_library_public_code";
				row["support_library_name"] = LocalToUtf8Text(
					GetJsonStringFieldLocalForAI(entry.dumpJson, "support_library_name"));
				row["file_path"] = LocalToUtf8Text(
					GetJsonStringFieldLocalForAI(entry.dumpJson, "file_path"));
				row["file_name"] = LocalToUtf8Text(
					GetJsonStringFieldLocalForAI(entry.dumpJson, "file_name"));
				row["md5"] = entry.md5;
				row["source_kind"] = entry.sourceKind;
				row["line_number"] = static_cast<int>(lineIndex) + 1;
				row["text"] = LocalToUtf8Text(line);
				if (!matchedKeywords.empty()) {
					nlohmann::json keywords = nlohmann::json::array();
					for (const auto& keyword : matchedKeywords) {
						keywords.push_back(LocalToUtf8Text(keyword));
					}
					row["matched_keywords"] = std::move(keywords);
				}
				if (spec.useRegex) {
					row["matched_regex"] = LocalToUtf8Text(spec.regexText);
				}
				if (!entry.md5.empty()) {
					row["cache_path"] = LocalToUtf8Text(GetSupportLibraryInfoCachePathForAI(entry.md5).string());
				}
				matches.push_back(std::move(row));
			}

			if (static_cast<int>(matches.size()) >= limit) {
				break;
			}
		}
	}

	nlohmann::json r;
	r["ok"] = true;
	r["match_count"] = matches.size();
	r["keyword_mode"] = spec.keywordMode;
	r["case_sensitive"] = spec.caseSensitive;
	r["target_types"] = nlohmann::json::array();
	if (spec.searchProject) {
		r["target_types"].push_back("project");
	}
	if (spec.searchProjectCache) {
		r["target_types"].push_back("project_cache");
	}
	if (spec.searchModules) {
		r["target_types"].push_back("module");
	}
	if (spec.searchSupportLibraries) {
		r["target_types"].push_back("support_library");
	}
	if (!spec.keywords.empty()) {
		nlohmann::json keywords = nlohmann::json::array();
		for (const auto& keyword : spec.keywords) {
			keywords.push_back(LocalToUtf8Text(keyword));
		}
		r["keywords"] = std::move(keywords);
	}
	if (spec.useRegex) {
		r["regex"] = LocalToUtf8Text(spec.regexText);
		if (!spec.regexFlags.empty()) {
			r["regex_flags"] = spec.regexFlags;
		}
	}
	std::string warning = "这里是完全搜索：默认会先刷新并搜索当前工程源码缓存，再补充当前 IDE 工程源码命中，以及模块公开声明文本、支持库公开声明文本。命中后请根据结果里的 target_type 与 read_tool 继续读取。若主要查当前工程源码并希望获得稳定页名和行号，优先用 search_project_source_cache；若要快速定位子程序、常量、数据类型、DLL命令、程序集变量、参数、局部变量、全局变量，可考虑使用类似“.子程序 XXXX”、“.常量 XXXX”、“.数据类型 XXXX”、“.DLL命令 XXXX”、“.参数 XXXX”、“.全局变量 XXXX”来进行。";
	for (const auto& extraWarning : warnings) {
		warning += " ";
		warning += extraWarning;
	}
	r["warning"] = LocalToUtf8Text(warning);
	r["matches"] = std::move(matches);
	outOk = true;
	return Utf8ToLocalText(r.dump());
}

void WarmupImportedModulePublicInfoCacheOnMainThread()
{
	static bool s_warmupDone = false;
	if (s_warmupDone) {
		return;
	}
	s_warmupDone = true;

	std::vector<std::string> paths;
	std::string listError;
	if (!TryListImportedModulePathsForAI(paths, &listError)) {
		OutputStringToELog(std::format(
			"[ModulePublicInfoWarmup] 枚举导入模块失败 error={}",
			listError.empty() ? "list_imported_modules_failed" : listError));
		return;
	}

	if (paths.empty()) {
		OutputStringToELog("[ModulePublicInfoWarmup] 当前项目没有导入模块，跳过预热");
		return;
	}

	int okCount = 0;
	int failCount = 0;
	int cacheHitCount = 0;
	for (const auto& path : paths) {
		ModulePublicInfoCacheEntry entry;
		std::string loadError;
		bool cacheHit = false;
		if (TryLoadModulePublicInfoCacheEntryForAI(path, entry, loadError, &cacheHit)) {
			++okCount;
			if (cacheHit) {
				++cacheHitCount;
			}
			continue;
		}

		++failCount;
		OutputStringToELog(std::format(
			"[ModulePublicInfoWarmup] 预热失败 module={} error={}",
			path,
			loadError.empty() ? "load_module_public_info_failed" : loadError));
	}

	OutputStringToELog(std::format(
		"[ModulePublicInfoWarmup] 预热完成 total={} ok={} cacheHit={} fail={} dir={}",
		paths.size(),
		okCount,
		cacheHitCount,
		failCount,
		GetModulePublicInfoDirectoryForAI().string()));
}

std::string BuildProgramSearchResultJsonOnMainThread(const std::string& argumentsJson, bool& outOk)
{
	nlohmann::json args;
	try {
		args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = std::string("invalid arguments json: ") + ex.what();
		return Utf8ToLocalText(r.dump());
	}

	std::string keyword;
	int limit = 50;
	if (args.contains("keyword") && args["keyword"].is_string()) {
		keyword = Utf8ToLocalText(args["keyword"].get<std::string>());
	}
	if (args.contains("limit") && args["limit"].is_number_integer()) {
		limit = (std::clamp)(args["limit"].get<int>(), 1, 200);
	}
	keyword = TrimAsciiCopy(keyword);
	if (keyword.empty()) {
		return R"({"ok":false,"error":"keyword is required"})";
	}

	bool dialogHandled = false;
	const auto hits = e571::DebugSearchDirectGlobalKeywordHiddenDetailed(
		keyword.c_str(),
		GetCurrentProcessImageBaseForAI(),
		&dialogHandled);

	std::vector<ProgramTreeItemInfo> items;
	std::string listError;
	TryListProgramTreeItemsForAI(items, &listError);

	const auto infos = BuildKeywordSearchResultInfosForAI(hits, items);
	nlohmann::json results = nlohmann::json::array();
	for (size_t i = 0; i < infos.size() && static_cast<int>(results.size()) < limit; ++i) {
		nlohmann::json row;
		row["index"] = static_cast<int>(i);
		row["page_name"] = LocalToUtf8Text(infos[i].pageName);
		row["page_type_key"] = infos[i].pageTypeKey;
		row["page_type_name"] = LocalToUtf8Text(infos[i].pageTypeName);
		row["line_number"] = infos[i].lineNumber;
		row["text"] = LocalToUtf8Text(infos[i].text);
		row["jump_token"] = BuildSearchJumpToken(hits[i]);
		row["read_tool"] = "read_project_search_result_code";
		row["source_kind"] = "ide_hidden_search";
		row["hit_index_in_page"] = infos[i].hitIndexInPage;
		row["hit_total_in_page"] = infos[i].hitTotalInPage;
		row["same_text_occurrence_index"] = infos[i].sameTextOccurrenceIndex;
		row["same_text_occurrence_total"] = infos[i].sameTextOccurrenceTotal;
		results.push_back(std::move(row));
	}

	nlohmann::json r;
	r["ok"] = true;
	r["keyword"] = LocalToUtf8Text(keyword);
	r["count"] = hits.size();
	r["dialog_handled"] = dialogHandled;
	r["source_kind"] = "ide_hidden_search";
	r["code_kind"] = "pseudo_reference";
	r["warning"] = LocalToUtf8Text("这里仅搜索当前 IDE 工程源码命中，不包含模块公开声明与支持库公开声明。搜索结果文本以及后续按页面名抓取到的代码，与IDE正常编辑页结构可能不同，仅可作为伪代码参考。若你想优先使用当前工程源码缓存并直接拿到稳定页名和行号，请改用 search_project_source_cache。");
	r["results"] = std::move(results);
	if (!listError.empty()) {
		r["page_type_lookup_error"] = listError;
	}
	outOk = true;
	return Utf8ToLocalText(r.dump());
}

std::string BuildReadProjectSearchResultCodeJsonOnMainThread(const std::string& argumentsJson, bool& outOk)
{
	nlohmann::json args;
	try {
		args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
	}
	catch (const std::exception& ex) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = std::string("invalid arguments json: ") + ex.what();
		return Utf8ToLocalText(r.dump());
	}

	const std::string jumpToken = GetJsonStringArgumentLocal(args, "jump_token");
	const std::string searchText = GetJsonStringArgumentLocal(args, "search_text");
	const int sameTextOccurrenceIndex = (std::max)(0, GetJsonIntArgument(args, "same_text_occurrence_index", 0));
	const int contextBefore = (std::clamp)(GetJsonIntArgument(args, "context_before", 3), 0, 50);
	const int contextAfter = (std::clamp)(GetJsonIntArgument(args, "context_after", 8), 0, 50);
	const bool refreshCache = GetJsonBoolArgument(args, "refresh_cache", true);
	if (TrimAsciiCopy(jumpToken).empty()) {
		return R"({"ok":false,"error":"jump_token is required"})";
	}
	if (TrimAsciiCopy(searchText).empty()) {
		return R"({"ok":false,"error":"search_text is required"})";
	}

	e571::DirectGlobalSearchDebugHit hit{};
	if (!ParseSearchJumpToken(jumpToken, hit)) {
		return R"({"ok":false,"error":"invalid jump_token"})";
	}

	std::string jumpTrace;
	if (!e571::DebugJumpToSearchHit(hit, GetCurrentProcessImageBaseForAI(), &jumpTrace)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = "jump to search result failed";
		r["trace"] = jumpTrace;
		return Utf8ToLocalText(r.dump());
	}

	std::string currentPageName;
	std::string currentPageType;
	std::string currentPageTrace;
	if (!IDEFacade::Instance().GetCurrentPageName(currentPageName, &currentPageType, &currentPageTrace) ||
		TrimAsciiCopy(currentPageName).empty()) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = "get current page after jump failed";
		r["trace"] = jumpTrace;
		r["current_page_trace"] = currentPageTrace;
		return Utf8ToLocalText(r.dump());
	}

	ProgramTreeItemInfo item;
	std::string itemError;
	if (!TryGetProgramItemByNameForAI(currentPageName, std::string(), item, itemError)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = itemError.empty() ? "program item lookup failed" : itemError;
		r["trace"] = jumpTrace;
		r["current_page_name"] = LocalToUtf8Text(currentPageName);
		r["current_page_type"] = LocalToUtf8Text(currentPageType);
		return Utf8ToLocalText(r.dump());
	}

	std::string code;
	PageCodeCacheEntry cacheEntry;
	bool fromCache = false;
	std::string codeTrace;
	if (!TryLoadRealPageCodeForReadForAI(item, refreshCache, code, cacheEntry, fromCache, codeTrace, itemError)) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = itemError.empty() ? "read real page code failed" : itemError;
		r["trace"] = LocalToUtf8Text(jumpTrace + "|" + codeTrace);
		r["current_page_name"] = LocalToUtf8Text(item.name);
		r["current_page_type"] = LocalToUtf8Text(item.typeName);
		return Utf8ToLocalText(r.dump());
	}

	const auto pageLines = SplitLinesCopyForAI(NormalizeLineBreaksForAI(code));
	const std::vector<std::string> candidates = BuildSearchTextCandidatesForAI(searchText, item.name);
	const int hintLineNumber = hit.outerIndex + 1;
	int resolvedLineNumber = -1;
	std::string resolvedBy;
	std::string matchedCandidate;
	int matchedCandidateTotal = 0;
	const bool resolved = TryResolveProjectSearchLineNumberForAI(
		pageLines,
		candidates,
		sameTextOccurrenceIndex,
		hintLineNumber,
		resolvedLineNumber,
		resolvedBy,
		matchedCandidate,
		matchedCandidateTotal);
	if (!resolved) {
		nlohmann::json r;
		r["ok"] = false;
		r["error"] = "resolve matched line failed";
		r["trace"] = LocalToUtf8Text(jumpTrace + "|" + codeTrace);
		r["page_name"] = LocalToUtf8Text(item.name);
		r["type_key"] = item.typeKey;
		r["type_name"] = LocalToUtf8Text(item.typeName);
		r["search_line_hint"] = hintLineNumber;
		r["resolve_method"] = resolvedBy;
		r["search_text"] = LocalToUtf8Text(searchText);
		return Utf8ToLocalText(r.dump());
	}

	const int totalLines = static_cast<int>(pageLines.size());
	const int startLine = (std::max)(1, resolvedLineNumber - contextBefore);
	const int endLine = (std::min)(totalLines, resolvedLineNumber + contextAfter);

	nlohmann::json lines = nlohmann::json::array();
	std::string text;
	for (int lineNo = startLine; lineNo <= endLine; ++lineNo) {
		const std::string& line = pageLines[static_cast<size_t>(lineNo - 1)];
		if (!text.empty()) {
			text += "\n";
		}
		text += line;
		nlohmann::json row;
		row["line_number"] = lineNo;
		row["text"] = LocalToUtf8Text(line);
		row["is_match_line"] = (lineNo == resolvedLineNumber);
		lines.push_back(std::move(row));
	}

	nlohmann::json r;
	r["ok"] = true;
	r["page_name"] = LocalToUtf8Text(item.name);
	r["type_key"] = item.typeKey;
	r["type_name"] = LocalToUtf8Text(item.typeName);
	r["item_data"] = item.itemData;
	r["jump_token"] = jumpToken;
	r["search_text"] = LocalToUtf8Text(searchText);
	r["same_text_occurrence_index"] = sameTextOccurrenceIndex;
	r["search_line_hint"] = hintLineNumber;
	r["resolved_line_number"] = resolvedLineNumber;
	r["resolve_method"] = resolvedBy;
	r["matched_candidate"] = LocalToUtf8Text(matchedCandidate);
	r["matched_candidate_total"] = matchedCandidateTotal;
	r["from_cache"] = fromCache;
	r["code_hash"] = cacheEntry.codeHash;
	r["total_lines"] = totalLines;
	r["start_line"] = startLine;
	r["end_line"] = endLine;
	r["returned_line_count"] = endLine - startLine + 1;
	r["trace"] = LocalToUtf8Text(jumpTrace + "|" + codeTrace);
	r["warning"] = LocalToUtf8Text("该结果来自 IDE 搜索命中后跳转并抓取真实页面源码，再用搜索结果文本在整页代码中回定位得到的行号。若提供了 same_text_occurrence_index，系统会优先按同页同文本第 N 次出现定位；只有在无法匹配该序号时，才会退回到旧的最近提示行号策略。");
	r["text"] = LocalToUtf8Text(text);
	r["lines"] = std::move(lines);
	outOk = true;
	return Utf8ToLocalText(r.dump());
}

std::string ExecuteToolCallOnMainThread(const std::string& toolName, const std::string& argumentsJson, bool& outOk)
{
	outOk = false;

	if (toolName == "get_current_page_code") {
		const std::string sourceFilePath = RefreshCurrentSourceFilePathForAI();
		std::string pageName;
		std::string pageType;
		std::string pageNameTrace;
		const bool nameOk = IDEFacade::Instance().GetCurrentPageName(pageName, &pageType, &pageNameTrace);

		std::string pageCode;
		std::string codeKind;
		std::string codeTrace;
		if (nameOk && !TrimAsciiCopy(pageName).empty()) {
			std::vector<std::string> lookupNames;
			lookupNames.push_back(pageName);

			const size_t slashPos = pageName.find(" / ");
			if (slashPos != std::string::npos) {
				const std::string primaryName = TrimAsciiCopy(pageName.substr(0, slashPos));
				if (!primaryName.empty() && primaryName != pageName) {
					lookupNames.push_back(primaryName);
				}
			}

			for (const auto& lookupName : lookupNames) {
				ProgramTreeItemInfo item;
				std::string lookupError;
				if (TryGetProgramItemByNameForAI(lookupName, std::string(), item, lookupError)) {
					e571::NativeRealPageAccessResult accessResult{};
					std::string readError;
					if (TryReadRealPageCodeForAI(item, pageCode, accessResult, readError)) {
						codeKind = "real_source";
						codeTrace = accessResult.trace;
						break;
					}
					codeTrace = !accessResult.trace.empty() ? accessResult.trace : readError;
				}
				else if (!lookupError.empty()) {
					codeTrace = lookupError;
				}
			}
		}

		if (pageCode.empty()) {
			if (!IDEFacade::Instance().GetCurrentPageCode(pageCode)) {
				nlohmann::json r;
				r["ok"] = false;
				r["error"] = "GetCurrentPageCode failed";
				r["page_name_ok"] = nameOk;
				r["page_name"] = LocalToUtf8Text(pageName);
				r["page_type"] = LocalToUtf8Text(pageType);
				r["page_name_trace"] = pageNameTrace;
				if (!codeTrace.empty()) {
					r["trace"] = LocalToUtf8Text(codeTrace);
				}
				return Utf8ToLocalText(r.dump());
			}
			codeKind = "editor_clipboard";
			if (codeTrace.empty()) {
				codeTrace = "clipboard";
			}
		}

		nlohmann::json r;
		r["ok"] = true;
		r["code"] = LocalToUtf8Text(pageCode);
		r["code_kind"] = codeKind;
		r["source_file_path"] = LocalToUtf8Text(sourceFilePath);
		r["page_name_ok"] = nameOk;
		r["page_name"] = LocalToUtf8Text(pageName);
		r["page_type"] = LocalToUtf8Text(pageType);
		r["page_name_trace"] = pageNameTrace;
		if (!codeTrace.empty()) {
			r["trace"] = LocalToUtf8Text(codeTrace);
		}
		LocalMcpServer::UpdateInstanceHints(sourceFilePath, pageName, pageType);
		outOk = true;
		return Utf8ToLocalText(r.dump());
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
			return Utf8ToLocalText(r.dump());
		}

		nlohmann::json r;
		r["ok"] = true;
		r["source_file_path"] = LocalToUtf8Text(sourceFilePath);
		r["page_name"] = LocalToUtf8Text(pageName);
		r["page_type"] = LocalToUtf8Text(pageType);
		r["page_name_trace"] = pageNameTrace;
		LocalMcpServer::UpdateInstanceHints(sourceFilePath, pageName, pageType);
		outOk = true;
		return Utf8ToLocalText(r.dump());
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
		r["mcp_running"] = LocalMcpServer::IsRunning();
		r["mcp_instance_id"] = LocalMcpServer::GetInstanceId();
		r["mcp_port"] = LocalMcpServer::GetBoundPort();
		r["mcp_endpoint"] = LocalMcpServer::GetEndpoint();
		outOk = true;
		return Utf8ToLocalText(r.dump());
	}

	if (toolName == "get_program_item_project_cache_code") {
		nlohmann::json args;
		try {
			args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return Utf8ToLocalText(r.dump());
		}

		const std::string pageName = GetJsonStringArgumentLocal(args, "page_name");
		const std::string kind = GetJsonStringArgumentLocal(args, "kind");
		const bool refreshCache = GetJsonBoolArgument(args, "refresh_cache", true);
		if (TrimAsciiCopy(pageName).empty()) {
			return R"({"ok":false,"error":"page_name is required"})";
		}

		ProgramTreeItemInfo item;
		std::string error;
		if (!TryGetProgramItemByNameForAI(pageName, kind, item, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "program item lookup failed" : error;
			return Utf8ToLocalText(r.dump());
		}

		std::string code;
		project_source_cache::Snapshot snapshot;
		const project_source_cache::Page* page = nullptr;
		bool refreshed = false;
		std::string trace;
		if (!TryGetProgramItemProjectCacheCodeForAI(
				item,
				refreshCache,
				code,
				snapshot,
				page,
				refreshed,
				trace,
				error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "get program item project cache code failed" : error;
			r["trace"] = LocalToUtf8Text(trace);
			r["page_name"] = LocalToUtf8Text(item.name);
			r["type_key"] = item.typeKey;
			r["type_name"] = LocalToUtf8Text(item.typeName);
			r["refresh_cache"] = refreshCache;
			return Utf8ToLocalText(r.dump());
		}

		nlohmann::json r;
		r["ok"] = true;
		r["page_name"] = LocalToUtf8Text(item.name);
		r["type_key"] = item.typeKey;
		r["type_name"] = LocalToUtf8Text(item.typeName);
		r["item_data"] = item.itemData;
		r["page_index"] = page != nullptr ? page->pageIndex : -1;
		r["code_kind"] = "project_cache_source";
		r["source_kind"] = BuildProjectCacheSourceKindForAI(snapshot);
		r["parsed_input_kind"] = snapshot.parsedInputKind;
		r["source_file_path"] = LocalToUtf8Text(snapshot.sourcePath);
		r["cache_revision"] = snapshot.revision;
		r["from_cache"] = !refreshed;
		r["refresh_cache"] = refreshCache;
		r["no_switch"] = true;
		r["trace"] = LocalToUtf8Text(trace);
		r["code_hash"] = BuildStableTextHashForRealCode(code);
		r["code"] = LocalToUtf8Text(code);
		r["warning"] = LocalToUtf8Text("该结果来自当前工程源码缓存：先把当前工程从 IDE 内存直序列化为 .e 二进制，再由 e2txt 解析得到对应页面代码。它不会切换 IDE 页面，但与编辑器当前页的真实复制结果可能存在格式差异。");
		outOk = true;
		return Utf8ToLocalText(r.dump());
	}

	if (toolName == "get_program_item_real_code") {
		nlohmann::json args;
		try {
			args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return Utf8ToLocalText(r.dump());
		}

		const std::string pageName = GetJsonStringArgumentLocal(args, "page_name");
		const std::string kind = GetJsonStringArgumentLocal(args, "kind");
		if (TrimAsciiCopy(pageName).empty()) {
			return R"({"ok":false,"error":"page_name is required"})";
		}

		ProgramTreeItemInfo item;
		std::string error;
		if (!TryGetProgramItemByNameForAI(pageName, kind, item, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "program item lookup failed" : error;
			return Utf8ToLocalText(r.dump());
		}

		std::string code;
		e571::NativeRealPageAccessResult accessResult{};
		if (!TryReadRealPageCodeForAI(item, code, accessResult, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "get real page code failed" : error;
			r["trace"] = LocalToUtf8Text(accessResult.trace);
			r["item_data"] = item.itemData;
			return Utf8ToLocalText(r.dump());
		}

		PageCodeCacheEntry cacheEntry;
		PutPageCodeCacheEntryForAI(item, code, nullptr, &cacheEntry);

		nlohmann::json r;
		r["ok"] = true;
		r["page_name"] = LocalToUtf8Text(item.name);
		r["type_key"] = item.typeKey;
		r["type_name"] = LocalToUtf8Text(item.typeName);
		r["item_data"] = item.itemData;
		r["code_kind"] = "real_source";
		r["trace"] = LocalToUtf8Text(accessResult.trace);
		r["captured_custom_format"] = accessResult.capturedCustomFormat;
		r["code_hash"] = cacheEntry.codeHash;
		r["code"] = LocalToUtf8Text(code);
		outOk = true;
		return Utf8ToLocalText(r.dump());
	}

	if (toolName == "read_program_item_real_code") {
		nlohmann::json args;
		try {
			args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return Utf8ToLocalText(r.dump());
		}

		const std::string pageName = GetJsonStringArgumentLocal(args, "page_name");
		const std::string kind = GetJsonStringArgumentLocal(args, "kind");
		const int offsetLines = (std::max)(0, GetJsonIntArgument(args, "offset_lines", 0));
		const int limitLines = GetJsonIntArgument(args, "limit_lines", 0);
		const bool withLineNumbers = GetJsonBoolArgument(args, "with_line_numbers", false);
		const bool refreshCache = GetJsonBoolArgument(args, "refresh_cache", false);
		if (TrimAsciiCopy(pageName).empty()) {
			return R"({"ok":false,"error":"page_name is required"})";
		}

		ProgramTreeItemInfo item;
		std::string error;
		if (!TryGetProgramItemByNameForAI(pageName, kind, item, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "program item lookup failed" : error;
			return Utf8ToLocalText(r.dump());
		}

		std::string code;
		PageCodeCacheEntry cacheEntry;
		bool fromCache = false;
		std::string trace;
		if (!TryLoadRealPageCodeForReadForAI(item, refreshCache, code, cacheEntry, fromCache, trace, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "read real page code failed" : error;
			r["trace"] = LocalToUtf8Text(trace);
			r["page_name"] = LocalToUtf8Text(item.name);
			r["type_key"] = item.typeKey;
			return Utf8ToLocalText(r.dump());
		}

		int totalLines = 0;
		int returnedLines = 0;
		int startLine = 1;
		const std::string view = BuildRealCodeView(
			code,
			offsetLines,
			limitLines,
			withLineNumbers,
			&totalLines,
			&returnedLines,
			&startLine);

		nlohmann::json r;
		r["ok"] = true;
		r["page_name"] = LocalToUtf8Text(item.name);
		r["type_key"] = item.typeKey;
		r["type_name"] = LocalToUtf8Text(item.typeName);
		r["item_data"] = item.itemData;
		r["code_kind"] = "real_source";
		r["trace"] = LocalToUtf8Text(trace);
		r["from_cache"] = fromCache;
		r["code_hash"] = cacheEntry.codeHash;
		r["total_lines"] = totalLines;
		r["returned_lines"] = returnedLines;
		r["start_line"] = startLine;
		r["is_partial"] = offsetLines > 0 || (limitLines > 0 && returnedLines < totalLines);
		r["with_line_numbers"] = withLineNumbers;
		r["code"] = LocalToUtf8Text(view);
		outOk = true;
		return Utf8ToLocalText(r.dump());
	}

	if (toolName == "edit_program_item_code") {
		nlohmann::json args;
		try {
			args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return Utf8ToLocalText(r.dump());
		}

		const std::string pageName = GetJsonStringArgumentLocal(args, "page_name");
		const std::string kind = GetJsonStringArgumentLocal(args, "kind");
		const std::string oldText = GetJsonStringArgumentLocal(args, "old_text");
		const std::string newText = GetJsonStringArgumentLocal(args, "new_text");
		if (TrimAsciiCopy(pageName).empty()) {
			return R"({"ok":false,"error":"page_name is required"})";
		}
		if (oldText.empty()) {
			return R"({"ok":false,"error":"old_text is required"})";
		}

		ProgramTreeItemInfo item;
		std::string error;
		if (!TryGetProgramItemByNameForAI(pageName, kind, item, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "program item lookup failed" : error;
			return Utf8ToLocalText(r.dump());
		}

		std::string baseCode;
		PageCodeCacheEntry cachedEntry;
		std::string trace;
		std::string currentCode;
		bool cacheRefreshed = false;
		if (!TryResolveRealPageWriteBaseForAI(item, std::string(), true, baseCode, cachedEntry, trace, currentCode, cacheRefreshed, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error;
			r["page_name"] = LocalToUtf8Text(item.name);
			r["type_key"] = item.typeKey;
			r["trace"] = LocalToUtf8Text(trace);
			r["cache_refreshed"] = cacheRefreshed;
			if (!currentCode.empty()) {
				r["code"] = LocalToUtf8Text(currentCode);
				r["code_hash"] = BuildStableTextHashForRealCode(currentCode);
			}
			return Utf8ToLocalText(r.dump());
		}

		size_t matchCount = 0;
		std::string replacedCode;
		if (!ReplaceExactlyOnceForAI(baseCode, oldText, newText, replacedCode, matchCount)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = matchCount == 0 ? "old_text not found in cached page" : "old_text matched multiple times in cached page";
			r["match_count"] = matchCount;
			r["page_name"] = LocalToUtf8Text(item.name);
			r["type_key"] = item.typeKey;
			return Utf8ToLocalText(r.dump());
		}

		if (BuildStableTextHashForRealCode(replacedCode) == BuildStableTextHashForRealCode(baseCode)) {
			nlohmann::json r;
			r["ok"] = true;
			r["no_changes"] = true;
			r["page_name"] = LocalToUtf8Text(item.name);
			r["type_key"] = item.typeKey;
			r["type_name"] = LocalToUtf8Text(item.typeName);
			r["item_data"] = item.itemData;
			r["match_count"] = matchCount;
			r["code_kind"] = "real_source";
			r["code_hash"] = BuildStableTextHashForRealCode(baseCode);
			r["code"] = LocalToUtf8Text(baseCode);
			outOk = true;
			return Utf8ToLocalText(r.dump());
		}

		PageCodeSnapshotEntry snapshot;
		std::string finalCode;
		bool rollbackAttempted = false;
		bool rollbackSucceeded = false;
		if (!TryWriteRealPageCodeForAI(
				item,
				baseCode,
				replacedCode,
				"edit_program_item_code",
				snapshot,
				finalCode,
				trace,
				rollbackAttempted,
				rollbackSucceeded,
				error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error;
			r["page_name"] = LocalToUtf8Text(item.name);
			r["type_key"] = item.typeKey;
			r["trace"] = LocalToUtf8Text(trace);
			r["rollback_attempted"] = rollbackAttempted;
			r["rollback_succeeded"] = rollbackSucceeded;
			return Utf8ToLocalText(r.dump());
		}

		const auto hunks = BuildRealPageStructuredPatch(baseCode, finalCode);

		nlohmann::json r;
		r["ok"] = true;
		r["page_name"] = LocalToUtf8Text(item.name);
		r["type_key"] = item.typeKey;
		r["type_name"] = LocalToUtf8Text(item.typeName);
		r["item_data"] = item.itemData;
		r["match_count"] = matchCount;
		r["trace"] = LocalToUtf8Text(trace);
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
		return Utf8ToLocalText(r.dump());
	}

	if (toolName == "multi_edit_program_item_code") {
		nlohmann::json args;
		try {
			args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return Utf8ToLocalText(r.dump());
		}

		const std::string pageName = GetJsonStringArgumentLocal(args, "page_name");
		const std::string kind = GetJsonStringArgumentLocal(args, "kind");
		const bool failOnUnmatched = GetJsonBoolArgument(args, "fail_on_unmatched", true);
		const bool atomic = GetJsonBoolArgument(args, "atomic", true);
		if (TrimAsciiCopy(pageName).empty()) {
			return R"({"ok":false,"error":"page_name is required"})";
		}

		std::vector<RealPageTextEditRequest> edits;
		std::string error;
		if (!TryParseRealPageTextEditsFromJson(args, "edits", edits, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error;
			return Utf8ToLocalText(r.dump());
		}
		if (edits.empty()) {
			return R"({"ok":false,"error":"edits is required"})";
		}

		ProgramTreeItemInfo item;
		if (!TryGetProgramItemByNameForAI(pageName, kind, item, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "program item lookup failed" : error;
			return Utf8ToLocalText(r.dump());
		}

		std::string baseCode;
		PageCodeCacheEntry cacheEntry;
		std::string trace;
		std::string currentCode;
		bool cacheRefreshed = false;
		if (!TryResolveRealPageWriteBaseForAI(item, std::string(), true, baseCode, cacheEntry, trace, currentCode, cacheRefreshed, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error;
			r["page_name"] = LocalToUtf8Text(item.name);
			r["type_key"] = item.typeKey;
			r["trace"] = LocalToUtf8Text(trace);
			r["cache_refreshed"] = cacheRefreshed;
			if (!currentCode.empty()) {
				r["code"] = LocalToUtf8Text(currentCode);
				r["code_hash"] = BuildStableTextHashForRealCode(currentCode);
			}
			return Utf8ToLocalText(r.dump());
		}

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
			r["code_hash"] = BuildStableTextHashForRealCode(baseCode);
			r["code"] = LocalToUtf8Text(baseCode);
			return Utf8ToLocalText(r.dump());
		}

		if (appliedCount == 0) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = "no edits applied";
			r["results"] = std::move(resultRows);
			return Utf8ToLocalText(r.dump());
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
			r["code_hash"] = BuildStableTextHashForRealCode(baseCode);
			r["code"] = LocalToUtf8Text(baseCode);
			outOk = true;
			return Utf8ToLocalText(r.dump());
		}

		PageCodeSnapshotEntry snapshot;
		std::string finalCode;
		bool rollbackAttempted = false;
		bool rollbackSucceeded = false;
		if (!TryWriteRealPageCodeForAI(
				item,
				baseCode,
				candidateCode,
				"multi_edit_program_item_code",
				snapshot,
				finalCode,
				trace,
				rollbackAttempted,
				rollbackSucceeded,
				error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error;
			r["results"] = std::move(resultRows);
			r["trace"] = LocalToUtf8Text(trace);
			r["rollback_attempted"] = rollbackAttempted;
			r["rollback_succeeded"] = rollbackSucceeded;
			return Utf8ToLocalText(r.dump());
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
		r["trace"] = LocalToUtf8Text(trace);
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
		return Utf8ToLocalText(r.dump());
	}

	if (toolName == "write_program_item_real_code") {
		nlohmann::json args;
		try {
			args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return Utf8ToLocalText(r.dump());
		}

		const std::string pageName = GetJsonStringArgumentLocal(args, "page_name");
		const std::string kind = GetJsonStringArgumentLocal(args, "kind");
		const std::string fullCode = GetJsonStringArgumentLocal(args, "full_code");
		const std::string expectedBaseHash = GetJsonStringArgumentLocal(args, "expected_base_hash");
		if (TrimAsciiCopy(pageName).empty()) {
			return R"({"ok":false,"error":"page_name is required"})";
		}
		if (fullCode.empty()) {
			return R"({"ok":false,"error":"full_code is required"})";
		}

		ProgramTreeItemInfo item;
		std::string error;
		if (!TryGetProgramItemByNameForAI(pageName, kind, item, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "program item lookup failed" : error;
			return Utf8ToLocalText(r.dump());
		}

		std::string baseCode;
		PageCodeCacheEntry cacheEntry;
		std::string trace;
		std::string currentCode;
		bool cacheRefreshed = false;
		if (!TryResolveRealPageWriteBaseForAI(item, expectedBaseHash, false, baseCode, cacheEntry, trace, currentCode, cacheRefreshed, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error;
			r["page_name"] = LocalToUtf8Text(item.name);
			r["type_key"] = item.typeKey;
			r["trace"] = LocalToUtf8Text(trace);
			r["cache_refreshed"] = cacheRefreshed;
			if (!currentCode.empty()) {
				r["code"] = LocalToUtf8Text(currentCode);
				r["code_hash"] = BuildStableTextHashForRealCode(currentCode);
			}
			return Utf8ToLocalText(r.dump());
		}

		const std::string normalizedFullCode = NormalizeRealCodeLineBreaksToCrLf(fullCode);
		if (BuildStableTextHashForRealCode(normalizedFullCode) == BuildStableTextHashForRealCode(baseCode)) {
			nlohmann::json r;
			r["ok"] = true;
			r["no_changes"] = true;
			r["page_name"] = LocalToUtf8Text(item.name);
			r["type_key"] = item.typeKey;
			r["type_name"] = LocalToUtf8Text(item.typeName);
			r["item_data"] = item.itemData;
			r["code_hash"] = BuildStableTextHashForRealCode(baseCode);
			r["code"] = LocalToUtf8Text(baseCode);
			outOk = true;
			return Utf8ToLocalText(r.dump());
		}

		PageCodeSnapshotEntry snapshot;
		std::string finalCode;
		bool rollbackAttempted = false;
		bool rollbackSucceeded = false;
		if (!TryWriteRealPageCodeForAI(
				item,
				baseCode,
				normalizedFullCode,
				"write_program_item_real_code",
				snapshot,
				finalCode,
				trace,
				rollbackAttempted,
				rollbackSucceeded,
				error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error;
			r["trace"] = LocalToUtf8Text(trace);
			r["rollback_attempted"] = rollbackAttempted;
			r["rollback_succeeded"] = rollbackSucceeded;
			return Utf8ToLocalText(r.dump());
		}

		const auto hunks = BuildRealPageStructuredPatch(baseCode, finalCode);
		nlohmann::json r;
		r["ok"] = true;
		r["page_name"] = LocalToUtf8Text(item.name);
		r["type_key"] = item.typeKey;
		r["type_name"] = LocalToUtf8Text(item.typeName);
		r["item_data"] = item.itemData;
		r["trace"] = LocalToUtf8Text(trace);
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
		return Utf8ToLocalText(r.dump());
	}

	if (toolName == "diff_program_item_code") {
		nlohmann::json args;
		try {
			args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return Utf8ToLocalText(r.dump());
		}

		const std::string pageName = GetJsonStringArgumentLocal(args, "page_name");
		const std::string kind = GetJsonStringArgumentLocal(args, "kind");
		const std::string newCodeInput = GetJsonStringArgumentLocal(args, "new_code");
		const std::string oldText = GetJsonStringArgumentLocal(args, "old_text");
		const std::string newText = GetJsonStringArgumentLocal(args, "new_text");
		const bool refreshCache = GetJsonBoolArgument(args, "refresh_cache", false);
		const bool failOnUnmatched = GetJsonBoolArgument(args, "fail_on_unmatched", true);
		if (TrimAsciiCopy(pageName).empty()) {
			return R"({"ok":false,"error":"page_name is required"})";
		}

		ProgramTreeItemInfo item;
		std::string error;
		if (!TryGetProgramItemByNameForAI(pageName, kind, item, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "program item lookup failed" : error;
			return Utf8ToLocalText(r.dump());
		}

		std::string baseCode;
		PageCodeCacheEntry cacheEntry;
		bool fromCache = false;
		std::string trace;
		if (!TryLoadRealPageCodeForReadForAI(item, refreshCache, baseCode, cacheEntry, fromCache, trace, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "read real page code failed" : error;
			r["trace"] = LocalToUtf8Text(trace);
			return Utf8ToLocalText(r.dump());
		}

		std::string candidateCode;
		nlohmann::json editResults = nlohmann::json::array();
		if (!newCodeInput.empty()) {
			candidateCode = NormalizeRealCodeLineBreaksToCrLf(newCodeInput);
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
				return Utf8ToLocalText(r.dump());
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
				return Utf8ToLocalText(r.dump());
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
				return Utf8ToLocalText(r.dump());
			}
			for (const auto& result : applyResults) {
				editResults.push_back(BuildRealPageTextEditResultJsonForAI(result));
			}
		}
		else {
			return R"({"ok":false,"error":"new_code or edit parameters are required"})";
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
		r["trace"] = LocalToUtf8Text(trace);
		r["from_cache"] = fromCache;
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
		return Utf8ToLocalText(r.dump());
	}

	if (toolName == "restore_program_item_code_snapshot") {
		nlohmann::json args;
		try {
			args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return Utf8ToLocalText(r.dump());
		}

		const std::string pageName = GetJsonStringArgumentLocal(args, "page_name");
		const std::string kind = GetJsonStringArgumentLocal(args, "kind");
		const std::string snapshotId = GetJsonStringArgumentLocal(args, "snapshot_id");
		const bool restoreLatest = GetJsonBoolArgument(args, "restore_latest", snapshotId.empty());
		if (TrimAsciiCopy(pageName).empty()) {
			return R"({"ok":false,"error":"page_name is required"})";
		}
		if (!restoreLatest && snapshotId.empty()) {
			return R"({"ok":false,"error":"snapshot_id is required when restore_latest is false"})";
		}

		ProgramTreeItemInfo item;
		std::string error;
		if (!TryGetProgramItemByNameForAI(pageName, kind, item, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "program item lookup failed" : error;
			return Utf8ToLocalText(r.dump());
		}

		PageCodeSnapshotEntry snapshot;
		const bool gotSnapshot = restoreLatest
			? PageCodeCacheManager::Instance().GetLatestSnapshot(item.name, item.typeKey, snapshot)
			: PageCodeCacheManager::Instance().GetSnapshot(item.name, item.typeKey, snapshotId, snapshot);
		if (!gotSnapshot) {
			return R"({"ok":false,"error":"snapshot not found"})";
		}

		std::string baseCode;
		PageCodeCacheEntry cacheEntry;
		std::string trace;
		std::string currentCode;
		bool cacheRefreshed = false;
		if (!TryResolveRealPageWriteBaseForAI(item, std::string(), true, baseCode, cacheEntry, trace, currentCode, cacheRefreshed, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error;
			r["trace"] = LocalToUtf8Text(trace);
			r["cache_refreshed"] = cacheRefreshed;
			if (!currentCode.empty()) {
				r["code"] = LocalToUtf8Text(currentCode);
				r["code_hash"] = BuildStableTextHashForRealCode(currentCode);
			}
			return Utf8ToLocalText(r.dump());
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
			return Utf8ToLocalText(r.dump());
		}

		PageCodeSnapshotEntry createdSnapshot;
		std::string finalCode;
		bool rollbackAttempted = false;
		bool rollbackSucceeded = false;
		if (!TryWriteRealPageCodeForAI(
				item,
				baseCode,
				snapshot.code,
				std::string("restore_program_item_code_snapshot:") + snapshot.snapshotId,
				createdSnapshot,
				finalCode,
				trace,
				rollbackAttempted,
				rollbackSucceeded,
				error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error;
			r["trace"] = LocalToUtf8Text(trace);
			r["rollback_attempted"] = rollbackAttempted;
			r["rollback_succeeded"] = rollbackSucceeded;
			return Utf8ToLocalText(r.dump());
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
		return Utf8ToLocalText(r.dump());
	}

	if (toolName == "search_program_item_real_code") {
		nlohmann::json args;
		try {
			args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return Utf8ToLocalText(r.dump());
		}

		const std::string pageName = GetJsonStringArgumentLocal(args, "page_name");
		const std::string kind = GetJsonStringArgumentLocal(args, "kind");
		const std::string keyword = GetJsonStringArgumentLocal(args, "keyword");
		const bool caseSensitive = GetJsonBoolArgument(args, "case_sensitive", false);
		const bool useRegex = GetJsonBoolArgument(args, "use_regex", false);
		const int contextLines = (std::clamp)(GetJsonIntArgument(args, "context_lines", 0), 0, 20);
		const size_t limit = GetJsonSizeArgument(args, "limit", 50);
		const bool refreshCache = GetJsonBoolArgument(args, "refresh_cache", false);
		if (TrimAsciiCopy(pageName).empty()) {
			return R"({"ok":false,"error":"page_name is required"})";
		}

		ProgramTreeItemInfo item;
		std::string error;
		if (!TryGetProgramItemByNameForAI(pageName, kind, item, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "program item lookup failed" : error;
			return Utf8ToLocalText(r.dump());
		}

		std::string code;
		PageCodeCacheEntry cacheEntry;
		bool fromCache = false;
		std::string trace;
		if (!TryLoadRealPageCodeForReadForAI(item, refreshCache, code, cacheEntry, fromCache, trace, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "read real page code failed" : error;
			r["trace"] = LocalToUtf8Text(trace);
			return Utf8ToLocalText(r.dump());
		}

		std::vector<RealPageSearchMatch> matches = SearchRealPageCode(code, keyword, caseSensitive, useRegex, contextLines, limit, &error);
		if (!error.empty()) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error;
			return Utf8ToLocalText(r.dump());
		}

		nlohmann::json rows = nlohmann::json::array();
		for (const auto& match : matches) {
			rows.push_back(BuildRealPageSearchMatchJsonForAI(match));
		}

		nlohmann::json r;
		r["ok"] = true;
		r["page_name"] = LocalToUtf8Text(item.name);
		r["type_key"] = item.typeKey;
		r["type_name"] = LocalToUtf8Text(item.typeName);
		r["item_data"] = item.itemData;
		r["trace"] = LocalToUtf8Text(trace);
		r["from_cache"] = fromCache;
		r["code_hash"] = cacheEntry.codeHash;
		r["match_count"] = matches.size();
		r["matches"] = std::move(rows);
		outOk = true;
		return Utf8ToLocalText(r.dump());
	}

	if (toolName == "list_program_item_symbols") {
		nlohmann::json args;
		try {
			args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return Utf8ToLocalText(r.dump());
		}

		const std::string pageName = GetJsonStringArgumentLocal(args, "page_name");
		const std::string kind = GetJsonStringArgumentLocal(args, "kind");
		const bool refreshCache = GetJsonBoolArgument(args, "refresh_cache", false);
		if (TrimAsciiCopy(pageName).empty()) {
			return R"({"ok":false,"error":"page_name is required"})";
		}

		ProgramTreeItemInfo item;
		std::string error;
		if (!TryGetProgramItemByNameForAI(pageName, kind, item, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "program item lookup failed" : error;
			return Utf8ToLocalText(r.dump());
		}

		std::string code;
		PageCodeCacheEntry cacheEntry;
		bool fromCache = false;
		std::string trace;
		if (!TryLoadRealPageCodeForReadForAI(item, refreshCache, code, cacheEntry, fromCache, trace, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "read real page code failed" : error;
			r["trace"] = LocalToUtf8Text(trace);
			return Utf8ToLocalText(r.dump());
		}

		const auto symbols = ParseRealPageSymbols(code);
		nlohmann::json rows = nlohmann::json::array();
		for (const auto& symbol : symbols) {
			rows.push_back(BuildRealPageSymbolJsonForAI(symbol));
		}

		nlohmann::json r;
		r["ok"] = true;
		r["page_name"] = LocalToUtf8Text(item.name);
		r["type_key"] = item.typeKey;
		r["type_name"] = LocalToUtf8Text(item.typeName);
		r["item_data"] = item.itemData;
		r["trace"] = LocalToUtf8Text(trace);
		r["from_cache"] = fromCache;
		r["code_hash"] = cacheEntry.codeHash;
		r["symbol_count"] = symbols.size();
		r["symbols"] = std::move(rows);
		outOk = true;
		return Utf8ToLocalText(r.dump());
	}

	if (toolName == "get_symbol_real_code") {
		nlohmann::json args;
		try {
			args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return Utf8ToLocalText(r.dump());
		}

		const std::string pageName = GetJsonStringArgumentLocal(args, "page_name");
		const std::string kind = GetJsonStringArgumentLocal(args, "kind");
		const std::string symbolName = GetJsonStringArgumentLocal(args, "symbol_name");
		const std::string symbolKind = GetJsonStringArgumentLocal(args, "symbol_kind");
		const std::string parentSymbolName = GetJsonStringArgumentLocal(args, "parent_symbol_name");
		const int occurrence = GetJsonIntArgument(args, "occurrence", 1);
		const bool refreshCache = GetJsonBoolArgument(args, "refresh_cache", false);
		if (TrimAsciiCopy(pageName).empty()) {
			return R"({"ok":false,"error":"page_name is required"})";
		}

		ProgramTreeItemInfo item;
		std::string error;
		if (!TryGetProgramItemByNameForAI(pageName, kind, item, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "program item lookup failed" : error;
			return Utf8ToLocalText(r.dump());
		}

		std::string code;
		PageCodeCacheEntry cacheEntry;
		bool fromCache = false;
		std::string trace;
		if (!TryLoadRealPageCodeForReadForAI(item, refreshCache, code, cacheEntry, fromCache, trace, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "read real page code failed" : error;
			r["trace"] = LocalToUtf8Text(trace);
			return Utf8ToLocalText(r.dump());
		}

		const auto symbols = ParseRealPageSymbols(code);
		RealPageSymbolInfo symbol;
		if (!FindRealPageSymbol(symbols, symbolName, symbolKind, parentSymbolName, occurrence, symbol, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error;
			return Utf8ToLocalText(r.dump());
		}

		nlohmann::json r;
		r["ok"] = true;
		r["page_name"] = LocalToUtf8Text(item.name);
		r["type_key"] = item.typeKey;
		r["type_name"] = LocalToUtf8Text(item.typeName);
		r["item_data"] = item.itemData;
		r["trace"] = LocalToUtf8Text(trace);
		r["from_cache"] = fromCache;
		r["code_hash"] = cacheEntry.codeHash;
		r["symbol"] = BuildRealPageSymbolJsonForAI(symbol);
		r["code"] = LocalToUtf8Text(symbol.code);
		outOk = true;
		return Utf8ToLocalText(r.dump());
	}

	if (toolName == "edit_symbol_real_code") {
		nlohmann::json args;
		try {
			args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return Utf8ToLocalText(r.dump());
		}

		const std::string pageName = GetJsonStringArgumentLocal(args, "page_name");
		const std::string kind = GetJsonStringArgumentLocal(args, "kind");
		const std::string symbolName = GetJsonStringArgumentLocal(args, "symbol_name");
		const std::string symbolKind = GetJsonStringArgumentLocal(args, "symbol_kind");
		const std::string parentSymbolName = GetJsonStringArgumentLocal(args, "parent_symbol_name");
		const int occurrence = GetJsonIntArgument(args, "occurrence", 1);
		const std::string newCode = GetJsonStringArgumentLocal(args, "new_code");
		if (TrimAsciiCopy(pageName).empty()) {
			return R"({"ok":false,"error":"page_name is required"})";
		}
		if (newCode.empty()) {
			return R"({"ok":false,"error":"new_code is required"})";
		}

		ProgramTreeItemInfo item;
		std::string error;
		if (!TryGetProgramItemByNameForAI(pageName, kind, item, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "program item lookup failed" : error;
			return Utf8ToLocalText(r.dump());
		}

		std::string baseCode;
		PageCodeCacheEntry cacheEntry;
		std::string trace;
		std::string currentCode;
		bool cacheRefreshed = false;
		if (!TryResolveRealPageWriteBaseForAI(item, std::string(), true, baseCode, cacheEntry, trace, currentCode, cacheRefreshed, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error;
			r["trace"] = LocalToUtf8Text(trace);
			r["cache_refreshed"] = cacheRefreshed;
			if (!currentCode.empty()) {
				r["code"] = LocalToUtf8Text(currentCode);
				r["code_hash"] = BuildStableTextHashForRealCode(currentCode);
			}
			return Utf8ToLocalText(r.dump());
		}

		const auto symbols = ParseRealPageSymbols(baseCode);
		RealPageSymbolInfo symbol;
		if (!FindRealPageSymbol(symbols, symbolName, symbolKind, parentSymbolName, occurrence, symbol, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error;
			return Utf8ToLocalText(r.dump());
		}

		std::string candidateCode;
		if (!ReplaceRealPageSymbolCode(baseCode, symbol, newCode, candidateCode, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error;
			return Utf8ToLocalText(r.dump());
		}

		if (BuildStableTextHashForRealCode(candidateCode) == BuildStableTextHashForRealCode(baseCode)) {
			nlohmann::json r;
			r["ok"] = true;
			r["no_changes"] = true;
			r["symbol"] = BuildRealPageSymbolJsonForAI(symbol);
			r["code_hash"] = BuildStableTextHashForRealCode(baseCode);
			r["code"] = LocalToUtf8Text(baseCode);
			outOk = true;
			return Utf8ToLocalText(r.dump());
		}

		PageCodeSnapshotEntry snapshot;
		std::string finalCode;
		bool rollbackAttempted = false;
		bool rollbackSucceeded = false;
		if (!TryWriteRealPageCodeForAI(
				item,
				baseCode,
				candidateCode,
				std::string("edit_symbol_real_code:") + symbol.kind + ":" + symbol.name,
				snapshot,
				finalCode,
				trace,
				rollbackAttempted,
				rollbackSucceeded,
				error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error;
			r["trace"] = LocalToUtf8Text(trace);
			r["rollback_attempted"] = rollbackAttempted;
			r["rollback_succeeded"] = rollbackSucceeded;
			return Utf8ToLocalText(r.dump());
		}

		const auto hunks = BuildRealPageStructuredPatch(baseCode, finalCode);
		nlohmann::json r;
		r["ok"] = true;
		r["page_name"] = LocalToUtf8Text(item.name);
		r["type_key"] = item.typeKey;
		r["type_name"] = LocalToUtf8Text(item.typeName);
		r["item_data"] = item.itemData;
		r["symbol"] = BuildRealPageSymbolJsonForAI(symbol);
		r["trace"] = LocalToUtf8Text(trace);
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
		return Utf8ToLocalText(r.dump());
	}

	if (toolName == "insert_program_item_code_block") {
		nlohmann::json args;
		try {
			args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return Utf8ToLocalText(r.dump());
		}

		const std::string pageName = GetJsonStringArgumentLocal(args, "page_name");
		const std::string kind = GetJsonStringArgumentLocal(args, "kind");
		const std::string codeBlock = GetJsonStringArgumentLocal(args, "code_block");
		if (TrimAsciiCopy(pageName).empty()) {
			return R"({"ok":false,"error":"page_name is required"})";
		}
		if (codeBlock.empty()) {
			return R"({"ok":false,"error":"code_block is required"})";
		}

		ProgramTreeItemInfo item;
		std::string error;
		if (!TryGetProgramItemByNameForAI(pageName, kind, item, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "program item lookup failed" : error;
			return Utf8ToLocalText(r.dump());
		}

		std::string baseCode;
		PageCodeCacheEntry cacheEntry;
		std::string trace;
		std::string currentCode;
		bool cacheRefreshed = false;
		if (!TryResolveRealPageWriteBaseForAI(item, std::string(), true, baseCode, cacheEntry, trace, currentCode, cacheRefreshed, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error;
			r["trace"] = LocalToUtf8Text(trace);
			r["cache_refreshed"] = cacheRefreshed;
			if (!currentCode.empty()) {
				r["code"] = LocalToUtf8Text(currentCode);
				r["code_hash"] = BuildStableTextHashForRealCode(currentCode);
			}
			return Utf8ToLocalText(r.dump());
		}

		const auto symbols = ParseRealPageSymbols(baseCode);
		RealPageInsertSpec spec;
		spec.mode = GetJsonStringArgumentLocal(args, "mode");
		spec.symbolName = GetJsonStringArgumentLocal(args, "symbol_name");
		spec.symbolKind = GetJsonStringArgumentLocal(args, "symbol_kind");
		spec.parentSymbolName = GetJsonStringArgumentLocal(args, "parent_symbol_name");
		spec.anchorText = GetJsonStringArgumentLocal(args, "anchor_text");
		spec.occurrence = GetJsonIntArgument(args, "occurrence", 1);

		std::string candidateCode;
		if (!InsertRealPageCodeBlock(baseCode, symbols, spec, codeBlock, candidateCode, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error;
			return Utf8ToLocalText(r.dump());
		}

		if (BuildStableTextHashForRealCode(candidateCode) == BuildStableTextHashForRealCode(baseCode)) {
			nlohmann::json r;
			r["ok"] = true;
			r["no_changes"] = true;
			r["code_hash"] = BuildStableTextHashForRealCode(baseCode);
			r["code"] = LocalToUtf8Text(baseCode);
			outOk = true;
			return Utf8ToLocalText(r.dump());
		}

		PageCodeSnapshotEntry snapshot;
		std::string finalCode;
		bool rollbackAttempted = false;
		bool rollbackSucceeded = false;
		if (!TryWriteRealPageCodeForAI(
				item,
				baseCode,
				candidateCode,
				std::string("insert_program_item_code_block:") + spec.mode,
				snapshot,
				finalCode,
				trace,
				rollbackAttempted,
				rollbackSucceeded,
				error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error;
			r["trace"] = LocalToUtf8Text(trace);
			r["rollback_attempted"] = rollbackAttempted;
			r["rollback_succeeded"] = rollbackSucceeded;
			return Utf8ToLocalText(r.dump());
		}

		const auto hunks = BuildRealPageStructuredPatch(baseCode, finalCode);
		nlohmann::json r;
		r["ok"] = true;
		r["page_name"] = LocalToUtf8Text(item.name);
		r["type_key"] = item.typeKey;
		r["type_name"] = LocalToUtf8Text(item.typeName);
		r["item_data"] = item.itemData;
		r["trace"] = LocalToUtf8Text(trace);
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
		return Utf8ToLocalText(r.dump());
	}

	if (toolName == "list_imported_modules") {
		return BuildListImportedModulesJsonOnMainThread(outOk);
	}

	if (toolName == "list_support_libraries") {
		return BuildListSupportLibrariesJsonOnMainThread(outOk);
	}

	if (toolName == "get_support_library_info") {
		return BuildGetSupportLibraryInfoJsonOnMainThread(argumentsJson, outOk);
	}

	if (toolName == "search_public_code") {
		return BuildSearchPublicCodeJsonOnMainThread(argumentsJson, outOk);
	}

	if (toolName == "search_support_library_public_code") {
		return BuildSearchSupportLibraryPublicCodeJsonOnMainThread(argumentsJson, outOk);
	}

	if (toolName == "read_support_library_public_code") {
		return BuildReadSupportLibraryPublicCodeJsonOnMainThread(argumentsJson, outOk);
	}

	if (toolName == "get_module_public_info") {
		return BuildGetModulePublicInfoJsonOnMainThread(argumentsJson, outOk);
	}

	if (toolName == "search_module_public_code") {
		return BuildSearchModulePublicCodeJsonOnMainThread(argumentsJson, outOk);
	}

	if (toolName == "read_module_public_code") {
		return BuildReadModulePublicCodeJsonOnMainThread(argumentsJson, outOk);
	}

	if (toolName == "list_program_items") {
		nlohmann::json args;
		try {
			args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return Utf8ToLocalText(r.dump());
		}

		const std::string kind = args.contains("kind") && args["kind"].is_string()
			? ToLowerAsciiCopyLocal(TrimAsciiCopy(Utf8ToLocalText(args["kind"].get<std::string>())))
			: std::string();
		const std::string nameContains = args.contains("name_contains") && args["name_contains"].is_string()
			? Utf8ToLocalText(args["name_contains"].get<std::string>())
			: std::string();
		const std::string exactName = args.contains("exact_name") && args["exact_name"].is_string()
			? Utf8ToLocalText(args["exact_name"].get<std::string>())
			: std::string();
		const bool includeCode = args.contains("include_code") && args["include_code"].is_boolean()
			? args["include_code"].get<bool>()
			: false;
		const int limit = args.contains("limit") && args["limit"].is_number_integer()
			? (std::clamp)(args["limit"].get<int>(), 1, 200)
			: 50;

		std::vector<ProgramTreeItemInfo> items;
		std::string error;
		if (!TryListProgramTreeItemsForAI(items, &error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "list program items failed" : error;
			return Utf8ToLocalText(r.dump());
		}

		nlohmann::json rows = nlohmann::json::array();
		for (const auto& item : items) {
			if (!MatchProgramItemKind(item, kind)) {
				continue;
			}
			if (!exactName.empty() && item.name != exactName) {
				continue;
			}
			if (!nameContains.empty() && item.name.find(nameContains) == std::string::npos) {
				continue;
			}
			if (static_cast<int>(rows.size()) >= limit) {
				break;
			}

			nlohmann::json row;
			row["name"] = LocalToUtf8Text(item.name);
			row["type_key"] = item.typeKey;
			row["type_name"] = LocalToUtf8Text(item.typeName);
			row["item_data"] = item.itemData;
			row["depth"] = item.depth;
			row["image"] = item.image;
			row["selected_image"] = item.selectedImage;

			if (includeCode) {
				std::string code;
				e571::RawSearchContextPageDumpDebugResult dumpResult;
				if (e571::DebugDumpCodePageByProgramTreeItemData(
						item.itemData,
						GetCurrentProcessImageBaseForAI(),
						&code,
						&dumpResult)) {
					row["code"] = LocalToUtf8Text(code);
					row["code_trace"] = dumpResult.trace;
					row["code_kind"] = "pseudo_reference";
					row["warning"] = LocalToUtf8Text("该代码来自程序树/逆向抓取，与IDE正常编辑页结构可能不同，仅可作为伪代码参考。");
				}
				else {
					row["code_error"] = dumpResult.trace.empty() ? "get page code failed" : dumpResult.trace;
				}
			}

			rows.push_back(std::move(row));
		}

		nlohmann::json r;
		r["ok"] = true;
		r["count"] = rows.size();
		r["code_kind"] = "pseudo_reference";
		r["warning"] = LocalToUtf8Text("程序树按名称抓取到的代码与IDE正常编辑页结构可能不同，仅可作为伪代码参考。");
		r["items"] = std::move(rows);
		outOk = true;
		return Utf8ToLocalText(r.dump());
	}

	if (toolName == "switch_to_program_item_page") {
		nlohmann::json args;
		try {
			args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return Utf8ToLocalText(r.dump());
		}

		const std::string name = args.contains("name") && args["name"].is_string()
			? Utf8ToLocalText(args["name"].get<std::string>())
			: std::string();
		const std::string kind = args.contains("kind") && args["kind"].is_string()
			? Utf8ToLocalText(args["kind"].get<std::string>())
			: std::string();
		if (TrimAsciiCopy(name).empty()) {
			return R"({"ok":false,"error":"name is required"})";
		}

		std::vector<ProgramTreeItemInfo> items;
		std::string error;
		if (!TryListProgramTreeItemsForAI(items, &error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "list program items failed" : error;
			return Utf8ToLocalText(r.dump());
		}

		std::vector<ProgramTreeItemInfo> matched;
		for (const auto& item : items) {
			if (item.name == name && MatchProgramItemKind(item, kind)) {
				matched.push_back(item);
			}
		}
		if (matched.empty()) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = "program item not found";
			return Utf8ToLocalText(r.dump());
		}
		if (matched.size() > 1) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = "program item name is ambiguous";
			return Utf8ToLocalText(r.dump());
		}

		std::string switchTrace;
		if (!e571::OpenProgramTreeItemPageByData(
				matched.front().itemData,
				&switchTrace)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = "open program item page failed";
			r["trace"] = LocalToUtf8Text(switchTrace);
			return Utf8ToLocalText(r.dump());
		}

		nlohmann::json r;
		r["ok"] = true;
		r["requested_name"] = LocalToUtf8Text(name);
		r["type_key"] = matched.front().typeKey;
		r["type_name"] = LocalToUtf8Text(matched.front().typeName);
		r["item_data"] = matched.front().itemData;
		r["trace"] = LocalToUtf8Text(switchTrace);
		outOk = true;
		return Utf8ToLocalText(r.dump());
	}

	if (toolName == "search_project_keyword") {
		return BuildProgramSearchResultJsonOnMainThread(argumentsJson, outOk);
	}

	if (toolName == "read_project_search_result_code") {
		return BuildReadProjectSearchResultCodeJsonOnMainThread(argumentsJson, outOk);
	}

	if (toolName == "refresh_project_source_cache") {
		return BuildRefreshProjectSourceCacheJsonOnMainThread(argumentsJson, outOk);
	}

	if (toolName == "search_project_source_cache") {
		return BuildSearchProjectSourceCacheJsonOnMainThread(argumentsJson, outOk);
	}

	if (toolName == "read_project_source_cache_code") {
		return BuildReadProjectSourceCacheCodeJsonOnMainThread(argumentsJson, outOk);
	}

	if (toolName == "jump_to_search_result") {
		nlohmann::json args;
		try {
			args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return Utf8ToLocalText(r.dump());
		}

		const std::string jumpToken = args.contains("jump_token") && args["jump_token"].is_string()
			? args["jump_token"].get<std::string>()
			: std::string();
		if (TrimAsciiCopy(jumpToken).empty()) {
			return R"({"ok":false,"error":"jump_token is required"})";
		}

		project_source_cache::HitToken projectToken;
		std::string projectTokenError;
		if (project_source_cache::ProjectSourceCacheManager::Instance().ParseHitToken(
				jumpToken,
				projectToken,
				&projectTokenError)) {
			project_source_cache::Snapshot snapshot;
			bool refreshed = false;
			std::string resolveError;
			std::string resolveTrace;
			if (!project_source_cache::ProjectSourceCacheManager::Instance().ResolveSnapshotForHit(
					projectToken,
					false,
					snapshot,
					&refreshed,
					&resolveError,
					&resolveTrace)) {
				nlohmann::json r;
				r["ok"] = false;
				r["error"] = resolveError.empty() ? "resolve project cache snapshot failed" : resolveError;
				r["trace"] = LocalToUtf8Text(resolveTrace);
				return Utf8ToLocalText(r.dump());
			}

			const auto* page = FindProjectCachePageForAI(snapshot, projectToken);
			if (page == nullptr) {
				nlohmann::json r;
				r["ok"] = false;
				r["error"] = "page not found in project cache snapshot";
				r["trace"] = LocalToUtf8Text(resolveTrace);
				return Utf8ToLocalText(r.dump());
			}

			ProgramTreeItemInfo item;
			std::string itemError;
			if (!TryGetProgramItemByNameForAI(page->name, page->typeKey, item, itemError) &&
				!TryGetProgramItemByNameForAI(page->name, std::string(), item, itemError)) {
				nlohmann::json r;
				r["ok"] = false;
				r["error"] = itemError.empty() ? "program item lookup failed" : itemError;
				r["trace"] = LocalToUtf8Text(resolveTrace);
				r["page_name"] = LocalToUtf8Text(page->name);
				r["type_key"] = page->typeKey;
				r["type_name"] = LocalToUtf8Text(page->typeName);
				return Utf8ToLocalText(r.dump());
			}

			std::string openTrace;
			if (!e571::OpenProgramTreeItemPageByData(item.itemData, &openTrace)) {
				nlohmann::json r;
				r["ok"] = false;
				r["error"] = "open program item page failed";
				r["trace"] = LocalToUtf8Text(resolveTrace + "|" + openTrace);
				return Utf8ToLocalText(r.dump());
			}

			std::string currentPageName;
			std::string currentPageType;
			std::string currentPageTrace;
			const bool currentPageOk = IDEFacade::Instance().GetCurrentPageName(
				currentPageName,
				&currentPageType,
				&currentPageTrace);

			nlohmann::json r;
			r["ok"] = true;
			r["source_kind"] = BuildProjectCacheSourceKindForAI(snapshot);
			r["trace"] = LocalToUtf8Text(resolveTrace + "|" + openTrace);
			r["line_number_hint"] = projectToken.lineNumber;
			r["current_page_ok"] = currentPageOk;
			r["current_page_name"] = LocalToUtf8Text(currentPageName);
			r["current_page_type"] = LocalToUtf8Text(currentPageType);
			r["current_page_trace"] = currentPageTrace;
			r["warning"] = LocalToUtf8Text("当前命中来自 e2txt 解析缓存，jump_to_search_result 这里只负责尽量打开对应页面，不再尝试把 IDE 光标精确移动到解析行号。若需要精确行内容，请继续调用 read_project_source_cache_code。");
			outOk = true;
			return Utf8ToLocalText(r.dump());
		}

		e571::DirectGlobalSearchDebugHit hit{};
		if (!ParseSearchJumpToken(jumpToken, hit)) {
			return R"({"ok":false,"error":"invalid jump_token"})";
		}

		std::string jumpTrace;
		if (!e571::DebugJumpToSearchHit(hit, GetCurrentProcessImageBaseForAI(), &jumpTrace)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = "jump to search result failed";
			r["trace"] = jumpTrace;
			return Utf8ToLocalText(r.dump());
		}

		std::string currentPageName;
		std::string currentPageType;
		std::string currentPageTrace;
		const bool currentPageOk = IDEFacade::Instance().GetCurrentPageName(
			currentPageName,
			&currentPageType,
			&currentPageTrace);

		nlohmann::json r;
		r["ok"] = true;
		r["trace"] = jumpTrace;
		r["current_page_ok"] = currentPageOk;
		r["current_page_name"] = LocalToUtf8Text(currentPageName);
		r["current_page_type"] = LocalToUtf8Text(currentPageType);
		r["current_page_trace"] = currentPageTrace;
		outOk = true;
		return Utf8ToLocalText(r.dump());
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
			return Utf8ToLocalText(r.dump());
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
			return Utf8ToLocalText(r.dump());
		}

		std::string normalizedPath;
		std::string diagnostics;
		if (!IDEFacade::Instance().CompileWithOutputPath(
				kind,
				outputPath,
				staticCompile,
				&normalizedPath,
				&diagnostics)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = diagnostics.empty() ? "compile_with_output_path_failed" : diagnostics;
			r["target"] = target;
			r["static_compile"] = staticCompile;
			return Utf8ToLocalText(r.dump());
		}

		nlohmann::json r;
		r["ok"] = true;
		r["target"] = target;
		r["static_compile"] = staticCompile;
		r["output_path"] = LocalToUtf8Text(normalizedPath);
		r["trace"] = diagnostics;
        r["warning"] = LocalToUtf8Text("This only suppresses the system save-file dialog by injecting the requested output path. Final compile success still needs to be confirmed from IDE output or artifacts.");
		outOk = true;
		return Utf8ToLocalText(r.dump());
	}

	nlohmann::json r;
	r["ok"] = false;
	r["error"] = "unknown tool: " + toolName;
	return Utf8ToLocalText(r.dump());
}

#endif
