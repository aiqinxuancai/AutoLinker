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
#include <mutex>
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
};

struct SupportLibraryDumpCacheEntry {
	std::string md5;
	nlohmann::json dumpJson = nlohmann::json::object();
	std::string error;
	bool ok = false;
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
};
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
		if (text.rfind("Class_", 0) == 0) {
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
		if (item.name.rfind("Class_", 0) == 0 && item.image >= 0) {
			classImage = item.image;
			break;
		}
	}
	for (auto& item : outItems) {
		item.typeKey = GetProgramTreeTypeKey(item.itemData, item.name, item.image, classImage);
		item.typeName = GetProgramTreeTypeName(item.typeKey);
	}
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

	std::vector<ProgramTreeItemInfo> items;
	if (!TryListProgramTreeItemsForAI(items, &outError)) {
		return false;
	}

	std::vector<ProgramTreeItemInfo> matched;
	for (const auto& item : items) {
		if (item.name == name && MatchProgramItemKind(item, kindFilter)) {
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
		outPageName = TrimAsciiCopy(displayText.substr(0, arrowPos));
		return !outPageName.empty();
	}

	const size_t parenPos = displayText.find(" (");
	if (parenPos != std::string::npos && parenPos > 0) {
		outPageName = TrimAsciiCopy(displayText.substr(0, parenPos));
		return !outPageName.empty();
	}

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

std::filesystem::path GetModulePublicInfoCachePathForAI(const std::string& md5)
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

bool TryLoadModulePublicInfoCachedForAI(
	const std::string& modulePath,
	e571::ModulePublicInfoDump& outDump,
	std::string& outError,
	bool* outCacheHit = nullptr)
{
	outDump = {};
	outError.clear();
	if (outCacheHit != nullptr) {
		*outCacheHit = false;
	}

	std::string md5;
	if (!TryGetFileMd5HexForAI(modulePath, md5) || md5.empty()) {
		return e571::LoadModulePublicInfoDump(
			modulePath,
			GetCurrentProcessImageBaseForAI(),
			&outDump,
			&outError);
	}

	const std::string cacheKey = BuildFileCacheKeyForAI(modulePath);
	{
		std::lock_guard<std::mutex> guard(g_modulePublicInfoCacheMutex);
		const auto it = g_modulePublicInfoCache.find(cacheKey);
		if (it != g_modulePublicInfoCache.end() && it->second.md5 == md5) {
			outDump = it->second.dump;
			outError = it->second.error;
			if (outCacheHit != nullptr) {
				*outCacheHit = true;
			}
			return it->second.ok;
		}
	}

	{
		nlohmann::json cacheJson;
		if (TryReadJsonCacheFileForAI(GetModulePublicInfoCachePathForAI(md5), cacheJson) &&
			cacheJson.is_object() &&
			cacheJson.value("schema", std::string()) == "module_public_info_v1" &&
			cacheJson.value("md5", std::string()) == md5 &&
			cacheJson.contains("dump")) {
			e571::ModulePublicInfoDump cachedDump;
			if (DeserializeModulePublicInfoDumpForAI(cacheJson["dump"], cachedDump)) {
				cachedDump.modulePath = modulePath;
				ModulePublicInfoCacheEntry entry;
				entry.md5 = md5;
				entry.dump = cachedDump;
				entry.ok = true;
				{
					std::lock_guard<std::mutex> guard(g_modulePublicInfoCacheMutex);
					g_modulePublicInfoCache[cacheKey] = entry;
				}
				outDump = std::move(cachedDump);
				if (outCacheHit != nullptr) {
					*outCacheHit = true;
				}
				return true;
			}
		}
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
	{
		std::lock_guard<std::mutex> guard(g_modulePublicInfoCacheMutex);
		g_modulePublicInfoCache[cacheKey] = std::move(entry);
	}

	if (ok) {
		nlohmann::json cacheJson = nlohmann::json::object();
		cacheJson["schema"] = "module_public_info_v1";
		cacheJson["md5"] = md5;
		cacheJson["dump"] = SerializeModulePublicInfoDumpForAI(loadedDump);
		TryWriteJsonCacheFileForAI(GetModulePublicInfoCachePathForAI(md5), cacheJson);
	}

	outDump = std::move(loadedDump);
	outError = loadError;
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

	nlohmann::json results = nlohmann::json::array();
	for (size_t i = 0; i < hits.size() && static_cast<int>(results.size()) < limit; ++i) {
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

		nlohmann::json row;
		row["index"] = static_cast<int>(i);
		row["page_name"] = LocalToUtf8Text(info.pageName);
		row["page_type_key"] = info.pageTypeKey;
		row["page_type_name"] = LocalToUtf8Text(info.pageTypeName);
		row["line_number"] = info.lineNumber;
		row["text"] = LocalToUtf8Text(info.text);
		row["jump_token"] = BuildSearchJumpToken(hits[i]);
		results.push_back(std::move(row));
	}

	nlohmann::json r;
	r["ok"] = true;
	r["keyword"] = LocalToUtf8Text(keyword);
	r["count"] = hits.size();
	r["dialog_handled"] = dialogHandled;
	r["code_kind"] = "pseudo_reference";
	r["warning"] = LocalToUtf8Text("搜索结果文本以及后续按页面名抓取到的代码，与IDE正常编辑页结构可能不同，仅可作为伪代码参考。");
	r["results"] = std::move(results);
	if (!listError.empty()) {
		r["page_type_lookup_error"] = listError;
	}
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

	if (toolName == "search_support_library_info") {
		return BuildSearchSupportLibraryInfoJsonOnMainThread(argumentsJson, outOk);
	}

	if (toolName == "get_module_public_info") {
		return BuildGetModulePublicInfoJsonOnMainThread(argumentsJson, outOk);
	}

	if (toolName == "search_module_public_info") {
		return BuildSearchModulePublicInfoJsonOnMainThread(argumentsJson, outOk);
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
