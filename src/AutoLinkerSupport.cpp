#include "AutoLinkerInternal.h"
#include <Windows.h>
#include <CommCtrl.h>
#include <TlHelp32.h>
#include <detver.h>
#include <fnshare.h>
#include <lang.h>
#include <lib2.h>
#include <algorithm>
#include <array>
#include <cctype>
#include <filesystem>
#include <format>
#include <fstream>
#include <mutex>
#include <string>
#include <unordered_set>
#include <vector>
#include "ECOMEx.h"
#include "Global.h"
#include "IDEFacade.h"
#include "PathHelper.h"
#include "WindowHelper.h"
#include "direct_global_search.hpp"
#if defined(_M_IX86)
#include "direct_global_search_debug.hpp"
#include "native_module_public_info.hpp"
#endif

std::mutex g_aiRoundtripLogMutex;
std::mutex g_addTabTestLogMutex;
std::mutex g_directGlobalSearchPageDumpLogMutex;
std::mutex g_programTreeListLogMutex;
std::mutex g_modulePublicInfoLogMutex;
std::mutex g_supportLibraryInfoLogMutex;
constexpr const char* kDirectGlobalSearchTestKeyword = "subWinHwnd";
constexpr const char* kTreeDirectPageDumpTestName = "Class_HWND";

std::filesystem::path GetAIRoundtripLogPath()
{
	std::filesystem::path dir = std::filesystem::path(GetBasePath()) / "AutoLinker";
	std::error_code ec;
	std::filesystem::create_directories(dir, ec);
	return dir / "ai_roundtrip_last.log";
}

std::filesystem::path GetAddTabTestLogPath()
{
	std::filesystem::path dir = std::filesystem::path(GetBasePath()) / "AutoLinker";
	std::error_code ec;
	std::filesystem::create_directories(dir, ec);
	return dir / "add_tab_test_last.log";
}

std::filesystem::path GetDirectGlobalSearchPageDumpLogPath()
{
	std::filesystem::path dir = std::filesystem::path(GetBasePath()) / "AutoLinker";
	std::error_code ec;
	std::filesystem::create_directories(dir, ec);
	return dir / "direct_global_search_page_last.txt";
}

std::filesystem::path GetProgramTreeListLogPath()
{
	std::filesystem::path dir = std::filesystem::path(GetBasePath()) / "AutoLinker";
	std::error_code ec;
	std::filesystem::create_directories(dir, ec);
	return dir / "program_tree_items_last.txt";
}

std::filesystem::path GetModulePublicInfoLogPath()
{
	std::filesystem::path dir = std::filesystem::path(GetBasePath()) / "AutoLinker";
	std::error_code ec;
	std::filesystem::create_directories(dir, ec);
	return dir / "module_public_info_last.txt";
}

std::filesystem::path GetSupportLibraryInfoLogPath()
{
	std::filesystem::path dir = std::filesystem::path(GetBasePath()) / "AutoLinker";
	std::error_code ec;
	std::filesystem::create_directories(dir, ec);
	return dir / "support_library_info_last.txt";
}

std::filesystem::path GetAutoRunModulePublicInfoTestMarkerPath()
{
	std::filesystem::path dir = std::filesystem::path(GetBasePath()) / "AutoLinker";
	std::error_code ec;
	std::filesystem::create_directories(dir, ec);
	return dir / "autorun_module_public_info_test.flag";
}

std::string EscapeOneLineForLog(std::string text)
{
	for (size_t i = 0; i < text.size(); ++i) {
		if (text[i] == '\r') {
			text.replace(i, 1, "\\r");
			++i;
			continue;
		}
		if (text[i] == '\n') {
			text.replace(i, 1, "\\n");
			++i;
			continue;
		}
	}
	return text;
}

bool IsValidUtf8ForLog(const std::string& text)
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

std::string BuildHexHead(const std::string& text, size_t maxBytes = 64)
{
	const size_t n = (std::min)(text.size(), maxBytes);
	std::string out;
	out.reserve(n * 3 + 16);
	constexpr char kHex[] = "0123456789ABCDEF";
	for (size_t i = 0; i < n; ++i) {
		const unsigned char ch = static_cast<unsigned char>(text[i]);
		out.push_back(kHex[(ch >> 4) & 0x0F]);
		out.push_back(kHex[ch & 0x0F]);
		if (i + 1 < n) {
			out.push_back(' ');
		}
	}
	if (text.size() > n) {
		out += " ...";
	}
	return out;
}

std::string ConvertUtf8ToLocalForEOutput(const std::string& text)
{
	if (text.empty()) {
		return std::string();
	}

	const int wideLen = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.c_str(), -1, nullptr, 0);
	if (wideLen <= 0) {
		return text;
	}

	std::wstring wide(static_cast<size_t>(wideLen), L'\0');
	if (MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.c_str(), -1, wide.data(), wideLen) <= 0) {
		return text;
	}

	const int localLen = WideCharToMultiByte(CP_ACP, 0, wide.c_str(), -1, nullptr, 0, nullptr, nullptr);
	if (localLen <= 0) {
		return text;
	}

	std::string local(static_cast<size_t>(localLen), '\0');
	if (WideCharToMultiByte(CP_ACP, 0, wide.c_str(), -1, local.data(), localLen, nullptr, nullptr) <= 0) {
		return text;
	}
	if (!local.empty() && local.back() == '\0') {
		local.pop_back();
	}
	return local;
}

std::string ConvertPossiblyUtf8ToLocalForEOutput(const std::string& text)
{
	if (text.empty()) {
		return std::string();
	}
	if (!IsValidUtf8ForLog(text)) {
		return text;
	}
	return ConvertUtf8ToLocalForEOutput(text);
}

void AppendAIRoundtripLogLineUnlocked(std::ofstream& out, const std::string& line)
{
	out << line << "\r\n";
}

void BeginAIRoundtripLogSession(uint64_t traceId, const std::string& scene, const std::string& taskName)
{
	const auto path = GetAIRoundtripLogPath();
	SYSTEMTIME st = {};
	GetLocalTime(&st);
	std::lock_guard<std::mutex> guard(g_aiRoundtripLogMutex);
	std::ofstream out(path, std::ios::trunc | std::ios::binary);
	if (!out.is_open()) {
		return;
	}

	AppendAIRoundtripLogLineUnlocked(
		out,
		"[AI-ROUNDTRIP] session-start trace=" + std::to_string(traceId) +
		" scene=" + scene +
		" task=\"" + EscapeOneLineForLog(taskName) + "\"" +
		" time=" + std::to_string(st.wYear) + "-" + std::to_string(st.wMonth) + "-" + std::to_string(st.wDay) +
		" " + std::to_string(st.wHour) + ":" + std::to_string(st.wMinute) + ":" + std::to_string(st.wSecond) +
		"." + std::to_string(st.wMilliseconds));
}

void AppendAIRoundtripLogLine(const std::string& line)
{
	const auto path = GetAIRoundtripLogPath();
	std::lock_guard<std::mutex> guard(g_aiRoundtripLogMutex);
	std::ofstream out(path, std::ios::app | std::ios::binary);
	if (!out.is_open()) {
		return;
	}
	AppendAIRoundtripLogLineUnlocked(out, line);
}

std::string PtrToHexText(UINT_PTR ptr)
{
	return std::format("0x{:X}", static_cast<unsigned long long>(ptr));
}

std::string DescribeWindowForAddTabLog(HWND hWnd)
{
	if (hWnd == nullptr) {
		return "hWnd=NULL";
	}

	const BOOL isValid = IsWindow(hWnd);
	char className[128] = {};
	char windowText[256] = {};
	HWND parent = nullptr;
	if (isValid) {
		GetClassNameA(hWnd, className, static_cast<int>(sizeof(className)));
		GetWindowTextA(hWnd, windowText, static_cast<int>(sizeof(windowText)));
		parent = GetParent(hWnd);
	}

	return std::format(
		"hWnd={} valid={} class=\"{}\" text=\"{}\" parent={}",
		PtrToHexText(reinterpret_cast<UINT_PTR>(hWnd)),
		isValid ? 1 : 0,
		EscapeOneLineForLog(className),
		EscapeOneLineForLog(windowText),
		PtrToHexText(reinterpret_cast<UINT_PTR>(parent)));
}

std::string GetWindowTextCopyA(HWND hWnd)
{
	char windowText[256] = {};
	if (hWnd != nullptr && IsWindow(hWnd)) {
		GetWindowTextA(hWnd, windowText, static_cast<int>(sizeof(windowText)));
	}
	return windowText;
}

std::string GetWindowClassCopyA(HWND hWnd)
{
	char className[128] = {};
	if (hWnd != nullptr && IsWindow(hWnd)) {
		GetClassNameA(hWnd, className, static_cast<int>(sizeof(className)));
	}
	return className;
}

std::string DescribeWindowChain(HWND hWnd, int maxDepth = 8)
{
	std::string out;
	HWND current = hWnd;
	for (int depth = 0; current != nullptr && depth < maxDepth; ++depth) {
		if (!out.empty()) {
			out += " <- ";
		}
		out += std::format(
			"[{} {} \"{}\"]",
			PtrToHexText(reinterpret_cast<UINT_PTR>(current)),
			EscapeOneLineForLog(GetWindowClassCopyA(current)),
			EscapeOneLineForLog(GetWindowTextCopyA(current)));
		current = GetParent(current);
	}
	return out;
}

BOOL CALLBACK EnumChildProcCollectTreeView(HWND hWnd, LPARAM lParam)
{
	auto* windows = reinterpret_cast<std::vector<HWND>*>(lParam);
	if (windows == nullptr) {
		return TRUE;
	}

	if (_stricmp(GetWindowClassCopyA(hWnd).c_str(), WC_TREEVIEWA) == 0 ||
		_stricmp(GetWindowClassCopyA(hWnd).c_str(), "SysTreeView32") == 0) {
		windows->push_back(hWnd);
	}
	return TRUE;
}

BOOL CALLBACK EnumChildProcCollectWindowsByClass(HWND hWnd, LPARAM lParam)
{
	auto* pair = reinterpret_cast<std::pair<const char*, std::vector<HWND>*>*>(lParam);
	if (pair == nullptr || pair->first == nullptr || pair->second == nullptr) {
		return TRUE;
	}

	const std::string className = GetWindowClassCopyA(hWnd);
	if (_stricmp(className.c_str(), pair->first) == 0) {
		pair->second->push_back(hWnd);
	}
	return TRUE;
}

std::vector<HWND> CollectChildWindowsByClass(HWND root, const char* className)
{
	std::vector<HWND> windows;
	if (root == nullptr || !IsWindow(root) || className == nullptr || className[0] == '\0') {
		return windows;
	}

	std::pair<const char*, std::vector<HWND>*> ctx{ className, &windows };
	EnumChildWindows(root, EnumChildProcCollectWindowsByClass, reinterpret_cast<LPARAM>(&ctx));
	return windows;
}

std::string ReadTabItemText(HWND tabHwnd, int index)
{
	if (tabHwnd == nullptr || !IsWindow(tabHwnd) || index < 0) {
		return std::string();
	}

	char textBuf[512] = {};
	TCITEMA item = {};
	item.mask = TCIF_TEXT;
	item.pszText = textBuf;
	item.cchTextMax = static_cast<int>(sizeof(textBuf));
	if (SendMessageA(tabHwnd, TCM_GETITEMA, static_cast<WPARAM>(index), reinterpret_cast<LPARAM>(&item)) == FALSE) {
		return std::string();
	}
	return textBuf;
}

void RunCurrentPageWindowProbeTest()
{
	OutputStringToELog("[CurrentPageProbe] 开始探测当前页窗口与页签");
	if (g_hwnd == nullptr || !IsWindow(g_hwnd)) {
		OutputStringToELog("[CurrentPageProbe] 中止：主窗口句柄无效");
		return;
	}

	const auto mdiClients = CollectChildWindowsByClass(g_hwnd, "MDIClient");
	OutputStringToELog(std::format("[CurrentPageProbe] mdiClientCount={}", mdiClients.size()));
	for (size_t i = 0; i < mdiClients.size(); ++i) {
		const HWND mdiHwnd = mdiClients[i];
		const BOOL activeMaximized = FALSE;
		HWND activeChild = reinterpret_cast<HWND>(SendMessageA(mdiHwnd, WM_MDIGETACTIVE, 0, 0));
		OutputStringToELog(std::format(
			"[CurrentPageProbe] mdi#{} hwnd={} text={} chain={} activeChild={} activeClass={} activeText={} activeChain={}",
			i,
			PtrToHexText(reinterpret_cast<UINT_PTR>(mdiHwnd)),
			EscapeOneLineForLog(GetWindowTextCopyA(mdiHwnd)),
			DescribeWindowChain(mdiHwnd),
			PtrToHexText(reinterpret_cast<UINT_PTR>(activeChild)),
			EscapeOneLineForLog(GetWindowClassCopyA(activeChild)),
			EscapeOneLineForLog(GetWindowTextCopyA(activeChild)),
			DescribeWindowChain(activeChild)));
		(void)activeMaximized;
	}

	const auto customTabs = CollectChildWindowsByClass(g_hwnd, "CCustomTabCtrl");
	OutputStringToELog(std::format("[CurrentPageProbe] customTabCount={}", customTabs.size()));
	for (size_t i = 0; i < customTabs.size(); ++i) {
		const HWND tabHwnd = customTabs[i];
		const int curSel = static_cast<int>(SendMessageA(tabHwnd, TCM_GETCURSEL, 0, 0));
		const int itemCount = static_cast<int>(SendMessageA(tabHwnd, TCM_GETITEMCOUNT, 0, 0));
		OutputStringToELog(std::format(
			"[CurrentPageProbe] tab#{} hwnd={} class={} text={} chain={} curSel={} itemCount={}",
			i,
			PtrToHexText(reinterpret_cast<UINT_PTR>(tabHwnd)),
			EscapeOneLineForLog(GetWindowClassCopyA(tabHwnd)),
			EscapeOneLineForLog(GetWindowTextCopyA(tabHwnd)),
			DescribeWindowChain(tabHwnd),
			curSel,
			itemCount));

		const int previewCount = (std::min)(itemCount, 12);
		for (int index = 0; index < previewCount; ++index) {
			OutputStringToELog(std::format(
				"[CurrentPageProbe] tab#{} item#{} selected={} text={}",
				i,
				index,
				(index == curSel) ? 1 : 0,
				EscapeOneLineForLog(ReadTabItemText(tabHwnd, index))));
		}
	}

	OutputStringToELog(std::format(
		"[CurrentPageProbe] mainWindowText={}",
		EscapeOneLineForLog(GetWindowTextCopyA(g_hwnd))));
}

void RunCurrentPageNameTest()
{
	OutputStringToELog("[CurrentPageNameTest] 开始获取当前页名称");

	std::string pageName;
	std::string typeText;
	std::string diagnostics;
	const bool ok = IDEFacade::Instance().GetCurrentPageName(pageName, &typeText, &diagnostics);
	OutputStringToELog(std::format(
		"[CurrentPageNameTest] ok={} name={} type={} trace={}",
		ok ? 1 : 0,
		EscapeOneLineForLog(pageName),
		EscapeOneLineForLog(typeText),
		EscapeOneLineForLog(diagnostics)));
}

std::vector<HWND> CollectTreeViewWindows(HWND root)
{
	std::vector<HWND> windows;
	if (root != nullptr && IsWindow(root)) {
		EnumChildWindows(root, EnumChildProcCollectTreeView, reinterpret_cast<LPARAM>(&windows));
	}
	return windows;
}

HTREEITEM GetTreeNextItem(HWND treeHwnd, HTREEITEM item, UINT code)
{
	return reinterpret_cast<HTREEITEM>(SendMessageA(treeHwnd, TVM_GETNEXTITEM, code, reinterpret_cast<LPARAM>(item)));
}

bool QueryTreeItemInfo(
	HWND treeHwnd,
	HTREEITEM item,
	std::string& outText,
	LPARAM& outParam,
	UINT& outState,
	int& outChildren,
	int& outImage,
	int& outSelectedImage)
{
	outText.clear();
	outParam = 0;
	outState = 0;
	outChildren = 0;
	outImage = -1;
	outSelectedImage = -1;
	if (treeHwnd == nullptr || item == nullptr) {
		return false;
	}

	char textBuf[512] = {};
	TVITEMA tvItem = {};
	tvItem.mask = TVIF_HANDLE | TVIF_TEXT | TVIF_PARAM | TVIF_STATE | TVIF_CHILDREN | TVIF_IMAGE | TVIF_SELECTEDIMAGE;
	tvItem.hItem = item;
	tvItem.pszText = textBuf;
	tvItem.cchTextMax = static_cast<int>(sizeof(textBuf));
	tvItem.stateMask = 0xFFFFFFFFu;
	if (SendMessageA(treeHwnd, TVM_GETITEMA, 0, reinterpret_cast<LPARAM>(&tvItem)) == FALSE) {
		return false;
	}

	outText = textBuf;
	outParam = tvItem.lParam;
	outState = tvItem.state;
	outChildren = tvItem.cChildren;
	outImage = tvItem.iImage;
	outSelectedImage = tvItem.iSelectedImage;
	return true;
}

void DumpTreeItemsRecursive(HWND treeHwnd, HTREEITEM firstItem, int depth, int maxDepth, int& remainingBudget)
{
	if (treeHwnd == nullptr || firstItem == nullptr || remainingBudget <= 0 || depth > maxDepth) {
		return;
	}

	for (HTREEITEM item = firstItem; item != nullptr && remainingBudget > 0; item = GetTreeNextItem(treeHwnd, item, TVGN_NEXT)) {
		std::string text;
		LPARAM itemData = 0;
		UINT state = 0;
		int childCount = 0;
		int image = -1;
		int selectedImage = -1;
		if (!QueryTreeItemInfo(treeHwnd, item, text, itemData, state, childCount, image, selectedImage)) {
			OutputStringToELog(std::format(
				"[TreeProbe] depth={} item={} query_failed",
				depth,
				PtrToHexText(reinterpret_cast<UINT_PTR>(item))));
			--remainingBudget;
			continue;
		}

		const unsigned int itemDataU = static_cast<unsigned int>(itemData);
		const unsigned int itemType = itemDataU >> 28;
		OutputStringToELog(std::format(
			"[TreeProbe] depth={} item={} text={} data=0x{:08X} typeNibble={} image={} selImage={} state=0x{:08X} children={}",
			depth,
			PtrToHexText(reinterpret_cast<UINT_PTR>(item)),
			EscapeOneLineForLog(text),
			itemDataU,
			itemType,
			image,
			selectedImage,
			state,
			childCount));
		--remainingBudget;

		if (depth < maxDepth) {
			HTREEITEM child = GetTreeNextItem(treeHwnd, item, TVGN_CHILD);
			if (child != nullptr) {
				DumpTreeItemsRecursive(treeHwnd, child, depth + 1, maxDepth, remainingBudget);
			}
		}
	}
}

void RunTreeViewProbeTest()
{
	OutputStringToELog("[TreeProbe] 开始枚举主窗口下的 SysTreeView32");
	if (g_hwnd == nullptr || !IsWindow(g_hwnd)) {
		OutputStringToELog("[TreeProbe] 中止：主窗口句柄无效");
		return;
	}

	const auto treeWindows = CollectTreeViewWindows(g_hwnd);
	OutputStringToELog(std::format("[TreeProbe] treeCount={}", treeWindows.size()));
	for (size_t index = 0; index < treeWindows.size(); ++index) {
		const HWND treeHwnd = treeWindows[index];
		const BOOL visible = IsWindowVisible(treeHwnd);
		const HTREEITEM rootItem = GetTreeNextItem(treeHwnd, nullptr, TVGN_ROOT);
		const HTREEITEM caretItem = GetTreeNextItem(treeHwnd, nullptr, TVGN_CARET);
		OutputStringToELog(std::format(
			"[TreeProbe] tree#{} hwnd={} visible={} enabled={} root={} caret={} chain={}",
			index,
			PtrToHexText(reinterpret_cast<UINT_PTR>(treeHwnd)),
			visible ? 1 : 0,
			IsWindowEnabled(treeHwnd) ? 1 : 0,
			PtrToHexText(reinterpret_cast<UINT_PTR>(rootItem)),
			PtrToHexText(reinterpret_cast<UINT_PTR>(caretItem)),
			DescribeWindowChain(treeHwnd)));

		int remainingBudget = 24;
		DumpTreeItemsRecursive(treeHwnd, rootItem, 0, 1, remainingBudget);
	}
}

HWND FindProgramDataTreeView()
{
	const auto treeWindows = CollectTreeViewWindows(g_hwnd);
	for (HWND treeHwnd : treeWindows) {
		const HTREEITEM rootItem = GetTreeNextItem(treeHwnd, nullptr, TVGN_ROOT);
		std::string text;
		LPARAM itemData = 0;
		UINT state = 0;
		int childCount = 0;
		int image = -1;
		int selectedImage = -1;
		if (!QueryTreeItemInfo(treeHwnd, rootItem, text, itemData, state, childCount, image, selectedImage)) {
			continue;
		}
		if (text == "程序数据") {
			return treeHwnd;
		}
	}
	return nullptr;
}

#if defined(_M_IX86)
std::uintptr_t GetCurrentProcessImageBase();
void WriteDirectGlobalSearchPageDump(const std::string& text);
void LogCurrentPageCodePreview(const std::string& pageCode, size_t maxLines);

bool TryListImportedModulePathsForTest(std::vector<std::string>& outPaths, std::string* outError)
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

void WriteModulePublicInfoLog(const std::vector<std::string>& lines)
{
	const auto path = GetModulePublicInfoLogPath();
	std::lock_guard<std::mutex> guard(g_modulePublicInfoLogMutex);
	std::ofstream out(path, std::ios::trunc | std::ios::binary);
	if (!out.is_open()) {
		OutputStringToELog("[ModulePublicInfoTest] 写日志文件失败");
		return;
	}
	for (const auto& line : lines) {
		out << line << "\r\n";
	}
}

void WriteSupportLibraryInfoLog(const std::vector<std::string>& lines)
{
	const auto path = GetSupportLibraryInfoLogPath();
	std::lock_guard<std::mutex> guard(g_supportLibraryInfoLogMutex);
	std::ofstream out(path, std::ios::trunc | std::ios::binary);
	if (!out.is_open()) {
		OutputStringToELog("[SupportLibraryInfoTest] 写日志文件失败");
		return;
	}
	for (const auto& line : lines) {
		out << line << "\r\n";
	}
}

struct SupportLibraryHeaderForTest {
	int index = -1;
	std::string rawText;
	std::string rawName;
	std::string name;
	std::string versionText;
	std::string fileName;
	std::string fileStem;
	std::string filePath;
	std::string resolveTrace;
};

struct LoadedSupportLibraryModuleForTest {
	std::string filePath;
	std::string fileName;
	std::string fileStem;
};

std::vector<std::string> SplitLinesForSupportLibraryTest(const std::string& text)
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

std::string NormalizeLineBreaksForSupportLibraryTest(std::string text)
{
	text.erase(std::remove(text.begin(), text.end(), '\r'), text.end());
	return text;
}

std::string FindPossibleSupportLibraryFileToken(const std::string& text)
{
	static const std::array<const char*, 3> exts = { ".fne", ".fnr", ".dll" };
	const std::string lower = ToLowerAsciiCopy(text);
	for (const char* ext : exts) {
		const std::string extLower = ToLowerAsciiCopy(ext);
		size_t pos = lower.find(extLower);
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

		size_t end = pos + extLower.size();
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

void ParseSupportLibraryHeaderText(
	int index,
	const std::string& rawText,
	SupportLibraryHeaderForTest& outInfo)
{
	outInfo = {};
	outInfo.index = index;
	outInfo.rawText = rawText;

	const auto lines = SplitLinesForSupportLibraryTest(NormalizeLineBreaksForSupportLibraryTest(rawText));
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
			std::string token = FindPossibleSupportLibraryFileToken(line);
			if (!token.empty()) {
				outInfo.fileName = std::filesystem::path(token).filename().string();
				outInfo.fileStem = std::filesystem::path(token).stem().string();
				if (token.find('\\') != std::string::npos || token.find('/') != std::string::npos) {
					outInfo.filePath = token;
				}
			}
		}
	}

	if (outInfo.name.empty()) {
		outInfo.name = std::format("support_library_{}", index);
		outInfo.rawName = outInfo.name;
	}
}

bool TryLoadSupportLibraryBasicInfoForTest(
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

	const auto* getInfoProc = reinterpret_cast<PFN_GET_LIB_INFO>(GetProcAddress(module, FUNCNAME_GET_LIB_INFO));
	if (getInfoProc == nullptr) {
		FreeLibrary(module);
		return false;
	}

	const auto* libInfo = getInfoProc();
	if (libInfo == nullptr) {
		FreeLibrary(module);
		return false;
	}

	if (libInfo->m_szName != nullptr) {
		outName = libInfo->m_szName;
	}
	outVersionText = std::format(
		"{}.{}.{}",
		libInfo->m_nMajorVersion,
		libInfo->m_nMinorVersion,
		libInfo->m_nBuildNumber);
	if (outGuid != nullptr && libInfo->m_szGuid != nullptr) {
		*outGuid = libInfo->m_szGuid;
	}

	FreeLibrary(module);
	return !outName.empty();
}

std::string ExtractLeadingAsciiStemForSupportLibraryTest(const std::string& text)
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

std::vector<LoadedSupportLibraryModuleForTest> EnumerateLoadedSupportLibraryModulesForTest()
{
	std::vector<LoadedSupportLibraryModuleForTest> modules;
	const std::filesystem::path libDir = std::filesystem::path(GetBasePath()) / "lib";
	std::error_code ec;
	if (!std::filesystem::exists(libDir, ec)) {
		return modules;
	}

	const std::string libDirLower = ToLowerAsciiCopy(libDir.string());
	HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE | TH32CS_SNAPMODULE32, GetCurrentProcessId());
	if (snapshot == INVALID_HANDLE_VALUE) {
		return modules;
	}

	MODULEENTRY32 entry = {};
	entry.dwSize = sizeof(entry);
	if (Module32First(snapshot, &entry) == FALSE) {
		CloseHandle(snapshot);
		return modules;
	}

	do {
		std::filesystem::path modulePath(entry.szExePath);
		const std::string fullPath = modulePath.string();
		const std::string fullPathLower = ToLowerAsciiCopy(fullPath);
		if (fullPathLower.rfind(libDirLower, 0) != 0) {
			continue;
		}

		const std::string extLower = ToLowerAsciiCopy(modulePath.extension().string());
		if (extLower != ".fne" && extLower != ".fnr" && extLower != ".dll") {
			continue;
		}

		LoadedSupportLibraryModuleForTest module;
		module.filePath = fullPath;
		module.fileName = modulePath.filename().string();
		module.fileStem = modulePath.stem().string();
		modules.push_back(std::move(module));
	} while (Module32Next(snapshot, &entry) != FALSE);

	CloseHandle(snapshot);
	return modules;
}

bool TryAssignSupportLibraryPathByLoadedModules(
	SupportLibraryHeaderForTest& info,
	const std::vector<LoadedSupportLibraryModuleForTest>& modules,
	std::unordered_set<size_t>& usedModuleIndexes)
{
	auto tryMatch = [&](const std::string& candidate, const char* trace) -> bool {
		const std::string needle = ToLowerAsciiCopy(TrimAsciiCopy(candidate));
		if (needle.empty()) {
			return false;
		}

		int matchedIndex = -1;
		for (size_t i = 0; i < modules.size(); ++i) {
			if (usedModuleIndexes.find(i) != usedModuleIndexes.end()) {
				continue;
			}
			if (ToLowerAsciiCopy(modules[i].fileName) == needle ||
				ToLowerAsciiCopy(modules[i].fileStem) == needle ||
				ToLowerAsciiCopy(modules[i].filePath) == needle) {
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

	const std::string prefixFromName = ExtractLeadingAsciiStemForSupportLibraryTest(info.name);
	if (tryMatch(prefixFromName, "loaded_module_ascii_prefix_name")) {
		return true;
	}

	const std::string prefixFromRaw = ExtractLeadingAsciiStemForSupportLibraryTest(info.rawText);
	if (tryMatch(prefixFromRaw, "loaded_module_ascii_prefix_raw")) {
		return true;
	}

	return false;
}

bool TryListSupportLibrariesForTest(std::vector<SupportLibraryHeaderForTest>& outInfos, std::string* outError)
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

		SupportLibraryHeaderForTest info;
		ParseSupportLibraryHeaderText(i, text, info);
		outInfos.push_back(std::move(info));
	}

	const auto loadedModules = EnumerateLoadedSupportLibraryModulesForTest();
	std::unordered_set<size_t> usedModuleIndexes;
	for (auto& info : outInfos) {
		if (!info.filePath.empty()) {
			info.resolveTrace = "header_text_path";
		}
		else {
			TryAssignSupportLibraryPathByLoadedModules(info, loadedModules, usedModuleIndexes);
		}

		if (!info.filePath.empty()) {
			std::string resolvedName;
			std::string resolvedVersion;
			if (TryLoadSupportLibraryBasicInfoForTest(info.filePath, resolvedName, resolvedVersion, nullptr)) {
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
			if (TryLoadSupportLibraryBasicInfoForTest(info.filePath, resolvedName, resolvedVersion, nullptr)) {
				info.name = resolvedName;
				info.versionText = resolvedVersion;
			}
		}
	}

	return true;
}

bool DumpSupportLibraryInfoFromFile(
	const std::string& filePath,
	std::vector<std::string>& outLines,
	std::string& outError)
{
	outLines.clear();
	outError.clear();

	HMODULE module = LoadLibraryExA(filePath.c_str(), nullptr, DONT_RESOLVE_DLL_REFERENCES);
	if (module == nullptr) {
		outError = "LoadLibraryEx failed";
		return false;
	}

	const auto* getInfoProc = reinterpret_cast<PFN_GET_LIB_INFO>(GetProcAddress(module, FUNCNAME_GET_LIB_INFO));
	if (getInfoProc == nullptr) {
		FreeLibrary(module);
		outError = "GetNewInf not found";
		return false;
	}

	const auto* libInfo = getInfoProc();
	if (libInfo == nullptr) {
		FreeLibrary(module);
		outError = "GetNewInf returned null";
		return false;
	}

	const auto readPtr = [](LPCSTR ptr) -> std::string {
		return ptr == nullptr ? std::string() : std::string(ptr);
	};
	const auto joinTexts = [](const std::vector<std::string>& values, const char* sep) -> std::string {
		std::string text;
		for (size_t i = 0; i < values.size(); ++i) {
			if (i > 0) {
				text += sep;
			}
			text += values[i];
		}
		return text;
	};
	const auto decodeDataType = [](DATA_TYPE type) -> std::string {
		const bool isArray = (type & DT_IS_ARY) != 0;
		const bool isVar = (type & DT_IS_VAR) != 0;
		const DATA_TYPE baseType = static_cast<DATA_TYPE>(type & ~DT_IS_ARY);

		std::string text;
		switch (baseType)
		{
		case _SDT_NULL: text = "空类型"; break;
		case _SDT_ALL: text = "通用型(all)"; break;
		case SDT_BYTE: text = "字节型(byte)"; break;
		case SDT_SHORT: text = "短整数型(short)"; break;
		case SDT_INT: text = "整数型(int)"; break;
		case SDT_INT64: text = "长整数型(int64)"; break;
		case SDT_FLOAT: text = "小数型(float)"; break;
		case SDT_DOUBLE: text = "双精度小数型(double)"; break;
		case SDT_BOOL: text = "逻辑型(bool)"; break;
		case SDT_DATE_TIME: text = "日期时间型(datetime)"; break;
		case SDT_TEXT: text = "文本型(text)"; break;
		case SDT_BIN: text = "字节集(bin)"; break;
		case SDT_SUB_PTR: text = "子程序指针(sub_ptr)"; break;
		case SDT_STATMENT: text = "子语句(statment)"; break;
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
	};
	const auto decodeCommandUserLevel = [](SHORT level) -> std::string {
		switch (level)
		{
		case LVL_SIMPLE: return "初级";
		case LVL_SECONDARY: return "中级";
		case LVL_HIGH: return "高级";
		default: return std::format("未知({})", static_cast<int>(level));
		}
	};
	const auto decodeLibOsText = [&](DWORD state) -> std::string {
		std::vector<std::string> items;
		if (_TEST_LIB_OS(state, __OS_WIN)) {
			items.emplace_back("Windows");
		}
		if (_TEST_LIB_OS(state, __OS_LINUX)) {
			items.emplace_back("Linux");
		}
		if (_TEST_LIB_OS(state, __OS_UNIX)) {
			items.emplace_back("Unix");
		}
		if (items.empty()) {
			return std::string("未声明");
		}
		return joinTexts(items, "/");
	};
	const auto decodeCmdOsText = [&](WORD state) -> std::string {
		std::vector<std::string> items;
		if (_TEST_CMD_OS(state, __OS_WIN)) {
			items.emplace_back("Windows");
		}
		if (_TEST_CMD_OS(state, __OS_LINUX)) {
			items.emplace_back("Linux");
		}
		if (_TEST_CMD_OS(state, __OS_UNIX)) {
			items.emplace_back("Unix");
		}
		if (items.empty()) {
			return std::string("未声明");
		}
		return joinTexts(items, "/");
	};
	const auto decodeDtOsText = [&](DWORD state) -> std::string {
		std::vector<std::string> items;
		if (_TEST_DT_OS(state, __OS_WIN)) {
			items.emplace_back("Windows");
		}
		if (_TEST_DT_OS(state, __OS_LINUX)) {
			items.emplace_back("Linux");
		}
		if (_TEST_DT_OS(state, __OS_UNIX)) {
			items.emplace_back("Unix");
		}
		if (items.empty()) {
			return std::string("未声明");
		}
		return joinTexts(items, "/");
	};
	const auto decodeArgFlags = [](DWORD state) -> std::vector<std::string> {
		std::vector<std::string> flags;
		if ((state & AS_HAS_DEFAULT_VALUE) != 0 || (state & AS_DEFAULT_VALUE_IS_EMPTY) != 0) {
			flags.emplace_back("可省略");
		}
		if ((state & AS_RECEIVE_VAR) != 0) {
			flags.emplace_back("仅变量");
		}
		if ((state & AS_RECEIVE_VAR_ARRAY) != 0) {
			flags.emplace_back("仅变量数组");
		}
		if ((state & AS_RECEIVE_VAR_OR_ARRAY) != 0) {
			flags.emplace_back("变量或数组");
		}
		if ((state & AS_RECEIVE_ARRAY_DATA) != 0) {
			flags.emplace_back("仅数组数据");
		}
		if ((state & AS_RECEIVE_ALL_TYPE_DATA) != 0) {
			flags.emplace_back("任意数组或非数组");
		}
		if ((state & AS_RECEIVE_VAR_OR_OTHER) != 0) {
			flags.emplace_back("变量或其它");
		}
		return flags;
	};
	const auto decodeCmdFlags = [](const CMD_INFO& cmd) -> std::vector<std::string> {
		std::vector<std::string> flags;
		if ((cmd.m_wState & CT_IS_HIDED) != 0) {
			flags.emplace_back("隐含");
		}
		if ((cmd.m_wState & CT_IS_ERROR) != 0) {
			flags.emplace_back("不可用");
		}
		if ((cmd.m_wState & CT_DISABLED_IN_RELEASE) != 0) {
			flags.emplace_back("Release禁用");
		}
		if ((cmd.m_wState & CT_ALLOW_APPEND_NEW_ARG) != 0) {
			flags.emplace_back("允许追加参数");
		}
		if ((cmd.m_wState & CT_RETRUN_ARY_TYPE_DATA) != 0) {
			flags.emplace_back("返回数组");
		}
		if ((cmd.m_wState & CT_IS_OBJ_COPY_CMD) != 0) {
			flags.emplace_back("对象复制");
		}
		if ((cmd.m_wState & CT_IS_OBJ_FREE_CMD) != 0) {
			flags.emplace_back("对象析构");
		}
		if ((cmd.m_wState & CT_IS_OBJ_CONSTURCT_CMD) != 0) {
			flags.emplace_back("对象构造");
		}
		if (cmd.m_shtCategory == -1) {
			flags.emplace_back("对象成员命令");
		}
		return flags;
	};
	const auto decodeDataTypeFlags = [](DWORD state) -> std::vector<std::string> {
		std::vector<std::string> flags;
		if ((state & LDT_IS_HIDED) != 0) {
			flags.emplace_back("隐含");
		}
		if ((state & LDT_IS_ERROR) != 0) {
			flags.emplace_back("不可用");
		}
		if ((state & LDT_WIN_UNIT) != 0) {
			flags.emplace_back("窗口组件");
		}
		if ((state & LDT_IS_CONTAINER) != 0) {
			flags.emplace_back("容器");
		}
		if ((state & LDT_IS_TAB_UNIT) != 0) {
			flags.emplace_back("Tab组件");
		}
		if ((state & LDT_IS_FUNCTION_PROVIDER) != 0) {
			flags.emplace_back("功能提供者");
		}
		if ((state & LDT_CANNOT_GET_FOCUS) != 0) {
			flags.emplace_back("不可聚焦");
		}
		if ((state & LDT_DEFAULT_NO_TABSTOP) != 0) {
			flags.emplace_back("默认无TabStop");
		}
		if ((state & LDT_ENUM) != 0) {
			flags.emplace_back("枚举");
		}
		if ((state & LDT_MSG_FILTER_CONTROL) != 0) {
			flags.emplace_back("消息过滤");
		}
		return flags;
	};
	const auto decodeElementFlags = [](DWORD state) -> std::vector<std::string> {
		std::vector<std::string> flags;
		if ((state & LES_HAS_DEFAULT_VALUE) != 0) {
			flags.emplace_back("有默认值");
		}
		if ((state & LES_HIDED) != 0) {
			flags.emplace_back("隐藏");
		}
		return flags;
	};
	const auto decodeConstType = [](SHORT type) -> std::string {
		switch (type)
		{
		case CT_NULL: return "空";
		case CT_NUM: return "数值";
		case CT_BOOL: return "逻辑";
		case CT_TEXT: return "文本";
		default: return std::format("未知({})", static_cast<int>(type));
		}
	};
	const auto formatCallSignature = [&](const CMD_INFO& cmd) -> std::string {
		std::string text;
		const std::string ret = decodeDataType(cmd.m_dtRetValType);
		if (cmd.m_dtRetValType == _SDT_NULL) {
			text += "〈无返回值〉 ";
		}
		else {
			text += "〈" + ret + "〉 ";
		}
		text += readPtr(cmd.m_szName);
		text += "（";
		for (int i = 0; i < cmd.m_nArgCount; ++i) {
			if (i > 0) {
				text += "，";
			}
			const ARG_INFO& arg = cmd.m_pBeginArgInfo[i];
			const bool optional = (arg.m_dwState & AS_HAS_DEFAULT_VALUE) != 0 ||
				(arg.m_dwState & AS_DEFAULT_VALUE_IS_EMPTY) != 0;
			if (optional) {
				text += "［";
			}
			text += decodeDataType(arg.m_dtType);
			text += " ";
			text += readPtr(arg.m_szName);
			if (optional) {
				text += "］";
			}
		}
		text += "）";
		return text;
	};
	std::vector<std::string> categories;
	if (libInfo->m_nCategoryCount > 0 && libInfo->m_szzCategory != nullptr) {
		const char* cursor = libInfo->m_szzCategory;
		for (int i = 0; i < libInfo->m_nCategoryCount; ++i) {
			const int bitmapIndex = *reinterpret_cast<const int*>(cursor);
			(void)bitmapIndex;
			cursor += sizeof(int);
			const std::string categoryName = readPtr(cursor);
			categories.push_back(categoryName);
			cursor += categoryName.size() + 1;
		}
	}
	const auto categoryNameOf = [&](SHORT category) -> std::string {
		if (category == -1) {
			return "对象成员";
		}
		if (category > 0 && static_cast<size_t>(category) <= categories.size()) {
			return categories[static_cast<size_t>(category) - 1];
		}
		return std::format("未分类({})", static_cast<int>(category));
	};
	const auto categoryPathOf = [&](SHORT category) -> std::string {
		const std::string libName = readPtr(libInfo->m_szName);
		const std::string categoryName = categoryNameOf(category);
		if (categoryName.empty()) {
			return libName;
		}
		return libName + "->" + categoryName;
	};
	const auto appendReadableCommandHelp = [&](int cmdIndex, const CMD_INFO& cmd) -> void {
		const std::string levelText = decodeCommandUserLevel(cmd.m_shtUserLevel);
		std::string explainText = readPtr(cmd.m_szExplain);
		if (!explainText.empty()) {
			const char last = explainText.back();
			if (last != '.' && last != '!' && last != '?') {
				explainText += "。";
			}
		}
		explainText += std::format("本命令为{}命令。", levelText);

		outLines.push_back(std::format("cmd_help#{}:", cmdIndex));
		outLines.push_back(std::format(
			"    调用格式： {} - {}",
			EscapeOneLineForLog(formatCallSignature(cmd)),
			EscapeOneLineForLog(categoryPathOf(cmd.m_shtCategory))));
		outLines.push_back(std::format(
			"    英文名称：{}",
			EscapeOneLineForLog(readPtr(cmd.m_szEgName))));
		outLines.push_back(std::format(
			"    {}",
			EscapeOneLineForLog(explainText)));

		for (int argIndex = 0; argIndex < cmd.m_nArgCount; ++argIndex) {
			const ARG_INFO& arg = cmd.m_pBeginArgInfo[argIndex];
			const auto argFlags = decodeArgFlags(arg.m_dwState);
			std::string argText = std::format(
				"    参数<{}>的名称为“{}”，类型为“{}”",
				argIndex + 1,
				EscapeOneLineForLog(readPtr(arg.m_szName)),
				EscapeOneLineForLog(decodeDataType(arg.m_dtType)));
			if (!argFlags.empty()) {
				argText += std::format("，{}", EscapeOneLineForLog(joinTexts(argFlags, "、")));
			}
			const std::string argExplain = readPtr(arg.m_szExplain);
			if (!argExplain.empty()) {
				argText += std::format("，说明：{}", EscapeOneLineForLog(argExplain));
			}
			if ((arg.m_dwState & AS_DEFAULT_VALUE_IS_EMPTY) != 0) {
				argText += "，默认值为空";
			}
			else if ((arg.m_dwState & AS_HAS_DEFAULT_VALUE) != 0) {
				argText += std::format("，默认值={}", arg.m_nDefault);
			}
			argText += "。";
			outLines.push_back(std::move(argText));
		}

		outLines.push_back(std::format(
			"    操作系统需求： {}",
			EscapeOneLineForLog(decodeCmdOsText(cmd.m_wState))));
	};

	outLines.push_back(std::format("file_path={}", filePath));
	outLines.push_back(std::format("support_library_name={}", readPtr(libInfo->m_szName)));
	outLines.push_back(std::format(
		"version={}.{}.{}",
		libInfo->m_nMajorVersion,
		libInfo->m_nMinorVersion,
		libInfo->m_nBuildNumber));
	outLines.push_back(std::format("guid={}", readPtr(libInfo->m_szGuid)));
	outLines.push_back(std::format("author={}", readPtr(libInfo->m_szAuthor)));
	outLines.push_back(std::format("explain={}", EscapeOneLineForLog(readPtr(libInfo->m_szExplain))));
	outLines.push_back(std::format(
		"required_eide_version={}.{}",
		libInfo->m_nRqSysMajorVer,
		libInfo->m_nRqSysMinorVer));
	outLines.push_back(std::format(
		"required_krnln_version={}.{}",
		libInfo->m_nRqSysKrnlLibMajorVer,
		libInfo->m_nRqSysKrnlLibMinorVer));
	outLines.push_back(std::format("supported_os={}", decodeLibOsText(libInfo->m_dwState)));
	outLines.push_back(std::format("category_count={}", libInfo->m_nCategoryCount));
	outLines.push_back(std::format("command_count={}", libInfo->m_nCmdCount));
	outLines.push_back(std::format("data_type_count={}", libInfo->m_nDataTypeCount));
	outLines.push_back(std::format("constant_count={}", libInfo->m_nLibConstCount));
	if (!categories.empty()) {
		outLines.push_back("");
		outLines.push_back("[Categories]");
		for (size_t i = 0; i < categories.size(); ++i) {
			outLines.push_back(std::format("category#{} name={}", i + 1, EscapeOneLineForLog(categories[i])));
		}
	}

	outLines.push_back("");
	outLines.push_back("[Commands]");
	for (int i = 0; i < libInfo->m_nCmdCount; ++i) {
		const CMD_INFO& cmd = libInfo->m_pBeginCmdInfo[i];
		const auto cmdFlags = decodeCmdFlags(cmd);
		outLines.push_back(std::format(
			"cmd#{} name={} eg={} category={} retType={} level={} argCount={} flags={} supported_os={}",
			i,
			EscapeOneLineForLog(readPtr(cmd.m_szName)),
			EscapeOneLineForLog(readPtr(cmd.m_szEgName)),
			EscapeOneLineForLog(categoryNameOf(cmd.m_shtCategory)),
			EscapeOneLineForLog(decodeDataType(cmd.m_dtRetValType)),
			EscapeOneLineForLog(decodeCommandUserLevel(cmd.m_shtUserLevel)),
			cmd.m_nArgCount,
			EscapeOneLineForLog(joinTexts(cmdFlags, ",")),
			EscapeOneLineForLog(decodeCmdOsText(cmd.m_wState))));
		outLines.push_back(std::format(
			"  call_format={}",
			EscapeOneLineForLog(formatCallSignature(cmd))));
		outLines.push_back(std::format(
			"  explain={}",
			EscapeOneLineForLog(readPtr(cmd.m_szExplain))));
		for (int argIndex = 0; argIndex < cmd.m_nArgCount; ++argIndex) {
			const ARG_INFO& arg = cmd.m_pBeginArgInfo[argIndex];
			const auto argFlags = decodeArgFlags(arg.m_dwState);
			outLines.push_back(std::format(
				"  arg#{} name={} type={} flags={} explain={}",
				argIndex,
				EscapeOneLineForLog(readPtr(arg.m_szName)),
				EscapeOneLineForLog(decodeDataType(arg.m_dtType)),
				EscapeOneLineForLog(joinTexts(argFlags, ",")),
				EscapeOneLineForLog(readPtr(arg.m_szExplain))));
		}
		appendReadableCommandHelp(i, cmd);
	}

	outLines.push_back("");
	outLines.push_back("[DataTypes]");
	for (int i = 0; i < libInfo->m_nDataTypeCount; ++i) {
		const LIB_DATA_TYPE_INFO& dataType = libInfo->m_pDataType[i];
		const auto flags = decodeDataTypeFlags(dataType.m_dwState);
		outLines.push_back(std::format(
			"type#{} name={} eg={} cmdCount={} elementCount={} flags={} supported_os={} explain={}",
			i,
			EscapeOneLineForLog(readPtr(dataType.m_szName)),
			EscapeOneLineForLog(readPtr(dataType.m_szEgName)),
			dataType.m_nCmdCount,
			dataType.m_nElementCount,
			EscapeOneLineForLog(joinTexts(flags, ",")),
			EscapeOneLineForLog(decodeDtOsText(dataType.m_dwState)),
			EscapeOneLineForLog(readPtr(dataType.m_szExplain))));
		for (int cmdIndex = 0; cmdIndex < dataType.m_nCmdCount; ++cmdIndex) {
			const int globalCmdIndex = dataType.m_pnCmdsIndex[cmdIndex];
			if (globalCmdIndex >= 0 && globalCmdIndex < libInfo->m_nCmdCount) {
				outLines.push_back(std::format(
					"  method#{} cmdIndex={} name={}",
					cmdIndex,
					globalCmdIndex,
					EscapeOneLineForLog(readPtr(libInfo->m_pBeginCmdInfo[globalCmdIndex].m_szName))));
			}
		}
		for (int elementIndex = 0; elementIndex < dataType.m_nElementCount; ++elementIndex) {
			const LIB_DATA_TYPE_ELEMENT& element = dataType.m_pElementBegin[elementIndex];
			const auto elementFlags = decodeElementFlags(element.m_dwState);
			outLines.push_back(std::format(
				"  element#{} name={} eg={} type={} flags={} explain={}",
				elementIndex,
				EscapeOneLineForLog(readPtr(element.m_szName)),
				EscapeOneLineForLog(readPtr(element.m_szEgName)),
				EscapeOneLineForLog(decodeDataType(element.m_dtType)),
				EscapeOneLineForLog(joinTexts(elementFlags, ",")),
				EscapeOneLineForLog(readPtr(element.m_szExplain))));
		}
	}

	outLines.push_back("");
	outLines.push_back("[Constants]");
	for (int i = 0; i < libInfo->m_nLibConstCount; ++i) {
		const LIB_CONST_INFO& item = libInfo->m_pLibConst[i];
		outLines.push_back(std::format(
			"const#{} name={} type={} textValue={} numericValue={} explain={}",
			i,
			EscapeOneLineForLog(readPtr(item.m_szName)),
			EscapeOneLineForLog(decodeConstType(item.m_shtType)),
			EscapeOneLineForLog(readPtr(item.m_szText)),
			item.m_dbValue,
			EscapeOneLineForLog(readPtr(item.m_szExplain))));
	}

	FreeLibrary(module);
	return true;
}

void RunImportedModuleListTest()
{
	OutputStringToELog("[ImportedModuleListTest] 开始枚举当前项目导入模块");

	std::vector<std::string> paths;
	std::string error;
	if (!TryListImportedModulePathsForTest(paths, &error)) {
		OutputStringToELog(std::format(
			"[ImportedModuleListTest] 枚举失败 error={}",
			EscapeOneLineForLog(error)));
		return;
	}

	OutputStringToELog(std::format(
		"[ImportedModuleListTest] 枚举完成 count={}",
		paths.size()));

	const size_t previewCount = (std::min)(paths.size(), static_cast<size_t>(20));
	for (size_t i = 0; i < previewCount; ++i) {
		OutputStringToELog(std::format(
			"[ImportedModuleListTest] #{} path={}",
			i,
			EscapeOneLineForLog(paths[i])));
	}
	if (paths.size() > previewCount) {
		OutputStringToELog(std::format(
			"[ImportedModuleListTest] 仅展示前 {} 条，剩余 {} 条未输出",
			previewCount,
			paths.size() - previewCount));
	}
}

void RunFirstImportedModulePublicInfoTest()
{
	OutputStringToELog("[ModulePublicInfoTest] 开始抓取首个导入模块公开信息");

	std::vector<std::string> paths;
	std::string error;
	if (!TryListImportedModulePathsForTest(paths, &error)) {
		OutputStringToELog(std::format(
			"[ModulePublicInfoTest] 获取模块列表失败 error={}",
			EscapeOneLineForLog(error)));
		return;
	}
	if (paths.empty()) {
		OutputStringToELog("[ModulePublicInfoTest] 当前项目没有导入任何模块");
		return;
	}

	const std::string modulePath = paths.front();
	e571::ModulePublicInfoDump dump;
	std::string loadError;
	if (!e571::LoadModulePublicInfoDump(
			modulePath,
			GetCurrentProcessImageBase(),
			&dump,
			&loadError)) {
		OutputStringToELog(std::format(
			"[ModulePublicInfoTest] 抓取失败 path={} error={} loaderError={} trace={}",
			EscapeOneLineForLog(modulePath),
			EscapeOneLineForLog(loadError),
			EscapeOneLineForLog(dump.loaderError),
			EscapeOneLineForLog(dump.trace)));
		return;
	}

	std::vector<std::string> lines;
	lines.push_back(std::format(
		"[ModulePublicInfoTest] path={} nativeResult={} recordCount={} sourceKind={} trace={}",
		modulePath,
		dump.nativeResult,
		dump.records.size(),
		dump.sourceKind,
		dump.trace));

	OutputStringToELog(std::format(
		"[ModulePublicInfoTest] 抓取成功 path={} nativeResult={} recordCount={} sourceKind={} trace={}",
		EscapeOneLineForLog(modulePath),
		dump.nativeResult,
		dump.records.size(),
		EscapeOneLineForLog(dump.sourceKind),
		EscapeOneLineForLog(dump.trace)));

	if (!dump.formattedText.empty()) {
		lines.push_back("");
		lines.push_back("[Formatted]");
		std::string currentLine;
		for (char ch : dump.formattedText) {
			if (ch == '\r') {
				continue;
			}
			if (ch == '\n') {
				lines.push_back(currentLine);
				currentLine.clear();
				continue;
			}
			currentLine.push_back(ch);
		}
		if (!currentLine.empty()) {
			lines.push_back(currentLine);
		}
	}

	const size_t previewCount = (std::min)(dump.records.size(), static_cast<size_t>(20));
	for (size_t i = 0; i < previewCount; ++i) {
		const auto& record = dump.records[i];
		std::string firstString;
		if (!record.extractedStrings.empty()) {
			firstString = EscapeOneLineForLog(record.extractedStrings.front());
		}

		OutputStringToELog(std::format(
			"[ModulePublicInfoTest] #{} tag={} kind={} name={} bodySize={} headerCount={} stringCount={} first={}",
			i,
			record.tag,
			EscapeOneLineForLog(record.kind),
			EscapeOneLineForLog(record.name),
			record.bodySize,
			record.headerInts.size(),
			record.extractedStrings.size(),
			firstString));

		lines.push_back(std::format(
			"#{} tag={} bodySize={} payloadOffset={} headerInts={} strings={} first={}",
			i,
			record.tag,
			record.bodySize,
			record.payloadOffset,
			record.headerInts.size(),
			record.extractedStrings.size(),
			firstString));
		if (!record.headerInts.empty()) {
			std::string headerIntsText;
			for (size_t h = 0; h < record.headerInts.size(); ++h) {
				if (!headerIntsText.empty()) {
					headerIntsText += ",";
				}
				headerIntsText += std::to_string(record.headerInts[h]);
			}
			lines.push_back(std::format(
				"  headerInts={}",
				headerIntsText));
		}
		for (size_t s = 0; s < record.extractedStrings.size() && s < 8; ++s) {
			lines.push_back(std::format(
				"  string#{}={}",
				s,
				EscapeOneLineForLog(record.extractedStrings[s])));
		}
	}
	if (dump.records.size() > previewCount) {
		OutputStringToELog(std::format(
			"[ModulePublicInfoTest] 仅展示前 {} 条，剩余 {} 条请查看文件",
			previewCount,
			dump.records.size() - previewCount));
	}

	WriteModulePublicInfoLog(lines);
	OutputStringToELog(std::format(
		"[ModulePublicInfoTest] 日志文件 path={}",
		GetModulePublicInfoLogPath().string()));
}

void RunFirstSupportLibraryInfoTest()
{
	OutputStringToELog("[SupportLibraryInfoTest] 开始抓取当前已选支持库信息");

	std::vector<SupportLibraryHeaderForTest> libs;
	std::string error;
	if (!TryListSupportLibrariesForTest(libs, &error)) {
		OutputStringToELog(std::format(
			"[SupportLibraryInfoTest] 获取支持库列表失败 error={}",
			EscapeOneLineForLog(error)));
		return;
	}

	std::vector<std::string> lines;
	lines.push_back(std::format("support_library_count={}", libs.size()));
	OutputStringToELog(std::format(
		"[SupportLibraryInfoTest] 枚举完成 count={}",
		libs.size()));

	const size_t previewCount = (std::min)(libs.size(), static_cast<size_t>(20));
	for (size_t i = 0; i < previewCount; ++i) {
		const auto& lib = libs[i];
		OutputStringToELog(std::format(
			"[SupportLibraryInfoTest] #{} index={} name={} version={} fileName={} filePath={} trace={} rawName={}",
			i,
			lib.index,
			EscapeOneLineForLog(lib.name),
			EscapeOneLineForLog(lib.versionText),
			EscapeOneLineForLog(lib.fileName),
			EscapeOneLineForLog(lib.filePath),
			EscapeOneLineForLog(lib.resolveTrace),
			EscapeOneLineForLog(lib.rawName)));
		lines.push_back(std::format(
			"#{} index={} name={} version={} fileName={} filePath={} trace={} rawName={}",
			i,
			lib.index,
			lib.name,
			lib.versionText,
			lib.fileName,
			lib.filePath,
			lib.resolveTrace,
			lib.rawName));
	}

	if (libs.empty()) {
		WriteSupportLibraryInfoLog(lines);
		OutputStringToELog(std::format(
			"[SupportLibraryInfoTest] 当前没有已选支持库，日志文件 path={}",
			GetSupportLibraryInfoLogPath().string()));
		return;
	}

	const auto it = std::find_if(libs.begin(), libs.end(), [](const SupportLibraryHeaderForTest& item) {
		return !item.filePath.empty();
	});
	if (it == libs.end()) {
		lines.push_back("");
		lines.push_back("[HeaderTextOnly]");
		lines.push_back(libs.front().rawText);
		WriteSupportLibraryInfoLog(lines);
		OutputStringToELog(std::format(
			"[SupportLibraryInfoTest] 未解析到支持库文件路径，仅输出IDE文本 path={}",
			GetSupportLibraryInfoLogPath().string()));
		return;
	}

	std::vector<std::string> detailLines;
	std::string dumpError;
	if (!DumpSupportLibraryInfoFromFile(it->filePath, detailLines, dumpError)) {
		OutputStringToELog(std::format(
			"[SupportLibraryInfoTest] 抓取失败 filePath={} error={}",
			EscapeOneLineForLog(it->filePath),
			EscapeOneLineForLog(dumpError)));
		lines.push_back("");
		lines.push_back(std::format("[DumpError] {}", dumpError));
		WriteSupportLibraryInfoLog(lines);
		OutputStringToELog(std::format(
			"[SupportLibraryInfoTest] 日志文件 path={}",
			GetSupportLibraryInfoLogPath().string()));
		return;
	}

	lines.push_back("");
	lines.push_back("[FirstSupportLibraryDump]");
	lines.insert(lines.end(), detailLines.begin(), detailLines.end());
	WriteSupportLibraryInfoLog(lines);
	OutputStringToELog(std::format(
		"[SupportLibraryInfoTest] 抓取成功 filePath={} 日志文件 path={}",
		EscapeOneLineForLog(it->filePath),
		GetSupportLibraryInfoLogPath().string()));
}

void RunStaticCompileWindowsExeTest()
{
	const std::string outputPath = (
		std::filesystem::path(GetBasePath()) /
		"AutoLinker" /
		"StaticCompileTest" /
		"AutoLinkerStaticTest.exe").string();

	std::string normalizedPath;
	std::string diagnostics;
	const bool ok = IDEFacade::Instance().CompileWithOutputPath(
		IDEFacade::CompileOutputKind::WinExe,
		outputPath,
		true,
		&normalizedPath,
		&diagnostics);

	if (!ok) {
		OutputStringToELog(std::format(
			"[StaticCompileExeTest] 触发失败 output={} diagnostics={}",
			EscapeOneLineForLog(outputPath),
			EscapeOneLineForLog(diagnostics)));
		return;
	}

	OutputStringToELog(std::format(
		"[StaticCompileExeTest] 已触发静态编译 output={} diagnostics={}",
		EscapeOneLineForLog(normalizedPath),
		EscapeOneLineForLog(diagnostics)));
}

void AutoRunModulePublicInfoTestThread(void* /*pParams*/)
{
	Sleep(10000);
	if (g_hwnd != nullptr && IsWindow(g_hwnd)) {
		PostMessageA(g_hwnd, WM_AUTOLINKER_RUN_MODULE_INFO_TEST, 0, 0);
	}
}

HTREEITEM FindTreeItemByExactTextRecursive(HWND treeHwnd, HTREEITEM firstItem, const std::string& targetText, int maxDepth, int depth)
{
	if (treeHwnd == nullptr || firstItem == nullptr || depth > maxDepth) {
		return nullptr;
	}

	for (HTREEITEM item = firstItem; item != nullptr; item = GetTreeNextItem(treeHwnd, item, TVGN_NEXT)) {
		std::string text;
		LPARAM itemData = 0;
		UINT state = 0;
		int childCount = 0;
		int image = -1;
		int selectedImage = -1;
		if (QueryTreeItemInfo(treeHwnd, item, text, itemData, state, childCount, image, selectedImage) && text == targetText) {
			return item;
		}

		if (depth < maxDepth) {
			if (HTREEITEM foundChild = FindTreeItemByExactTextRecursive(
					treeHwnd,
					GetTreeNextItem(treeHwnd, item, TVGN_CHILD),
					targetText,
					maxDepth,
					depth + 1)) {
				return foundChild;
			}
		}
	}
	return nullptr;
}

struct ProgramTreePageItemInfo
{
	int depth = 0;
	std::string text;
	unsigned int itemData = 0;
	int image = -1;
	int selectedImage = -1;
	std::string typeName;
};

struct ProgramTreeVisualHints
{
	int classImage = -1;
	int windowAssemblyImage = -1;
};

bool IsProgramTreeType1Item(const ProgramTreePageItemInfo& item)
{
	return (item.itemData >> 28) == 1u;
}

ProgramTreeVisualHints BuildProgramTreeVisualHints(const std::vector<ProgramTreePageItemInfo>& items)
{
	ProgramTreeVisualHints hints;
	for (const auto& item : items) {
		if (!IsProgramTreeType1Item(item) || item.image < 0) {
			continue;
		}
		if (hints.classImage < 0 && item.text.rfind("Class_", 0) == 0) {
			hints.classImage = item.image;
		}
		if (hints.windowAssemblyImage < 0 && item.text.rfind("窗口程序集_", 0) == 0) {
			hints.windowAssemblyImage = item.image;
		}
		if (hints.classImage >= 0 && hints.windowAssemblyImage >= 0) {
			break;
		}
	}
	return hints;
}

std::string DescribeProgramTreeItemType(
	unsigned int itemData,
	const std::string& text,
	int image,
	const ProgramTreeVisualHints& hints)
{
	const unsigned int typeNibble = itemData >> 28;
	switch (typeNibble) {
	case 1:
		if (text.rfind("Class_", 0) == 0) {
			return "类模块";
		}
		if (image >= 0) {
			if (hints.classImage >= 0 && image == hints.classImage) {
				return "类模块";
			}
		}
		return "程序集";
	case 2:
		return "全局变量";
	case 3:
		return "自定义数据类型";
	case 4:
		return "DLL命令";
	case 5:
		return "窗口/表单";
	case 6:
		return "常量资源";
	case 7:
		return ((itemData & 0x0FFFFFFFu) == 1u) ? "图片资源" : "声音资源";
	case 15:
		return "分组";
	default:
		return "未知";
	}
}

void CollectProgramTreePageItemsRecursive(
	HWND treeHwnd,
	HTREEITEM firstItem,
	int depth,
	int maxDepth,
	std::vector<ProgramTreePageItemInfo>& outItems)
{
	if (treeHwnd == nullptr || firstItem == nullptr || depth > maxDepth) {
		return;
	}

	for (HTREEITEM item = firstItem; item != nullptr; item = GetTreeNextItem(treeHwnd, item, TVGN_NEXT)) {
		std::string text;
		LPARAM itemData = 0;
		UINT state = 0;
		int childCount = 0;
		int image = -1;
		int selectedImage = -1;
		if (!QueryTreeItemInfo(treeHwnd, item, text, itemData, state, childCount, image, selectedImage)) {
			continue;
		}

		const unsigned int itemDataU = static_cast<unsigned int>(itemData);
		const unsigned int typeNibble = itemDataU >> 28;
		if (typeNibble != 0 && typeNibble != 15) {
			ProgramTreePageItemInfo info;
			info.depth = depth;
			info.text = text;
			info.itemData = itemDataU;
			info.image = image;
			info.selectedImage = selectedImage;
			outItems.push_back(std::move(info));
			continue;
		}

		if (depth < maxDepth) {
			CollectProgramTreePageItemsRecursive(
				treeHwnd,
				GetTreeNextItem(treeHwnd, item, TVGN_CHILD),
				depth + 1,
				maxDepth,
				outItems);
		}
	}
}

void WriteProgramTreeListLog(const std::vector<ProgramTreePageItemInfo>& items)
{
	const auto path = GetProgramTreeListLogPath();
	std::lock_guard<std::mutex> guard(g_programTreeListLogMutex);
	std::ofstream out(path, std::ios::trunc | std::ios::binary);
	if (!out.is_open()) {
		OutputStringToELog("[ProgramTreeListTest] 写程序树清单文件失败");
		return;
	}

	for (size_t i = 0; i < items.size(); ++i) {
		out << std::format(
			"#{} depth={} type={} data=0x{:08X} image={} selImage={} text={}\r\n",
			i,
			items[i].depth,
			items[i].typeName,
			items[i].itemData,
			items[i].image,
			items[i].selectedImage,
			items[i].text);
	}
}

void RunProgramTreeListTest()
{
	OutputStringToELog("[ProgramTreeListTest] 开始枚举程序树中的页面节点");

	const HWND treeHwnd = FindProgramDataTreeView();
	if (treeHwnd == nullptr) {
		OutputStringToELog("[ProgramTreeListTest] 未找到程序树 TreeView");
		return;
	}

	const HTREEITEM rootItem = GetTreeNextItem(treeHwnd, nullptr, TVGN_ROOT);
	const HTREEITEM firstChild = GetTreeNextItem(treeHwnd, rootItem, TVGN_CHILD);
	std::vector<ProgramTreePageItemInfo> items;
	CollectProgramTreePageItemsRecursive(treeHwnd, firstChild, 0, 8, items);
	const ProgramTreeVisualHints hints = BuildProgramTreeVisualHints(items);
	for (auto& item : items) {
		item.typeName = DescribeProgramTreeItemType(item.itemData, item.text, item.image, hints);
	}
	WriteProgramTreeListLog(items);

	OutputStringToELog(std::format(
		"[ProgramTreeListTest] 枚举完成 count={} classImage={} windowAssemblyImage={} path={}",
		items.size(),
		hints.classImage,
		hints.windowAssemblyImage,
		GetProgramTreeListLogPath().string()));

	const size_t previewCount = (std::min)(items.size(), static_cast<size_t>(20));
	for (size_t i = 0; i < previewCount; ++i) {
		OutputStringToELog(std::format(
			"[ProgramTreeListTest] #{} depth={} type={} data=0x{:08X} image={} selImage={} text={}",
			i,
			items[i].depth,
			items[i].typeName,
			items[i].itemData,
			items[i].image,
			items[i].selectedImage,
			EscapeOneLineForLog(items[i].text)));
	}
	if (items.size() > previewCount) {
		OutputStringToELog(std::format(
			"[ProgramTreeListTest] 仅展示前 {} 条，剩余 {} 条请查看文件",
			previewCount,
			items.size() - previewCount));
	}
}

void RunProgramTreeDirectPageDumpTest()
{
	OutputStringToELog(std::format(
		"[TreeDirectPageDumpTest] 开始在程序树中查找并抓取 name={}",
		kTreeDirectPageDumpTestName));

	const HWND treeHwnd = FindProgramDataTreeView();
	if (treeHwnd == nullptr) {
		OutputStringToELog("[TreeDirectPageDumpTest] 未找到程序树 TreeView");
		return;
	}

	const HTREEITEM rootItem = GetTreeNextItem(treeHwnd, nullptr, TVGN_ROOT);
	const HTREEITEM item = FindTreeItemByExactTextRecursive(treeHwnd, rootItem, kTreeDirectPageDumpTestName, 4, 0);
	if (item == nullptr) {
		OutputStringToELog(std::format(
			"[TreeDirectPageDumpTest] 未在程序树中找到目标 name={} tree={}",
			kTreeDirectPageDumpTestName,
			PtrToHexText(reinterpret_cast<UINT_PTR>(treeHwnd))));
		return;
	}

	std::string itemText;
	LPARAM itemData = 0;
	UINT state = 0;
	int childCount = 0;
	int image = -1;
	int selectedImage = -1;
	if (!QueryTreeItemInfo(treeHwnd, item, itemText, itemData, state, childCount, image, selectedImage)) {
		OutputStringToELog("[TreeDirectPageDumpTest] 读取目标节点信息失败");
		return;
	}

	std::string currentPageCode;
	e571::RawSearchContextPageDumpDebugResult dumpResult;
	if (!e571::DebugDumpCodePageByProgramTreeItemData(
			static_cast<unsigned int>(itemData),
			GetCurrentProcessImageBase(),
			&currentPageCode,
			&dumpResult)) {
		OutputStringToELog(std::format(
			"[TreeDirectPageDumpTest] 抓取失败 text={} data=0x{:08X} image={} selImage={} type={} resolvedIndex={} bucketData={} outerCount={} lineCount={} fetchFailures={} trace={}",
			EscapeOneLineForLog(itemText),
			static_cast<unsigned int>(itemData),
			image,
			selectedImage,
			dumpResult.type,
			dumpResult.resolvedIndex,
			dumpResult.bucketData,
			dumpResult.outerCount,
			dumpResult.lineCount,
			dumpResult.fetchFailures,
			EscapeOneLineForLog(dumpResult.trace)));
		return;
	}

	WriteDirectGlobalSearchPageDump(currentPageCode);
	const size_t lineCount = static_cast<size_t>(std::count(currentPageCode.begin(), currentPageCode.end(), '\n')) +
		(currentPageCode.empty() ? 0u : 1u);
	OutputStringToELog(std::format(
		"[TreeDirectPageDumpTest] 抓取成功 text={} data=0x{:08X} image={} selImage={} type={} resolvedIndex={} bucketData={} outerCount={} lines={} fetchFailures={} trace={} path={}",
		EscapeOneLineForLog(itemText),
		static_cast<unsigned int>(itemData),
		image,
		selectedImage,
		dumpResult.type,
		dumpResult.resolvedIndex,
		dumpResult.bucketData,
		dumpResult.outerCount,
		lineCount,
		dumpResult.fetchFailures,
		EscapeOneLineForLog(dumpResult.trace),
		GetDirectGlobalSearchPageDumpLogPath().string()));
	LogCurrentPageCodePreview(currentPageCode, 10);
}
#else
void RunProgramTreeDirectPageDumpTest()
{
	OutputStringToELog("[TreeDirectPageDumpTest] 当前仅支持 x86 配置");
}
#endif

void WriteAddTabTestLog(const std::vector<std::string>& lines)
{
	const auto path = GetAddTabTestLogPath();
	std::lock_guard<std::mutex> guard(g_addTabTestLogMutex);
	std::ofstream out(path, std::ios::trunc | std::ios::binary);
	if (!out.is_open()) {
		OutputStringToELog("[AutoLinker][ADD_TAB_TEST] 写日志文件失败");
		return;
	}
	for (const auto& line : lines) {
		out << line << "\r\n";
	}
}

HWND EnsureAddTabTestWindow()
{
	static HWND s_testWnd = nullptr;
	if (s_testWnd != nullptr && IsWindow(s_testWnd)) {
		return s_testWnd;
	}

	HWND parent = g_hwnd;
	HINSTANCE module = GetModuleHandleA(nullptr);
	s_testWnd = CreateWindowExA(
		WS_EX_CLIENTEDGE,
		"EDIT",
		"AutoLinker FN_ADD_TAB Test Content",
		WS_CHILD | ES_MULTILINE | ES_AUTOVSCROLL | WS_VSCROLL | ES_READONLY,
		0, 0, 540, 320,
		parent,
		nullptr,
		module,
		nullptr);
	return s_testWnd;
}

void RunFnAddTabStructPassThroughTest()
{
	std::vector<std::string> logs;
	const auto appendLog = [&logs](const std::string& line) {
		logs.push_back(line);
		OutputStringToELog(line);
	};

	appendLog("[AutoLinker][ADD_TAB_TEST] 开始测试 FN_ADD_TAB（传 ADD_TAB_INF*）");
	HWND testWnd = EnsureAddTabTestWindow();
	appendLog("[AutoLinker][ADD_TAB_TEST] " + DescribeWindowForAddTabLog(testWnd));
	if (testWnd == nullptr || !IsWindow(testWnd)) {
		appendLog("[AutoLinker][ADD_TAB_TEST] 中止：测试窗口创建失败");
		WriteAddTabTestLog(logs);
		return;
	}

	const uint32_t seed = static_cast<uint32_t>(GetTickCount());
	const std::string captionUtf8 = std::format("AutoLinker AddTab Test {}", seed);
	const std::string tooltipUtf8 = std::format("trace-{}", seed);

	ADD_TAB_INF tabInf = {};
	tabInf.m_hWnd = testWnd;
	tabInf.m_hIcon = nullptr;
#ifdef UNICODE
	std::wstring captionW = std::wstring(captionUtf8.begin(), captionUtf8.end());
	std::wstring tooltipW = std::wstring(tooltipUtf8.begin(), tooltipUtf8.end());
	tabInf.m_szCaption = const_cast<LPWSTR>(captionW.c_str());
	tabInf.m_szToolTip = const_cast<LPWSTR>(tooltipW.c_str());
#else
	tabInf.m_szCaption = const_cast<LPSTR>(captionUtf8.c_str());
	tabInf.m_szToolTip = const_cast<LPSTR>(tooltipUtf8.c_str());
#endif

	const HWND beforeWnd = tabInf.m_hWnd;
	const HICON beforeIcon = tabInf.m_hIcon;
	const LPTSTR beforeCaptionPtr = tabInf.m_szCaption;
	const LPTSTR beforeTooltipPtr = tabInf.m_szToolTip;

	appendLog(std::format(
		"[AutoLinker][ADD_TAB_TEST] before struct={} m_hWnd={} m_hIcon={} m_szCaption={} m_szToolTip={} caption=\"{}\" tooltip=\"{}\"",
		PtrToHexText(reinterpret_cast<UINT_PTR>(&tabInf)),
		PtrToHexText(reinterpret_cast<UINT_PTR>(tabInf.m_hWnd)),
		PtrToHexText(reinterpret_cast<UINT_PTR>(tabInf.m_hIcon)),
		PtrToHexText(reinterpret_cast<UINT_PTR>(tabInf.m_szCaption)),
		PtrToHexText(reinterpret_cast<UINT_PTR>(tabInf.m_szToolTip)),
		EscapeOneLineForLog(captionUtf8),
		EscapeOneLineForLog(tooltipUtf8)));

	const INT rawRet = IDEFacade::Instance().RunFunctionRaw(
		FN_ADD_TAB,
		static_cast<DWORD>(reinterpret_cast<UINT_PTR>(&tabInf)),
		0);
	const bool ok = (rawRet != FALSE);

	appendLog(std::format(
		"[AutoLinker][ADD_TAB_TEST] after rawRet={} ok={} struct={} m_hWnd={} m_hIcon={} m_szCaption={} m_szToolTip={}",
		rawRet,
		ok ? 1 : 0,
		PtrToHexText(reinterpret_cast<UINT_PTR>(&tabInf)),
		PtrToHexText(reinterpret_cast<UINT_PTR>(tabInf.m_hWnd)),
		PtrToHexText(reinterpret_cast<UINT_PTR>(tabInf.m_hIcon)),
		PtrToHexText(reinterpret_cast<UINT_PTR>(tabInf.m_szCaption)),
		PtrToHexText(reinterpret_cast<UINT_PTR>(tabInf.m_szToolTip))));

	appendLog(std::format(
		"[AutoLinker][ADD_TAB_TEST] delta hWndChanged={} hIconChanged={} captionPtrChanged={} tooltipPtrChanged={}",
		(tabInf.m_hWnd != beforeWnd) ? 1 : 0,
		(tabInf.m_hIcon != beforeIcon) ? 1 : 0,
		(tabInf.m_szCaption != beforeCaptionPtr) ? 1 : 0,
		(tabInf.m_szToolTip != beforeTooltipPtr) ? 1 : 0));
	appendLog("[AutoLinker][ADD_TAB_TEST] after-window " + DescribeWindowForAddTabLog(tabInf.m_hWnd));

	WriteAddTabTestLog(logs);
	appendLog("[AutoLinker][ADD_TAB_TEST] 详细日志已写入 add_tab_test_last.log");
}

#if defined(_M_IX86)
std::uintptr_t GetCurrentProcessImageBase()
{
	HMODULE module = GetModuleHandleW(nullptr);
	return reinterpret_cast<std::uintptr_t>(module);
}

void LogDirectGlobalSearchResults(
	const e571::HiddenBuiltinSearchDebugResult& result,
	size_t maxCount)
{
	const size_t count = (std::min)(result.previewLines.size(), maxCount);
	for (size_t i = 0; i < count; ++i) {
		OutputStringToELog(std::format(
			"[DirectGlobalSearchTest] #{} text={}",
			i,
			EscapeOneLineForLog(result.previewLines[i])));
	}

	if (result.hits > count) {
		OutputStringToELog(std::format(
			"[DirectGlobalSearchTest] 仅展示前 {} 条，剩余 {} 条未输出",
			count,
			result.hits - count));
	}
}

void RunDirectGlobalSearchKeywordTest()
{
	OutputStringToELog(std::format(
		"[DirectGlobalSearchTest] 开始固定搜索 keyword={}",
		kDirectGlobalSearchTestKeyword));

	const auto result = e571::DebugSearchDirectGlobalKeywordHidden(
		kDirectGlobalSearchTestKeyword,
		GetCurrentProcessImageBase());
	OutputStringToELog(std::format(
		"[DirectGlobalSearchTest] 搜索完成 keyword={} hits={} dialogHandled={} first={}",
		kDirectGlobalSearchTestKeyword,
		result.hits,
		result.dialogHandled ? 1 : 0,
		EscapeOneLineForLog(result.firstResultText)));
	LogDirectGlobalSearchResults(result, 10);
}

void RunDirectGlobalSearchLocateKeywordTest()
{
	OutputStringToELog(std::format(
		"[DirectGlobalSearchLocateTest] 开始固定搜索并定位 keyword={}",
		kDirectGlobalSearchTestKeyword));

	e571::HiddenBuiltinSearchDebugResult result;
	const bool ok = e571::DebugLocateFirstDirectGlobalKeywordHidden(
		kDirectGlobalSearchTestKeyword,
		GetCurrentProcessImageBase(),
		&result);
	OutputStringToELog(std::format(
		"[DirectGlobalSearchLocateTest] 搜索完成 keyword={} hits={} dialogHandled={}",
		kDirectGlobalSearchTestKeyword,
		result.hits,
		result.dialogHandled ? 1 : 0));
	if (result.hasRawFirstHit) {
		OutputStringToELog(std::format(
			"[DirectGlobalSearchLocateTest] rawFirst type={} extra={} outer={} inner={} offset={} rawHits={} resolveOk={} resolvedIndex={} bucketData={}",
			result.rawFirstHit.type,
			result.rawFirstHit.extra,
			result.rawFirstHit.outerIndex,
			result.rawFirstHit.innerIndex,
			result.rawFirstHit.matchOffset,
			result.rawHitCount,
			result.rawResolveOk ? 1 : 0,
			result.rawResolvedIndex,
			result.rawBucketData));
	}
	if (result.hasDirectFirstHit) {
		OutputStringToELog(std::format(
			"[DirectGlobalSearchLocateTest] directFirst type={} extra={} outer={} inner={} offset={} directHits={} text={}",
			result.directFirstHit.type,
			result.directFirstHit.extra,
			result.directFirstHit.outerIndex,
			result.directFirstHit.innerIndex,
			result.directFirstHit.matchOffset,
			result.directHitCount,
			EscapeOneLineForLog(result.directFirstHit.displayText)));
	}
	if (result.hits == 0) {
		OutputStringToELog("[DirectGlobalSearchLocateTest] 未找到可定位结果");
		return;
	}

	OutputStringToELog(std::format(
		"[DirectGlobalSearchLocateTest] jump={} first={} trace={}",
		ok ? 1 : 0,
		EscapeOneLineForLog(result.firstResultText),
		EscapeOneLineForLog(result.jumpTrace)));
}

void WriteDirectGlobalSearchPageDump(const std::string& text)
{
	const auto path = GetDirectGlobalSearchPageDumpLogPath();
	std::lock_guard<std::mutex> guard(g_directGlobalSearchPageDumpLogMutex);
	std::ofstream out(path, std::ios::trunc | std::ios::binary);
	if (!out.is_open()) {
		OutputStringToELog("[DirectGlobalSearchPageDumpTest] 写页面代码文件失败");
		return;
	}
	out.write(text.data(), static_cast<std::streamsize>(text.size()));
}

void LogCurrentPageCodePreview(const std::string& pageCode, size_t maxLines)
{
	size_t lineBegin = 0;
	size_t lineIndex = 0;
	while (lineBegin <= pageCode.size() && lineIndex < maxLines) {
		size_t lineEnd = pageCode.find('\n', lineBegin);
		std::string line = pageCode.substr(
			lineBegin,
			lineEnd == std::string::npos ? std::string::npos : lineEnd - lineBegin);
		if (!line.empty() && line.back() == '\r') {
			line.pop_back();
		}
		OutputStringToELog(std::format(
			"[DirectGlobalSearchPageDumpTest] pageLine#{} {}",
			lineIndex,
			EscapeOneLineForLog(ConvertPossiblyUtf8ToLocalForEOutput(line))));
		++lineIndex;
		if (lineEnd == std::string::npos) {
			break;
		}
		lineBegin = lineEnd + 1;
	}
}

void RunDirectGlobalSearchLocateAndDumpCurrentPageTest()
{
	OutputStringToELog(std::format(
		"[DirectGlobalSearchPageDumpTest] 开始固定搜索、定位并抓取当前页代码 keyword={}",
		kDirectGlobalSearchTestKeyword));

	e571::HiddenBuiltinSearchDebugResult result;
	const bool located = e571::DebugLocateFirstDirectGlobalKeywordHidden(
		kDirectGlobalSearchTestKeyword,
		GetCurrentProcessImageBase(),
		&result);
	OutputStringToELog(std::format(
		"[DirectGlobalSearchPageDumpTest] 定位完成 keyword={} hits={} dialogHandled={} jump={} trace={}",
		kDirectGlobalSearchTestKeyword,
		result.hits,
		result.dialogHandled ? 1 : 0,
		located ? 1 : 0,
		EscapeOneLineForLog(result.jumpTrace)));
	if (!located) {
		OutputStringToELog("[DirectGlobalSearchPageDumpTest] 中止：定位失败，未抓取当前页代码");
		return;
	}

	std::string currentPageCode;
	e571::DirectGlobalSearchDebugHit dumpHit = {};
	if (result.hasRawFirstHit) {
		dumpHit = result.rawFirstHit;
	}
	else if (result.hasDirectFirstHit) {
		dumpHit = result.directFirstHit;
	}
	else {
		OutputStringToELog("[DirectGlobalSearchPageDumpTest] 中止：没有可用的命中记录用于抓取页面");
		return;
	}

	e571::NativeEditorPageDumpDebugResult dumpResult;
	if (!e571::DebugDumpCodePageForSearchHit(
			dumpHit,
			GetCurrentProcessImageBase(),
			&currentPageCode,
			&dumpResult)) {
		OutputStringToELog(std::format(
			"[DirectGlobalSearchPageDumpTest] DebugDumpCodePageForSearchHit failed editor=0x{:X} outerCount={} lineCount={} fetchFailures={} trace={}",
			static_cast<unsigned long long>(dumpResult.editorObject),
			dumpResult.outerCount,
			dumpResult.lineCount,
			dumpResult.fetchFailures,
			EscapeOneLineForLog(dumpResult.trace)));
		return;
	}

	WriteDirectGlobalSearchPageDump(currentPageCode);
	const size_t lineCount = static_cast<size_t>(std::count(currentPageCode.begin(), currentPageCode.end(), '\n')) +
		(currentPageCode.empty() ? 0u : 1u);
	OutputStringToELog(std::format(
		"[DirectGlobalSearchPageDumpTest] 当前页代码抓取成功 bytes={} lines={} editor=0x{:X} outerCount={} fetchFailures={} trace={} path={}",
		currentPageCode.size(),
		lineCount,
		static_cast<unsigned long long>(dumpResult.editorObject),
		dumpResult.outerCount,
		dumpResult.fetchFailures,
		EscapeOneLineForLog(dumpResult.trace),
		GetDirectGlobalSearchPageDumpLogPath().string()));
	LogCurrentPageCodePreview(currentPageCode, 10);
}
#else
void RunDirectGlobalSearchKeywordTest()
{
	OutputStringToELog("[DirectGlobalSearchTest] 当前仅支持 x86 配置");
}

void RunDirectGlobalSearchLocateKeywordTest()
{
	OutputStringToELog("[DirectGlobalSearchLocateTest] 当前仅支持 x86 配置");
}

void RunDirectGlobalSearchLocateAndDumpCurrentPageTest()
{
	OutputStringToELog("[DirectGlobalSearchPageDumpTest] 当前仅支持 x86 配置");
}
#endif

void AppendAIRoundtripLogBlock(const std::string& title, const std::string& text)
{
	const auto path = GetAIRoundtripLogPath();
	std::lock_guard<std::mutex> guard(g_aiRoundtripLogMutex);
	std::ofstream out(path, std::ios::app | std::ios::binary);
	if (!out.is_open()) {
		return;
	}

	AppendAIRoundtripLogLineUnlocked(
		out,
		"[BLOCK-BEGIN] title=" + title +
		" bytes=" + std::to_string(text.size()) +
		" utf8Valid=" + std::to_string(IsValidUtf8ForLog(text) ? 1 : 0) +
		" hexHead=" + BuildHexHead(text));
	out << text << "\r\n";
	AppendAIRoundtripLogLineUnlocked(out, "[BLOCK-END] title=" + title);
}
