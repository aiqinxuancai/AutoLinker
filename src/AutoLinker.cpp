#include "AutoLinker.h"
#include <vector>
#include <Windows.h>
#include <CommCtrl.h>
#include <format>
#include <algorithm>
#include <atomic>
#include <cctype>
#include "ConfigManager.h"
#include <regex>
#include "PathHelper.h"
#include "LinkerManager.h"
#include "InlineHook.h"
#include "ModelManager.h"
#include "Global.h"
#include "StringHelper.h"
#include <thread>
#include <chrono>
#include "MouseBack.h"
#include <PublicIDEFunctions.h>
#include "ECOMEx.h"
#include "WindowHelper.h"
#include <future>
#include "WinINetUtil.h"
#include "Version.h"
#include "IDEFacade.h"
#include "AIService.h"
#include "AIConfigDialog.h"
#include "AIChatFeature.h"
#if defined(_M_IX86)
#include "direct_global_search_debug.hpp"
#endif
#include <memory>
#include <new>
#include <process.h>
#include <cstdint>
#include <unordered_map>
#include <filesystem>
#include <fstream>
#include <mutex>

#pragma comment(lib, "comctl32.lib")

//当前打开的易源文件路径
std::string g_nowOpenSourceFilePath;

//配置文件，当前路径的e源码对应的编译器名
ConfigManager g_configManager;

//管理当前所有的link文件
LinkerManager g_linkerManager;

//管理模块 调试 -> 编译 的管理器
ModelManager g_modelManager;


//e主窗口句柄
HWND g_hwnd = NULL;

//工具条句柄
HWND g_toolBarHwnd = NULL;

//准备开始调试 废弃
bool g_preDebugging;

//准备开始编译 废弃
bool g_preCompiling;

bool g_initStarted = false;

HMENU g_topLinkerSubMenu = NULL;
std::unordered_map<UINT, std::string> g_topLinkerCommandMap;

void UpdateCurrentOpenSourceFile();
void OutputCurrentSourceLinker();

namespace {
bool g_isContextMenuRegistered = false;
constexpr UINT IDM_AUTOLINKER_CTX_COPY_FUNC = 31001; // Keep < 0x8000 to avoid signed-ID issues in some hosts.
constexpr UINT IDM_AUTOLINKER_CTX_AI_OPTIMIZE_FUNC = 31101;
constexpr UINT IDM_AUTOLINKER_CTX_AI_COMMENT_FUNC = 31102;
constexpr UINT IDM_AUTOLINKER_CTX_AI_TRANSLATE_FUNC = 31103;
constexpr UINT IDM_AUTOLINKER_CTX_AI_TRANSLATE_TEXT = 31104;
constexpr UINT IDM_AUTOLINKER_CTX_AI_ADD_BY_PAGE = 31106;
constexpr UINT IDM_AUTOLINKER_LINKER_BASE = 34000;
constexpr UINT IDM_AUTOLINKER_LINKER_MAX = 34999;
constexpr UINT WM_AUTOLINKER_AI_TASK_DONE = WM_USER + 1001;
constexpr UINT WM_AUTOLINKER_AI_APPLY_RESULT = WM_USER + 1002;

std::atomic_bool g_aiTaskInProgress = false;
std::atomic_uint64_t g_aiPerfTraceSeed = 1;
thread_local uint64_t g_aiPerfTraceIdTLS = 0;

enum class AIAsyncUiAction {
	ReplaceCurrentFunction,
	OutputTranslation,
	InsertAtPageBottom
};

struct AIAsyncRequest {
	uint64_t traceId = 0;
	AIAsyncUiAction action = AIAsyncUiAction::ReplaceCurrentFunction;
	AITaskKind taskKind = AITaskKind::OptimizeFunction;
	AISettings settings = {};
	std::string inputText;
	std::string displayName;
	std::string pageCodeSnapshot;
	std::string sourceFunctionCode;
	int targetCaretRow = -1;
	int targetCaretCol = -1;
};

struct AIAsyncResult {
	uint64_t traceId = 0;
	AIAsyncUiAction action = AIAsyncUiAction::ReplaceCurrentFunction;
	std::string displayName;
	std::string pageCodeSnapshot;
	std::string sourceFunctionCode;
	AIResult taskResult = {};
	int targetCaretRow = -1;
	int targetCaretCol = -1;
};

struct AIApplyRequest {
	AIAsyncUiAction action = AIAsyncUiAction::ReplaceCurrentFunction;
	std::string text;
	std::string sourceFunctionCode;
	int addedCount = 0;
	int skippedCount = 0;
	int targetCaretRow = -1;
	int targetCaretCol = -1;
};

void HandleAiTaskCompletionMessage(LPARAM lParam);
void HandleAiApplyMessage(LPARAM lParam);

std::string TrimAsciiCopy(const std::string& text)
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

std::string TruncateForPerfLog(const std::string& text, size_t maxLen = 120)
{
	if (text.size() <= maxLen) {
		return text;
	}
	return text.substr(0, maxLen) + "...";
}

using PerfClock = std::chrono::steady_clock;

long long ElapsedMs(const PerfClock::time_point& start)
{
	return static_cast<long long>(
		std::chrono::duration_cast<std::chrono::milliseconds>(PerfClock::now() - start).count());
}

std::mutex g_aiRoundtripLogMutex;
std::mutex g_addTabTestLogMutex;
std::mutex g_directGlobalSearchPageDumpLogMutex;
std::mutex g_programTreeListLogMutex;
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

class ScopedAIPerfTrace {
public:
	explicit ScopedAIPerfTrace(uint64_t traceId)
		: m_prev(GetCurrentAIPerfTraceId())
	{
		SetCurrentAIPerfTraceId(traceId);
	}

	~ScopedAIPerfTrace()
	{
		SetCurrentAIPerfTraceId(m_prev);
	}

	ScopedAIPerfTrace(const ScopedAIPerfTrace&) = delete;
	ScopedAIPerfTrace& operator=(const ScopedAIPerfTrace&) = delete;

private:
	uint64_t m_prev = 0;
};

std::vector<std::string> SplitLinesNormalized(const std::string& text)
{
	std::string normalized = text;
	for (size_t pos = 0; (pos = normalized.find("\r\n", pos)) != std::string::npos; ) {
		normalized.replace(pos, 2, "\n");
		++pos;
	}

	std::vector<std::string> lines;
	size_t start = 0;
	while (start <= normalized.size()) {
		size_t end = normalized.find('\n', start);
		if (end == std::string::npos) {
			lines.push_back(normalized.substr(start));
			break;
		}
		lines.push_back(normalized.substr(start, end - start));
		start = end + 1;
	}
	return lines;
}

std::string JoinLinesCrLf(const std::vector<std::string>& lines)
{
	std::string output;
	for (const std::string& line : lines) {
		output += line;
		output += "\r\n";
	}
	return output;
}

std::string NormalizeCodeForEIDE(const std::string& text)
{
	if (text.empty()) {
		return std::string();
	}

	std::string expanded = text;
	const bool hasRealLineBreak = expanded.find('\n') != std::string::npos || expanded.find('\r') != std::string::npos;
	if (!hasRealLineBreak && (expanded.find("\\n") != std::string::npos || expanded.find("\\r\\n") != std::string::npos)) {
		std::string unescaped;
		unescaped.reserve(expanded.size());
		for (size_t i = 0; i < expanded.size(); ++i) {
			if (expanded[i] == '\\' && i + 1 < expanded.size()) {
				if (expanded[i + 1] == 'n') {
					unescaped.push_back('\n');
					++i;
					continue;
				}
				if (expanded[i + 1] == 'r') {
					if (i + 3 < expanded.size() && expanded[i + 2] == '\\' && expanded[i + 3] == 'n') {
						unescaped.push_back('\n');
						i += 3;
						continue;
					}
					unescaped.push_back('\n');
					++i;
					continue;
				}
			}
			unescaped.push_back(expanded[i]);
		}
		expanded.swap(unescaped);
	}

	std::string normalized;
	normalized.reserve(expanded.size() + 16);
	for (size_t i = 0; i < expanded.size(); ++i) {
		const char ch = expanded[i];
		if (ch == '\0') {
			continue;
		}
		if (ch == '\r') {
			if (i + 1 < expanded.size() && expanded[i + 1] == '\n') {
				++i;
			}
			normalized += "\r\n";
			continue;
		}
		if (ch == '\n') {
			normalized += "\r\n";
			continue;
		}
		normalized.push_back(ch);
	}
	return normalized;
}

std::string ToLowerAsciiCopy(const std::string& text)
{
	std::string lowered = text;
	std::transform(lowered.begin(), lowered.end(), lowered.begin(), [](unsigned char c) {
		return static_cast<char>(std::tolower(c));
	});
	return lowered;
}

bool StartsWithAscii(const std::string& text, const std::string& prefix)
{
	return text.rfind(prefix, 0) == 0;
}

std::string ExtractNameAfterPrefix(const std::string& line, const std::string& prefix)
{
	if (!StartsWithAscii(line, prefix)) {
		return std::string();
	}
	size_t pos = prefix.size();
	while (pos < line.size() && std::isspace(static_cast<unsigned char>(line[pos])) != 0) {
		++pos;
	}
	size_t end = pos;
	while (end < line.size()) {
		const unsigned char c = static_cast<unsigned char>(line[end]);
		if (c == ',' || c == '\t' || c == ' ') {
			break;
		}
		++end;
	}
	return TrimAsciiCopy(line.substr(pos, end - pos));
}

bool IsLikelyChineseText(const std::string& text)
{
	for (unsigned char ch : text) {
		if (ch >= 0x80) {
			return true;
		}
	}
	return false;
}

std::string DescribeActivePageType(IDEFacade::ActiveWindowType type)
{
	switch (type)
	{
	case IDEFacade::ActiveWindowType::Module:
		return "程序集/类页面";
	case IDEFacade::ActiveWindowType::UserDataType:
		return "数据类型页面";
	case IDEFacade::ActiveWindowType::DllCommand:
		return "DLL命令声明页面";
	case IDEFacade::ActiveWindowType::GlobalVar:
		return "全局变量页面";
	case IDEFacade::ActiveWindowType::ConstResource:
		return "常量资源页面";
	case IDEFacade::ActiveWindowType::PictureResource:
		return "图片资源页面";
	case IDEFacade::ActiveWindowType::SoundResource:
		return "声音资源页面";
	default:
		return "未知/通用代码页面";
	}
}

std::string BuildAddByPageTypeOutputRule(IDEFacade::ActiveWindowType type)
{
	switch (type)
	{
	case IDEFacade::ActiveWindowType::Module:
		return "优先新增 .子程序（包含必要 .参数/.局部变量/实现语句），避免重复已有子程序名。";
	case IDEFacade::ActiveWindowType::DllCommand:
		return "优先新增 .DLL命令 + .参数 声明，避免重复已有命令名。";
	case IDEFacade::ActiveWindowType::UserDataType:
		return "优先新增 .数据类型 / .成员 声明，避免重复已有类型名。";
	default:
		return "根据用户需求新增最合适的易语言代码片段，保持可直接追加到页底。";
	}
}

void OutputMultiline(const std::string& title, const std::string& body)
{
	OutputStringToELog(title);
	auto lines = SplitLinesNormalized(body);
	for (const std::string& line : lines) {
		IDEFacade::Instance().AppendOutputWindowLine(line);
	}
}

bool EnsureAISettingsReady(AISettings& settings)
{
	AIService::LoadSettings(g_configManager, settings);
	std::string missing;
	if (AIService::HasRequiredSettings(settings, missing)) {
		return true;
	}

	OutputStringToELog("AI配置缺失，准备打开配置窗口");
	if (!ShowAIConfigDialog(g_hwnd, settings)) {
		OutputStringToELog("AI配置已取消，本次操作终止");
		return false;
	}
	AIService::SaveSettings(g_configManager, settings);
	OutputStringToELog("AI配置已保存");
	return true;
}

bool TryBeginAiTask()
{
	bool expected = false;
	if (!g_aiTaskInProgress.compare_exchange_strong(expected, true)) {
		OutputStringToELog("[AI]已有任务在执行，请稍候");
		return false;
	}
	return true;
}

void EndAiTask()
{
	g_aiTaskInProgress.store(false);
}

void PostAiTaskResult(AIAsyncResult* result)
{
	if (result == nullptr) {
		EndAiTask();
		return;
	}

	if (g_hwnd == NULL || !IsWindow(g_hwnd) || PostMessage(g_hwnd, WM_AUTOLINKER_AI_TASK_DONE, 0, reinterpret_cast<LPARAM>(result)) == FALSE) {
		if (!result->taskResult.ok) {
			OutputStringToELog("[AI]请求失败：" + result->taskResult.error);
		}
		else {
			OutputStringToELog("[AI]任务完成，但回调消息投递失败");
		}
		delete result;
		EndAiTask();
	}
}

void PostAiApplyRequest(AIApplyRequest* request)
{
	if (request == nullptr) {
		return;
	}
	if (g_hwnd == NULL || !IsWindow(g_hwnd) || PostMessage(g_hwnd, WM_AUTOLINKER_AI_APPLY_RESULT, 0, reinterpret_cast<LPARAM>(request)) == FALSE) {
		HandleAiApplyMessage(reinterpret_cast<LPARAM>(request));
	}
}

void RunAiTaskWorker(void* pParams)
{
	std::unique_ptr<AIAsyncRequest> request(reinterpret_cast<AIAsyncRequest*>(pParams));
	if (!request) {
		EndAiTask();
		return;
	}
	ScopedAIPerfTrace traceScope(request->traceId);

	std::unique_ptr<AIAsyncResult> result(new (std::nothrow) AIAsyncResult());
	if (!result) {
		OutputStringToELog("[AI]内存不足，无法执行请求");
		EndAiTask();
		return;
	}

	result->action = request->action;
	result->displayName = request->displayName;
	result->pageCodeSnapshot = request->pageCodeSnapshot;
	result->sourceFunctionCode = request->sourceFunctionCode;
	result->targetCaretRow = request->targetCaretRow;
	result->targetCaretCol = request->targetCaretCol;
	result->traceId = request->traceId;

	try {
		const auto networkStart = PerfClock::now();
		result->taskResult = AIService::ExecuteTask(request->taskKind, request->inputText, request->settings);
		AppendAIRoundtripLogLine(
			"[RESPONSE] ok=" + std::to_string(result->taskResult.ok ? 1 : 0) +
			" httpStatus=" + std::to_string(result->taskResult.httpStatus) +
			" contentBytes=" + std::to_string(result->taskResult.content.size()) +
			" error=\"" + EscapeOneLineForLog(result->taskResult.error) + "\"");
		AppendAIRoundtripLogBlock("ai_raw_response_content", result->taskResult.content);
		LogAIPerfCost(
			request->traceId,
			"RunAiTaskWorker.execute_task_total",
			ElapsedMs(networkStart),
			"ok=" + std::to_string(result->taskResult.ok ? 1 : 0) + " http=" + std::to_string(result->taskResult.httpStatus));
	}
	catch (const std::exception& ex) {
		result->taskResult.ok = false;
		result->taskResult.error = std::string("后台任务异常：") + ex.what();
		AppendAIRoundtripLogLine(
			"[RESPONSE] exception_std what=\"" + EscapeOneLineForLog(ex.what()) + "\"");
	}
	catch (...) {
		result->taskResult.ok = false;
		result->taskResult.error = "后台任务发生未知异常";
		AppendAIRoundtripLogLine("[RESPONSE] exception_unknown");
	}

	PostAiTaskResult(result.release());
}

void RunAiFunctionReplaceTask(AITaskKind kind)
{
	const uint64_t traceId = AllocateAIPerfTraceId();
	ScopedAIPerfTrace traceScope(traceId);
	const auto totalStart = PerfClock::now();
	auto logTotal = [&](const std::string& status) {
		LogAIPerfCost(
			traceId,
			"RunAiFunctionReplaceTask.total",
			ElapsedMs(totalStart),
			"status=" + status + " kind=" + std::to_string(static_cast<int>(kind)));
	};

	try {
		const std::string displayName = AIService::BuildTaskDisplayName(kind);
		OutputStringToELog("[AI]开始执行：" + displayName);
		BeginAIRoundtripLogSession(traceId, "function_replace", displayName);
		AppendAIRoundtripLogLine(
			"[STATE] begin kind=" + std::to_string(static_cast<int>(kind)) +
			" displayName=\"" + EscapeOneLineForLog(displayName) + "\"");
		if (!TryBeginAiTask()) {
			AppendAIRoundtripLogLine("[STATE] abort reason=busy");
			logTotal("busy");
			return;
		}

		IDEFacade& ide = IDEFacade::Instance();
		std::string functionCode;
		std::string functionDiag;
		{
			const auto t0 = PerfClock::now();
			const bool ok = ide.GetCurrentFunctionCode(functionCode, &functionDiag);
			LogAIPerfCost(
				traceId,
				"GetCurrentFunctionCode.total",
				ElapsedMs(t0),
				"ok=" + std::to_string(ok ? 1 : 0) + " diag=" + TruncateForPerfLog(functionDiag));
			if (!ok) {
				OutputStringToELog("[AI]无法获取当前函数代码");
				if (!functionDiag.empty()) {
					OutputStringToELog("[AI]函数代码获取诊断：" + functionDiag);
				}
				AppendAIRoundtripLogLine(
					"[STATE] get_function_failed diag=\"" + EscapeOneLineForLog(functionDiag) + "\"");
				EndAiTask();
				logTotal("get_function_failed");
				return;
			}
		}
		if (functionCode.empty()) {
			OutputStringToELog("[AI]无法获取当前函数代码");
			AppendAIRoundtripLogLine("[STATE] get_function_failed reason=empty_function_code");
			EndAiTask();
			logTotal("empty_function");
			return;
		}
		int caretRow = -1;
		int caretCol = -1;
		ide.GetCaretPosition(caretRow, caretCol);
		AppendAIRoundtripLogLine(
			"[STATE] caret row=" + std::to_string(caretRow) +
			" col=" + std::to_string(caretCol) +
			" functionDiag=\"" + EscapeOneLineForLog(functionDiag) + "\"");
		AppendAIRoundtripLogBlock("source_function_code", functionCode);

		AISettings settings = {};
		{
			const auto t0 = PerfClock::now();
			const bool ok = EnsureAISettingsReady(settings);
			LogAIPerfCost(
				traceId,
				"EnsureAISettingsReady.total",
				ElapsedMs(t0),
				"ok=" + std::to_string(ok ? 1 : 0));
			if (!ok) {
				AppendAIRoundtripLogLine("[STATE] abort reason=settings_not_ready");
				EndAiTask();
				logTotal("settings_not_ready");
				return;
			}
		}
		AppendAIRoundtripLogLine(
			"[SETTINGS] baseUrl=\"" + EscapeOneLineForLog(settings.baseUrl) +
			"\" model=\"" + EscapeOneLineForLog(settings.model) +
			"\" timeoutMs=" + std::to_string(settings.timeoutMs) +
			" temperature=" + std::format("{:.3f}", settings.temperature));

		const std::string userInput =
			"请处理以下易语言函数代码，并严格返回完整可替换代码：\n```e\n" +
			functionCode + "\n```";
		std::string finalInput = userInput;
		if (kind == AITaskKind::OptimizeFunction) {
			std::string extraRequirement;
			const auto inputStart = PerfClock::now();
			const bool accepted = ShowAITextInputDialog(
				g_hwnd,
				"AI优化函数 - 输入优化要求",
				"请输入你希望优化的方向（可为空）：",
				extraRequirement);
			LogAIPerfCost(
				traceId,
				"ShowAITextInputDialog.modal",
				ElapsedMs(inputStart),
				"ok=" + std::to_string(accepted ? 1 : 0) + " scene=optimize");
			if (!accepted) {
				OutputStringToELog("[AI]已取消输入优化要求");
				AppendAIRoundtripLogLine("[STATE] user_cancel_optimize_requirement");
				EndAiTask();
				logTotal("user_cancel_optimize_requirement");
				return;
			}
			extraRequirement = TrimAsciiCopy(extraRequirement);
			AppendAIRoundtripLogLine(
				"[INPUT] optimize_requirement bytes=" + std::to_string(extraRequirement.size()) +
				" text=\"" + EscapeOneLineForLog(extraRequirement) + "\"");
			if (!extraRequirement.empty()) {
				finalInput = std::string("额外优化要求：\n") + extraRequirement + "\n\n" + userInput;
			}
		}
		AppendAIRoundtripLogBlock("ai_input_prompt", finalInput);

		std::unique_ptr<AIAsyncRequest> request(new (std::nothrow) AIAsyncRequest());
		if (!request) {
			OutputStringToELog("[AI]内存不足，无法发起任务");
			AppendAIRoundtripLogLine("[STATE] abort reason=alloc_request_failed");
			EndAiTask();
			logTotal("alloc_request_failed");
			return;
		}
		request->traceId = traceId;
		request->action = AIAsyncUiAction::ReplaceCurrentFunction;
		request->taskKind = kind;
		request->settings = settings;
		request->inputText = finalInput;
		request->displayName = displayName;
		request->sourceFunctionCode = functionCode;
		request->targetCaretRow = caretRow;
		request->targetCaretCol = caretCol;
		AppendAIRoundtripLogLine(
			"[REQUEST] queued action=ReplaceCurrentFunction trace=" + std::to_string(traceId) +
			" inputBytes=" + std::to_string(request->inputText.size()) +
			" sourceBytes=" + std::to_string(request->sourceFunctionCode.size()));

		OutputStringToELog("[AI]正在请求模型（后台）...");
		uintptr_t threadId = _beginthread(RunAiTaskWorker, 0, request.get());
		if (threadId == static_cast<uintptr_t>(-1L)) {
			OutputStringToELog("[AI]启动后台任务失败");
			AppendAIRoundtripLogLine("[STATE] abort reason=start_worker_failed");
			EndAiTask();
			logTotal("start_worker_failed");
			return;
		}
		request.release();
		AppendAIRoundtripLogLine("[STATE] worker_started");
		logTotal("queued");
	}
	catch (const std::exception& ex) {
		OutputStringToELog(std::string("[AI]发生异常：") + ex.what());
		AppendAIRoundtripLogLine(
			"[STATE] exception_std what=\"" + EscapeOneLineForLog(ex.what()) + "\"");
		EndAiTask();
		logTotal("exception_std");
	}
	catch (...) {
		OutputStringToELog("[AI]发生未知异常");
		AppendAIRoundtripLogLine("[STATE] exception_unknown");
		EndAiTask();
		logTotal("exception_unknown");
	}
}

void RunAiTranslateSelectedTextTask()
{
	const uint64_t traceId = AllocateAIPerfTraceId();
	ScopedAIPerfTrace traceScope(traceId);
	const auto totalStart = PerfClock::now();
	auto logTotal = [&](const std::string& status) {
		LogAIPerfCost(
			traceId,
			"RunAiTranslateSelectedTextTask.total",
			ElapsedMs(totalStart),
			"status=" + status);
	};

	try {
		OutputStringToELog("[AI]开始执行：AI翻译选中文本");
		if (!TryBeginAiTask()) {
			logTotal("busy");
			return;
		}
		IDEFacade& ide = IDEFacade::Instance();
		std::string selectedText;
		if (!ide.GetSelectedText(selectedText)) {
			OutputStringToELog("[AI]未检测到有效选中文本");
			EndAiTask();
			logTotal("no_selection");
			return;
		}

		AISettings settings = {};
		{
			const auto t0 = PerfClock::now();
			const bool ok = EnsureAISettingsReady(settings);
			LogAIPerfCost(
				traceId,
				"EnsureAISettingsReady.total",
				ElapsedMs(t0),
				"ok=" + std::to_string(ok ? 1 : 0) + " scene=translate_selected");
			if (!ok) {
				EndAiTask();
				logTotal("settings_not_ready");
				return;
			}
		}

		const bool sourceIsChinese = IsLikelyChineseText(selectedText);
		const std::string direction = sourceIsChinese ? "请把以下中文翻译为英文，仅输出翻译结果：" : "请把以下英文翻译为中文，仅输出翻译结果：";
		const std::string input = direction + "\n" + selectedText;
		std::unique_ptr<AIAsyncRequest> request(new (std::nothrow) AIAsyncRequest());
		if (!request) {
			OutputStringToELog("[AI]内存不足，无法发起任务");
			EndAiTask();
			logTotal("alloc_request_failed");
			return;
		}
		request->traceId = traceId;
		request->action = AIAsyncUiAction::OutputTranslation;
		request->taskKind = AITaskKind::TranslateText;
		request->settings = settings;
		request->inputText = input;
		request->displayName = "AI翻译选中文本";

		OutputStringToELog("[AI]正在请求模型（后台）...");
		uintptr_t threadId = _beginthread(RunAiTaskWorker, 0, request.get());
		if (threadId == static_cast<uintptr_t>(-1L)) {
			OutputStringToELog("[AI]启动后台任务失败");
			EndAiTask();
			logTotal("start_worker_failed");
			return;
		}
		request.release();
		logTotal("queued");
	}
	catch (const std::exception& ex) {
		OutputStringToELog(std::string("[AI]发生异常：") + ex.what());
		EndAiTask();
		logTotal("exception_std");
	}
	catch (...) {
		OutputStringToELog("[AI]发生未知异常");
		EndAiTask();
		logTotal("exception_unknown");
	}
}

void HandleAiTaskCompletionMessage(LPARAM lParam)
{
	struct AiTaskEndGuard {
		~AiTaskEndGuard() { EndAiTask(); }
	} endGuard;

	std::unique_ptr<AIAsyncResult> result(reinterpret_cast<AIAsyncResult*>(lParam));
	if (!result) {
		return;
	}

	try {
		if (!result->taskResult.ok) {
			OutputStringToELog("[AI]请求失败：" + result->taskResult.error);
			AppendAIRoundtripLogLine(
				"[STATE] request_failed error=\"" + EscapeOneLineForLog(result->taskResult.error) + "\"");
			return;
		}

		switch (result->action)
		{
		case AIAsyncUiAction::ReplaceCurrentFunction: {
			std::string generatedCode = AIService::NormalizeModelOutputToCode(result->taskResult.content);
			generatedCode = AIService::Trim(generatedCode);
			if (generatedCode.empty()) {
				OutputStringToELog("[AI]模型返回为空");
				AppendAIRoundtripLogLine("[STATE] empty_generated_code_after_normalize");
				return;
			}
			generatedCode = NormalizeCodeForEIDE(generatedCode);
			AppendAIRoundtripLogBlock("ai_output_normalized_code", generatedCode);

			const std::string title = result->displayName.empty() ? "AI结果预览" : (result->displayName + " - 结果预览");
			const AIPreviewAction previewAction = ShowAIPreviewDialogEx(
				g_hwnd,
				title,
				generatedCode,
				"复制到剪贴板",
				"替换（不稳定）");
			if (previewAction == AIPreviewAction::Cancel) {
				OutputStringToELog("[AI]用户取消应用结果");
				AppendAIRoundtripLogLine("[PREVIEW] action=cancel");
				return;
			}
			if (previewAction == AIPreviewAction::SecondaryConfirm) {
				std::unique_ptr<AIApplyRequest> request(new (std::nothrow) AIApplyRequest());
				if (!request) {
					OutputStringToELog("[AI]内存不足，无法执行替换");
					AppendAIRoundtripLogLine("[PREVIEW] action=replace_unstable alloc_request_failed");
					return;
				}
				request->action = AIAsyncUiAction::ReplaceCurrentFunction;
				request->text = generatedCode;
				request->sourceFunctionCode = result->sourceFunctionCode;
				request->targetCaretRow = result->targetCaretRow;
				request->targetCaretCol = result->targetCaretCol;
				OutputStringToELog("[AI]用户选择：替换（不稳定）");
				AppendAIRoundtripLogLine(
					"[PREVIEW] action=replace_unstable targetCaretRow=" + std::to_string(request->targetCaretRow) +
					" targetCaretCol=" + std::to_string(request->targetCaretCol));
				AppendAIRoundtripLogBlock("replace_request_payload", request->text);
				PostAiApplyRequest(request.release());
				return;
			}

			OutputStringToELog("[AI]用户选择：复制到剪贴板");
			AppendAIRoundtripLogLine("[PREVIEW] action=copy_to_clipboard");
			AppendAIRoundtripLogBlock("clipboard_payload", generatedCode);
			if (!IDEFacade::Instance().SetClipboardText(generatedCode)) {
				OutputStringToELog("[AI]复制到剪贴板失败");
				AppendAIRoundtripLogLine("[STATE] copy_to_clipboard_failed");
				return;
			}
			OutputStringToELog("[AI]已复制到剪贴板，请手动替换");
			AppendAIRoundtripLogLine("[STATE] copy_to_clipboard_done");
			return;
		}
		case AIAsyncUiAction::OutputTranslation: {
			std::string translated = AIService::NormalizeModelOutputToCode(result->taskResult.content);
			translated = AIService::Trim(translated);
			if (translated.empty()) {
				OutputStringToELog("[AI]翻译结果为空");
				return;
			}
			OutputMultiline("[AI]翻译结果：", translated);
			return;
		}
		case AIAsyncUiAction::InsertAtPageBottom: {
			std::string generatedText = AIService::NormalizeModelOutputToCode(result->taskResult.content);
			generatedText = AIService::Trim(generatedText);
			if (generatedText.empty()) {
				OutputStringToELog("[AI]模型未返回可新增内容");
				return;
			}
			generatedText = NormalizeCodeForEIDE(generatedText);

			if (!ShowAIPreviewDialog(g_hwnd, "AI按页类型添加代码 - 结果预览", generatedText, "插入")) {
				OutputStringToELog("[AI]用户取消插入");
				return;
			}

			std::unique_ptr<AIApplyRequest> request(new (std::nothrow) AIApplyRequest());
			if (!request) {
				OutputStringToELog("[AI]内存不足，无法执行插入");
				return;
			}
			request->action = AIAsyncUiAction::InsertAtPageBottom;
			request->text = generatedText;
			PostAiApplyRequest(request.release());
			return;
		}
		default:
			OutputStringToELog("[AI]未知任务结果类型");
			return;
		}
	}
	catch (const std::exception& ex) {
		OutputStringToELog(std::string("[AI]处理结果异常：") + ex.what());
	}
	catch (...) {
		OutputStringToELog("[AI]处理结果发生未知异常");
	}
}

void HandleAiApplyMessage(LPARAM lParam)
{
	std::unique_ptr<AIApplyRequest> request(reinterpret_cast<AIApplyRequest*>(lParam));
	if (!request) {
		return;
	}

	try {
		IDEFacade& ide = IDEFacade::Instance();
		constexpr bool kAiApplyPreCompile = false;
		switch (request->action)
		{
		case AIAsyncUiAction::ReplaceCurrentFunction: {
			if (g_hwnd != NULL && IsWindow(g_hwnd)) {
				SetForegroundWindow(g_hwnd);
			}
			if (request->targetCaretRow >= 0 && request->targetCaretCol >= 0) {
				ide.MoveCaret(request->targetCaretRow, request->targetCaretCol);
			}
			AppendAIRoundtripLogLine(
				"[APPLY] begin action=ReplaceCurrentFunction targetCaretRow=" + std::to_string(request->targetCaretRow) +
				" targetCaretCol=" + std::to_string(request->targetCaretCol));
			AppendAIRoundtripLogBlock("apply_request_text_raw", request->text);
			const std::string newFunctionCode = NormalizeCodeForEIDE(request->text);
			AppendAIRoundtripLogBlock("apply_request_text_normalized", newFunctionCode);
			if (ide.ReplaceCurrentFunctionCode(newFunctionCode, kAiApplyPreCompile)) {
				OutputStringToELog("[AI]替换完成");
				AppendAIRoundtripLogLine("[APPLY] result=ok");
				return;
			}
			OutputStringToELog("[AI]替换当前函数失败（当前坐标未定位到可替换子程序）");
			AppendAIRoundtripLogLine("[APPLY] result=failed reason=replace_current_function_failed");
			return;
		}

		case AIAsyncUiAction::InsertAtPageBottom:
			if (g_hwnd != NULL && IsWindow(g_hwnd)) {
				SetForegroundWindow(g_hwnd);
			}
			if (!ide.InsertCodeAtPageBottom(request->text, kAiApplyPreCompile)) {
				OutputStringToELog("[AI]追加到页底失败");
				return;
			}
			OutputStringToELog("[AI]追加到页底完成");
			return;

		default:
			return;
		}
	}
	catch (const std::exception& ex) {
		OutputStringToELog(std::string("[AI]执行结果异常：") + ex.what());
	}
	catch (...) {
		OutputStringToELog("[AI]执行结果发生未知异常");
	}
}

void RunAiAddByCurrentPageTypeTask()
{
	const uint64_t traceId = AllocateAIPerfTraceId();
	ScopedAIPerfTrace traceScope(traceId);
	const auto totalStart = PerfClock::now();
	auto logTotal = [&](const std::string& status) {
		LogAIPerfCost(
			traceId,
			"RunAiAddByCurrentPageTypeTask.total",
			ElapsedMs(totalStart),
			"status=" + status);
	};

	try {
		OutputStringToELog("[AI]开始执行：AI按当前页类型添加代码");
		if (!TryBeginAiTask()) {
			logTotal("busy");
			return;
		}

		IDEFacade& ide = IDEFacade::Instance();
		std::string currentPageCode;
		{
			const auto t0 = PerfClock::now();
			const bool ok = ide.GetCurrentPageCode(currentPageCode);
			LogAIPerfCost(
				traceId,
				"GetCurrentPageCode.total",
				ElapsedMs(t0),
				"ok=" + std::to_string(ok ? 1 : 0) + " scene=add_by_page");
			if (!ok) {
				OutputStringToELog("[AI]无法获取当前页代码");
				EndAiTask();
				logTotal("get_page_failed");
				return;
			}
		}
		if (currentPageCode.empty()) {
			OutputStringToELog("[AI]无法获取当前页代码");
			EndAiTask();
			logTotal("empty_page_code");
			return;
		}

		const IDEFacade::ActiveWindowType pageType = ide.GetActiveWindowType();
		const std::string pageTypeText = DescribeActivePageType(pageType);
		const std::string outputRule = BuildAddByPageTypeOutputRule(pageType);

		std::string requirement;
		const auto inputStart = PerfClock::now();
		const bool accepted = ShowAITextInputDialog(
			g_hwnd,
			"AI按当前页类型添加代码",
			"请输入添加需求（例如：数量、命名风格、用途、限制）：",
			requirement);
		LogAIPerfCost(
			traceId,
			"ShowAITextInputDialog.modal",
			ElapsedMs(inputStart),
			"ok=" + std::to_string(accepted ? 1 : 0) + " scene=add_by_page");
		if (!accepted) {
			OutputStringToELog("[AI]已取消输入添加需求");
			EndAiTask();
			logTotal("user_cancel_requirement");
			return;
		}
		requirement = TrimAsciiCopy(requirement);
		if (requirement.empty()) {
			OutputStringToELog("[AI]添加需求为空，已取消");
			EndAiTask();
			logTotal("empty_requirement");
			return;
		}

		AISettings settings = {};
		{
			const auto t0 = PerfClock::now();
			const bool ok = EnsureAISettingsReady(settings);
			LogAIPerfCost(
				traceId,
				"EnsureAISettingsReady.total",
				ElapsedMs(t0),
				"ok=" + std::to_string(ok ? 1 : 0) + " scene=add_by_page");
			if (!ok) {
				EndAiTask();
				logTotal("settings_not_ready");
				return;
			}
		}

		const std::string input =
			"当前页类型：" + pageTypeText + "\n"
			"按页类型输出规则：" + outputRule + "\n"
			"输出硬性要求：仅返回新增信息，不要返回整页、不解释。\n"
			"用户添加需求：\n" + requirement + "\n\n"
			"当前页完整代码：\n```e\n" + currentPageCode + "\n```";

		std::unique_ptr<AIAsyncRequest> request(new (std::nothrow) AIAsyncRequest());
		if (!request) {
			OutputStringToELog("[AI]内存不足，无法发起任务");
			EndAiTask();
			logTotal("alloc_request_failed");
			return;
		}
		request->traceId = traceId;
		request->action = AIAsyncUiAction::InsertAtPageBottom;
		request->taskKind = AITaskKind::AddByCurrentPageType;
		request->settings = settings;
		request->inputText = input;
		request->displayName = "AI按当前页类型添加代码";
		request->pageCodeSnapshot = currentPageCode;

		OutputStringToELog("[AI]正在请求模型（后台）...");
		uintptr_t threadId = _beginthread(RunAiTaskWorker, 0, request.get());
		if (threadId == static_cast<uintptr_t>(-1L)) {
			OutputStringToELog("[AI]启动后台任务失败");
			EndAiTask();
			logTotal("start_worker_failed");
			return;
		}
		request.release();
		logTotal("queued");
	}
	catch (const std::exception& ex) {
		OutputStringToELog(std::string("[AI]发生异常：") + ex.what());
		EndAiTask();
		logTotal("exception_std");
	}
	catch (...) {
		OutputStringToELog("[AI]发生未知异常");
		EndAiTask();
		logTotal("exception_unknown");
	}
}

void TryCopyCurrentFunctionCode()
{
	const uint64_t traceId = AllocateAIPerfTraceId();
	ScopedAIPerfTrace traceScope(traceId);
	const auto totalStart = PerfClock::now();
	const bool ok = IDEFacade::Instance().CopyCurrentFunctionCodeToClipboard();
	LogAIPerfCost(
		traceId,
		"TryCopyCurrentFunctionCode.total",
		ElapsedMs(totalStart),
		"ok=" + std::to_string(ok ? 1 : 0));

	if (ok) {
		OutputStringToELog("已复制当前函数代码到剪贴板");
	}
	else {
		OutputStringToELog("复制当前函数代码失败，当前位置可能不在代码函数中");
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
			if (itemTitle.find(L"编译") != std::wstring::npos ||
				itemTitle.find(L"发布") != std::wstring::npos ||
				itemTitle.find(L"静态") != std::wstring::npos ||
				itemTitle.find(L"Build") != std::wstring::npos ||
				itemTitle.find(L"Release") != std::wstring::npos) {
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

	// 顶级标题不可读时，回退到子项关键词判断。
	return popupKeywordMatch();
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
	AppendMenuW(hTargetMenu, MF_POPUP | MF_STRING, reinterpret_cast<UINT_PTR>(g_topLinkerSubMenu), L"链接器切换");
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
		EnsureTopLinkerSubMenuAttached(hMenu);
		RebuildTopLinkerSubMenu();
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
}


static auto originalCreateFileA = CreateFileA;
static auto originalGetSaveFileNameA = GetSaveFileNameA;
static auto originalCreateProcessA = CreateProcessA;
static auto originalMessageBoxA = MessageBoxA;

//开始调试 5.71-5.95
typedef int(__thiscall* OriginalEStartDebugFuncType)(DWORD* thisPtr, int a2, int a3);
OriginalEStartDebugFuncType originalEStartDebugFunc = (OriginalEStartDebugFuncType)0x40A080; //int __thiscall sub_40A080(int this, int a2, int a3)

//开始编译 5.71-5.95
typedef int(__thiscall* OriginalEStartCompileFuncType)(DWORD* thisPtr, int a2);
OriginalEStartCompileFuncType originalEStartCompileFunc = (OriginalEStartCompileFuncType)0x40A9F1;  //int __thiscall sub_40A9F1(_DWORD *this, int a2)

int __fastcall MyEStartCompileFunc(DWORD* thisPtr, int dummy, int a2) {
	OutputStringToELog("编译开始#2");
	//ChangeVMProtectModel(true);
	RunChangeECOM(true);
	return originalEStartCompileFunc(thisPtr, a2);
}

int __fastcall MyEStartDebugFunc(DWORD* thisPtr, int dummy, int a2, int a3) {
	OutputStringToELog("调试开始#2");
	//ChangeVMProtectModel(false);
	RunChangeECOM(false);
	return originalEStartDebugFunc(thisPtr, a2, a3);
}

HANDLE WINAPI MyCreateFileA(
	LPCSTR lpFileName,
	DWORD dwDesiredAccess,
	DWORD dwShareMode,
	LPSECURITY_ATTRIBUTES lpSecurityAttributes,
	DWORD dwCreationDisposition,
	DWORD dwFlagsAndAttributes,
	HANDLE hTemplateFile) {
	if (std::string(lpFileName).find("\\Temp\\e_debug\\") != std::string::npos) {
		OutputStringToELog("结束预编译代码（调试）");
		g_preDebugging = false;
	}

	if (g_preDebugging) {
		
	}

	std::filesystem::path currentPath = GetBasePath();
	std::filesystem::path autoLinkerPath = currentPath / "tools" / "link.ini";
	if (autoLinkerPath.string() == std::string(lpFileName)) {
		g_preCompiling = false;

		//EC编译阶段结束
		auto linkName = g_configManager.getValue(g_nowOpenSourceFilePath);

		if (!linkName.empty() ) {
			auto linkConfig = g_linkerManager.getConfig(linkName);

			if (std::filesystem::exists(linkConfig.path)) {
				//切换路径
				OutputStringToELog("切换为Linker:" + linkConfig.name + " " + linkConfig.path);

				return originalCreateFileA(linkConfig.path.c_str(), dwDesiredAccess, dwShareMode, lpSecurityAttributes,
					dwCreationDisposition,
					dwFlagsAndAttributes,
					hTemplateFile);
			}
			else {
				OutputStringToELog("无法切换Linker，Linker文件不存在#1");
			}

		}
		else {
			OutputStringToELog("未设置此源文件的Linker，使用默认");
		}
	}


	return originalCreateFileA(lpFileName,dwDesiredAccess,dwShareMode,lpSecurityAttributes,
								dwCreationDisposition,
								dwFlagsAndAttributes,
								hTemplateFile);
}

BOOL APIENTRY MyGetSaveFileNameA(LPOPENFILENAMEA item) {
	if (g_preCompiling) {
		OutputStringToELog("结束预编译代码");
		g_preCompiling = false;
	}

	return originalGetSaveFileNameA(item);
}


BOOL WINAPI MyCreateProcessA(
	LPCSTR lpApplicationName,
	LPSTR lpCommandLine,
	LPSECURITY_ATTRIBUTES lpProcessAttributes,
	LPSECURITY_ATTRIBUTES lpThreadAttributes,
	BOOL bInheritHandles,
	DWORD dwCreationFlags,
	LPVOID lpEnvironment,
	LPCSTR lpCurrentDirectory,
	LPSTARTUPINFOA lpStartupInfo,
	LPPROCESS_INFORMATION lpProcessInformation
) {
	std::string commandLine = lpCommandLine;

	//检查加入提前链接
	auto krnlnPath = GetLinkerCommandKrnlnFileName(lpCommandLine);
	OutputStringToELog(krnlnPath);

	if (!krnlnPath.empty()) {
		//核心库代码优先使用黑月的，然后再使用核心库的
		//std::string bklib = "D:\\git\\KanAutoControls\\Release\\TestCore.lib";
		
		//将把这个Lib插入的链接器的前方，添加/FORCE强制链接，忽略重定义
		std::string libFilePath = std::format("{}\\AutoLinker\\ForceLinkLib.txt", GetBasePath());
		auto libList = ReadFileAndSplitLines(libFilePath);



		if (libList.size() > 0) {
			//和链接器关联
			auto currentLinkerName = g_configManager.getValue(g_nowOpenSourceFilePath);

			OutputStringToELog(std::format("当前指定的链接器：{}", currentLinkerName));

			std::string libCmd;
			for (const auto& line : libList) {
				auto lines = SplitStringTwo(line, '=');
				auto libPath = line;
				std::string linkerName;
				if (lines.size() == 2) {
					linkerName = lines[0];
					libPath = lines[1];
				}

				//OutputStringToELog(std::format("找到设定的强制链接Lib：{} -> {}", linkerName, libPath));

				if (!linkerName.empty()) {
					//要求必须指定Linker才可使用（包含名称既可）
					if (currentLinkerName.find(linkerName) != std::string::npos) {
						//可使用，link名称一致
						if (std::filesystem::exists(libPath)) {
							OutputStringToELog(std::format("强制链接Lib：{} -> {}", linkerName, libPath));
							if (!libCmd.empty()) {
								libCmd += " ";
							}
							libCmd += "\"" + libPath + "\"";
						}
					} else {
						
						OutputStringToELog(std::format("链接器{}不符合当前的链接器{}，不链接：{}", linkerName, currentLinkerName, libPath));
					}

				} else {
					if (std::filesystem::exists(libPath)) {
						OutputStringToELog(std::format("强制链接Lib：{}", libPath));
						if (!libCmd.empty()) {
							libCmd += " ";
						}
						libCmd += "\"" + libPath + "\"";
					}
				}


			}

			std::string newLibs = libCmd + " \"" + krnlnPath + "\" /FORCE";
			commandLine = ReplaceSubstring(commandLine, "\"" + krnlnPath + "\"", newLibs);


		}
	}

	auto outFileName = GetLinkerCommandOutFileName(lpCommandLine);
	if (!outFileName.empty()) {
		if (commandLine.find("/pdb:\"build.pdb\"") != std::string::npos) {
			//PDB更名为当前编译的程序的名字+.pdb，看起来更正规
			std::string newPdbCommand = std::format("/pdb:\"{}.pdb\"", outFileName);
			commandLine = ReplaceSubstring(commandLine, "/pdb:\"build.pdb\"", newPdbCommand);
		}
		if (commandLine.find("/map:\"build.map\"") != std::string::npos) {
			//Map更名
			std::string newPdbCommand = std::format("/map:\"{}.map\"", outFileName);
			commandLine = ReplaceSubstring(commandLine, "/map:\"build.map\"", newPdbCommand);
		}
	}
	OutputStringToELog("启动命令行：" + commandLine);
	return originalCreateProcessA(lpApplicationName, (char *)commandLine.c_str(), lpProcessAttributes, lpThreadAttributes, bInheritHandles, dwCreationFlags, lpEnvironment, lpCurrentDirectory, lpStartupInfo, lpProcessInformation);
}

int WINAPI MyMessageBoxA(
	HWND hWnd,
	LPCSTR lpText,
	LPCSTR lpCaption,
	UINT uType) {
	//自动返回确认编译
	if (std::string(lpCaption).find("链接器输出了大量错误或警告信息") != std::string::npos) {
		return IDNO;
	}
	if (std::string(lpText).find("链接器输出了大量错误或警告信息") != std::string::npos) {
		return IDNO;
	}
	return originalMessageBoxA(hWnd, lpText, lpCaption, uType);
}

void StartHookCreateFileA() {
	DetourRestoreAfterWith();
	DetourTransactionBegin();
	DetourUpdateThread(GetCurrentThread());
	DetourAttach(&(PVOID&)originalCreateFileA, MyCreateFileA);
	DetourAttach(&(PVOID&)originalGetSaveFileNameA, MyGetSaveFileNameA);
	DetourAttach(&(PVOID&)originalCreateProcessA, MyCreateProcessA);
	DetourAttach(&(PVOID&)originalMessageBoxA, MyMessageBoxA);


	if (g_debugStartAddress != -1 && g_compileStartAddress != -1) {
		originalEStartCompileFunc = (OriginalEStartCompileFuncType)g_compileStartAddress;
		originalEStartDebugFunc = (OriginalEStartDebugFuncType)g_debugStartAddress;

		//用于自动在编译和调试之间切换Lib与Dll模块
		DetourAttach(&(PVOID&)originalEStartCompileFunc, MyEStartCompileFunc);
		DetourAttach(&(PVOID&)originalEStartDebugFunc, MyEStartDebugFunc);
	}
	else {
		//无法启用
	}
	DetourTransactionCommit();
}



/// <summary>
/// 输出文本
/// </summary>
/// <param name="szbuf"></param>
void OutputStringToELog(const std::string& szbuf) {
	const std::string line = "[AutoLinker]" + szbuf;
	OutputDebugStringA((line + "\n").c_str());
	IDEFacade::Instance().AppendOutputWindowLine(line);
}

uint64_t AllocateAIPerfTraceId()
{
	return g_aiPerfTraceSeed.fetch_add(1);
}

void SetCurrentAIPerfTraceId(uint64_t traceId)
{
	g_aiPerfTraceIdTLS = traceId;
}

uint64_t GetCurrentAIPerfTraceId()
{
	return g_aiPerfTraceIdTLS;
}

bool IsAIPerfLogEnabled()
{
	const std::string raw = ToLowerAsciiCopy(TrimAsciiCopy(g_configManager.getValue("ai.perf_log_enabled")));
	if (raw.empty()) {
		return false;
	}
	return raw == "1" || raw == "true" || raw == "on" || raw == "yes";
}

int GetAIPerfLogThresholdMs()
{
	const std::string raw = TrimAsciiCopy(g_configManager.getValue("ai.perf_log_threshold_ms"));
	if (raw.empty()) {
		return 30;
	}
	try {
		const int parsed = std::stoi(raw);
		return (std::max)(0, (std::min)(parsed, 600000));
	}
	catch (...) {
		return 30;
	}
}

bool IsAICodeFetchDebugEnabled()
{
	const std::string raw = ToLowerAsciiCopy(TrimAsciiCopy(g_configManager.getValue("ai.code_fetch_debug_log_enabled")));
	if (raw.empty()) {
		return IsAIPerfLogEnabled();
	}
	return raw == "1" || raw == "true" || raw == "on" || raw == "yes";
}

void LogAIPerfCost(uint64_t traceId, const std::string& step, long long costMs, const std::string& extra, bool force)
{
	if (!IsAIPerfLogEnabled()) {
		return;
	}
	if (traceId == 0) {
		traceId = GetCurrentAIPerfTraceId();
	}
	if (traceId == 0) {
		return;
	}
	if (!force && costMs < static_cast<long long>(GetAIPerfLogThresholdMs())) {
		return;
	}

	std::string line = std::format("[AI-PERF] trace={} step={} cost={}ms", traceId, step, costMs);
	if (!extra.empty()) {
		line += " ";
		line += extra;
	}
	OutputStringToELog(line);
}


void UpdateCurrentOpenSourceFile() {
	std::string sourceFile = GetSourceFilePath();

	if (sourceFile.empty()) {
		//OutputStringToELog("无法获取源文件路径");
	}

	if (g_nowOpenSourceFilePath != sourceFile) {
		OutputStringToELog(sourceFile);
	}
	g_nowOpenSourceFilePath = sourceFile;
}

void OutputCurrentSourceLinker()
{
	UpdateCurrentOpenSourceFile();

	std::string linkerName = "默认";
	if (!g_nowOpenSourceFilePath.empty()) {
		std::string configured = g_configManager.getValue(g_nowOpenSourceFilePath);
		if (!configured.empty()) {
			linkerName = configured;
		}
	}

	OutputStringToELog("当前源码链接器：" + linkerName);
}

//工具条子类过程
LRESULT CALLBACK ToolbarSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
	switch (uMsg)
	{
		case WM_INITMENUPOPUP:
			HandleInitMenuPopup(reinterpret_cast<HMENU>(wParam));
			break;
	}

	return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}


/// <summary>
/// 废弃，使用新的配置来切换了
/// </summary>
/// <param name="isLib"></param>
void ChangeVMProtectModel(bool isLib) {
	if (isLib) {
		
		int sdk = FindECOMNameIndex("VMPSDK");
		if (sdk != -1) {
			RemoveECOM(sdk); //移除
		}
		int sdkLib = FindECOMNameIndex("VMPSDK_LIB");
		if (sdkLib == -1) {
			OutputStringToELog("切换到静态VMP模块");
			char buffer[MAX_PATH] = { 0 };
			NotifySys(NAS_GET_PATH, 1004, (DWORD)buffer);
			std::string cmd = std::format("{}VMPSDK_LIB.ec", buffer);

			OutputStringToELog(cmd);
			AddECOM2(cmd);
			PeekAllMessage();
		}
	}
	else {
		int sdk = FindECOMNameIndex("VMPSDK_LIB");
		if (sdk != -1) {
			RemoveECOM(sdk); //移除
		}
		int sdkLib = FindECOMNameIndex("VMPSDK");
		if (sdkLib == -1) {
			OutputStringToELog("切换到动态VMP模块");
			char buffer[MAX_PATH] = { 0 };
			NotifySys(NAS_GET_PATH, 1004, (DWORD)buffer);
			std::string cmd = std::format("{}VMPSDK.ec", buffer);
			OutputStringToELog(cmd);
			AddECOM2(cmd);
			PeekAllMessage();
		}
	}
	//
}

#define WM_AUTOLINKER_INIT (WM_USER + 1000)

bool FneInit();

//主窗口子类过程
LRESULT CALLBACK MainWindowSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
	//if (uMsg == 20707) {
	//	//此消息配合AutoLinkerBuild，已废弃！
	//	if (wParam) {
	//		g_preCompiling = true;
	//		OutputStringToELog("开始编译");
	//		//ChangeVMPModel(true);
	//	}
	//	else {
	//		g_preDebugging = true;
	//		OutputStringToELog("开始调试");
	//		//ChangeVMPModel(false);
	//	}
	//	return 0;
	//}

	if (uMsg == WM_AUTOLINKER_AI_TASK_DONE) {
		HandleAiTaskCompletionMessage(lParam);
		return 0;
	}
	if (uMsg == WM_AUTOLINKER_AI_APPLY_RESULT) {
		HandleAiApplyMessage(lParam);
		return 0;
	}
	if (AIChatFeature::HandleMainWindowMessage(uMsg, wParam, lParam)) {
		return 0;
	}

	if (uMsg == WM_COMMAND) {
		UINT cmd = LOWORD(wParam);
		if (HandleTopLinkerMenuCommand(cmd)) {
			return 0;
		}
		if (IDEFacade::Instance().HandleMainWindowCommand(wParam)) {
			return 0;
		}
	}

	if (uMsg == WM_AUTOLINKER_INIT) {
		OutputStringToELog("收到初始化消息，尝试初始化");
		if (FneInit()) {
			OutputStringToELog("初始化成功");
			return 1;
		}
		return 0;
	}

	if (uMsg == 20708) {
		BOOL result = SetWindowSubclass((HWND)wParam, EditViewSubclassProc, 0, 0);
		return result ? 1 : 0;
	}

	if (uMsg == WM_INITMENUPOPUP) {
		LRESULT result = DefSubclassProc(hWnd, uMsg, wParam, lParam);
		IDEFacade::Instance().RefreshContextMenuEnabledState(reinterpret_cast<HMENU>(wParam));
		HandleInitMenuPopup(reinterpret_cast<HMENU>(wParam));
		return result;
	}

	return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

INT NESRUNFUNC(INT code, DWORD p1, DWORD p2) {
	return IDEFacade::Instance().RunFunctionRaw(code, p1, p2);
}

INT WINAPI fnAddInFunc(INT nAddInFnIndex) {

	switch (nAddInFnIndex) {
		case 0: { //TODO 打开项目目录
			std::string cmd = std::format("/select,{}", g_nowOpenSourceFilePath);
			ShellExecute(NULL, "open", "explorer.exe", cmd.c_str(), NULL, SW_SHOWDEFAULT);
			break;
		}
		case 1: { //Auto 目录
			std::string cmd = std::format("{}\\AutoLinker", GetBasePath());
			ShellExecute(NULL, "open", "explorer.exe", cmd.c_str(), NULL, SW_SHOWDEFAULT);
			break;
		}
		case 2: { //E 目录
			std::string cmd = std::format("{}", GetBasePath());
			ShellExecute(NULL, "open", "explorer.exe", cmd.c_str(), NULL, SW_SHOWDEFAULT);
			break;
		}
		case 3: { // 复制当前函数代码
			TryCopyCurrentFunctionCode();
			break;
		}
		case 4: { // AutoLinker AI接口设置
			AISettings settings = {};
			AIService::LoadSettings(g_configManager, settings);
			if (!ShowAIConfigDialog(g_hwnd, settings)) {
				OutputStringToELog("AI配置已取消");
				break;
			}
			AIService::SaveSettings(g_configManager, settings);
			OutputStringToELog("AI配置已保存");
			break;
		}
		case 5: { // FN_ADD_TAB 结构传递测试
			RunFnAddTabStructPassThroughTest();
			break;
		}
		case 6: { // 测试整体搜索 subWinHwnd
			RunDirectGlobalSearchKeywordTest();
			break;
		}
		case 7: { // 测试定位 subWinHwnd 首个结果
			RunDirectGlobalSearchLocateKeywordTest();
			break;
		}
		case 8: { // 测试定位后抓取当前页代码
			RunDirectGlobalSearchLocateAndDumpCurrentPageTest();
			break;
		}
		case 9: { // 测试枚举左侧 TreeView
			RunTreeViewProbeTest();
			break;
		}
		case 10: { // 测试程序树按名称直接抓页面代码
			RunProgramTreeDirectPageDumpTest();
			break;
		}
		case 11: { // 测试枚举程序树全部页面
			RunProgramTreeListTest();
			break;
		}
		case 12: { // 测试当前页窗口与页签
			RunCurrentPageWindowProbeTest();
			break;
		}
		case 13: { // 测试获取当前页名称
			RunCurrentPageNameTest();
			break;
		}
		//case 4: { //切换到VMPSDK静态（自用）
		//	ChangeVMProtectModel(true);
		//	break;
		//}
		//case 5: { //切换到VMPSDK动态（自用）
		//	ChangeVMProtectModel(false);
		//	break;
		//}
			  
		default: {

		}
	}

	return 0;
}


void FneCheckNewVersion(void* pParams) {
	Sleep(1000);

	OutputStringToELog("AutoLinker开源下载地址：https://github.com/aiqinxuancai/AutoLinker");
	std::string url = "https://api.github.com/repos/aiqinxuancai/AutoLinker/releases";
	//std::string customHeaders = "user-agent: Mozilla/5.0";
	auto response = PerformGetRequest(url);

	std::string currentVersion = AUTOLINKER_VERSION;


	if (response.second == 200) {
		std::string nowGithubVersion = "0.0.0";
		

		if (strcmp(AUTOLINKER_VERSION, "0.0.0") == 0) {
			//自行编译，无需检查版本更新
			OutputStringToELog(std::format("自编译版本，不检查更新，当前版本：{}", currentVersion));
		} else {
			if (!response.first.empty()) {

				try
				{
					auto releases = json::parse(response.first);
					for (const auto& release : releases) {
						if (!release["prerelease"].get<bool>()) {
							nowGithubVersion = release["tag_name"];
							break;
						}
					}

					Version nowGithubVersionObj(nowGithubVersion);
					Version currentVersionObj(AUTOLINKER_VERSION);

					if (nowGithubVersionObj > currentVersionObj) {
						OutputStringToELog(std::format("有新版本：{}", nowGithubVersion));
					}
					else {

					}
				}
				catch (const std::exception& e) {
					OutputStringToELog(std::format("检查新版本失败，当前版本：{} 错误：{}", currentVersion, e.what()));
				}


			}
		}

		
	}
	else {
		//OutputStringToELog(std::format("检查新版本失败，当前版本：{} 错误码：{}", currentVersion, response.second));
	}

	//return false;
}

bool FneInit() {
	OutputStringToELog("开始初始化");

	// g_hwnd 已经在外部获取并子类化
	if (g_hwnd == NULL) {
		OutputStringToELog("g_hwnd 为空");
		return false;
	}

	g_toolBarHwnd = FindMenuBar(g_hwnd);

	DWORD processID = GetCurrentProcessId();
	std::string s = std::format("E进程ID{} 主句柄{} 菜单栏句柄{}", processID, (int)g_hwnd, (int)g_toolBarHwnd);
	OutputStringToELog(s);

	if (g_toolBarHwnd != NULL)
	{
		StartEditViewSubclassTask();
		RebuildTopLinkerSubMenu();
		OutputCurrentSourceLinker();
		AIChatFeature::EnsureTabCreated();

		OutputStringToELog("找到工具条");
		SetWindowSubclass(g_toolBarHwnd, ToolbarSubclassProc, 0, 0);
		StartHookCreateFileA();
		PostAppMessageA(g_toolBarHwnd, WM_PRINT, 0, 0);
		OutputStringToELog("初始化完成");

		//初始化Lib相关库的状态

		//启动版本检查线程
		uintptr_t threadID = _beginthread(FneCheckNewVersion, 0, NULL);

		return true;
	}
	else
	{
		OutputStringToELog(std::format("初始化失败，未找到工具条窗口 {}", (int)g_toolBarHwnd));
	}

	return false;
}

/*-----------------支持库消息处理函数------------------*/

// 初始化重试线程函数
void InitRetryThread(void* pParams) {
	const int MAX_RETRY_COUNT = 5;
	int retryCount = 0;

	while (retryCount < MAX_RETRY_COUNT) {
		Sleep(1000); // 每秒重试一次
		retryCount++;

		OutputStringToELog(std::format("初始化重试第 {}/{} 次", retryCount, MAX_RETRY_COUNT));

		// 发送自定义消息到主窗口进行初始化
		LRESULT result = SendMessage(g_hwnd, WM_AUTOLINKER_INIT, 0, 0);

		if (result == 1) {
			// 初始化成功
			OutputStringToELog("初始化线程完成"); 
			return;
		}
	}

	OutputStringToELog(std::format("初始化失败，已重试 {} 次", MAX_RETRY_COUNT));
}

EXTERN_C INT WINAPI AutoLinker_MessageNotify(INT nMsg, DWORD dwParam1, DWORD dwParam2)
{
	std::string s = std::format("AutoLinker_MessageNotify {0} {1} {2}", (int)nMsg, dwParam1, dwParam2);
	OutputStringToELog(s);

#ifndef __E_STATIC_LIB
	if (nMsg == NL_GET_CMD_FUNC_NAMES) // 返回所有命令实现函数的的函数名称数组(char*[]), 支持静态编译的动态库必须处理
		return NULL;
	else if (nMsg == NL_GET_NOTIFY_LIB_FUNC_NAME) // 返回处理系统通知的函数名称(PFN_NOTIFY_LIB函数名称), 支持静态编译的动态库必须处理
		return (INT)LIBARAYNAME;
	else if (nMsg == NL_GET_DEPENDENT_LIBS) return (INT)NULL;
	//else if (nMsg == NL_GET_DEPENDENT_LIBS) return (INT)NULL;
	// 返回静态库所依赖的其它静态库文件名列表(格式为\0分隔的文本,结尾两个\0), 支持静态编译的动态库必须处理
	// kernel32.lib user32.lib gdi32.lib 等常用的系统库不需要放在此列表中
	// 返回NULL或NR_ERR表示不指定依赖文件  

	else if (nMsg == NL_SYS_NOTIFY_FUNCTION) {
		if (dwParam1) {
			if (!g_initStarted) {
				// 获取主窗口句柄
				g_hwnd = GetMainWindowByProcessId();

				if (g_hwnd) {
					g_initStarted = true;
					SetWindowSubclass(g_hwnd, MainWindowSubclassProc, 0, 0);
					AIChatFeature::Initialize(g_hwnd, &g_configManager);
					RegisterIDEContextMenu();
					OutputStringToELog("主窗口子类化完成，启动初始化线程");
					uintptr_t threadID = _beginthread(InitRetryThread, 0, NULL);
				} else {
					OutputStringToELog("无法获取主窗口句柄");
				}
			}

		}
	}
	else if (nMsg == NL_RIGHT_POPUP_MENU_SHOW) {
		IDEFacade::Instance().HandleNotifyMessage(nMsg, dwParam1, dwParam2);
	}

#endif
	return ProcessNotifyLib(nMsg, dwParam1, dwParam2);
}
/*定义支持库基本信息*/
#ifndef __E_STATIC_LIB
static LIB_INFOX LibInfo =
{
	/* { 库格式号, GUID串号, 主版本号, 次版本号, 构建版本号, 系统主版本号, 系统次版本号, 核心库主版本号, 核心库次版本号,
	支持库名, 支持库语言, 支持库描述, 支持库状态,
	作者姓名, 邮政编码, 通信地址, 电话号码, 传真号码, 电子邮箱, 主页地址, 其它信息,
	类型数量, 类型指针, 类别数量, 命令类别, 命令总数, 命令指针, 命令入口,
	附加功能, 功能描述, 消息指针, 超级模板, 模板描述,
	常量数量, 常量指针, 外部文件} */
	LIB_FORMAT_VER,
	_T(LIB_GUID_STR),
	LIB_MajorVersion,
	LIB_MinorVersion,
	LIB_BuildNumber,
	LIB_SysMajorVer,
	LIB_SysMinorVer,
	LIB_KrnlLibMajorVer,
	LIB_KrnlLibMinorVer,
	_T(LIB_NAME_STR),
	__GBK_LANG_VER,
	_WT(LIB_DESCRIPTION_STR),
	LBS_IDE_PLUGIN | LBS_LIB_INFO2, //_LIB_OS(__OS_WIN), //#LBS_IDE_PLUGIN  LBS_LIB_INFO2
	_WT(LIB_Author),
	_WT(LIB_ZipCode),
	_WT(LIB_Address),
	_WT(LIB_Phone), 
	_WT(LIB_Fax),
	_WT(LIB_Email),
	_WT(LIB_HomePage),
	_WT(LIB_Other),
	0,
	NULL,
	LIB_TYPE_COUNT,
	_WT(LIB_TYPE_STR),
	0,
	NULL,
	NULL,
	fnAddInFunc,
	_T("打开项目目录\0这是个用作测试的辅助工具功能。\0打开AutoLinker配置目录\0这是个用作测试的辅助工具功能。\0打开E语言目录\0这是个用作测试的辅助工具功能。\0复制当前函数代码\0复制当前光标所在子程序完整代码到剪贴板。\0AutoLinker AI接口设置\0编辑AI接口地址、API Key、模型和提示词等配置。\0FN_ADD_TAB结构传递测试\0构造ADD_TAB_INF调用FN_ADD_TAB，并打印调用前后结构体字段。\0测试整体搜索subWinHwnd\0调用direct_global_search固定搜索subWinHwnd，并输出命中结果到E输出窗口。\0测试定位subWinHwnd首个结果\0调用direct_global_search固定搜索subWinHwnd，并跳转到首个命中位置。\0测试定位后抓取当前页代码\0先定位到subWinHwnd首个命中，再抓取当前代码页完整代码并写入AutoLinker目录。\0测试枚举左侧TreeView\0枚举主窗口下所有SysTreeView32，并输出前几层节点文本与item data特征。\0测试程序树按名称抓代码\0在程序树中固定查找Class_HWND，并根据tree item data直接抓取整页代码。\0测试枚举程序树页面\0枚举程序树中所有页面节点，输出名称、类型和item data，并写入文件。\0测试当前页窗口与页签\0探测MDIClient当前活动子页与CCustomTabCtrl当前选中项文本，用于定位当前页名称来源。\0测试获取当前页名称\0调用IDEFacade当前页名称接口，输出当前页名称、类型和来源链路。\0\0") ,
	AutoLinker_MessageNotify,
	NULL,
	NULL,
	0,
	NULL,
	NULL,
	//-----------------
	NULL,
	0,
	NULL,
	"You",

};

PLIB_INFOX WINAPI GetNewInf()
{
	//LibInfo.m_szLicenseToUserName = "You"
	return (&LibInfo);
};
#endif
