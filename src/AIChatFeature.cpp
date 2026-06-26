#include "AIChatFeature.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <cctype>
#include <chrono>
#include <CommCtrl.h>
#include <condition_variable>
#include <cstring>
#include <ctime>
#include <filesystem>
#include <fstream>
#include <format>
#include <memory>
#include <mutex>
#include <new>
#include <process.h>
#include <Shellapi.h>
#include <string>
#include <unordered_map>
#include <unordered_set>
#include <vector>
#include <tlhelp32.h>
#include <wincrypt.h>
#include <wrl.h>

#include "..\\thirdparty\\json.hpp"
#include "..\\thirdparty\\WebView2.h"

#include "AIConfigDialog.h"
#include "AIJsonConfig.h"
#include "AIChatSessionStore.h"
#include "AIService.h"
#include "ConfigManager.h"
#include "AIChatTooling.h"
#include "AIChatToolingInternal.h"
#include "Global.h"
#include "IDEFacade.h"
#include "PathHelper.h"
#include "ProjectSourceCacheManager.h"
#include "ResourceTextLoader.h"
#include "WorkspaceMirror.h"
#include "WinINetUtil.h"
#include "resource.h"
#include "..\\elib\\lib2.h"
#if defined(_M_IX86)
#include "direct_global_search_debug.hpp"
#include "native_module_public_info.hpp"
#endif

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "advapi32.lib")
#pragma comment(lib, "shell32.lib")
#if defined _M_IX86
#pragma comment(linker, "\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_X64
#pragma comment(linker, "\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
#pragma comment(linker, "\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif

namespace {
constexpr UINT WM_AUTOLINKER_AI_CHAT_REFRESH = WM_APP + 203;
constexpr UINT WM_AUTOLINKER_AI_CHAT_DEFER_LAYOUT = WM_APP + 204;
constexpr UINT WM_AUTOLINKER_AI_CHAT_SUBMIT = WM_APP + 205;
constexpr UINT WM_AUTOLINKER_AI_CHAT_CLEAR = WM_APP + 206;
constexpr UINT WM_AUTOLINKER_AI_CHAT_STOP = WM_APP + 207;
constexpr UINT WM_AUTOLINKER_AI_CHAT_OPEN_SETTINGS = WM_APP + 208;
constexpr UINT WM_AUTOLINKER_AI_CHAT_RESTORE_LAST = WM_APP + 209;
constexpr UINT WM_AUTOLINKER_AI_CHAT_CLEAR_CONFIRMED = WM_APP + 210;
constexpr UINT WM_AUTOLINKER_AI_CHAT_CLEAR_CANCEL = WM_APP + 211;
constexpr UINT WM_AUTOLINKER_AI_CHAT_RESTORE_CONFIRMED = WM_APP + 212;
constexpr UINT WM_AUTOLINKER_AI_CHAT_RESTORE_CANCEL = WM_APP + 213;
constexpr UINT WM_AUTOLINKER_AI_CHAT_UPDATE_TAG = WM_APP + 214;
constexpr UINT_PTR kHistoryWebViewFlushTimerId = 0xA17;
constexpr UINT_PTR kSessionTimingTimerId = 0xA18;

constexpr int IDC_AI_CHAT_HISTORY = 2101;
constexpr int IDC_AI_CHAT_INPUT = 2102;
constexpr int IDC_AI_CHAT_SEND = 32553;
constexpr int IDC_AI_CHAT_CLEAR_HISTORY = 32554;
constexpr int IDC_AI_CHAT_STOP = 32555;
constexpr int IDC_AI_CHAT_OPEN_SETTINGS = 32556;
constexpr int IDC_AI_CHAT_RESTORE_SESSION = 32557;
constexpr int IDC_AI_CHAT_CLEAR_CONFIRM_TEXT = 32558;
constexpr int IDC_AI_CHAT_CLEAR_CONFIRM_APPLY = 32559;
constexpr int IDC_AI_CHAT_CLEAR_CONFIRM_CANCEL = 32560;
constexpr int IDC_AI_CHAT_MCP_GUIDE_LINK = 32561;
constexpr int IDC_AI_CHAT_CONTEXT_USAGE = 32562;
constexpr int IDC_AI_CHAT_SESSION_ELAPSED = 32563;
constexpr int IDC_AI_CHAT_SESSION_STATUS = 32564;

constexpr UINT_PTR kEditSubclassId = 1;
constexpr UINT_PTR kActionControlSubclassId = 2;
constexpr UINT_PTR kLeftWorkAreaHostSubclassId = 3;
constexpr int kLeftWorkAreaPageBottomInset = 4;
constexpr DWORD_PTR kEditFlagNone = 0;
constexpr DWORD_PTR kEditFlagSubmitOnEnter = 1;
constexpr UINT_PTR kActionSubmit = 1;
constexpr UINT_PTR kActionClear = 2;
constexpr UINT_PTR kActionStop = 3;
constexpr UINT_PTR kActionOpenSettings = 4;
constexpr UINT_PTR kActionRestoreSession = 5;
constexpr UINT_PTR kActionClearConfirm = 6;
constexpr UINT_PTR kActionClearCancel = 7;

constexpr const char* kChatMcpGuideUrl =
	"https://github.com/aiqinxuancai/AutoLinker/blob/master/CONFIG.md#%E5%A4%96%E9%83%A8-agent-mcp-%E9%85%8D%E7%BD%AE";
constexpr const char* kChatAgentWhitepaperUrl =
	"https://github.com/aiqinxuancai/Awesome-E-Agent";
constexpr const char* kChatHomeUrl =
	"https://github.com/aiqinxuancai/AutoLinker";
constexpr const char* kChatReleasesUrl =
	"https://github.com/aiqinxuancai/AutoLinker/releases";

enum class SessionRole {
	System,
	User,
	Assistant,
	Tool
};

struct SessionMessage {
	SessionRole role = SessionRole::System;
	std::string content;
	bool includeInContext = true;
	bool visibleInHistory = true;
	std::string reasoningContent;
	std::string rawMessageJsonUtf8;
};

struct AIChatRequestCancellation {
	std::atomic_bool cancelled = false;
	HttpRequestCancellation httpRequest;

	bool RequestCancel()
	{
		const bool wasCancelled = cancelled.exchange(true);
		httpRequest.Cancel();
		return !wasCancelled;
	}

	bool IsCancelled() const
	{
		return cancelled.load();
	}
};

struct AIChatSessionState {
	std::mutex mutex;
	std::vector<SessionMessage> messages;
	std::string rollingSummary;
	std::string streamingAssistantPreview;
	std::vector<std::string> agentActivityLines;
	std::string activeSessionId;
	std::filesystem::path activeSessionFilePath;
	std::string sourceFilePathLocal;
	std::string sourceFileNameLocal;
	long long createdAtUnixMs = 0;
	long long accumulatedElapsedMs = 0;
	long long activeRequestStartedAtUnixMs = 0;
	bool requestInFlight = false;
	unsigned long long activeRequestId = 0;
	unsigned long long nextRequestId = 1;
	std::shared_ptr<AIChatRequestCancellation> cancellation;
	int lastInputTokens = 0;          // 上一轮响应的真实输入 token（无则 0）
	bool hasLastUsage = false;        // 是否已有真实 usage
	int effectiveContextWindow = 0;   // 请求开始时按当前 settings 解析的上下文窗口
};

struct ChatDialogContext {
	HWND hHistory = nullptr;
	HWND hHistoryHost = nullptr;
	HWND hInput = nullptr;
	HWND hSend = nullptr;
	HWND hStop = nullptr;
	HWND hClearHistory = nullptr;
	HWND hRestoreSession = nullptr;
	HWND hOpenSettings = nullptr;
	HWND hClearConfirmText = nullptr;
	HWND hClearConfirmApply = nullptr;
	HWND hClearConfirmCancel = nullptr;
	HWND hMcpGuideLink = nullptr;
	HWND hSessionElapsed = nullptr;
	HWND hSessionStatus = nullptr;
	HWND hContextUsage = nullptr;
	int inputRowsVisible = 1;
	bool sessionTimingInProgress = false;
	bool sessionTimingVisible = false;
	bool clearConfirmVisible = false;
	bool restoreConfirmVisible = false;
	bool hasPendingRestoreSession = false;
	bool webViewDesired = false;
	bool webViewReady = false;
	bool webViewContentReady = false;
	bool webViewFlushScheduled = false;
	std::string pendingHistoryHtml;
	AIChatStoredSessionListEntry pendingRestoreSession;
	std::vector<AIChatStoredSessionListEntry> lastSessionMenuEntries;
	Microsoft::WRL::ComPtr<ICoreWebView2Environment> webViewEnvironment;
	Microsoft::WRL::ComPtr<ICoreWebView2Controller> webViewController;
	Microsoft::WRL::ComPtr<ICoreWebView2> webView;
};

struct AIChatAsyncRequest {
	unsigned long long requestId = 0;
	AISettings settings = {};
	std::vector<AIChatMessage> contextMessages;
	std::shared_ptr<AIChatRequestCancellation> cancellation;
};

struct AIChatAsyncResult {
	unsigned long long requestId = 0;
	AIChatResult chatResult = {};
};

struct ContextUsageSnapshot {
	bool available = false;
	bool estimated = true;
	int percent = 0;
	long long usedTokens = 0;
	int contextWindow = 0;
	std::string labelLocal;
	std::string titleLocal;
};

struct SessionTimingSnapshot {
	bool visible = false;
	bool inProgress = false;
	long long elapsedMs = 0;
	std::string elapsedLabelLocal;
	std::string statusLabelLocal;
	std::string titleLocal;
};

struct ChatHistoryGarbage {
	std::vector<SessionMessage> messages;
	std::string rollingSummary;
};

HWND g_mainWindow = nullptr;
ConfigManager* g_configManager = nullptr;
AIJsonConfig* g_aiJsonConfig = nullptr;
HWND g_chatDialog = nullptr;
bool g_chatTabAdded = false;
std::atomic_bool g_updateAvailable = false;
enum class ChatHostMode {
	None,
	LeftWorkArea,
	OutputTab
};
struct LeftWorkAreaHostState {
	HWND tabHwnd = nullptr;
	HWND hostHwnd = nullptr;
	int tabIndex = -1;
	int imageIndex = -1;
	bool pageVisible = false;
	bool subclassInstalled = false;
	std::vector<HWND> hiddenNativeChildren;
};
ChatHostMode g_chatHostMode = ChatHostMode::None;
LeftWorkAreaHostState g_leftWorkAreaHost;
AIChatSessionState g_session;
UINT g_msgAIChatDone = 0;
UINT g_msgAIChatToolDialog = 0;
UINT g_msgAIChatToolExec = 0;
std::atomic_bool g_clearHistoryInProgress = false;
bool g_webView2RuntimeChecked = false;
bool g_webView2RuntimeAvailable = false;

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

void RefreshChatDialog(HWND hWnd);
void RequestClearChatHistoryAsync();
void HandleChatSubmitUi(HWND hWnd, ChatDialogContext* ctx, const std::string& text);
void HandleChatClearUi(HWND hWnd, ChatDialogContext* ctx);
void HandleChatClearConfirmedUi(HWND hWnd, ChatDialogContext* ctx);
void HandleChatClearCancelUi(HWND hWnd, ChatDialogContext* ctx);
void HandleChatStopUi(HWND hWnd, ChatDialogContext* ctx);
void HandleChatOpenSettingsUi(HWND hWnd, ChatDialogContext* ctx);
void HandleChatRestoreSessionUi(HWND hWnd, ChatDialogContext* ctx);
void HandleChatShowRecentSessionsUi(HWND hWnd, ChatDialogContext* ctx);
void HandleChatRestoreSessionByIdUi(HWND hWnd, ChatDialogContext* ctx, const std::string& sessionId);
void HandleChatRestoreSessionConfirmedUi(HWND hWnd, ChatDialogContext* ctx);
void HandleChatRestoreSessionCancelUi(HWND hWnd, ChatDialogContext* ctx);
std::vector<AIChatStoredSessionListEntry> LoadRecentChatSessionsForCurrentSource();
bool BeginRestoreStoredChatSessionUi(HWND hWnd, ChatDialogContext* ctx, const AIChatStoredSessionListEntry& selected);
bool RestoreStoredChatSessionEntry(HWND hWnd, const AIChatStoredSessionListEntry& selected);
bool TryPromptAndRestoreRecentChatSession(HWND hWnd);
bool IsStopRequestedLocked(const AIChatSessionState& state);
std::string DescribeMissingChatSettingField(const std::string& missingField);
bool QueryChatSettingsState(AISettings& outSettings, std::string* outMissingField = nullptr);
void PostRefreshDialog();
void AppendAgentActivity(unsigned long long requestId, const std::string& line);
std::wstring WideFromUtf8Text(const std::string& text);

std::string GetWindowTextCopyLocalA(HWND hWnd)
{
	char buffer[512] = {};
	if (hWnd != nullptr && IsWindow(hWnd)) {
		GetWindowTextA(hWnd, buffer, static_cast<int>(sizeof(buffer)));
	}
	return buffer;
}

std::string GetWindowClassCopyLocalA(HWND hWnd)
{
	char buffer[128] = {};
	if (hWnd != nullptr && IsWindow(hWnd)) {
		GetClassNameA(hWnd, buffer, static_cast<int>(sizeof(buffer)));
	}
	return buffer;
}

BOOL CALLBACK EnumChildProcCollectClass(HWND hWnd, LPARAM lParam)
{
	auto* pair = reinterpret_cast<std::pair<const char*, std::vector<HWND>*>*>(lParam);
	if (pair == nullptr || pair->first == nullptr || pair->second == nullptr) {
		return TRUE;
	}

	const std::string className = GetWindowClassCopyLocalA(hWnd);
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
	EnumChildWindows(root, EnumChildProcCollectClass, reinterpret_cast<LPARAM>(&ctx));
	return windows;
}

std::string LocalFromWide(const wchar_t* text);

std::string ToLowerAsciiCopySimple(std::string text)
{
	for (char& ch : text) {
		ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
	}
	return text;
}

std::string ReadTabItemTextLocal(HWND tabHwnd, int index)
{
	if (tabHwnd == nullptr || !IsWindow(tabHwnd) || index < 0) {
		return std::string();
	}

	wchar_t wideBuf[256] = {};
	TCITEMW wideItem = {};
	wideItem.mask = TCIF_TEXT;
	wideItem.pszText = wideBuf;
	wideItem.cchTextMax = static_cast<int>(sizeof(wideBuf) / sizeof(wideBuf[0]));
	if (SendMessageW(tabHwnd, TCM_GETITEMW, static_cast<WPARAM>(index), reinterpret_cast<LPARAM>(&wideItem)) != FALSE) {
		return LocalFromWide(wideBuf);
	}

	char textBuf[256] = {};
	TCITEMA ansiItem = {};
	ansiItem.mask = TCIF_TEXT;
	ansiItem.pszText = textBuf;
	ansiItem.cchTextMax = static_cast<int>(sizeof(textBuf));
	if (SendMessageA(tabHwnd, TCM_GETITEMA, static_cast<WPARAM>(index), reinterpret_cast<LPARAM>(&ansiItem)) == FALSE) {
		return std::string();
	}
	return textBuf;
}

bool IsKnownEideModuleStem(const std::string& stemLower)
{
	static constexpr std::array<const char*, 7> kKnownEideModules = {
		"idraw",
		"icontrols",
		"iwindowsize",
		"iconfig",
		"ievent",
		"iresource",
		"itheme"
	};
	return std::find(kKnownEideModules.begin(), kKnownEideModules.end(), stemLower) != kKnownEideModules.end();
}

bool DetectEideCompatibilityEnvironment()
{
	HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE | TH32CS_SNAPMODULE32, GetCurrentProcessId());
	if (snapshot == INVALID_HANDLE_VALUE) {
		return false;
	}

	MODULEENTRY32 entry = {};
	entry.dwSize = sizeof(entry);
	if (Module32First(snapshot, &entry) == FALSE) {
		CloseHandle(snapshot);
		return false;
	}

	bool found = false;
	do {
		const std::filesystem::path modulePath(entry.szExePath);
		const std::string fullPathLower = ToLowerAsciiCopySimple(modulePath.string());
		const std::string stemLower = ToLowerAsciiCopySimple(modulePath.stem().string());
		if (!IsKnownEideModuleStem(stemLower)) {
			continue;
		}
		if (fullPathLower.find("\\eide\\") == std::string::npos &&
			fullPathLower.find("/eide/") == std::string::npos) {
			continue;
		}
		found = true;
		break;
	} while (Module32Next(snapshot, &entry) != FALSE);

	CloseHandle(snapshot);
	return found;
}

bool IsEideCompatibilityEnvironment()
{
	static bool s_checked = false;
	static bool s_detected = false;
	if (!s_checked) {
		s_detected = DetectEideCompatibilityEnvironment();
		s_checked = true;
		if (s_detected) {
			OutputStringToELog("[AI Chat][Compat] detected eide modules, left work area host disabled");
		}
	}
	return s_detected;
}

int GetTabDirectionLocal(HWND tabHwnd)
{
	if (tabHwnd == nullptr || !IsWindow(tabHwnd)) {
		return 1;
	}

	const DWORD style = static_cast<DWORD>(GetWindowLongPtrW(tabHwnd, GWL_STYLE));
	const bool isBottom = (style & TCS_BOTTOM) == TCS_BOTTOM;
	int direction = (style & TCS_VERTICAL) == TCS_VERTICAL ? 1 : 0;
	if (direction != 0) {
		direction = isBottom ? 2 : 0;
	}
	else {
		direction = isBottom ? 3 : 1;
	}
	return direction;
}

RECT CalcTabPageRectLocal(HWND tabHwnd)
{
	RECT rc = {};
	if (tabHwnd == nullptr || !IsWindow(tabHwnd)) {
		return rc;
	}

	RECT rcTabControl = {};
	GetClientRect(tabHwnd, &rcTabControl);

	RECT rcItem = {};
	if (TabCtrl_GetItemRect(tabHwnd, 0, &rcItem) == FALSE) {
		rc = rcTabControl;
		TabCtrl_AdjustRect(tabHwnd, FALSE, &rc);
		return rc;
	}

	const int cxClient = rcTabControl.right - rcTabControl.left;
	const int cyClient = rcTabControl.bottom - rcTabControl.top;
	switch (GetTabDirectionLocal(tabHwnd))
	{
	case 0:
		rc.left = rcItem.right;
		rc.top = 0;
		rc.right = cxClient - 1;
		rc.bottom = cyClient - 1;
		break;
	case 1:
		rc.left = 1;
		rc.top = rcItem.bottom - 1;
		rc.right = cxClient - 1;
		rc.bottom = cyClient - 1;
		break;
	case 2:
		rc.left = 1;
		rc.top = 0;
		rc.right = rcItem.left;
		rc.bottom = cyClient - 1;
		break;
	default:
		rc.left = 1;
		rc.top = 0;
		rc.right = cxClient - 1;
		rc.bottom = rcItem.top;
		break;
	}

	if (rc.right < rc.left) {
		rc.right = rc.left;
	}
	if (rc.bottom < rc.top) {
		rc.bottom = rc.top;
	}
	return rc;
}

HMODULE GetCurrentModuleHandle()
{
	HMODULE module = nullptr;
	GetModuleHandleExW(
		GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
		reinterpret_cast<LPCWSTR>(&GetCurrentModuleHandle),
		&module);
	return module;
}

HICON LoadAppIconHandle(int cx, int cy)
{
	const HMODULE module = GetCurrentModuleHandle();
	if (module == nullptr) {
		return nullptr;
	}
	return reinterpret_cast<HICON>(LoadImageA(
		module,
		MAKEINTRESOURCEA(IDI_APP_ICON),
		IMAGE_ICON,
		cx,
		cy,
		LR_DEFAULTCOLOR));
}

HICON GetAppIconLarge()
{
	static HICON s_largeIcon = LoadAppIconHandle(GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON));
	return s_largeIcon;
}

HICON GetAppIconSmall()
{
	static HICON s_smallIcon = LoadAppIconHandle(GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON));
	return s_smallIcon;
}

void ApplyWindowIcon(HWND hWnd)
{
	if (hWnd == nullptr) {
		return;
	}
	SendMessageA(hWnd, WM_SETICON, ICON_BIG, reinterpret_cast<LPARAM>(GetAppIconLarge()));
	SendMessageA(hWnd, WM_SETICON, ICON_SMALL, reinterpret_cast<LPARAM>(GetAppIconSmall()));
}

void EnsureWindowTitle(HWND hWnd, const std::string& fallbackTitle)
{
	if (hWnd == nullptr) {
		return;
	}
	if (GetWindowTextLengthA(hWnd) > 0) {
		return;
	}
	if (fallbackTitle.empty()) {
		SetWindowTextA(hWnd, "AutoLinker");
		return;
	}
	SetWindowTextA(hWnd, fallbackTitle.c_str());
}

class ComCtl6ActivationScope {
public:
	ComCtl6ActivationScope()
	{
		const HMODULE module = GetCurrentModuleHandle();
		if (module == nullptr) {
			return;
		}

		ACTCTXW act = {};
		act.cbSize = sizeof(act);
		act.dwFlags = ACTCTX_FLAG_HMODULE_VALID | ACTCTX_FLAG_RESOURCE_NAME_VALID;
		act.hModule = module;
		act.lpResourceName = MAKEINTRESOURCEW(2);

		m_actCtx = CreateActCtxW(&act);
		if (m_actCtx == INVALID_HANDLE_VALUE) {
			m_actCtx = nullptr;
			return;
		}

		ULONG_PTR localCookie = 0;
		if (ActivateActCtx(m_actCtx, &localCookie) != FALSE) {
			m_cookie = localCookie;
			m_active = true;
		}
	}

	~ComCtl6ActivationScope()
	{
		if (m_active) {
			DeactivateActCtx(0, m_cookie);
		}
		if (m_actCtx != nullptr && m_actCtx != INVALID_HANDLE_VALUE) {
			ReleaseActCtx(m_actCtx);
		}
	}

	ComCtl6ActivationScope(const ComCtl6ActivationScope&) = delete;
	ComCtl6ActivationScope& operator=(const ComCtl6ActivationScope&) = delete;

private:
	HANDLE m_actCtx = nullptr;
	ULONG_PTR m_cookie = 0;
	bool m_active = false;
};

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

std::string NormalizeCodeForEIDE(const std::string& text)
{
	std::string normalized;
	normalized.reserve(text.size() + 16);
	for (size_t i = 0; i < text.size(); ++i) {
		const char ch = text[i];
		if (ch == '\r') {
			if (i + 1 < text.size() && text[i + 1] == '\n') {
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

std::string Utf8FromWide(const wchar_t* text)
{
	if (text == nullptr || *text == L'\0') {
		return std::string();
	}
	const int size = WideCharToMultiByte(CP_UTF8, 0, text, -1, nullptr, 0, nullptr, nullptr);
	if (size <= 0) {
		return std::string();
	}
	std::string out(static_cast<size_t>(size), '\0');
	if (WideCharToMultiByte(CP_UTF8, 0, text, -1, out.data(), size, nullptr, nullptr) <= 0) {
		return std::string();
	}
	if (!out.empty() && out.back() == '\0') {
		out.pop_back();
	}
	return out;
}

std::wstring WideFromLocal(const std::string& text)
{
	if (text.empty()) {
		return std::wstring();
	}
	const int size = MultiByteToWideChar(CP_ACP, 0, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (size <= 0) {
		return std::wstring();
	}
	std::wstring out(static_cast<size_t>(size), L'\0');
	if (MultiByteToWideChar(CP_ACP, 0, text.data(), static_cast<int>(text.size()), out.data(), size) <= 0) {
		return std::wstring();
	}
	return out;
}

std::wstring EscapeJsSingleQuotedWide(const std::wstring& text)
{
	std::wstring out;
	out.reserve(text.size() + 32);
	for (wchar_t ch : text) {
		switch (ch)
		{
		case L'\\': out += L"\\\\"; break;
		case L'\'': out += L"\\'"; break;
		case L'\r': out += L"\\r"; break;
		case L'\n': out += L"\\n"; break;
		case 0x2028: out += L"\\u2028"; break;
		case 0x2029: out += L"\\u2029"; break;
		default: out.push_back(ch); break;
		}
	}
	return out;
}

std::wstring EscapeJsDoubleQuotedWide(const std::wstring& text)
{
	std::wstring out;
	out.reserve(text.size() + 32);
	for (wchar_t ch : text) {
		switch (ch)
		{
		case L'\\': out += L"\\\\"; break;
		case L'"': out += L"\\\""; break;
		case L'\r': out += L"\\r"; break;
		case L'\n': out += L"\\n"; break;
		case 0x2028: out += L"\\u2028"; break;
		case 0x2029: out += L"\\u2029"; break;
		default: out.push_back(ch); break;
		}
	}
	return out;
}

std::string NormalizeNewlinesLf(const std::string& text)
{
	std::string normalized;
	normalized.reserve(text.size());
	for (size_t i = 0; i < text.size(); ++i) {
		const char ch = text[i];
		if (ch == '\r') {
			if (i + 1 < text.size() && text[i + 1] == '\n') {
				++i;
			}
			normalized.push_back('\n');
			continue;
		}
		normalized.push_back(ch);
	}
	return normalized;
}

std::string EscapeHtml(const std::string& text)
{
	std::string out;
	out.reserve(text.size() + 32);
	for (char ch : text) {
		switch (ch)
		{
		case '&': out += "&amp;"; break;
		case '<': out += "&lt;"; break;
		case '>': out += "&gt;"; break;
		case '"': out += "&quot;"; break;
		default: out.push_back(ch); break;
		}
	}
	return out;
}

std::string EscapeHtmlAttribute(const std::string& text)
{
	std::string out;
	out.reserve(text.size() + 32);
	for (char ch : text) {
		switch (ch)
		{
		case '&': out += "&amp;"; break;
		case '<': out += "&lt;"; break;
		case '>': out += "&gt;"; break;
		case '"': out += "&quot;"; break;
		case '\'': out += "&#39;"; break;
		default: out.push_back(ch); break;
		}
	}
	return out;
}

bool IsWebView2RuntimeAvailable()
{
	if (g_webView2RuntimeChecked) {
		return g_webView2RuntimeAvailable;
	}
	g_webView2RuntimeChecked = true;

	LPWSTR version = nullptr;
	const HRESULT hr = GetAvailableCoreWebView2BrowserVersionString(nullptr, &version);
	g_webView2RuntimeAvailable = SUCCEEDED(hr) && version != nullptr && version[0] != L'\0';
	if (version != nullptr) {
		CoTaskMemFree(version);
	}
	return g_webView2RuntimeAvailable;
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

long long GetCurrentUnixTimeMsForChat()
{
	const auto now = std::chrono::system_clock::now();
	return static_cast<long long>(
		std::chrono::duration_cast<std::chrono::milliseconds>(now.time_since_epoch()).count());
}

std::string FormatUnixTimeLocalForChat(long long unixMs)
{
	if (unixMs <= 0) {
		return std::string();
	}

	const time_t unixSeconds = static_cast<time_t>(unixMs / 1000);
	std::tm localTm = {};
	if (localtime_s(&localTm, &unixSeconds) != 0) {
		return std::string();
	}

	char buffer[64] = {};
	if (std::strftime(buffer, sizeof(buffer), "%Y-%m-%d %H:%M:%S", &localTm) == 0) {
		return std::string();
	}
	return buffer;
}

std::string FormatSessionElapsedLocal(long long elapsedMs)
{
	if (elapsedMs < 0) {
		elapsedMs = 0;
	}

	const long long totalSeconds = elapsedMs / 1000;
	const long long hours = totalSeconds / 3600;
	const long long minutes = (totalSeconds % 3600) / 60;
	const long long seconds = totalSeconds % 60;
	if (hours > 0) {
		return std::format("{}h {}m {}s", hours, minutes, seconds);
	}
	return std::format("{}m {}s", minutes, seconds);
}

long long CalculateSessionElapsedMsLocked(const AIChatSessionState& state, long long nowMs)
{
	long long elapsedMs = state.accumulatedElapsedMs;
	if (elapsedMs < 0) {
		elapsedMs = 0;
	}
	if (state.requestInFlight && state.activeRequestStartedAtUnixMs > 0 && nowMs > state.activeRequestStartedAtUnixMs) {
		elapsedMs += nowMs - state.activeRequestStartedAtUnixMs;
	}
	return elapsedMs;
}

void FinishActiveSessionTimingLocked(AIChatSessionState& state, long long nowMs)
{
	if (state.activeRequestStartedAtUnixMs <= 0) {
		return;
	}
	if (nowMs > state.activeRequestStartedAtUnixMs) {
		state.accumulatedElapsedMs += nowMs - state.activeRequestStartedAtUnixMs;
	}
	if (state.accumulatedElapsedMs < 0) {
		state.accumulatedElapsedMs = 0;
	}
	state.activeRequestStartedAtUnixMs = 0;
}

SessionTimingSnapshot BuildSessionTimingSnapshotLocked(const AIChatSessionState& state, long long nowMs)
{
	SessionTimingSnapshot snapshot = {};
	snapshot.visible = state.requestInFlight ||
		state.accumulatedElapsedMs > 0 ||
		!state.messages.empty() ||
		!TrimAsciiCopy(state.rollingSummary).empty() ||
		!TrimAsciiCopy(state.streamingAssistantPreview).empty();
	snapshot.inProgress = state.requestInFlight;
	snapshot.elapsedMs = CalculateSessionElapsedMsLocked(state, nowMs);
	snapshot.elapsedLabelLocal = FormatSessionElapsedLocal(snapshot.elapsedMs);
	snapshot.statusLabelLocal = snapshot.inProgress
		? LocalFromWide(L"\u5de5\u4f5c\u4e2d")
		: LocalFromWide(L"\u5df2\u5b8c\u6210");
	snapshot.titleLocal = snapshot.inProgress
		? LocalFromWide(L"\u4f1a\u8bdd\u6b63\u5728\u5de5\u4f5c\uff0c\u7528\u65f6\u6301\u7eed\u589e\u52a0")
		: LocalFromWide(L"\u4f1a\u8bdd\u5df2\u5b8c\u6210\u7684\u5b9e\u9645\u7528\u65f6");
	return snapshot;
}

ContextUsageSnapshot BuildContextUsageSnapshotLocked(const AIChatSessionState& state, int fallbackContextWindow)
{
	ContextUsageSnapshot snapshot = {};
	snapshot.contextWindow = state.effectiveContextWindow > 0 ? state.effectiveContextWindow : fallbackContextWindow;
	if (snapshot.contextWindow <= 0) {
		snapshot.labelLocal = LocalFromWide(L"\u4e0a\u4e0b\u6587 --");
		snapshot.titleLocal = LocalFromWide(L"\u6682\u65e0\u53ef\u7528\u7684\u4e0a\u4e0b\u6587\u7a97\u53e3\u914d\u7f6e");
		return snapshot;
	}

	size_t contextChars = state.rollingSummary.size();
	for (const auto& message : state.messages) {
		if (message.includeInContext) {
			contextChars += message.content.size();
		}
	}

	snapshot.available = true;
	snapshot.estimated = !state.hasLastUsage || state.lastInputTokens <= 0;
	snapshot.usedTokens = snapshot.estimated
		? static_cast<long long>(contextChars / 4)
		: static_cast<long long>(state.lastInputTokens);
	if (snapshot.usedTokens < 0) {
		snapshot.usedTokens = 0;
	}

	const long long rawPercent =
		(snapshot.usedTokens * 100 + snapshot.contextWindow / 2) / snapshot.contextWindow;
	snapshot.percent = static_cast<int>((std::min<long long>)(999, (std::max<long long>)(0, rawPercent)));
	snapshot.labelLocal = std::format(
		"{} {}%",
		LocalFromWide(L"\u4e0a\u4e0b\u6587"),
		snapshot.percent);
	snapshot.titleLocal = std::format(
		"{}{} / {} tokens",
		snapshot.estimated ? LocalFromWide(L"\u4f30\u7b97\uff1a") : std::string(),
		snapshot.usedTokens,
		snapshot.contextWindow);
	return snapshot;
}

std::string GetCurrentChatSourceFilePathLocal()
{
	try {
		if (TrimAsciiCopy(g_nowOpenSourceFilePath).empty()) {
			return std::string();
		}
		std::filesystem::path path(g_nowOpenSourceFilePath);
		path = path.lexically_normal();
		if (path.is_relative()) {
			path = std::filesystem::absolute(path);
		}
		return path.lexically_normal().string();
	}
	catch (...) {
		return TrimAsciiCopy(g_nowOpenSourceFilePath);
	}
}

std::string GetCurrentChatSourceFileNameLocal()
{
	const std::string path = GetCurrentChatSourceFilePathLocal();
	if (path.empty()) {
		return std::string();
	}
	try {
		return std::filesystem::path(path).filename().string();
	}
	catch (...) {
		return path;
	}
}

HFONT GetDialogFont()
{
	static HFONT s_dialogFont = nullptr;
	if (s_dialogFont != nullptr) {
		return s_dialogFont;
	}

	NONCLIENTMETRICSA ncm = {};
	ncm.cbSize = sizeof(ncm);
	if (SystemParametersInfoA(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0) != FALSE) {
		s_dialogFont = CreateFontIndirectA(&ncm.lfMessageFont);
	}
	if (s_dialogFont == nullptr) {
		s_dialogFont = reinterpret_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
	}
	return s_dialogFont;
}

void SetDefaultFont(HWND hWnd)
{
	if (hWnd != nullptr) {
		SendMessageA(hWnd, WM_SETFONT, reinterpret_cast<WPARAM>(GetDialogFont()), TRUE);
	}
}

AIChatStoredSession BuildStoredSessionFromLockedState(const AIChatSessionState& state)
{
	AIChatStoredSession stored = {};
	const long long nowMs = GetCurrentUnixTimeMsForChat();
	stored.schemaVersion = 1;
	stored.sessionId = state.activeSessionId;
	stored.sourceFileNameLocal = state.sourceFileNameLocal;
	stored.sourceFilePathHintLocal = state.sourceFilePathLocal;
	stored.createdAtUnixMs = state.createdAtUnixMs;
	stored.updatedAtUnixMs = nowMs;
	stored.elapsedMs = CalculateSessionElapsedMsLocked(state, nowMs);
	stored.createdAtDisplayLocal = FormatUnixTimeLocalForChat(stored.createdAtUnixMs);
	stored.updatedAtDisplayLocal = FormatUnixTimeLocalForChat(stored.updatedAtUnixMs);
	stored.rollingSummaryLocal = state.rollingSummary;
	stored.sessionFilePath = state.activeSessionFilePath;
	for (const auto& message : state.messages) {
		AIChatStoredMessage row = {};
		switch (message.role)
		{
		case SessionRole::User:
			row.role = "user";
			break;
		case SessionRole::Assistant:
			row.role = "assistant";
			break;
		case SessionRole::Tool:
			row.role = "tool";
			break;
		default:
			row.role = "system";
			break;
		}
		row.contentLocal = message.content;
		row.includeInContext = message.includeInContext;
		row.visibleInHistory = message.visibleInHistory;
		row.reasoningContentUtf8 = message.reasoningContent;
		row.rawMessageJsonUtf8 = message.rawMessageJsonUtf8;
		stored.messages.push_back(std::move(row));
	}
	return stored;
}

SessionRole ParseStoredMessageRole(const std::string& role)
{
	if (_stricmp(role.c_str(), "user") == 0) {
		return SessionRole::User;
	}
	if (_stricmp(role.c_str(), "assistant") == 0) {
		return SessionRole::Assistant;
	}
	if (_stricmp(role.c_str(), "tool") == 0) {
		return SessionRole::Tool;
	}
	return SessionRole::System;
}

void EnsureChatSessionBindingLocked(AIChatSessionState& state)
{
	const std::string currentSourcePath = GetCurrentChatSourceFilePathLocal();
	const std::string currentSourceName = GetCurrentChatSourceFileNameLocal();
	if (state.sourceFilePathLocal.empty()) {
		state.sourceFilePathLocal = currentSourcePath;
	}
	if (state.sourceFileNameLocal.empty()) {
		state.sourceFileNameLocal = currentSourceName;
	}
	if (!state.activeSessionId.empty()) {
		if (state.activeSessionFilePath.empty()) {
			state.activeSessionFilePath = ResolveAIChatSessionFilePath(state.sourceFilePathLocal, state.activeSessionId);
		}
		return;
	}

	state.activeSessionId = CreateAIChatSessionId();
	state.createdAtUnixMs = GetCurrentUnixTimeMsForChat();
	if (state.sourceFilePathLocal.empty()) {
		state.sourceFilePathLocal = currentSourcePath;
	}
	if (state.sourceFileNameLocal.empty()) {
		state.sourceFileNameLocal = currentSourceName;
	}
	state.activeSessionFilePath = ResolveAIChatSessionFilePath(state.sourceFilePathLocal, state.activeSessionId);
}

bool HasAnyChatHistoryLocked(const AIChatSessionState& state)
{
	return !state.messages.empty() ||
		!TrimAsciiCopy(state.rollingSummary).empty() ||
		!TrimAsciiCopy(state.streamingAssistantPreview).empty();
}

bool PersistChatSessionSnapshotLocked(AIChatSessionState& state)
{
	if (!HasAnyChatHistoryLocked(state)) {
		return true;
	}
	EnsureChatSessionBindingLocked(state);

	std::string saveError;
	const AIChatStoredSession stored = BuildStoredSessionFromLockedState(state);
	const bool ok = SaveAIChatStoredSession(stored, &saveError);
	if (!ok) {
		OutputStringToELog("[AI Chat][Session] save failed: " + saveError);
	}
	return ok;
}

void SaveChatSessionSnapshotNow()
{
	AIChatStoredSession stored = {};
	bool shouldSave = false;

	{
		std::lock_guard<std::mutex> guard(g_session.mutex);
		if (!HasAnyChatHistoryLocked(g_session)) {
			return;
		}
		EnsureChatSessionBindingLocked(g_session);
		stored = BuildStoredSessionFromLockedState(g_session);
		shouldSave = true;
	}
	if (!shouldSave) {
		return;
	}

	std::string saveError;
	if (!SaveAIChatStoredSession(stored, &saveError)) {
		OutputStringToELog("[AI Chat][Session] save failed: " + saveError);
	}
}

bool ReplaceChatSessionStateFromStoredSession(const AIChatStoredSession& stored)
{
	std::lock_guard<std::mutex> guard(g_session.mutex);
	if (g_session.requestInFlight) {
		return false;
	}

	g_session.messages.clear();
	g_session.rollingSummary = stored.rollingSummaryLocal;
	g_session.streamingAssistantPreview.clear();
	g_session.agentActivityLines.clear();
	g_session.activeSessionId = stored.sessionId;
	g_session.activeSessionFilePath = stored.sessionFilePath;
	g_session.sourceFilePathLocal = stored.sourceFilePathHintLocal.empty()
		? GetCurrentChatSourceFilePathLocal()
		: stored.sourceFilePathHintLocal;
	g_session.sourceFileNameLocal = stored.sourceFileNameLocal.empty()
		? GetCurrentChatSourceFileNameLocal()
		: stored.sourceFileNameLocal;
	g_session.createdAtUnixMs = stored.createdAtUnixMs;
	g_session.accumulatedElapsedMs = stored.elapsedMs > 0 ? stored.elapsedMs : 0;
	g_session.activeRequestStartedAtUnixMs = 0;
	g_session.requestInFlight = false;
	g_session.activeRequestId = 0;
	g_session.cancellation.reset();
	g_session.nextRequestId = 1;
	g_session.lastInputTokens = 0;
	g_session.hasLastUsage = false;
	g_session.effectiveContextWindow = 0;

	for (const auto& row : stored.messages) {
		g_session.messages.push_back(SessionMessage{
			ParseStoredMessageRole(row.role),
			row.contentLocal,
			row.includeInContext,
			row.visibleInHistory,
			row.reasoningContentUtf8,
			row.rawMessageJsonUtf8
		});
	}
	return true;
}

void ResetChatSessionBindingLocked(AIChatSessionState& state)
{
	state.activeSessionId.clear();
	state.activeSessionFilePath.clear();
	state.sourceFilePathLocal = GetCurrentChatSourceFilePathLocal();
	state.sourceFileNameLocal = GetCurrentChatSourceFileNameLocal();
	state.createdAtUnixMs = 0;
	state.accumulatedElapsedMs = 0;
	state.activeRequestStartedAtUnixMs = 0;
}

void RebindChatSessionToCurrentSourceIfNeeded()
{
	const std::string currentSourcePath = GetCurrentChatSourceFilePathLocal();
	std::string previousSourcePath;
	bool hadHistory = false;
	{
		std::lock_guard<std::mutex> guard(g_session.mutex);
		if (g_session.requestInFlight) {
			return;
		}
		if (_stricmp(g_session.sourceFilePathLocal.c_str(), currentSourcePath.c_str()) == 0) {
			if (g_session.sourceFileNameLocal.empty()) {
				g_session.sourceFileNameLocal = GetCurrentChatSourceFileNameLocal();
			}
			return;
		}
		previousSourcePath = g_session.sourceFilePathLocal;
		hadHistory = HasAnyChatHistoryLocked(g_session);
		if (hadHistory) {
			PersistChatSessionSnapshotLocked(g_session);
		}
		g_session.messages.clear();
		g_session.rollingSummary.clear();
		g_session.streamingAssistantPreview.clear();
		g_session.agentActivityLines.clear();
		ResetChatSessionBindingLocked(g_session);
	}
	WorkspaceMirror::ResetAndCleanup();
	if (!previousSourcePath.empty() || !currentSourcePath.empty()) {
		OutputStringToELog(std::format(
			"[AI Chat][Session] source rebind old={} new={} history_reset={}",
			previousSourcePath.empty() ? "<empty>" : previousSourcePath,
			currentSourcePath.empty() ? "<empty>" : currentSourcePath,
			hadHistory ? 1 : 0));
	}
	PostRefreshDialog();
}

void PrepareWorkspaceMirrorForChat(unsigned long long requestId)
{
	std::string error;
	if (WorkspaceMirror::EnsureMirrorFresh(error)) {
		OutputStringToELog("[WorkspaceMirror] chat workspace mirror prepared");
		AppendAgentActivity(requestId, LocalFromWide(L"\u5de5\u7a0b\u955c\u50cf\u51c6\u5907\u5b8c\u6210"));
		return;
	}
	if (!TrimAsciiCopy(error).empty()) {
		OutputStringToELog("[WorkspaceMirror] prepare chat workspace mirror failed: " + error);
		std::string displayError = TrimAsciiCopy(error);
		if (displayError.size() > 300) {
			displayError.resize(300);
			displayError += "...";
		}
		AppendAgentActivity(
			requestId,
			LocalFromWide(L"\u5de5\u7a0b\u955c\u50cf\u51c6\u5907\u5931\u8d25\uff0c\u540e\u7eed\u5de5\u5177\u4f1a\u6309\u9700\u91cd\u8bd5\uff1a") + displayError);
	}
}

void ScrollEditToBottom(HWND hEdit)
{
	if (hEdit == nullptr) {
		return;
	}
	const int len = GetWindowTextLengthA(hEdit);
	SendMessageA(hEdit, EM_SETSEL, static_cast<WPARAM>(len), static_cast<LPARAM>(len));
	SendMessageA(hEdit, EM_SCROLLCARET, 0, 0);
	SendMessageA(hEdit, WM_VSCROLL, SB_BOTTOM, 0);
}

LRESULT CALLBACK EditControlSubclassProc(
	HWND hWnd,
	UINT uMsg,
	WPARAM wParam,
	LPARAM lParam,
	UINT_PTR uIdSubclass,
	DWORD_PTR dwRefData)
{
	switch (uMsg)
	{
	case WM_KEYDOWN: {
		const bool ctrlDown = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
		const bool shiftDown = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
		if (ctrlDown && (wParam == 'A' || wParam == 'a')) {
			SendMessageA(hWnd, EM_SETSEL, 0, -1);
			return 0;
		}

		if ((dwRefData & kEditFlagSubmitOnEnter) != 0 && wParam == VK_RETURN) {
			if (ctrlDown || shiftDown) {
				SendMessageA(hWnd, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>("\r\n"));
				HWND hParent = GetParent(hWnd);
				if (hParent != nullptr) {
					PostMessageA(hParent, WM_AUTOLINKER_AI_CHAT_DEFER_LAYOUT, 0, 0);
				}
			}
			else {
				HWND hParent = GetParent(hWnd);
				if (hParent != nullptr) {
					PostMessageA(hParent, WM_AUTOLINKER_AI_CHAT_SUBMIT, 0, 0);
				}
			}
			return 0;
		}
		break;
	}
	case WM_NCDESTROY:
		RemoveWindowSubclass(hWnd, EditControlSubclassProc, uIdSubclass);
		break;
	default:
		break;
	}
	return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

void InstallEditHotkeys(HWND hEdit, DWORD_PTR flags)
{
	if (hEdit != nullptr) {
		SetWindowSubclass(hEdit, EditControlSubclassProc, kEditSubclassId, flags);
	}
}

void PostChatAction(HWND hWnd, UINT_PTR action)
{
	if (hWnd == nullptr) {
		return;
	}
	HWND hParent = GetParent(hWnd);
	if (hParent == nullptr) {
		return;
	}
	if (action == kActionClear) {
		OutputStringToELog("[AI Chat][UI] click action: new_session");
		PostMessageA(hParent, WM_AUTOLINKER_AI_CHAT_CLEAR_CONFIRMED, 0, 0);
	}
	else if (action == kActionClearConfirm) {
		auto* ctx = reinterpret_cast<ChatDialogContext*>(GetWindowLongPtrA(hParent, GWLP_USERDATA));
		if (ctx != nullptr && ctx->restoreConfirmVisible) {
			OutputStringToELog("[AI Chat][UI] click action: restore_confirmed");
			PostMessageA(hParent, WM_AUTOLINKER_AI_CHAT_RESTORE_CONFIRMED, 0, 0);
		}
		else {
			OutputStringToELog("[AI Chat][UI] click action: clear_confirmed");
			PostMessageA(hParent, WM_AUTOLINKER_AI_CHAT_CLEAR_CONFIRMED, 0, 0);
		}
	}
	else if (action == kActionClearCancel) {
		auto* ctx = reinterpret_cast<ChatDialogContext*>(GetWindowLongPtrA(hParent, GWLP_USERDATA));
		if (ctx != nullptr && ctx->restoreConfirmVisible) {
			OutputStringToELog("[AI Chat][UI] click action: restore_cancel");
			PostMessageA(hParent, WM_AUTOLINKER_AI_CHAT_RESTORE_CANCEL, 0, 0);
		}
		else {
			OutputStringToELog("[AI Chat][UI] click action: clear_cancel");
			PostMessageA(hParent, WM_AUTOLINKER_AI_CHAT_CLEAR_CANCEL, 0, 0);
		}
	}
	else if (action == kActionStop) {
		OutputStringToELog("[AI Chat][UI] click action: stop");
		PostMessageA(hParent, WM_AUTOLINKER_AI_CHAT_STOP, 0, 0);
	}
	else if (action == kActionOpenSettings) {
		OutputStringToELog("[AI Chat][UI] click action: open_settings");
		PostMessageA(hParent, WM_AUTOLINKER_AI_CHAT_OPEN_SETTINGS, 0, 0);
	}
	else if (action == kActionRestoreSession) {
		OutputStringToELog("[AI Chat][UI] click action: restore_session");
		PostMessageA(hParent, WM_AUTOLINKER_AI_CHAT_RESTORE_LAST, 0, 0);
	}
	else {
		OutputStringToELog("[AI Chat][UI] click action: submit");
		PostMessageA(hParent, WM_AUTOLINKER_AI_CHAT_SUBMIT, 0, 0);
	}
}

LRESULT CALLBACK ActionControlSubclassProc(
	HWND hWnd,
	UINT uMsg,
	WPARAM wParam,
	LPARAM lParam,
	UINT_PTR uIdSubclass,
	DWORD_PTR dwRefData)
{
	constexpr const char* kPropPressed = "AutoLinker.AIChat.ClickPressed";
	switch (uMsg)
	{
	case WM_LBUTTONDOWN:
		SetCapture(hWnd);
		SetPropA(hWnd, kPropPressed, reinterpret_cast<HANDLE>(1));
		return 0;
	case WM_LBUTTONUP: {
		const bool wasPressed = GetPropA(hWnd, kPropPressed) != nullptr;
		RemovePropA(hWnd, kPropPressed);
		if (GetCapture() == hWnd) {
			ReleaseCapture();
		}
		if (wasPressed) {
			RECT rc = {};
			GetClientRect(hWnd, &rc);
			const POINT pt = {
				static_cast<short>(LOWORD(lParam)),
				static_cast<short>(HIWORD(lParam))
			};
			if (PtInRect(&rc, pt) != FALSE && IsWindowEnabled(hWnd)) {
				PostChatAction(hWnd, static_cast<UINT_PTR>(dwRefData));
			}
		}
		return 0;
	}
	case WM_CANCELMODE:
	case WM_CAPTURECHANGED:
		RemovePropA(hWnd, kPropPressed);
		if (GetCapture() == hWnd) {
			ReleaseCapture();
		}
		return 0;
	case WM_SETCURSOR:
		SetCursor(LoadCursor(nullptr, IDC_HAND));
		return TRUE;
	case WM_NCDESTROY:
		RemovePropA(hWnd, kPropPressed);
		RemoveWindowSubclass(hWnd, ActionControlSubclassProc, uIdSubclass);
		break;
	default:
		break;
	}
	return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

void InstallChatActionControl(HWND hControl, UINT_PTR action)
{
	if (hControl != nullptr) {
		SetWindowSubclass(
			hControl,
			ActionControlSubclassProc,
			kActionControlSubclassId,
			static_cast<DWORD_PTR>(action));
	}
}

bool HasVisibleInlineConfirm(const ChatDialogContext* ctx)
{
	return ctx != nullptr && (ctx->clearConfirmVisible || ctx->restoreConfirmVisible);
}

void UpdateNativeInlineConfirmText(ChatDialogContext* ctx)
{
	if (ctx == nullptr) {
		return;
	}
	if (ctx->hClearConfirmText != nullptr) {
		SetWindowTextW(
			ctx->hClearConfirmText,
			ctx->restoreConfirmVisible
				? L"\u5f53\u524d\u5df2\u6709\u5bf9\u8bdd\uff0c\u6062\u590d\u5386\u53f2\u4f1a\u8bdd\u5c06\u8986\u76d6\u5f53\u524d\u5185\u5bb9\uff0c\u662f\u5426\u7ee7\u7eed\uff1f"
				: L"\u5f53\u524d AI \u5bf9\u8bdd\u5c06\u4fdd\u5b58\u5230\u5386\u53f2\uff0c\u5e76\u5f00\u542f\u65b0\u4f1a\u8bdd\u3002");
	}
	if (ctx->hClearConfirmApply != nullptr) {
		SetWindowTextW(
			ctx->hClearConfirmApply,
			ctx->restoreConfirmVisible ? L"\u8986\u76d6\u6062\u590d" : L"\u65b0\u5efa");
	}
}

void ApplyMinTrackSize(HWND hWnd, MINMAXINFO* mmi, int minClientWidth, int minClientHeight)
{
	if (hWnd == nullptr || mmi == nullptr) {
		return;
	}

	RECT rc = { 0, 0, minClientWidth, minClientHeight };
	const DWORD style = static_cast<DWORD>(GetWindowLongPtrA(hWnd, GWL_STYLE));
	const DWORD exStyle = static_cast<DWORD>(GetWindowLongPtrA(hWnd, GWL_EXSTYLE));
	AdjustWindowRectEx(&rc, style, FALSE, exStyle);
	mmi->ptMinTrackSize.x = rc.right - rc.left;
	mmi->ptMinTrackSize.y = rc.bottom - rc.top;
}

void LayoutAIChatDialog(HWND hWnd, ChatDialogContext* ctx)
{
	if (hWnd == nullptr || ctx == nullptr) {
		return;
	}

	RECT rc = {};
	GetClientRect(hWnd, &rc);
	const int clientWidth = rc.right - rc.left;
	const int clientHeight = rc.bottom - rc.top;

	if (ctx->webViewContentReady && ctx->hHistoryHost != nullptr) {
		ctx->clearConfirmVisible = false;
		ctx->restoreConfirmVisible = false;
		if (ctx->hHistory != nullptr) {
			MoveWindow(ctx->hHistory, 0, 0, (std::max)(120, clientWidth), (std::max)(80, clientHeight), TRUE);
		}
		MoveWindow(ctx->hHistoryHost, 0, 0, (std::max)(120, clientWidth), (std::max)(80, clientHeight), TRUE);
		if (ctx->webViewController != nullptr) {
			RECT webRc = { 0, 0, (std::max)(120, clientWidth), (std::max)(80, clientHeight) };
			ctx->webViewController->put_Bounds(webRc);
		}
		if (ctx->hClearHistory != nullptr) {
			ShowWindow(ctx->hClearHistory, SW_HIDE);
		}
		if (ctx->hRestoreSession != nullptr) {
			ShowWindow(ctx->hRestoreSession, SW_HIDE);
		}
		if (ctx->hOpenSettings != nullptr) {
			ShowWindow(ctx->hOpenSettings, SW_HIDE);
		}
		if (ctx->hMcpGuideLink != nullptr) {
			ShowWindow(ctx->hMcpGuideLink, SW_HIDE);
		}
		if (ctx->hSessionElapsed != nullptr) {
			ShowWindow(ctx->hSessionElapsed, SW_HIDE);
		}
		if (ctx->hSessionStatus != nullptr) {
			ShowWindow(ctx->hSessionStatus, SW_HIDE);
		}
		if (ctx->hContextUsage != nullptr) {
			ShowWindow(ctx->hContextUsage, SW_HIDE);
		}
		if (ctx->hClearConfirmText != nullptr) {
			ShowWindow(ctx->hClearConfirmText, SW_HIDE);
		}
		if (ctx->hClearConfirmApply != nullptr) {
			ShowWindow(ctx->hClearConfirmApply, SW_HIDE);
		}
		if (ctx->hClearConfirmCancel != nullptr) {
			ShowWindow(ctx->hClearConfirmCancel, SW_HIDE);
		}
		if (ctx->hInput != nullptr) {
			ShowWindow(ctx->hInput, SW_HIDE);
		}
		if (ctx->hSend != nullptr) {
			ShowWindow(ctx->hSend, SW_HIDE);
		}
		if (ctx->hStop != nullptr) {
			ShowWindow(ctx->hStop, SW_HIDE);
		}
		return;
	}

	const int margin = 0;
	const int bottomMargin = 4;
	const int gap = 6;
	const int actionRowHeight = 22;
	const int mcpGuideHeight = 36;
	const int mcpGuideGap = 2;
	const int inputHeightSingle = 30;
	const int inputHeightDouble = 54;
	const int sendWidth = 92;
	const int sessionElapsedWidth = 86;
	const int sessionStatusWidth = 54;
	const int contextUsageWidth = 82;
	const int clearHistoryWidth = 26;
	const int restoreHistoryWidth = 26;
	const int openSettingsWidth = 26;
	const int clearConfirmApplyWidth = ctx->restoreConfirmVisible ? 72 : 52;
	const int clearConfirmCancelWidth = 52;
	const int inputRowsVisible = ctx->inputRowsVisible >= 2 ? 2 : 1;
	const int inputHeight = inputRowsVisible >= 2 ? inputHeightDouble : inputHeightSingle;
	const bool showInlineConfirm = HasVisibleInlineConfirm(ctx);
	const bool showSessionTiming = ctx->sessionTimingVisible;

	const int contentWidth = (std::max)(120, clientWidth - margin * 2);
	const int sessionTimingReservedWidth = showSessionTiming
		? (sessionElapsedWidth + gap + sessionStatusWidth + gap)
		: 0;
	const int inputWidth = (std::max)(
		80,
		contentWidth - sessionTimingReservedWidth - contextUsageWidth - sendWidth - gap * 2);
	int nextInputSideX = margin + inputWidth + gap;
	const int sessionElapsedX = nextInputSideX;
	if (showSessionTiming) {
		nextInputSideX += sessionElapsedWidth + gap;
	}
	const int sessionStatusX = nextInputSideX;
	if (showSessionTiming) {
		nextInputSideX += sessionStatusWidth + gap;
	}
	const int contextUsageX = nextInputSideX;
	const int sendX = contextUsageX + contextUsageWidth + gap;
	const int inputY = clientHeight - bottomMargin - inputHeight;
	const int mcpGuideY = inputY - mcpGuideGap - mcpGuideHeight;
	const int actionRowY = mcpGuideY - gap - actionRowHeight;
	const int historyY = margin;
	const int historyHeight = (std::max)(80, actionRowY - gap - historyY);

	if (ctx->hHistory != nullptr) {
		MoveWindow(ctx->hHistory, margin, historyY, contentWidth, historyHeight, TRUE);
	}
	if (ctx->hHistoryHost != nullptr) {
		MoveWindow(ctx->hHistoryHost, margin, historyY, contentWidth, historyHeight, TRUE);
		if (ctx->webViewController != nullptr) {
			RECT rc = { 0, 0, contentWidth, historyHeight };
			ctx->webViewController->put_Bounds(rc);
		}
	}
	if (ctx->hClearHistory != nullptr) {
		ShowWindow(ctx->hClearHistory, showInlineConfirm ? SW_HIDE : SW_SHOW);
		MoveWindow(ctx->hClearHistory, margin, actionRowY, clearHistoryWidth, actionRowHeight, TRUE);
	}
	if (ctx->hRestoreSession != nullptr) {
		ShowWindow(ctx->hRestoreSession, showInlineConfirm ? SW_HIDE : SW_SHOW);
		MoveWindow(ctx->hRestoreSession, margin + clearHistoryWidth + gap, actionRowY, restoreHistoryWidth, actionRowHeight, TRUE);
	}
	if (ctx->hOpenSettings != nullptr) {
		ShowWindow(ctx->hOpenSettings, showInlineConfirm ? SW_HIDE : SW_SHOW);
		MoveWindow(ctx->hOpenSettings, margin + clearHistoryWidth + gap + restoreHistoryWidth + gap, actionRowY, openSettingsWidth, actionRowHeight, TRUE);
	}
	UpdateNativeInlineConfirmText(ctx);
	if (ctx->hClearConfirmText != nullptr) {
		const int textWidth = (std::max)(
			120,
			contentWidth - clearConfirmApplyWidth - clearConfirmCancelWidth - gap * 2);
		ShowWindow(ctx->hClearConfirmText, showInlineConfirm ? SW_SHOW : SW_HIDE);
		MoveWindow(ctx->hClearConfirmText, margin, actionRowY, textWidth, actionRowHeight, TRUE);
	}
	if (ctx->hClearConfirmApply != nullptr) {
		ShowWindow(ctx->hClearConfirmApply, showInlineConfirm ? SW_SHOW : SW_HIDE);
		MoveWindow(
			ctx->hClearConfirmApply,
			margin + contentWidth - clearConfirmApplyWidth - clearConfirmCancelWidth - gap,
			actionRowY,
			clearConfirmApplyWidth,
			actionRowHeight,
			TRUE);
	}
	if (ctx->hClearConfirmCancel != nullptr) {
		ShowWindow(ctx->hClearConfirmCancel, showInlineConfirm ? SW_SHOW : SW_HIDE);
		MoveWindow(
			ctx->hClearConfirmCancel,
			margin + contentWidth - clearConfirmCancelWidth,
			actionRowY,
			clearConfirmCancelWidth,
			actionRowHeight,
			TRUE);
	}
	if (ctx->hMcpGuideLink != nullptr) {
		ShowWindow(ctx->hMcpGuideLink, SW_SHOW);
		MoveWindow(ctx->hMcpGuideLink, margin, mcpGuideY, contentWidth, mcpGuideHeight, TRUE);
	}
	if (ctx->hInput != nullptr) {
		ShowWindow(ctx->hInput, SW_SHOW);
		MoveWindow(ctx->hInput, margin, inputY, inputWidth, inputHeight, TRUE);
	}
	if (ctx->hSessionElapsed != nullptr) {
		ShowWindow(ctx->hSessionElapsed, showSessionTiming ? SW_SHOW : SW_HIDE);
		MoveWindow(ctx->hSessionElapsed, sessionElapsedX, inputY, sessionElapsedWidth, inputHeight, TRUE);
	}
	if (ctx->hSessionStatus != nullptr) {
		ShowWindow(ctx->hSessionStatus, showSessionTiming ? SW_SHOW : SW_HIDE);
		MoveWindow(ctx->hSessionStatus, sessionStatusX, inputY, sessionStatusWidth, inputHeight, TRUE);
	}
	if (ctx->hContextUsage != nullptr) {
		ShowWindow(ctx->hContextUsage, SW_SHOW);
		MoveWindow(ctx->hContextUsage, contextUsageX, inputY, contextUsageWidth, inputHeight, TRUE);
	}
	if (ctx->hSend != nullptr) {
		ShowWindow(ctx->hSend, SW_SHOW);
		MoveWindow(ctx->hSend, sendX, inputY, sendWidth, inputHeight, TRUE);
	}
	if (ctx->hStop != nullptr) {
		MoveWindow(ctx->hStop, sendX, inputY, sendWidth, inputHeight, TRUE);
	}
}

void SyncHistoryPresentation(ChatDialogContext* ctx)
{
	if (ctx == nullptr) {
		return;
	}

	const bool showWebView = ctx->webViewReady && ctx->webViewContentReady && ctx->hHistoryHost != nullptr && IsWindow(ctx->hHistoryHost);
	if (ctx->webViewController != nullptr) {
		ctx->webViewController->put_IsVisible(showWebView ? TRUE : FALSE);
	}
	if (ctx->hHistory != nullptr && IsWindow(ctx->hHistory)) {
		ShowWindow(ctx->hHistory, showWebView ? SW_HIDE : SW_SHOW);
	}
	if (ctx->hHistoryHost != nullptr && IsWindow(ctx->hHistoryHost)) {
		ShowWindow(ctx->hHistoryHost, showWebView ? SW_SHOW : SW_HIDE);
	}
}

std::string BuildHistoryWebViewShellHtml();

void ExecuteWebViewScript(ChatDialogContext* ctx, const std::wstring& script)
{
	if (ctx == nullptr || !ctx->webViewReady || !ctx->webViewContentReady || ctx->webView == nullptr || script.empty()) {
		return;
	}
	ctx->webView->ExecuteScript(script.c_str(), nullptr);
}

void UpdateWebViewComposerState(ChatDialogContext* ctx, bool busy, bool stopRequested = false)
{
	if (ctx == nullptr) {
		return;
	}
	std::wstring script = L"window.autolinkerSetBusy(";
	script += busy ? L"true" : L"false";
	script += L",";
	script += stopRequested ? L"true" : L"false";
	script += L");";
	ExecuteWebViewScript(ctx, script);
}

void UpdateWebViewContextUsage(ChatDialogContext* ctx, const ContextUsageSnapshot& snapshot)
{
	if (ctx == nullptr) {
		return;
	}
	std::wstring script = L"window.autolinkerSetContextUsage(";
	script += snapshot.available ? L"true" : L"false";
	script += L",";
	script += std::to_wstring(snapshot.percent);
	script += L",\"";
	script += EscapeJsDoubleQuotedWide(WideFromLocal(snapshot.labelLocal));
	script += L"\",\"";
	script += EscapeJsDoubleQuotedWide(WideFromLocal(snapshot.titleLocal));
	script += L"\",";
	script += snapshot.estimated ? L"true" : L"false";
	script += L");";
	ExecuteWebViewScript(ctx, script);
}

void UpdateWebViewSessionTiming(ChatDialogContext* ctx, const SessionTimingSnapshot& snapshot)
{
	if (ctx == nullptr) {
		return;
	}
	std::wstring script = L"window.autolinkerSetSessionTiming(\"";
	script += EscapeJsDoubleQuotedWide(WideFromLocal(snapshot.elapsedLabelLocal));
	script += L"\",\"";
	script += EscapeJsDoubleQuotedWide(WideFromLocal(snapshot.statusLabelLocal));
	script += L"\",\"";
	script += EscapeJsDoubleQuotedWide(WideFromLocal(snapshot.titleLocal));
	script += L"\",";
	script += snapshot.inProgress ? L"true" : L"false";
	script += L",";
	script += snapshot.visible ? L"true" : L"false";
	script += L");";
	ExecuteWebViewScript(ctx, script);
}

void UpdateWebViewUpdateTag(ChatDialogContext* ctx)
{
	if (ctx == nullptr) {
		return;
	}
	std::wstring script = L"window.autolinkerSetUpdateTag(";
	script += g_updateAvailable.load() ? L"true" : L"false";
	script += L");";
	ExecuteWebViewScript(ctx, script);
}

void UpdateNativeContextUsage(ChatDialogContext* ctx, const ContextUsageSnapshot& snapshot)
{
	if (ctx == nullptr || ctx->hContextUsage == nullptr) {
		return;
	}
	const std::wstring label = WideFromLocal(snapshot.labelLocal.empty()
		? LocalFromWide(L"\u4e0a\u4e0b\u6587 --")
		: snapshot.labelLocal);
	SetWindowTextW(ctx->hContextUsage, label.c_str());
}

void UpdateNativeSessionTiming(ChatDialogContext* ctx, const SessionTimingSnapshot& snapshot)
{
	if (ctx == nullptr) {
		return;
	}
	ctx->sessionTimingVisible = snapshot.visible;
	ctx->sessionTimingInProgress = snapshot.inProgress;
	if (ctx->hSessionElapsed != nullptr) {
		SetWindowTextW(ctx->hSessionElapsed, WideFromLocal(snapshot.elapsedLabelLocal).c_str());
	}
	if (ctx->hSessionStatus != nullptr) {
		SetWindowTextW(ctx->hSessionStatus, WideFromLocal(snapshot.statusLabelLocal).c_str());
	}
}

void SyncSessionTimingTimer(HWND hWnd, bool inProgress)
{
	if (hWnd == nullptr || !IsWindow(hWnd)) {
		return;
	}
	if (inProgress) {
		SetTimer(hWnd, kSessionTimingTimerId, 1000, nullptr);
	}
	else {
		KillTimer(hWnd, kSessionTimingTimerId);
	}
}

void RefreshSessionTimingOnly(HWND hWnd, ChatDialogContext* ctx)
{
	if (ctx == nullptr) {
		return;
	}

	SessionTimingSnapshot snapshot;
	{
		std::lock_guard<std::mutex> guard(g_session.mutex);
		snapshot = BuildSessionTimingSnapshotLocked(g_session, GetCurrentUnixTimeMsForChat());
	}
	const bool timingVisibilityChanged = ctx->sessionTimingVisible != snapshot.visible;
	UpdateNativeSessionTiming(ctx, snapshot);
	UpdateWebViewSessionTiming(ctx, snapshot);
	SyncSessionTimingTimer(hWnd, snapshot.inProgress);
	if (timingVisibilityChanged) {
		LayoutAIChatDialog(hWnd, ctx);
	}
}

void ClearWebViewInput(ChatDialogContext* ctx)
{
	ExecuteWebViewScript(ctx, L"window.autolinkerClearInput();");
}

void FocusWebViewInput(ChatDialogContext* ctx)
{
	ExecuteWebViewScript(ctx, L"window.autolinkerFocusInput();");
}

bool IsAllowedChatExternalUrl(const std::string& url)
{
	return _stricmp(url.c_str(), kChatMcpGuideUrl) == 0 ||
		_stricmp(url.c_str(), kChatAgentWhitepaperUrl) == 0 ||
		_stricmp(url.c_str(), kChatHomeUrl) == 0 ||
		_stricmp(url.c_str(), kChatReleasesUrl) == 0;
}

void OpenChatExternalUrl(HWND owner, const std::string& url)
{
	if (!IsAllowedChatExternalUrl(url)) {
		return;
	}
	ShellExecuteA(owner != nullptr ? owner : g_mainWindow, "open", url.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
}

void FocusChatComposerInput(ChatDialogContext* ctx)
{
	if (ctx == nullptr) {
		return;
	}
	if (ctx->webViewDesired && ctx->webViewContentReady) {
		FocusWebViewInput(ctx);
	}
	else if (ctx->hInput != nullptr) {
		SetFocus(ctx->hInput);
	}
}

bool QueryClearChatHistoryAvailability(bool& needConfirm)
{
	needConfirm = false;
	std::lock_guard<std::mutex> guard(g_session.mutex);
	if (g_session.requestInFlight) {
		return false;
	}
	needConfirm = HasAnyChatHistoryLocked(g_session);
	return true;
}

bool QueryRestoreChatSessionAvailability(bool& needConfirm)
{
	needConfirm = false;
	std::lock_guard<std::mutex> guard(g_session.mutex);
	if (g_session.requestInFlight) {
		return false;
	}
	needConfirm = HasAnyChatHistoryLocked(g_session);
	return true;
}

void ClearPendingRestoreSession(ChatDialogContext* ctx)
{
	if (ctx == nullptr) {
		return;
	}
	ctx->hasPendingRestoreSession = false;
	ctx->pendingRestoreSession = AIChatStoredSessionListEntry();
}

void HideClearChatConfirmInPage(ChatDialogContext* ctx)
{
	if (ctx == nullptr) {
		return;
	}
	ctx->clearConfirmVisible = false;
	ExecuteWebViewScript(ctx, L"window.autolinkerHideClearConfirm();");
}

void HideRestoreChatConfirmInPage(ChatDialogContext* ctx, bool clearPending)
{
	if (ctx == nullptr) {
		return;
	}
	ctx->restoreConfirmVisible = false;
	ExecuteWebViewScript(ctx, L"window.autolinkerHideRestoreConfirm();");
	if (clearPending) {
		ClearPendingRestoreSession(ctx);
	}
}

void HideChatConfirmInPage(ChatDialogContext* ctx)
{
	HideClearChatConfirmInPage(ctx);
	HideRestoreChatConfirmInPage(ctx, true);
}

void ShowClearChatConfirmInPage(HWND hWnd, ChatDialogContext* ctx)
{
	if (ctx == nullptr) {
		return;
	}
	HideRestoreChatConfirmInPage(ctx, true);
	if (ctx->webViewDesired && ctx->webViewContentReady) {
		ExecuteWebViewScript(ctx, L"window.autolinkerShowClearConfirm();");
		return;
	}

	ctx->clearConfirmVisible = true;
	LayoutAIChatDialog(hWnd, ctx);
	if (ctx->hClearConfirmApply != nullptr) {
		SetFocus(ctx->hClearConfirmApply);
	}
}

void ShowRestoreChatConfirmInPage(HWND hWnd, ChatDialogContext* ctx)
{
	if (ctx == nullptr) {
		return;
	}
	HideClearChatConfirmInPage(ctx);
	if (ctx->webViewDesired && ctx->webViewContentReady) {
		ExecuteWebViewScript(ctx, L"window.autolinkerShowRestoreConfirm();");
		return;
	}

	ctx->restoreConfirmVisible = true;
	UpdateNativeInlineConfirmText(ctx);
	LayoutAIChatDialog(hWnd, ctx);
	if (ctx->hClearConfirmApply != nullptr) {
		SetFocus(ctx->hClearConfirmApply);
	}
}

void HideWebViewSessionMenu(ChatDialogContext* ctx)
{
	ExecuteWebViewScript(ctx, L"window.autolinkerHideSessionMenu();");
}

void ShowWebViewSessionMenu(ChatDialogContext* ctx, const std::vector<AIChatStoredSessionListEntry>& entries)
{
	if (ctx == nullptr) {
		return;
	}

	nlohmann::json payload = nlohmann::json::array();
	for (const auto& entry : entries) {
		payload.push_back({
			{"sessionId", entry.sessionId},
			{"title", LocalToUtf8Text(entry.titleLocal)},
			{"updatedAt", LocalToUtf8Text(entry.updatedAtDisplayLocal)}
		});
	}

	std::wstring script = L"window.autolinkerShowSessionMenu(";
	script += WideFromUtf8Text(payload.dump());
	script += L");";
	ExecuteWebViewScript(ctx, script);
}

const AIChatStoredSessionListEntry* FindSessionMenuEntryById(const ChatDialogContext* ctx, const std::string& sessionId)
{
	if (ctx == nullptr) {
		return nullptr;
	}
	for (const auto& entry : ctx->lastSessionMenuEntries) {
		if (_stricmp(entry.sessionId.c_str(), sessionId.c_str()) == 0) {
			return &entry;
		}
	}
	return nullptr;
}

void UpdateHistoryWebViewHtml(ChatDialogContext* ctx, const std::string& htmlLocal)
{
	if (ctx == nullptr) {
		return;
	}
	ctx->pendingHistoryHtml = htmlLocal;
	if (!ctx->webViewReady || ctx->webView == nullptr) {
		return;
	}
	ctx->webViewFlushScheduled = true;
	if (g_chatDialog != nullptr && IsWindow(g_chatDialog)) {
		SetTimer(g_chatDialog, kHistoryWebViewFlushTimerId, 80, nullptr);
	}
}

void FlushHistoryWebViewHtml(ChatDialogContext* ctx)
{
	if (ctx == nullptr || !ctx->webViewReady || ctx->webView == nullptr || !ctx->webViewContentReady) {
		return;
	}

	ctx->webViewFlushScheduled = false;
	const std::wstring htmlWide = WideFromLocal(ctx->pendingHistoryHtml);
	if (htmlWide.empty()) {
		return;
	}
	std::wstring script = L"window.autolinkerSetChatHtml('";
	script += EscapeJsSingleQuotedWide(htmlWide);
	script += L"');";
	ExecuteWebViewScript(ctx, script);
}

void TryInitializeHistoryWebView(HWND hWnd, ChatDialogContext* ctx)
{
	if (hWnd == nullptr || ctx == nullptr || ctx->hHistoryHost == nullptr || !IsWindow(ctx->hHistoryHost) || !ctx->webViewDesired) {
		return;
	}

	using Microsoft::WRL::Callback;
	const std::wstring webViewUserDataFolder = GetWebView2UserDataFolderPath();
	const HRESULT hr = CreateCoreWebView2EnvironmentWithOptions(
		nullptr,
		webViewUserDataFolder.empty() ? nullptr : webViewUserDataFolder.c_str(),
		nullptr,
		Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
			[hWnd](HRESULT envResult, ICoreWebView2Environment* environment) -> HRESULT {
				auto* currentCtx = reinterpret_cast<ChatDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
				if (currentCtx == nullptr || !IsWindow(hWnd)) {
					return S_OK;
				}
								if (FAILED(envResult) || environment == nullptr) {
									currentCtx->webViewDesired = false;
									SyncHistoryPresentation(currentCtx);
									LayoutAIChatDialog(hWnd, currentCtx);
									OutputStringToELog(std::format("[AI Chat][WebView2] create environment failed hr=0x{:08X}", static_cast<unsigned int>(envResult)));
									return S_OK;
								}

				currentCtx->webViewEnvironment = environment;
				return environment->CreateCoreWebView2Controller(
					currentCtx->hHistoryHost,
					Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
						[hWnd](HRESULT controllerResult, ICoreWebView2Controller* controller) -> HRESULT {
							auto* innerCtx = reinterpret_cast<ChatDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
							if (innerCtx == nullptr || !IsWindow(hWnd)) {
								return S_OK;
							}
							if (FAILED(controllerResult) || controller == nullptr) {
								innerCtx->webViewDesired = false;
								SyncHistoryPresentation(innerCtx);
								LayoutAIChatDialog(hWnd, innerCtx);
								OutputStringToELog(std::format("[AI Chat][WebView2] create controller failed hr=0x{:08X}", static_cast<unsigned int>(controllerResult)));
								return S_OK;
							}

							innerCtx->webViewController = controller;
							innerCtx->webViewController->get_CoreWebView2(&innerCtx->webView);
							if (innerCtx->webView != nullptr) {
								Microsoft::WRL::ComPtr<ICoreWebView2Settings> settings;
								if (SUCCEEDED(innerCtx->webView->get_Settings(&settings)) && settings != nullptr) {
									settings->put_IsStatusBarEnabled(FALSE);
									settings->put_AreDevToolsEnabled(FALSE);
									settings->put_IsZoomControlEnabled(FALSE);
								}
								innerCtx->webView->add_WebMessageReceived(
									Callback<ICoreWebView2WebMessageReceivedEventHandler>(
										[hWnd](ICoreWebView2* sender, ICoreWebView2WebMessageReceivedEventArgs* args) -> HRESULT {
											(void)sender;
											auto* msgCtx = reinterpret_cast<ChatDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
											if (msgCtx == nullptr || !IsWindow(hWnd) || args == nullptr) {
												return S_OK;
											}

											LPWSTR rawMessage = nullptr;
											if (FAILED(args->TryGetWebMessageAsString(&rawMessage)) || rawMessage == nullptr) {
												return S_OK;
											}

											const std::string utf8Message = Utf8FromWide(rawMessage);
											CoTaskMemFree(rawMessage);
											try {
												const auto payload = nlohmann::json::parse(utf8Message);
												const std::string action = payload.contains("action") && payload["action"].is_string()
													? payload["action"].get<std::string>()
													: std::string();
												if (action == "submit") {
													const std::string text = payload.contains("text") && payload["text"].is_string()
														? Utf8ToLocalText(payload["text"].get<std::string>())
														: std::string();
													HandleChatSubmitUi(hWnd, msgCtx, text);
												}
												else if (action == "new_session") {
													HandleChatClearConfirmedUi(hWnd, msgCtx);
												}
												else if (action == "clear") {
													HandleChatClearUi(hWnd, msgCtx);
												}
												else if (action == "clear_confirmed") {
													HandleChatClearConfirmedUi(hWnd, msgCtx);
												}
												else if (action == "clear_cancel") {
													HandleChatClearCancelUi(hWnd, msgCtx);
												}
												else if (action == "restore_confirmed") {
													HandleChatRestoreSessionConfirmedUi(hWnd, msgCtx);
												}
												else if (action == "restore_cancel") {
													HandleChatRestoreSessionCancelUi(hWnd, msgCtx);
												}
												else if (action == "show_sessions") {
													HandleChatShowRecentSessionsUi(hWnd, msgCtx);
												}
												else if (action == "restore_session") {
													const std::string sessionId = payload.contains("sessionId") && payload["sessionId"].is_string()
														? payload["sessionId"].get<std::string>()
														: std::string();
													HandleChatRestoreSessionByIdUi(hWnd, msgCtx, sessionId);
												}
												else if (action == "stop") {
													HandleChatStopUi(hWnd, msgCtx);
												}
												else if (action == "open_settings") {
													PostMessageA(hWnd, WM_AUTOLINKER_AI_CHAT_OPEN_SETTINGS, 0, 0);
												}
												else if (action == "open_url") {
													const std::string url = payload.contains("url") && payload["url"].is_string()
														? payload["url"].get<std::string>()
														: std::string();
													OpenChatExternalUrl(hWnd, url);
												}
											}
											catch (...) {
											}
											return S_OK;
										}).Get(),
									nullptr);
								innerCtx->webView->add_NavigationCompleted(
									Callback<ICoreWebView2NavigationCompletedEventHandler>(
										[hWnd](ICoreWebView2* sender, ICoreWebView2NavigationCompletedEventArgs* args) -> HRESULT {
											(void)sender;
											auto* navCtx = reinterpret_cast<ChatDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
											if (navCtx == nullptr || !IsWindow(hWnd) || args == nullptr) {
												return S_OK;
											}

											BOOL isSuccess = FALSE;
											args->get_IsSuccess(&isSuccess);
											COREWEBVIEW2_WEB_ERROR_STATUS webErrorStatus = COREWEBVIEW2_WEB_ERROR_STATUS_UNKNOWN;
											args->get_WebErrorStatus(&webErrorStatus);
											if (isSuccess == TRUE) {
												navCtx->webViewContentReady = true;
												SyncHistoryPresentation(navCtx);
												LayoutAIChatDialog(hWnd, navCtx);
												if (!navCtx->pendingHistoryHtml.empty()) {
													UpdateHistoryWebViewHtml(navCtx, navCtx->pendingHistoryHtml);
													FlushHistoryWebViewHtml(navCtx);
												}
												bool inFlight = false;
												bool stopRequested = false;
												ContextUsageSnapshot usageSnapshot;
												SessionTimingSnapshot timingSnapshot;
												{
													AISettings settings = {};
													const int fallbackContextWindow =
														QueryChatSettingsState(settings, nullptr)
															? AIService::ResolveContextWindowTokens(settings)
															: 200000;
													std::lock_guard<std::mutex> guard(g_session.mutex);
													inFlight = g_session.requestInFlight;
													stopRequested = IsStopRequestedLocked(g_session);
													usageSnapshot = BuildContextUsageSnapshotLocked(g_session, fallbackContextWindow);
													timingSnapshot = BuildSessionTimingSnapshotLocked(g_session, GetCurrentUnixTimeMsForChat());
												}
												UpdateWebViewComposerState(navCtx, inFlight, stopRequested);
												UpdateWebViewContextUsage(navCtx, usageSnapshot);
												UpdateWebViewSessionTiming(navCtx, timingSnapshot);
												UpdateWebViewUpdateTag(navCtx);
												FocusWebViewInput(navCtx);
												return S_OK;
											}

											if (webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_OPERATION_CANCELED ||
												webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_ABORTED ||
												webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_RESET) {
												OutputStringToELog(std::format(
													"[AI Chat][WebView2] navigation superseded errorStatus={}",
													static_cast<int>(webErrorStatus)));
												return S_OK;
											}

											navCtx->webViewContentReady = false;
											navCtx->webViewDesired = false;
											SyncHistoryPresentation(navCtx);
											LayoutAIChatDialog(hWnd, navCtx);
											OutputStringToELog(std::format(
												"[AI Chat][WebView2] navigation failed, fallback to edit errorStatus={}",
												static_cast<int>(webErrorStatus)));
											if (navCtx->hHistory != nullptr && IsWindow(navCtx->hHistory)) {
												SetWindowTextA(navCtx->hHistory, "WebView2 navigation failed, fallback to edit.\r\n");
											}
											return S_OK;
										}).Get(),
									nullptr);
							}

							RECT rc = {};
							GetClientRect(innerCtx->hHistoryHost, &rc);
							innerCtx->webViewController->put_Bounds(rc);
							innerCtx->webViewReady = innerCtx->webView != nullptr;
							innerCtx->webViewContentReady = false;
							SyncHistoryPresentation(innerCtx);
							if (innerCtx->webViewReady) {
								const std::wstring shellHtml = WideFromLocal(BuildHistoryWebViewShellHtml());
								if (!shellHtml.empty()) {
									innerCtx->webView->NavigateToString(shellHtml.c_str());
								}
							}
							return S_OK;
						}).Get());
			}).Get());

	if (FAILED(hr)) {
		ctx->webViewDesired = false;
		SyncHistoryPresentation(ctx);
		OutputStringToELog(std::format("[AI Chat][WebView2] bootstrap failed hr=0x{:08X}", static_cast<unsigned int>(hr)));
	}
}

std::string GetEditTextA(HWND hEdit)
{
	if (hEdit == nullptr) {
		return std::string();
	}
	const int len = GetWindowTextLengthA(hEdit);
	if (len <= 0) {
		return std::string();
	}

	std::string text(static_cast<size_t>(len + 1), '\0');
	GetWindowTextA(hEdit, &text[0], len + 1);
	text.resize(static_cast<size_t>(len));
	return text;
}

int CountInputTextLines(const std::string& text)
{
	if (text.empty()) {
		return 1;
	}

	int lines = 1;
	for (char ch : text) {
		if (ch == '\n') {
			++lines;
		}
	}
	return lines;
}

void UpdateInputRowsAndLayout(HWND hWnd, ChatDialogContext* ctx, bool forceLayout)
{
	if (hWnd == nullptr || ctx == nullptr || ctx->hInput == nullptr) {
		return;
	}

	const std::string text = GetEditTextA(ctx->hInput);
	const int lineCount = CountInputTextLines(text);
	const int targetRows = lineCount > 1 ? 2 : 1;
	if (!forceLayout && targetRows == ctx->inputRowsVisible) {
		return;
	}

	ctx->inputRowsVisible = targetRows;
	LayoutAIChatDialog(hWnd, ctx);
}

bool RunModalWindow(HWND owner, HWND hDialog)
{
	if (hDialog == nullptr) {
		return false;
	}
	if (owner != nullptr && IsWindow(owner)) {
		EnableWindow(owner, FALSE);
	}

	ShowWindow(hDialog, SW_SHOW);
	UpdateWindow(hDialog);

	MSG msg = {};
	while (IsWindow(hDialog)) {
		const BOOL ok = GetMessageA(&msg, nullptr, 0, 0);
		if (ok <= 0) {
			break;
		}
		if (!IsDialogMessageA(hDialog, &msg)) {
			TranslateMessage(&msg);
			DispatchMessageA(&msg);
		}
	}

	if (owner != nullptr && IsWindow(owner)) {
		EnableWindow(owner, TRUE);
		SetForegroundWindow(owner);
	}
	return true;
}

std::string RoleLabel(SessionRole role)
{
	switch (role)
	{
	case SessionRole::User:
		return LocalFromWide(L"\u7528\u6237");
	case SessionRole::Assistant:
		return "AI";
	case SessionRole::Tool:
		return LocalFromWide(L"\u5de5\u5177");
	default:
		return LocalFromWide(L"\u7cfb\u7edf");
	}
}

std::string RenderInlineMarkdownToHtml(const std::string& text)
{
	std::string out;
	out.reserve(text.size() + 32);
	for (size_t i = 0; i < text.size();) {
		if (text.compare(i, 2, "**") == 0) {
			const size_t end = text.find("**", i + 2);
			if (end != std::string::npos) {
				out += "<strong>";
				out += EscapeHtml(text.substr(i + 2, end - (i + 2)));
				out += "</strong>";
				i = end + 2;
				continue;
			}
		}
		if (text[i] == '`') {
			const size_t end = text.find('`', i + 1);
			if (end != std::string::npos) {
				out += "<code>";
				out += EscapeHtml(text.substr(i + 1, end - (i + 1)));
				out += "</code>";
				i = end + 1;
				continue;
			}
		}
		if (text[i] == '[') {
			const size_t mid = text.find("](", i + 1);
			const size_t end = mid == std::string::npos ? std::string::npos : text.find(')', mid + 2);
			if (mid != std::string::npos && end != std::string::npos) {
				out += "<a href=\"";
				out += EscapeHtmlAttribute(text.substr(mid + 2, end - (mid + 2)));
				out += "\">";
				out += EscapeHtml(text.substr(i + 1, mid - (i + 1)));
				out += "</a>";
				i = end + 1;
				continue;
			}
		}

		switch (text[i])
		{
		case '&': out += "&amp;"; break;
		case '<': out += "&lt;"; break;
		case '>': out += "&gt;"; break;
		case '"': out += "&quot;"; break;
		default: out.push_back(text[i]); break;
		}
		++i;
	}
	return out;
}

std::string RenderPlainTextToHtml(const std::string& text)
{
	const std::string normalized = NormalizeNewlinesLf(text);
	std::string out;
	out.reserve(normalized.size() + 32);
	for (char ch : normalized) {
		if (ch == '\n') {
			out += "<br/>";
			continue;
		}
		switch (ch)
		{
		case '&': out += "&amp;"; break;
		case '<': out += "&lt;"; break;
		case '>': out += "&gt;"; break;
		case '"': out += "&quot;"; break;
		default: out.push_back(ch); break;
		}
	}
	return out;
}

std::string RenderMarkdownToHtml(const std::string& markdown)
{
	const std::string normalized = NormalizeNewlinesLf(markdown);
	std::vector<std::string> lines;
	lines.reserve(64);
	size_t begin = 0;
	while (begin <= normalized.size()) {
		const size_t end = normalized.find('\n', begin);
		if (end == std::string::npos) {
			lines.push_back(normalized.substr(begin));
			break;
		}
		lines.push_back(normalized.substr(begin, end - begin));
		begin = end + 1;
	}

	std::string html;
	bool inCodeBlock = false;
	bool inList = false;
	bool inParagraph = false;
	bool firstParagraphLine = true;

	auto closeParagraph = [&]() {
		if (inParagraph) {
			html += "</p>";
			inParagraph = false;
			firstParagraphLine = true;
		}
	};
	auto closeList = [&]() {
		if (inList) {
			html += "</ul>";
			inList = false;
		}
	};
	auto openCodeBlock = [&]() {
		html += "<div class=\"code-block\"><pre><code>";
		inCodeBlock = true;
	};
	auto closeCodeBlock = [&]() {
		if (inCodeBlock) {
			html += "</code></pre></div>";
			inCodeBlock = false;
		}
	};

	for (const std::string& line : lines) {
		const std::string trimmed = TrimAsciiCopy(line);
		if (line.rfind("```", 0) == 0) {
			closeParagraph();
			closeList();
			if (!inCodeBlock) {
				openCodeBlock();
			}
			else {
				closeCodeBlock();
			}
			continue;
		}

		if (inCodeBlock) {
			html += EscapeHtml(line);
			html += "\n";
			continue;
		}

		if (trimmed.empty()) {
			closeParagraph();
			closeList();
			continue;
		}

		size_t headingLevel = 0;
		while (headingLevel < line.size() && line[headingLevel] == '#') {
			++headingLevel;
		}
		if (headingLevel > 0 && headingLevel <= 6 &&
			headingLevel < line.size() && line[headingLevel] == ' ') {
			closeParagraph();
			closeList();
			html += "<h" + std::to_string(headingLevel) + ">";
			html += RenderInlineMarkdownToHtml(TrimAsciiCopy(line.substr(headingLevel + 1)));
			html += "</h" + std::to_string(headingLevel) + ">";
			continue;
		}

		if (line.rfind("- ", 0) == 0 || line.rfind("* ", 0) == 0) {
			closeParagraph();
			if (!inList) {
				html += "<ul>";
				inList = true;
			}
			html += "<li>";
			html += RenderInlineMarkdownToHtml(line.substr(2));
			html += "</li>";
			continue;
		}

		closeList();
		if (!inParagraph) {
			html += "<p>";
			inParagraph = true;
			firstParagraphLine = true;
		}
		if (!firstParagraphLine) {
			html += "<br/>";
		}
		html += RenderInlineMarkdownToHtml(line);
		firstParagraphLine = false;
	}

	closeParagraph();
	closeList();
	closeCodeBlock();
	return html;
}

void CompactHistoryLocked(AIChatSessionState& state)
{
	auto cleanupInvisibleContextMessages = [&state]() {
		const size_t keepInvisibleCount = 24;
		size_t trailingInvisibleCount = 0;
		for (auto it = state.messages.rbegin(); it != state.messages.rend(); ++it) {
			if (it->visibleInHistory || it->includeInContext) {
				continue;
			}
			++trailingInvisibleCount;
			if (trailingInvisibleCount > keepInvisibleCount) {
				it->content.clear();
				it->reasoningContent.clear();
				it->rawMessageJsonUtf8.clear();
			}
		}
		state.messages.erase(
			std::remove_if(
				state.messages.begin(),
				state.messages.end(),
				[](const SessionMessage& msg) {
					return !msg.visibleInHistory &&
						!msg.includeInContext &&
						msg.content.empty() &&
						msg.reasoningContent.empty() &&
						msg.rawMessageJsonUtf8.empty();
				}),
			state.messages.end());
	};

	size_t contextChars = 0;
	size_t contextCount = 0;
	for (const auto& message : state.messages) {
		if (message.includeInContext) {
			contextChars += message.content.size();
			++contextCount;
		}
	}

	// 触发判定（移植自 openai/codex）：以「模型上下文窗口 × 95%」为可用上限，
	// 优先用上一轮的真实输入 token，无则按 字符数/4 估算；另保留宽松的字符数二级兜底。
	const int window = state.effectiveContextWindow > 0 ? state.effectiveContextWindow : 200000;
	const long long limit = static_cast<long long>(window) * 95 / 100;
	const long long currentTokens = state.hasLastUsage
		? static_cast<long long>(state.lastInputTokens)
		: static_cast<long long>(contextChars / 4);
	const bool overTokenLimit = currentTokens >= limit;
	const bool overCharSafetyNet = contextChars > 200000;

	if (!overTokenLimit && !overCharSafetyNet) {
		cleanupInvisibleContextMessages();
		return;
	}

	const size_t keepContextCount = (std::min)(contextCount, static_cast<size_t>(24));
	const size_t compactContextCount = contextCount - keepContextCount;
	if (compactContextCount == 0) {
		return;
	}

	std::string summaryAppend;
	summaryAppend.reserve(2048);
	size_t compactedCount = 0;
	for (auto& msg : state.messages) {
		if (!msg.includeInContext) {
			continue;
		}

		if (compactedCount < compactContextCount) {
			if (msg.visibleInHistory) {
				std::string line = TrimAsciiCopy(msg.content);
				if (line.size() > 120) {
					line.resize(120);
					line += "...";
				}
				if (summaryAppend.size() <= 4000) {
					summaryAppend += "[";
					summaryAppend += RoleLabel(msg.role);
					summaryAppend += "] ";
					summaryAppend += line;
					summaryAppend += "\n";
				}
			}
			msg.includeInContext = false;
			msg.reasoningContent.clear();
			msg.rawMessageJsonUtf8.clear();
			++compactedCount;
		}

		if (compactedCount >= compactContextCount) {
			break;
		}
	}

	if (!summaryAppend.empty()) {
		if (!state.rollingSummary.empty()) {
			state.rollingSummary += "\n";
		}
		state.rollingSummary += summaryAppend;
		if (state.rollingSummary.size() > 12000) {
			state.rollingSummary.erase(0, state.rollingSummary.size() - 12000);
		}
	}

	cleanupInvisibleContextMessages();
}

std::vector<AIChatMessage> BuildContextMessagesLocked(const AIChatSessionState& state)
{
	std::vector<AIChatMessage> out;
	out.reserve(40);

	if (!TrimAsciiCopy(state.rollingSummary).empty()) {
		out.push_back(AIChatMessage{
			"system",
			LocalFromWide(L"\u5386\u53f2\u6458\u8981\uff08\u81ea\u52a8\u538b\u7f29\uff09\uff1a\n") + state.rollingSummary,
			"",
			""
		});
	}

	std::vector<SessionMessage> contextMsgs;
	contextMsgs.reserve(state.messages.size());
	for (const auto& msg : state.messages) {
		if (msg.includeInContext) {
			contextMsgs.push_back(msg);
		}
	}

	const size_t keep = (std::min)(contextMsgs.size(), static_cast<size_t>(24));
	const size_t begin = contextMsgs.size() > keep ? (contextMsgs.size() - keep) : 0;
	for (size_t i = begin; i < contextMsgs.size(); ++i) {
		const auto& msg = contextMsgs[i];
		std::string role = "system";
		if (msg.role == SessionRole::User) {
			role = "user";
		}
		else if (msg.role == SessionRole::Assistant) {
			role = "assistant";
		}
		else if (msg.role == SessionRole::Tool) {
			role = "tool";
		}
		out.push_back(AIChatMessage{ role, msg.content, msg.reasoningContent, msg.rawMessageJsonUtf8 });
	}
	return out;
}

std::wstring WideFromUtf8Text(const std::string& text)
{
	if (text.empty()) {
		return std::wstring();
	}
	const int size = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (size <= 0) {
		return WideFromLocal(Utf8ToLocalText(text));
	}
	std::wstring out(static_cast<size_t>(size), L'\0');
	if (MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), out.data(), size) <= 0) {
		return WideFromLocal(Utf8ToLocalText(text));
	}
	return out;
}

bool IsStopRequestedLocked(const AIChatSessionState& state)
{
	return state.requestInFlight &&
		state.cancellation != nullptr &&
		state.cancellation->IsCancelled();
}

std::string BuildHistoryHtmlLocked(
	const AIChatSessionState& state,
	bool settingsReady,
	const std::string& missingField)
{
	auto appendMessageCard = [](std::string& html, SessionRole role, const std::string& roleText, const std::string& content, bool renderMarkdown) {
		std::string roleClass = "system";
		switch (role)
		{
		case SessionRole::User: roleClass = "user"; break;
		case SessionRole::Assistant: roleClass = "assistant"; break;
		case SessionRole::Tool: roleClass = "tool"; break;
		default: break;
		}

		html += "<section class=\"msg ";
		html += roleClass;
		html += "\"><div class=\"role\">";
		html += EscapeHtml(roleText);
		html += "</div><div class=\"body\">";
		html += renderMarkdown ? RenderMarkdownToHtml(content) : RenderPlainTextToHtml(content);
		html += "</div></section>";
	};
	auto appendRawCard = [](std::string& html, SessionRole role, const std::string& roleText, const std::string& bodyHtml) {
		std::string roleClass = "system";
		switch (role)
		{
		case SessionRole::User: roleClass = "user"; break;
		case SessionRole::Assistant: roleClass = "assistant"; break;
		case SessionRole::Tool: roleClass = "tool"; break;
		default: break;
		}

		html += "<section class=\"msg ";
		html += roleClass;
		html += " ai-config-callout\"><div class=\"role\">";
		html += EscapeHtml(roleText);
		html += "</div><div class=\"body\">";
		html += bodyHtml;
		html += "</div></section>";
	};

	std::string body;
	body.reserve(4096);
	if (!settingsReady) {
		std::string calloutHtml;
		calloutHtml += "<p>";
		calloutHtml += EscapeHtml(LocalFromWide(L"AI 配置未完成，当前无法发起对话。"));
		calloutHtml += "</p><p class=\"muted\">";
		calloutHtml += EscapeHtml(
			LocalFromWide(L"缺少必填项：") + DescribeMissingChatSettingField(missingField) +
			LocalFromWide(L"。请先完成 Key / 接口 / 模型配置。"));
		calloutHtml += "</p><div class=\"inline-actions\"><button class=\"btn primary ai-config-btn\" type=\"button\" data-chat-action=\"open_settings\">";
		calloutHtml += EscapeHtml(LocalFromWide(L"前往设置 Key"));
		calloutHtml += "</button></div>";
		appendRawCard(body, SessionRole::System, LocalFromWide(L"系统"), calloutHtml);
	}
	for (const auto& msg : state.messages) {
		if (!msg.visibleInHistory) {
			continue;
		}
		appendMessageCard(
			body,
			msg.role,
			RoleLabel(msg.role),
			msg.content,
			msg.role == SessionRole::Assistant);
	}

	if (state.requestInFlight) {
		const bool stopRequested = IsStopRequestedLocked(state);
		const std::string preview = TrimAsciiCopy(state.streamingAssistantPreview);
		std::string activityText;
		for (const std::string& line : state.agentActivityLines) {
			if (!activityText.empty()) {
				activityText += "\r\n";
			}
			activityText += line;
		}
		if (!preview.empty()) {
			appendMessageCard(body, SessionRole::Assistant, "AI", state.streamingAssistantPreview, true);
			appendMessageCard(
				body,
				SessionRole::System,
				LocalFromWide(L"系统"),
				stopRequested
					? LocalFromWide(L"已发出停止请求，等待当前步骤结束...")
					: (activityText.empty()
						? LocalFromWide(L"AI 正在生成（流式）...")
						: activityText),
				false);
		}
		else {
			appendMessageCard(
				body,
				SessionRole::System,
				LocalFromWide(L"系统"),
				stopRequested
					? LocalFromWide(L"正在中断 AI 请求...")
					: (activityText.empty()
						? LocalFromWide(L"等待 AI 返回...")
						: activityText),
				false);
		}
	}

	if (TrimAsciiCopy(body).empty()) {
		body += "<div class=\"empty-state\"><div class=\"empty-state-title\">";
		body += EscapeHtml(LocalFromWide(L"需要做点什么？"));
		body += "</div><div class=\"empty-state-links\">";
		body += "<a class=\"mcp-guide-link\" href=\"https://github.com/aiqinxuancai/AutoLinker/blob/master/CONFIG.md#%E5%A4%96%E9%83%A8-agent-mcp-%E9%85%8D%E7%BD%AE\">";
		body += EscapeHtml(LocalFromWide(L"Codex连接AutoLinker MCP"));
		body += "</a><a class=\"mcp-guide-link\" href=\"https://github.com/aiqinxuancai/Awesome-E-Agent\">";
		body += EscapeHtml(LocalFromWide(L"易语言 × AI Agent 实践白皮书"));
		body += "</a><a class=\"mcp-guide-link\" href=\"https://github.com/aiqinxuancai/AutoLinker\">";
		body += EscapeHtml(LocalFromWide(L"前往AutoLinker项目Github"));
		body += "</a></div></div>";
	}

	std::string html;
	return body;
}

std::string BuildHistoryWebViewShellHtml()
{
	std::string html = LoadUtf8HtmlResourceText(IDR_HTML_AI_CHAT_HISTORY);
	if (!html.empty()) {
		return html;
	}
	return "<!doctype html><html><head><meta charset=\"utf-8\"></head><body>WebView shell resource missing.</body></html>";
}

std::string BuildHistoryTextLocked(
	const AIChatSessionState& state,
	bool settingsReady,
	const std::string& missingField)
{
	std::string text;
	if (!settingsReady) {
		text += "[" + LocalFromWide(L"系统") + "]\r\n";
		text += LocalFromWide(L"AI 配置未完成，当前无法发起对话。") + "\r\n";
		text += LocalFromWide(L"缺少必填项：") + DescribeMissingChatSettingField(missingField) +
			LocalFromWide(L"\u3002\u8bf7\u70b9\u51fb\u8bbe\u7f6e\u6309\u94ae\u5b8c\u6210\u8bbe\u7f6e\u3002") + "\r\n\r\n";
	}
	for (const auto& msg : state.messages) {
		if (!msg.visibleInHistory) {
			continue;
		}
		text += "[";
		text += RoleLabel(msg.role);
		text += "]\r\n";
		text += msg.content;
		text += "\r\n\r\n";
	}
	if (state.requestInFlight) {
		const bool stopRequested = IsStopRequestedLocked(state);
		const std::string preview = TrimAsciiCopy(state.streamingAssistantPreview);
		std::string activityText;
		for (const std::string& line : state.agentActivityLines) {
			if (!activityText.empty()) {
				activityText += "\r\n";
			}
			activityText += line;
		}
		if (!preview.empty()) {
			text += "[AI]\r\n";
			text += state.streamingAssistantPreview;
			text += "\r\n\r\n";
			text += "[" + LocalFromWide(L"\u7cfb\u7edf") + "]\r\n"
				+ (stopRequested
					? LocalFromWide(L"\u5df2\u53d1\u51fa\u505c\u6b62\u8bf7\u6c42\uff0c\u7b49\u5f85\u5f53\u524d\u6b65\u9aa4\u7ed3\u675f...")
					: (activityText.empty()
						? LocalFromWide(L"AI \u6b63\u5728\u751f\u6210\uff08\u6d41\u5f0f\uff09...")
						: activityText)) + "\r\n";
		}
		else {
			text += "[" + LocalFromWide(L"\u7cfb\u7edf") + "]\r\n"
				+ (stopRequested
					? LocalFromWide(L"\u6b63\u5728\u4e2d\u65ad AI \u8bf7\u6c42...")
					: (activityText.empty()
						? LocalFromWide(L"\u7b49\u5f85 AI \u8fd4\u56de...")
						: activityText)) + "\r\n";
		}
	}
	return text;
}

void PostRefreshDialog()
{
	if (g_chatDialog != nullptr && IsWindow(g_chatDialog)) {
		PostMessage(g_chatDialog, WM_AUTOLINKER_AI_CHAT_REFRESH, 0, 0);
	}
}

bool RequestStopCurrentChat()
{
	std::shared_ptr<AIChatRequestCancellation> cancellation;
	{
		std::lock_guard<std::mutex> guard(g_session.mutex);
		if (!g_session.requestInFlight || g_session.activeRequestId == 0 || g_session.cancellation == nullptr) {
			return false;
		}
		cancellation = g_session.cancellation;
	}

	const bool started = cancellation->RequestCancel();
	if (started) {
		OutputStringToELog("[AI Chat] stop requested for current dialog session");
	}
	PostRefreshDialog();
	return started;
}

std::string DescribeMissingChatSettingField(const std::string& missingField)
{
	if (missingField == "apiKey") {
		return LocalFromWide(L"API Key");
	}
	if (missingField == "baseUrl") {
		return LocalFromWide(L"接口地址");
	}
	if (missingField == "model") {
		return LocalFromWide(L"模型");
	}
	return LocalFromWide(L"AI 配置");
}

bool QueryChatSettingsState(AISettings& outSettings, std::string* outMissingField)
{
	outSettings = {};
	if (g_configManager == nullptr || g_aiJsonConfig == nullptr) {
		if (outMissingField != nullptr) {
			*outMissingField = "configUnavailable";
		}
		return false;
	}

	AIService::LoadSettings(*g_aiJsonConfig, g_configManager, outSettings);
	std::string missingField;
	const bool ready = AIService::HasRequiredSettings(outSettings, missingField);
	if (outMissingField != nullptr) {
		*outMissingField = missingField;
	}
	return ready;
}

void RecoverInFlightIfNeeded(const std::string& reason)
{
	bool recovered = false;
	{
		std::lock_guard<std::mutex> guard(g_session.mutex);
		if (!g_session.requestInFlight) {
			return;
		}
		FinishActiveSessionTimingLocked(g_session, GetCurrentUnixTimeMsForChat());
		g_session.requestInFlight = false;
		g_session.activeRequestId = 0;
		g_session.cancellation.reset();
		g_session.streamingAssistantPreview.clear();
		g_session.agentActivityLines.clear();
		g_session.messages.push_back(SessionMessage{
			SessionRole::System,
			"Chat request auto-recovered: " + reason,
			false,
			true,
			"",
			""
		});
		recovered = true;
	}
	if (recovered) {
		SaveChatSessionSnapshotNow();
	}
}

std::string BuildAgentActivityLine(const std::string& toolName, bool done, bool ok)
{
	std::string displayName = TrimAsciiCopy(toolName);
	if (displayName.empty()) {
		displayName = "<unknown>";
	}
	if (!done) {
		return LocalFromWide(L"\u6b63\u5728\u8c03\u7528\u5de5\u5177\uff1a") + displayName;
	}
	return LocalFromWide(L"\u5de5\u5177\u5b8c\u6210\uff1a") + displayName + (ok ? " (ok)" : " (failed)");
}

void AppendAgentActivity(unsigned long long requestId, const std::string& line)
{
	const std::string trimmed = TrimAsciiCopy(line);
	if (trimmed.empty()) {
		return;
	}
	{
		std::lock_guard<std::mutex> guard(g_session.mutex);
		if (!g_session.requestInFlight || g_session.activeRequestId != requestId) {
			return;
		}
		g_session.agentActivityLines.push_back(trimmed);
		if (g_session.agentActivityLines.size() > 12) {
			g_session.agentActivityLines.erase(
				g_session.agentActivityLines.begin(),
				g_session.agentActivityLines.begin() + (g_session.agentActivityLines.size() - 12));
		}
	}
	PostRefreshDialog();
}

void AppendStreamingAssistantDelta(unsigned long long requestId, const std::string& deltaLocalText)
{
	const std::string normalized = NormalizeCodeForEIDE(deltaLocalText);
	if (normalized.empty()) {
		return;
	}

	{
		std::lock_guard<std::mutex> guard(g_session.mutex);
		if (!g_session.requestInFlight || g_session.activeRequestId != requestId) {
			return;
		}
		g_session.streamingAssistantPreview += normalized;
	}
	PostRefreshDialog();
}

bool EnsureChatSettingsReady(AISettings& settings)
{
	std::string missing;
	if (QueryChatSettingsState(settings, &missing)) {
		return true;
	}

	OutputStringToELog("[AI Chat] AI settings missing, opening config dialog");
	if (!ShowAIConfigDialog(g_mainWindow, *g_aiJsonConfig, settings)) {
		OutputStringToELog("[AI Chat] AI config cancelled");
		return false;
	}
	OutputStringToELog("[AI Chat] AI config saved");
	return true;
}

bool RequestConfirmationFromMainThread(
	const std::string& title,
	const std::string& content,
	const std::string& primaryText,
	const std::string& secondaryText,
	bool& outAccepted,
	bool& outSecondaryAccepted)
{
	outAccepted = false;
	outSecondaryAccepted = false;
	if (g_mainWindow == nullptr || !IsWindow(g_mainWindow)) {
		return false;
	}

	ToolDialogRequest request = {};
	request.kind = ToolDialogRequest::Kind::Confirmation;
	request.title = title;
	request.content = content;
	request.primaryText = primaryText;
	request.secondaryText = secondaryText;
	if (g_msgAIChatToolDialog == 0) {
		return false;
	}
	if (PostMessage(g_mainWindow, g_msgAIChatToolDialog, 0, reinterpret_cast<LPARAM>(&request)) == FALSE) {
		return false;
	}

	std::unique_lock<std::mutex> lock(request.mutex);
	if (!request.cv.wait_for(lock, std::chrono::minutes(20), [&request]() { return request.done; })) {
		return false;
	}

	outAccepted = request.accepted;
	outSecondaryAccepted = request.secondaryAccepted;
	return true;
}

std::string ToLowerAsciiCopyLocal(std::string text)
{
	for (char& ch : text) {
		ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
	}
	return text;
}

void PostChatResult(AIChatAsyncResult* result)
{
	if (result == nullptr) {
		return;
	}
	if (g_msgAIChatDone == 0) {
		RecoverInFlightIfNeeded("chat message id not initialized");
		PostRefreshDialog();
		delete result;
		return;
	}
	if (g_mainWindow == nullptr || !IsWindow(g_mainWindow) ||
		PostMessage(g_mainWindow, g_msgAIChatDone, 0, reinterpret_cast<LPARAM>(result)) == FALSE) {
		RecoverInFlightIfNeeded("result callback post failed");
		PostRefreshDialog();
		delete result;
	}
}

void RunAIChatWorker(void* pParams)
{
	std::unique_ptr<AIChatAsyncRequest> request(reinterpret_cast<AIChatAsyncRequest*>(pParams));
	if (!request) {
		return;
	}

	std::unique_ptr<AIChatAsyncResult> result(new (std::nothrow) AIChatAsyncResult());
	if (!result) {
		return;
	}
	result->requestId = request->requestId;
	try {
		const auto isCancelled = [&request]() {
			return request->cancellation != nullptr && request->cancellation->IsCancelled();
		};
		if (isCancelled()) {
			result->chatResult.ok = false;
			result->chatResult.error = LocalFromWide(L"\u5df2\u53d6\u6d88\uff0c\u672a\u53d1\u9001 AI \u8bf7\u6c42\u3002");
		}
		else {
			PrepareWorkspaceMirrorForChat(request->requestId);
			if (isCancelled()) {
				result->chatResult.ok = false;
				result->chatResult.error = LocalFromWide(L"\u5df2\u53d6\u6d88\uff0c\u672a\u53d1\u9001 AI \u8bf7\u6c42\u3002");
			}
			else {
				result->chatResult = AIService::ExecuteChatWithTools(
					request->contextMessages,
					request->settings,
					[
						cancellation = request->cancellation,
						requestId = request->requestId
					](const std::string& toolName, const std::string& argumentsJson, bool& outOk) -> std::string {
						AppendAgentActivity(requestId, BuildAgentActivityLine(toolName, false, false));
						std::string toolResult = ExecuteToolCall(
							toolName,
							argumentsJson,
							outOk,
							true,
							[cancellation]() {
								return cancellation != nullptr && cancellation->IsCancelled();
							});
						AppendAgentActivity(requestId, BuildAgentActivityLine(toolName, true, outOk));
						return toolResult;
					},
					[requestId = request->requestId](const std::string& deltaText) {
						AppendStreamingAssistantDelta(requestId, deltaText);
					},
					[cancellation = request->cancellation]() {
						return cancellation != nullptr && cancellation->IsCancelled();
					},
					request->cancellation != nullptr ? &request->cancellation->httpRequest : nullptr);
			}
		}
	}
	catch (const std::exception& ex) {
		result->chatResult.ok = false;
		result->chatResult.error = std::string("chat worker exception: ") + ex.what();
	}
	catch (...) {
		result->chatResult.ok = false;
		result->chatResult.error = "chat worker unknown exception";
	}

	PostChatResult(result.release());
}

bool StartChatRequest(const std::string& userInput)
{
	const std::string trimmed = TrimAsciiCopy(userInput);
	if (trimmed.empty()) {
		return false;
	}
	RebindChatSessionToCurrentSourceIfNeeded();

	AISettings settings = {};
	if (!EnsureChatSettingsReady(settings)) {
		return false;
	}

	std::unique_ptr<AIChatAsyncRequest> request(new (std::nothrow) AIChatAsyncRequest());
	if (!request) {
		return false;
	}

	{
		std::lock_guard<std::mutex> guard(g_session.mutex);
		if (g_session.requestInFlight) {
			OutputStringToELog("[AI Chat] request is already in progress");
			return false;
		}

		const long long requestStartedAtMs = GetCurrentUnixTimeMsForChat();
		g_session.sourceFilePathLocal = GetCurrentChatSourceFilePathLocal();
		g_session.sourceFileNameLocal = GetCurrentChatSourceFileNameLocal();
		EnsureChatSessionBindingLocked(g_session);
		g_session.messages.push_back(SessionMessage{ SessionRole::User, trimmed, true, true, "", "" });
		g_session.effectiveContextWindow = AIService::ResolveContextWindowTokens(settings);
		CompactHistoryLocked(g_session);

		request->requestId = g_session.nextRequestId++;
		request->settings = settings;
		request->contextMessages = BuildContextMessagesLocked(g_session);
		request->cancellation = std::make_shared<AIChatRequestCancellation>();
		g_session.streamingAssistantPreview.clear();
		g_session.agentActivityLines.clear();
		g_session.agentActivityLines.push_back(LocalFromWide(L"\u6b63\u5728\u51c6\u5907\u5de5\u7a0b\u955c\u50cf..."));
		g_session.requestInFlight = true;
		g_session.activeRequestId = request->requestId;
		g_session.activeRequestStartedAtUnixMs = requestStartedAtMs;
		g_session.cancellation = request->cancellation;
	}
	SaveChatSessionSnapshotNow();

	const uintptr_t threadId = _beginthread(RunAIChatWorker, 0, request.get());
	if (threadId == static_cast<uintptr_t>(-1L)) {
		{
			std::lock_guard<std::mutex> guard(g_session.mutex);
			FinishActiveSessionTimingLocked(g_session, GetCurrentUnixTimeMsForChat());
			g_session.requestInFlight = false;
			g_session.activeRequestId = 0;
			g_session.cancellation.reset();
			g_session.streamingAssistantPreview.clear();
			g_session.agentActivityLines.clear();
			g_session.messages.push_back(SessionMessage{ SessionRole::System, "Failed to start background chat task.", false, true, "", "" });
		}
		PostRefreshDialog();
		SaveChatSessionSnapshotNow();
		return false;
	}

	request.release();
	PostRefreshDialog();
	return true;
}

void HandleChatSubmitUi(HWND hWnd, ChatDialogContext* ctx, const std::string& text)
{
	if (ctx == nullptr) {
		return;
	}
	HideWebViewSessionMenu(ctx);
	HideChatConfirmInPage(ctx);
	LayoutAIChatDialog(hWnd, ctx);

	const std::string trimmed = TrimAsciiCopy(text);
	if (trimmed.empty()) {
		FocusChatComposerInput(ctx);
		return;
	}

	if (!ctx->webViewDesired && ctx->hInput != nullptr) {
		SetWindowTextA(ctx->hInput, "");
		ctx->inputRowsVisible = 1;
		LayoutAIChatDialog(hWnd, ctx);
	}

	if (ctx->webViewDesired) {
		ClearWebViewInput(ctx);
	}

	EnableWindow(ctx->hSend, FALSE);
	UpdateWebViewComposerState(ctx, true);
	if (!StartChatRequest(trimmed)) {
		RefreshChatDialog(hWnd);
	}
}

void HandleChatClearUi(HWND hWnd, ChatDialogContext* ctx)
{
	if (ctx == nullptr) {
		return;
	}
	HideWebViewSessionMenu(ctx);
	HandleChatClearConfirmedUi(hWnd, ctx);
}

void HandleChatClearConfirmedUi(HWND hWnd, ChatDialogContext* ctx)
{
	if (ctx == nullptr) {
		return;
	}
	HideWebViewSessionMenu(ctx);
	HideChatConfirmInPage(ctx);
	LayoutAIChatDialog(hWnd, ctx);

	bool needConfirm = false;
	if (!QueryClearChatHistoryAvailability(needConfirm)) {
		FocusChatComposerInput(ctx);
		return;
	}

	if (!ctx->webViewDesired && ctx->hHistory != nullptr) {
		SetWindowTextA(ctx->hHistory, "");
	}
	RequestClearChatHistoryAsync();
	UpdateWebViewComposerState(ctx, false);
	FocusChatComposerInput(ctx);
	(void)hWnd;
}

void HandleChatClearCancelUi(HWND hWnd, ChatDialogContext* ctx)
{
	if (ctx == nullptr) {
		return;
	}
	HideChatConfirmInPage(ctx);
	RefreshChatDialog(hWnd);
	LayoutAIChatDialog(hWnd, ctx);
	FocusChatComposerInput(ctx);
}

void HandleChatRestoreSessionUi(HWND hWnd, ChatDialogContext* ctx)
{
	if (ctx != nullptr) {
		HideChatConfirmInPage(ctx);
		LayoutAIChatDialog(hWnd, ctx);
	}
	TryPromptAndRestoreRecentChatSession(hWnd);
}

void HandleChatShowRecentSessionsUi(HWND hWnd, ChatDialogContext* ctx)
{
	RebindChatSessionToCurrentSourceIfNeeded();
	if (ctx != nullptr) {
		HideChatConfirmInPage(ctx);
		LayoutAIChatDialog(hWnd, ctx);
	}
	const std::vector<AIChatStoredSessionListEntry> sessions = LoadRecentChatSessionsForCurrentSource();
	if (ctx != nullptr) {
		ctx->lastSessionMenuEntries = sessions;
	}
	if (ctx != nullptr && ctx->webViewDesired && ctx->webViewContentReady) {
		ShowWebViewSessionMenu(ctx, sessions);
		return;
	}
	if (sessions.empty()) {
		MessageBoxA(hWnd, "当前项目暂无历史会话。", "AI Chat", MB_ICONINFORMATION | MB_OK);
		return;
	}
	TryPromptAndRestoreRecentChatSession(hWnd);
}

void HandleChatRestoreSessionByIdUi(HWND hWnd, ChatDialogContext* ctx, const std::string& sessionId)
{
	if (ctx == nullptr) {
		return;
	}
	const AIChatStoredSessionListEntry* entry = FindSessionMenuEntryById(ctx, sessionId);
	if (entry == nullptr) {
		MessageBoxA(hWnd, "未找到对应的历史会话。", "AI Chat", MB_ICONWARNING | MB_OK);
		return;
	}
	HideWebViewSessionMenu(ctx);
	HideChatConfirmInPage(ctx);
	LayoutAIChatDialog(hWnd, ctx);
	BeginRestoreStoredChatSessionUi(hWnd, ctx, *entry);
}

void HandleChatStopUi(HWND hWnd, ChatDialogContext* ctx)
{
	if (ctx == nullptr) {
		return;
	}
	HideChatConfirmInPage(ctx);
	LayoutAIChatDialog(hWnd, ctx);
	RequestStopCurrentChat();
	RefreshChatDialog(hWnd);
}

void HandleChatOpenSettingsUi(HWND hWnd, ChatDialogContext* ctx)
{
	AISettings settings = {};
	std::string missingField;
	QueryChatSettingsState(settings, &missingField);
	if (ctx != nullptr) {
		HideChatConfirmInPage(ctx);
		LayoutAIChatDialog(hWnd, ctx);
	}
	OutputStringToELog("[AI Chat] open config dialog from chat page");
	if (!ShowAIConfigDialog(g_mainWindow != nullptr ? g_mainWindow : hWnd, *g_aiJsonConfig, settings)) {
		OutputStringToELog("[AI Chat] AI config cancelled from chat page");
		if (ctx != nullptr) {
			if (ctx->webViewDesired) {
				FocusWebViewInput(ctx);
			}
			else if (ctx->hInput != nullptr) {
				SetFocus(ctx->hInput);
			}
		}
		return;
	}
	OutputStringToELog("[AI Chat] AI config saved from chat page");
	RefreshChatDialog(hWnd);
	if (ctx != nullptr) {
		if (ctx->webViewDesired) {
			FocusWebViewInput(ctx);
		}
		else if (ctx->hInput != nullptr) {
			SetFocus(ctx->hInput);
		}
	}
}

std::vector<AIChatStoredSessionListEntry> LoadRecentChatSessionsForCurrentSource()
{
	const std::string sourceFilePath = GetCurrentChatSourceFilePathLocal();
	return ListRecentAIChatStoredSessions(sourceFilePath, 10);
}

bool BeginRestoreStoredChatSessionUi(HWND hWnd, ChatDialogContext* ctx, const AIChatStoredSessionListEntry& selected)
{
	bool needConfirm = false;
	if (!QueryRestoreChatSessionAvailability(needConfirm)) {
		MessageBoxA(hWnd, "当前有进行中的 AI 请求，暂时无法恢复会话。", "AI Chat", MB_ICONWARNING | MB_OK);
		return false;
	}
	if (needConfirm) {
		if (ctx == nullptr) {
			return false;
		}
		ctx->pendingRestoreSession = selected;
		ctx->hasPendingRestoreSession = true;
		ShowRestoreChatConfirmInPage(hWnd, ctx);
		return false;
	}

	return RestoreStoredChatSessionEntry(hWnd, selected);
}

void HandleChatRestoreSessionConfirmedUi(HWND hWnd, ChatDialogContext* ctx)
{
	if (ctx == nullptr || !ctx->hasPendingRestoreSession) {
		if (ctx != nullptr) {
			HideRestoreChatConfirmInPage(ctx, true);
			LayoutAIChatDialog(hWnd, ctx);
			FocusChatComposerInput(ctx);
		}
		return;
	}

	const AIChatStoredSessionListEntry selected = ctx->pendingRestoreSession;
	HideRestoreChatConfirmInPage(ctx, true);
	LayoutAIChatDialog(hWnd, ctx);
	RestoreStoredChatSessionEntry(hWnd, selected);
}

void HandleChatRestoreSessionCancelUi(HWND hWnd, ChatDialogContext* ctx)
{
	if (ctx == nullptr) {
		return;
	}
	HideRestoreChatConfirmInPage(ctx, true);
	RefreshChatDialog(hWnd);
	LayoutAIChatDialog(hWnd, ctx);
	FocusChatComposerInput(ctx);
}

bool RestoreStoredChatSessionEntry(HWND hWnd, const AIChatStoredSessionListEntry& selected)
{
	AIChatStoredSession stored;
	std::string loadError;
	if (!LoadAIChatStoredSession(selected.sessionFilePath, stored, &loadError)) {
		MessageBoxA(hWnd, ("恢复会话失败: " + loadError).c_str(), "AI Chat", MB_ICONERROR | MB_OK);
		return false;
	}

	if (!ReplaceChatSessionStateFromStoredSession(stored)) {
		MessageBoxA(hWnd, "当前有进行中的 AI 请求，暂时无法恢复会话。", "AI Chat", MB_ICONWARNING | MB_OK);
		return false;
	}

	RefreshChatDialog(g_chatDialog != nullptr ? g_chatDialog : hWnd);
	auto* dialogCtx = reinterpret_cast<ChatDialogContext*>(GetWindowLongPtrA(
		g_chatDialog != nullptr ? g_chatDialog : hWnd,
		GWLP_USERDATA));
	if (dialogCtx != nullptr && dialogCtx->webViewDesired) {
		FocusWebViewInput(dialogCtx);
	}
	else if (dialogCtx != nullptr && dialogCtx->hInput != nullptr) {
		SetFocus(dialogCtx->hInput);
	}
	return true;
}

bool TryPromptAndRestoreRecentChatSession(HWND hWnd)
{
	RebindChatSessionToCurrentSourceIfNeeded();

	const std::vector<AIChatStoredSessionListEntry> sessions = LoadRecentChatSessionsForCurrentSource();
	if (sessions.empty()) {
		MessageBoxA(hWnd, "当前项目暂无历史会话。", "AI Chat", MB_ICONINFORMATION | MB_OK);
		return false;
	}

	HMENU menu = CreatePopupMenu();
	if (menu == nullptr) {
		return false;
	}

	constexpr UINT kRestoreSessionMenuBase = 47000;
	for (size_t i = 0; i < sessions.size(); ++i) {
		std::string text = sessions[i].updatedAtDisplayLocal;
		if (text.empty()) {
			text = "未知时间";
		}
		text += "  ";
		text += sessions[i].titleLocal.empty() ? "未命名会话" : sessions[i].titleLocal;
		AppendMenuA(menu, MF_STRING, kRestoreSessionMenuBase + static_cast<UINT>(i), text.c_str());
	}

	RECT anchor = {};
	auto* ctx = reinterpret_cast<ChatDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
	if (ctx != nullptr && ctx->hRestoreSession != nullptr && IsWindow(ctx->hRestoreSession)) {
		GetWindowRect(ctx->hRestoreSession, &anchor);
	}
	else if (g_chatDialog != nullptr && IsWindow(g_chatDialog)) {
		GetWindowRect(g_chatDialog, &anchor);
	}
	const int command = TrackPopupMenu(
		menu,
		TPM_RETURNCMD | TPM_LEFTALIGN | TPM_TOPALIGN,
		anchor.left,
		anchor.bottom + 4,
		0,
		g_chatDialog != nullptr ? g_chatDialog : hWnd,
		nullptr);
	DestroyMenu(menu);

	if (command < static_cast<int>(kRestoreSessionMenuBase) ||
		command >= static_cast<int>(kRestoreSessionMenuBase + sessions.size())) {
		return false;
	}
	return BeginRestoreStoredChatSessionUi(
		hWnd,
		ctx,
		sessions[static_cast<size_t>(command - kRestoreSessionMenuBase)]);
}

void ClearChatHistory()
{
	std::vector<SessionMessage> oldMessages;
	std::string oldSummary;
	{
		std::lock_guard<std::mutex> guard(g_session.mutex);
		if (HasAnyChatHistoryLocked(g_session)) {
			PersistChatSessionSnapshotLocked(g_session);
		}
		oldMessages.swap(g_session.messages);
		oldSummary.swap(g_session.rollingSummary);
		g_session.streamingAssistantPreview.clear();
		g_session.agentActivityLines.clear();
		g_session.requestInFlight = false;
		g_session.activeRequestId = 0;
		g_session.cancellation.reset();
		g_session.lastInputTokens = 0;
		g_session.hasLastUsage = false;
		g_session.effectiveContextWindow = 0;
		ResetChatSessionBindingLocked(g_session);
	}
	WorkspaceMirror::ResetAndCleanup();
}

void RequestClearChatHistoryAsync()
{
	if (g_clearHistoryInProgress.exchange(true)) {
		return;
	}
	ClearChatHistory();
	g_clearHistoryInProgress = false;
	PostRefreshDialog();
}

void RefreshProjectSourceCacheAfterChatTurn()
{
#if defined(_M_IX86)
	project_source_cache::Snapshot snapshot;
	bool refreshed = false;
	std::string error;
	std::string trace;
	if (project_source_cache::ProjectSourceCacheManager::Instance().EnsureCurrentSourceLatest(
			snapshot,
			true,
			&refreshed,
			&error,
			&trace)) {
		OutputStringToELog(std::format(
			"[ProjectSourceCache] chat_turn_refresh_ok source={} revision={} pages={} refreshed={} trace={}",
			snapshot.sourcePath,
			snapshot.revision,
			snapshot.pages.size(),
			refreshed ? 1 : 0,
			trace));
	}
	else if (!error.empty()) {
		OutputStringToELog(std::format(
			"[ProjectSourceCache] chat_turn_refresh_skip error={} trace={}",
			error,
			trace));
	}
#endif
}

std::string BuildToolEventHistoryLine(const AIChatToolEvent& evt)
{
	std::string line =
		LocalFromWide(L"\u8c03\u7528 ") + evt.name +
		LocalFromWide(L"\uff0c\u8fd4\u56de\uff1a") + evt.resultJson;
	if (line.size() > 1200) {
		line.resize(1200);
		line += "...";
	}
	return line;
}

std::string BuildToolEventSummaryLine(const std::vector<AIChatToolEvent>& events)
{
	std::string line = LocalFromWide(L"\u672c\u8f6e\u5171\u8c03\u7528\u5de5\u5177 ") +
		std::to_string(events.size()) +
		LocalFromWide(L" \u6b21\uff0c\u5df2\u4fdd\u7559\u6458\u8981\u3002\u6700\u8fd1\u5de5\u5177\uff1a");
	const size_t keepCount = (std::min)(events.size(), static_cast<size_t>(5));
	const size_t begin = events.size() - keepCount;
	for (size_t i = begin; i < events.size(); ++i) {
		if (i > begin) {
			line += ", ";
		}
		line += events[i].name.empty() ? "<unknown>" : events[i].name;
		line += events[i].ok ? "(ok)" : "(failed)";
	}
	if (begin > 0) {
		line += LocalFromWide(L"\uff08\u524d\u9762 ");
		line += std::to_string(begin);
		line += LocalFromWide(L" \u6b21\u5df2\u7701\u7565\uff09");
	}
	return line;
}

void AppendToolEventHistoryMessagesLocked(AIChatSessionState& state, const AIChatResult& chatResult)
{
	if (chatResult.toolEvents.empty()) {
		return;
	}

	const bool shouldSummarize =
		chatResult.toolRoundsExceeded ||
		chatResult.cancelled ||
		!chatResult.ok ||
		chatResult.toolEvents.size() > 16;
	if (shouldSummarize && chatResult.toolEvents.size() > 8) {
		state.messages.push_back(SessionMessage{
			SessionRole::Tool,
			BuildToolEventSummaryLine(chatResult.toolEvents),
			false,
			true,
			"",
			""
		});
		return;
	}

	for (const auto& evt : chatResult.toolEvents) {
		state.messages.push_back(SessionMessage{
			SessionRole::Tool,
			BuildToolEventHistoryLine(evt),
			false,
			true,
			"",
			""
		});
	}
}

void HandleChatTaskDone(LPARAM lParam)
{
	std::unique_ptr<AIChatAsyncResult> result(reinterpret_cast<AIChatAsyncResult*>(lParam));
	if (!result) {
		return;
	}

	{
		std::lock_guard<std::mutex> guard(g_session.mutex);
		if (!g_session.requestInFlight || g_session.activeRequestId != result->requestId) {
			return;
		}

		FinishActiveSessionTimingLocked(g_session, GetCurrentUnixTimeMsForChat());
		g_session.requestInFlight = false;
		g_session.activeRequestId = 0;
		g_session.cancellation.reset();
		g_session.streamingAssistantPreview.clear();
		g_session.agentActivityLines.clear();

		if (result->chatResult.hasUsage && result->chatResult.promptTokens > 0) {
			g_session.lastInputTokens = result->chatResult.promptTokens;
			g_session.hasLastUsage = true;
		}

		AppendToolEventHistoryMessagesLocked(g_session, result->chatResult);

		if (result->chatResult.ok) {
			for (const auto& rawMessageJsonUtf8 : result->chatResult.contextPrefixRawMessagesUtf8) {
				nlohmann::json parsedMessage;
				try {
					parsedMessage = nlohmann::json::parse(rawMessageJsonUtf8);
				}
				catch (...) {
					continue;
				}

				if (!parsedMessage.is_object()) {
					continue;
				}

				const std::string role = ToLowerAsciiCopySimple(AIService::Trim(parsedMessage.value("role", "")));
				const std::string itemType = ToLowerAsciiCopySimple(AIService::Trim(parsedMessage.value("type", "")));
				if (role == "assistant") {
					std::string contentLocal;
					if (parsedMessage.contains("content") && parsedMessage["content"].is_string()) {
						contentLocal = NormalizeCodeForEIDE(parsedMessage["content"].get<std::string>());
					}
					std::string reasoningUtf8;
					if (parsedMessage.contains("reasoning_content") && parsedMessage["reasoning_content"].is_string()) {
						reasoningUtf8 = parsedMessage["reasoning_content"].get<std::string>();
					}
					g_session.messages.push_back(SessionMessage{
						SessionRole::Assistant,
						contentLocal,
						true,
						false,
						reasoningUtf8,
						rawMessageJsonUtf8
					});
				}
				else if (role == "tool") {
					g_session.messages.push_back(SessionMessage{
						SessionRole::Tool,
						"",
						true,
						false,
						"",
						rawMessageJsonUtf8
					});
				}
				else if (!itemType.empty()) {
					g_session.messages.push_back(SessionMessage{
						SessionRole::Tool,
						"",
						true,
						false,
						"",
						rawMessageJsonUtf8
					});
				}
			}
		}

		if (result->chatResult.cancelled) {
			const std::string partial = NormalizeCodeForEIDE(result->chatResult.content);
			if (!TrimAsciiCopy(partial).empty()) {
				g_session.messages.push_back(SessionMessage{
					SessionRole::Assistant,
					partial,
					true,
					true,
					result->chatResult.reasoningContent,
					""
				});
			}
			g_session.messages.push_back(SessionMessage{
				SessionRole::System,
				TrimAsciiCopy(partial).empty()
					? LocalFromWide(L"已停止本次 AI 对话。")
					: LocalFromWide(L"已停止本次 AI 对话，已保留已生成内容。"),
				false,
				true,
				"",
				""
			});
		}
		else if (result->chatResult.ok) {
			g_session.messages.push_back(SessionMessage{
				SessionRole::Assistant,
				NormalizeCodeForEIDE(result->chatResult.content),
				true,
				true,
				result->chatResult.reasoningContent,
				""
			});
		}
		else {
			std::string err;
			if (result->chatResult.toolRoundsExceeded) {
				err = LocalFromWide(L"AI\u5bf9\u8bdd\u5931\u8d25: ") + result->chatResult.error +
					LocalFromWide(L"\u3002\u5df2\u4fdd\u7559\u5f53\u524d\u4f1a\u8bdd\u548c\u672c\u8f6e\u7528\u6237\u8f93\u5165\uff0c\u4e0b\u6b21\u53d1\u9001\u4f1a\u7ee7\u7eed\u5e26\u5165\u3002\u53ef\u5728 AI Key \u8bbe\u7f6e\u4e2d\u8c03\u9ad8\u6700\u5927\u5de5\u5177\u8f6e\u6570\uff0c\u6216\u8865\u5145\u66f4\u660e\u786e\u7684\u505c\u6b62\u6761\u4ef6\u3002");
			}
			else {
				err = result->chatResult.error.empty()
					? LocalFromWide(L"AI\u5bf9\u8bdd\u5931\u8d25")
					: (LocalFromWide(L"AI\u5bf9\u8bdd\u5931\u8d25: ") + result->chatResult.error);
			}
			g_session.messages.push_back(SessionMessage{ SessionRole::System, err, false, true, "", "" });
			OutputStringToELog("[" + LocalFromWide(L"AI\u5bf9\u8bdd") + "]" + err);
		}

		CompactHistoryLocked(g_session);
	}

	PostRefreshDialog();
	SaveChatSessionSnapshotNow();
	RefreshProjectSourceCacheAfterChatTurn();
}

void RefreshChatDialog(HWND hWnd)
{
	RebindChatSessionToCurrentSourceIfNeeded();

	auto* ctx = reinterpret_cast<ChatDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
	if (ctx == nullptr) {
		return;
	}

	AISettings currentSettings = {};
	std::string missingField;
	const bool settingsReady = QueryChatSettingsState(currentSettings, &missingField);
	std::string history;
	std::string historyHtml;
	bool inFlight = false;
	bool stopRequested = false;
	ContextUsageSnapshot usageSnapshot;
	SessionTimingSnapshot timingSnapshot;
	const int fallbackContextWindow = settingsReady
		? AIService::ResolveContextWindowTokens(currentSettings)
		: 200000;
	{
		std::lock_guard<std::mutex> guard(g_session.mutex);
		if (g_session.requestInFlight && g_session.activeRequestId == 0) {
			FinishActiveSessionTimingLocked(g_session, GetCurrentUnixTimeMsForChat());
			g_session.requestInFlight = false;
			g_session.cancellation.reset();
			g_session.streamingAssistantPreview.clear();
		}
		history = BuildHistoryTextLocked(g_session, settingsReady, missingField);
		if (ctx->webViewDesired) {
			historyHtml = BuildHistoryHtmlLocked(g_session, settingsReady, missingField);
		}
		inFlight = g_session.requestInFlight;
		stopRequested = IsStopRequestedLocked(g_session);
		usageSnapshot = BuildContextUsageSnapshotLocked(g_session, fallbackContextWindow);
		timingSnapshot = BuildSessionTimingSnapshotLocked(g_session, GetCurrentUnixTimeMsForChat());
	}

	const bool hideInlineConfirmForBusy = inFlight && HasVisibleInlineConfirm(ctx);
	if (hideInlineConfirmForBusy) {
		HideChatConfirmInPage(ctx);
	}
	const bool timingVisibilityChanged = ctx->sessionTimingVisible != timingSnapshot.visible;

	SetWindowTextA(ctx->hHistory, history.c_str());
	ScrollEditToBottom(ctx->hHistory);
	if (ctx->webViewDesired) {
		UpdateHistoryWebViewHtml(ctx, historyHtml);
	}
	const bool nativeComposerVisible = !ctx->webViewContentReady;
	if (ctx->hClearHistory != nullptr) {
		EnableWindow(ctx->hClearHistory, inFlight ? FALSE : TRUE);
	}
	if (ctx->hRestoreSession != nullptr) {
		EnableWindow(ctx->hRestoreSession, inFlight ? FALSE : TRUE);
	}
	if (ctx->hOpenSettings != nullptr) {
		EnableWindow(ctx->hOpenSettings, inFlight ? FALSE : TRUE);
	}
	if (ctx->hClearConfirmApply != nullptr) {
		EnableWindow(ctx->hClearConfirmApply, inFlight ? FALSE : TRUE);
	}
	if (ctx->hClearConfirmCancel != nullptr) {
		EnableWindow(ctx->hClearConfirmCancel, inFlight ? FALSE : TRUE);
	}
	if (ctx->hSend != nullptr) {
		EnableWindow(ctx->hSend, inFlight ? FALSE : TRUE);
		ShowWindow(ctx->hSend, (nativeComposerVisible && !inFlight) ? SW_SHOW : SW_HIDE);
	}
	if (ctx->hStop != nullptr) {
		EnableWindow(ctx->hStop, (inFlight && !stopRequested) ? TRUE : FALSE);
		SetWindowTextW(ctx->hStop, stopRequested ? L"\u505c\u6b62\u4e2d..." : L"\u505c\u6b62");
		ShowWindow(ctx->hStop, (nativeComposerVisible && inFlight) ? SW_SHOW : SW_HIDE);
	}
	if (ctx->hContextUsage != nullptr) {
		ShowWindow(ctx->hContextUsage, nativeComposerVisible ? SW_SHOW : SW_HIDE);
		UpdateNativeContextUsage(ctx, usageSnapshot);
	}
	if (ctx->hSessionElapsed != nullptr) {
		ShowWindow(ctx->hSessionElapsed, (nativeComposerVisible && timingSnapshot.visible) ? SW_SHOW : SW_HIDE);
	}
	if (ctx->hSessionStatus != nullptr) {
		ShowWindow(ctx->hSessionStatus, (nativeComposerVisible && timingSnapshot.visible) ? SW_SHOW : SW_HIDE);
	}
	UpdateNativeSessionTiming(ctx, timingSnapshot);
	UpdateWebViewComposerState(ctx, inFlight, stopRequested);
	UpdateWebViewContextUsage(ctx, usageSnapshot);
	UpdateWebViewSessionTiming(ctx, timingSnapshot);
	SyncSessionTimingTimer(hWnd, timingSnapshot.inProgress);
	if (hideInlineConfirmForBusy || timingVisibilityChanged) {
		LayoutAIChatDialog(hWnd, ctx);
	}
}

LRESULT CALLBACK AIChatDialogProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	auto* ctx = reinterpret_cast<ChatDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
	switch (uMsg)
	{
	case WM_NCCREATE: {
		auto* created = new (std::nothrow) ChatDialogContext();
		if (created == nullptr) {
			return FALSE;
		}
		SetWindowLongPtrA(hWnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(created));
		return TRUE;
	}

	case WM_CREATE: {
		if (ctx == nullptr) {
			return -1;
		}

		ctx->hHistory = CreateWindowExA(WS_EX_CLIENTEDGE, "EDIT", "",
			WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY | WS_VSCROLL,
			14, 14, 752, 420, hWnd, reinterpret_cast<HMENU>(IDC_AI_CHAT_HISTORY), nullptr, nullptr);
		ctx->webViewDesired = IsWebView2RuntimeAvailable();
		if (ctx->webViewDesired) {
			ctx->hHistoryHost = CreateWindowExA(
				WS_EX_CLIENTEDGE,
				"STATIC",
				"",
				WS_CHILD,
				14, 14, 752, 420,
				hWnd,
				nullptr,
				nullptr,
				nullptr);
		}
		ctx->hClearHistory = CreateWindowW(L"STATIC", L"+",
			WS_CHILD | WS_VISIBLE | SS_NOTIFY | SS_CENTER | SS_CENTERIMAGE,
			14, 442, 26, 26, hWnd, reinterpret_cast<HMENU>(IDC_AI_CHAT_CLEAR_HISTORY), nullptr, nullptr);
		ctx->hRestoreSession = CreateWindowW(L"STATIC", L"\u21BB",
			WS_CHILD | WS_VISIBLE | SS_NOTIFY | SS_CENTER | SS_CENTERIMAGE,
			44, 442, 26, 26, hWnd, reinterpret_cast<HMENU>(IDC_AI_CHAT_RESTORE_SESSION), nullptr, nullptr);
		ctx->hOpenSettings = CreateWindowW(L"STATIC", L"\u2699",
			WS_CHILD | WS_VISIBLE | SS_NOTIFY | SS_CENTER | SS_CENTERIMAGE,
			76, 442, 26, 26, hWnd, reinterpret_cast<HMENU>(IDC_AI_CHAT_OPEN_SETTINGS), nullptr, nullptr);
		ctx->hSessionElapsed = CreateWindowW(L"STATIC", L"0m 0s",
			WS_CHILD | WS_VISIBLE | SS_CENTER | SS_CENTERIMAGE,
			424, 476, 86, 32, hWnd, reinterpret_cast<HMENU>(IDC_AI_CHAT_SESSION_ELAPSED), nullptr, nullptr);
		ctx->hSessionStatus = CreateWindowW(L"STATIC", L"\u5df2\u5b8c\u6210",
			WS_CHILD | WS_VISIBLE | SS_CENTER | SS_CENTERIMAGE,
			516, 476, 54, 32, hWnd, reinterpret_cast<HMENU>(IDC_AI_CHAT_SESSION_STATUS), nullptr, nullptr);
		ctx->hContextUsage = CreateWindowW(L"STATIC", L"\u4e0a\u4e0b\u6587 --",
			WS_CHILD | WS_VISIBLE | SS_CENTER | SS_CENTERIMAGE,
			586, 476, 82, 32, hWnd, reinterpret_cast<HMENU>(IDC_AI_CHAT_CONTEXT_USAGE), nullptr, nullptr);
		ctx->hClearConfirmText = CreateWindowW(L"STATIC", L"\u5f53\u524d AI \u5bf9\u8bdd\u5c06\u4fdd\u5b58\u5230\u5386\u53f2\uff0c\u5e76\u5f00\u542f\u65b0\u4f1a\u8bdd\u3002",
			WS_CHILD | SS_LEFT | SS_CENTERIMAGE,
			14, 442, 540, 26, hWnd, reinterpret_cast<HMENU>(IDC_AI_CHAT_CLEAR_CONFIRM_TEXT), nullptr, nullptr);
		ctx->hClearConfirmApply = CreateWindowW(L"STATIC", L"\u65b0\u5efa",
			WS_CHILD | SS_NOTIFY | SS_CENTER | SS_CENTERIMAGE,
			560, 442, 52, 26, hWnd, reinterpret_cast<HMENU>(IDC_AI_CHAT_CLEAR_CONFIRM_APPLY), nullptr, nullptr);
		ctx->hClearConfirmCancel = CreateWindowW(L"STATIC", L"\u53d6\u6d88",
			WS_CHILD | SS_NOTIFY | SS_CENTER | SS_CENTERIMAGE,
			618, 442, 52, 26, hWnd, reinterpret_cast<HMENU>(IDC_AI_CHAT_CLEAR_CONFIRM_CANCEL), nullptr, nullptr);
		ctx->hMcpGuideLink = CreateWindowExW(
			0,
			L"SysLink",
			L"<a href=\"https://github.com/aiqinxuancai/AutoLinker/blob/master/CONFIG.md#%E5%A4%96%E9%83%A8-agent-mcp-%E9%85%8D%E7%BD%AE\">\u63a8\u8350\u7528Codex\u8fde\u63a5MCP</a>\r\n"
			L"<a href=\"https://github.com/aiqinxuancai/Awesome-E-Agent\">\u5fc5\u8bfb\uff1a\u6613\u8bed\u8a00 \u00d7 AI Agent \u5b9e\u8df5\u767d\u76ae\u4e66</a>",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP,
			14, 468, 652, 36,
			hWnd,
			reinterpret_cast<HMENU>(IDC_AI_CHAT_MCP_GUIDE_LINK),
			nullptr,
			nullptr);
		ctx->hInput = CreateWindowExA(WS_EX_CLIENTEDGE, "EDIT", "",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_MULTILINE | ES_AUTOVSCROLL | WS_VSCROLL,
			14, 476, 652, 32, hWnd, reinterpret_cast<HMENU>(IDC_AI_CHAT_INPUT), nullptr, nullptr);
		ctx->hSend = CreateWindowW(L"STATIC", L"\u53d1\u9001",
			WS_CHILD | WS_VISIBLE | SS_NOTIFY | SS_CENTER | SS_CENTERIMAGE,
			674, 476, 92, 32, hWnd, reinterpret_cast<HMENU>(IDC_AI_CHAT_SEND), nullptr, nullptr);
		ctx->hStop = CreateWindowW(L"STATIC", L"\u505c\u6b62",
			WS_CHILD | SS_NOTIFY | SS_CENTER | SS_CENTERIMAGE,
			674, 476, 92, 32, hWnd, reinterpret_cast<HMENU>(IDC_AI_CHAT_STOP), nullptr, nullptr);

		SetDefaultFont(ctx->hHistory);
		SetDefaultFont(ctx->hClearHistory);
		SetDefaultFont(ctx->hRestoreSession);
		SetDefaultFont(ctx->hOpenSettings);
		SetDefaultFont(ctx->hSessionElapsed);
		SetDefaultFont(ctx->hSessionStatus);
		SetDefaultFont(ctx->hContextUsage);
		SetDefaultFont(ctx->hClearConfirmText);
		SetDefaultFont(ctx->hClearConfirmApply);
		SetDefaultFont(ctx->hClearConfirmCancel);
		SetDefaultFont(ctx->hMcpGuideLink);
		SetDefaultFont(ctx->hInput);
		SetDefaultFont(ctx->hSend);
		SetDefaultFont(ctx->hStop);
		InstallEditHotkeys(ctx->hHistory, kEditFlagNone);
		InstallEditHotkeys(ctx->hInput, kEditFlagSubmitOnEnter);
		InstallChatActionControl(ctx->hSend, kActionSubmit);
		InstallChatActionControl(ctx->hStop, kActionStop);
		InstallChatActionControl(ctx->hClearHistory, kActionClear);
		InstallChatActionControl(ctx->hRestoreSession, kActionRestoreSession);
		InstallChatActionControl(ctx->hOpenSettings, kActionOpenSettings);
		InstallChatActionControl(ctx->hClearConfirmApply, kActionClearConfirm);
		InstallChatActionControl(ctx->hClearConfirmCancel, kActionClearCancel);
		ctx->inputRowsVisible = 1;
		LayoutAIChatDialog(hWnd, ctx);
		SyncHistoryPresentation(ctx);
		PostMessageA(hWnd, WM_AUTOLINKER_AI_CHAT_DEFER_LAYOUT, 0, 0);

		RefreshChatDialog(hWnd);
		TryInitializeHistoryWebView(hWnd, ctx);
		SetFocus(ctx->hInput);
		return 0;
	}

	case WM_GETMINMAXINFO: {
		auto* mmi = reinterpret_cast<MINMAXINFO*>(lParam);
		ApplyMinTrackSize(hWnd, mmi, 560, 420);
		return 0;
	}

	case WM_SIZE:
		if (ctx != nullptr) {
			UpdateInputRowsAndLayout(hWnd, ctx, true);
			ScrollEditToBottom(ctx->hHistory);
		}
		return 0;

	case WM_WINDOWPOSCHANGED:
		if (ctx != nullptr) {
			const auto* pos = reinterpret_cast<WINDOWPOS*>(lParam);
			if (pos != nullptr && (pos->flags & SWP_NOSIZE) == 0) {
				UpdateInputRowsAndLayout(hWnd, ctx, true);
				ScrollEditToBottom(ctx->hHistory);
			}
		}
		break;

	case WM_NOTIFY: {
		const auto* hdr = reinterpret_cast<const NMHDR*>(lParam);
		if (hdr != nullptr &&
			hdr->idFrom == IDC_AI_CHAT_MCP_GUIDE_LINK &&
			(hdr->code == NM_CLICK || hdr->code == NM_RETURN)) {
			const auto* link = reinterpret_cast<const NMLINK*>(lParam);
			const std::string url = link != nullptr ? Utf8FromWide(link->item.szUrl) : std::string();
			OpenChatExternalUrl(hWnd, url);
			return 0;
		}
		break;
	}

	case WM_CTLCOLORSTATIC:
		if (ctx != nullptr) {
			HWND hStatic = reinterpret_cast<HWND>(lParam);
			HDC hdc = reinterpret_cast<HDC>(wParam);
			if (hStatic == ctx->hClearConfirmText) {
				SetTextColor(hdc, RGB(55, 65, 81));
				SetBkMode(hdc, TRANSPARENT);
				return reinterpret_cast<LRESULT>(GetSysColorBrush(COLOR_BTNFACE));
			}
			if (hStatic == ctx->hSessionElapsed) {
				SetTextColor(hdc, RGB(76, 79, 105));
				SetBkMode(hdc, TRANSPARENT);
				return reinterpret_cast<LRESULT>(GetSysColorBrush(COLOR_BTNFACE));
			}
			if (hStatic == ctx->hSessionStatus) {
				SetTextColor(hdc, ctx->sessionTimingInProgress ? RGB(30, 102, 245) : RGB(46, 125, 50));
				SetBkMode(hdc, TRANSPARENT);
				return reinterpret_cast<LRESULT>(GetSysColorBrush(COLOR_BTNFACE));
			}
			if (hStatic != ctx->hClearHistory &&
				hStatic != ctx->hRestoreSession &&
				hStatic != ctx->hOpenSettings &&
				hStatic != ctx->hClearConfirmApply &&
				hStatic != ctx->hClearConfirmCancel &&
				hStatic != ctx->hSend &&
				hStatic != ctx->hStop) {
				break;
			}
			SetTextColor(hdc, hStatic == ctx->hClearConfirmApply ? RGB(180, 35, 24) : RGB(0, 102, 204));
			SetBkMode(hdc, TRANSPARENT);
			return reinterpret_cast<LRESULT>(GetSysColorBrush(COLOR_BTNFACE));
		}
		break;

	case WM_COMMAND: {
		if (ctx == nullptr) {
			return 0;
		}
		const int id = LOWORD(wParam);
		const int notifyCode = HIWORD(wParam);
		if (id == IDC_AI_CHAT_INPUT && notifyCode == EN_CHANGE) {
			UpdateInputRowsAndLayout(hWnd, ctx, false);
			return 0;
		}
		return 0;
	}

	case WM_AUTOLINKER_AI_CHAT_DEFER_LAYOUT:
		if (ctx != nullptr) {
			UpdateInputRowsAndLayout(hWnd, ctx, true);
			ScrollEditToBottom(ctx->hHistory);
		}
		return 0;

	case WM_TIMER:
		if (ctx != nullptr && wParam == kHistoryWebViewFlushTimerId) {
			KillTimer(hWnd, kHistoryWebViewFlushTimerId);
			if (ctx->webViewFlushScheduled) {
				FlushHistoryWebViewHtml(ctx);
			}
			return 0;
		}
		if (ctx != nullptr && wParam == kSessionTimingTimerId) {
			RefreshSessionTimingOnly(hWnd, ctx);
			return 0;
		}
		break;

	case WM_AUTOLINKER_AI_CHAT_SUBMIT:
		if (ctx != nullptr) {
			OutputStringToELog("[AI Chat][UI] submit begin");
			HandleChatSubmitUi(hWnd, ctx, GetEditTextA(ctx->hInput));
			OutputStringToELog("[AI Chat][UI] submit end");
		}
		return 0;

	case WM_AUTOLINKER_AI_CHAT_CLEAR:
		if (ctx != nullptr) {
			HandleChatClearUi(hWnd, ctx);
		}
		return 0;

	case WM_AUTOLINKER_AI_CHAT_CLEAR_CONFIRMED:
		if (ctx != nullptr) {
			HandleChatClearConfirmedUi(hWnd, ctx);
		}
		return 0;

	case WM_AUTOLINKER_AI_CHAT_CLEAR_CANCEL:
		if (ctx != nullptr) {
			HandleChatClearCancelUi(hWnd, ctx);
		}
		return 0;

	case WM_AUTOLINKER_AI_CHAT_RESTORE_CONFIRMED:
		if (ctx != nullptr) {
			HandleChatRestoreSessionConfirmedUi(hWnd, ctx);
		}
		return 0;

	case WM_AUTOLINKER_AI_CHAT_RESTORE_CANCEL:
		if (ctx != nullptr) {
			HandleChatRestoreSessionCancelUi(hWnd, ctx);
		}
		return 0;

	case WM_AUTOLINKER_AI_CHAT_STOP:
		if (ctx != nullptr) {
			HandleChatStopUi(hWnd, ctx);
		}
		return 0;

	case WM_AUTOLINKER_AI_CHAT_OPEN_SETTINGS:
		HandleChatOpenSettingsUi(hWnd, ctx);
		return 0;

	case WM_AUTOLINKER_AI_CHAT_RESTORE_LAST:
		if (ctx != nullptr) {
			HandleChatRestoreSessionUi(hWnd, ctx);
		}
		return 0;

	case WM_AUTOLINKER_AI_CHAT_REFRESH:
		RefreshChatDialog(hWnd);
		return 0;

	case WM_AUTOLINKER_AI_CHAT_UPDATE_TAG:
		if (ctx != nullptr) {
			UpdateWebViewUpdateTag(ctx);
		}
		return 0;

	case WM_CLOSE:
		DestroyWindow(hWnd);
		return 0;

	case WM_DESTROY:
		if (g_chatDialog == hWnd) {
			g_chatDialog = nullptr;
			g_chatTabAdded = false;
		}
		KillTimer(hWnd, kHistoryWebViewFlushTimerId);
		KillTimer(hWnd, kSessionTimingTimerId);
		ctx->webView = nullptr;
		ctx->webViewController = nullptr;
		ctx->webViewEnvironment = nullptr;
		delete ctx;
		SetWindowLongPtrA(hWnd, GWLP_USERDATA, 0);
		return 0;

	default:
		break;
	}
	return DefWindowProcA(hWnd, uMsg, wParam, lParam);
}

bool HandleToolDialogRequest(LPARAM lParam)
{
	auto* request = reinterpret_cast<ToolDialogRequest*>(lParam);
	if (request == nullptr) {
		return true;
	}

	bool accepted = false;
	bool secondaryAccepted = false;
	if (request->kind == ToolDialogRequest::Kind::Confirmation) {
		const AIPreviewAction action = ShowAIPreviewDialogEx(
			g_mainWindow,
			request->title.empty() ? "AI Tool Confirmation" : request->title,
			request->content,
			request->primaryText,
			request->secondaryText);
		accepted = action == AIPreviewAction::PrimaryConfirm;
		secondaryAccepted = action == AIPreviewAction::SecondaryConfirm;
	}

	{
		std::lock_guard<std::mutex> guard(request->mutex);
		request->accepted = accepted;
		request->secondaryAccepted = secondaryAccepted;
		request->done = true;
	}
	request->cv.notify_one();
	return true;
}

bool HandleToolExecRequest(LPARAM lParam)
{
	auto* request = reinterpret_cast<ToolExecutionRequest*>(lParam);
	if (request == nullptr) {
		return true;
	}

	bool ok = false;
	const std::string resultJson = ExecuteToolCallOnMainThread(request->toolName, request->argumentsJson, ok);
	{
		std::lock_guard<std::mutex> guard(request->mutex);
		request->ok = ok;
		request->resultJson = resultJson;
		request->done = true;
	}
	request->cv.notify_one();
	return true;
}
} // namespace

HWND GetAIChatMainWindowForTooling()
{
	return g_mainWindow;
}

ConfigManager* GetAIChatConfigManagerForTooling()
{
	return g_configManager;
}

AIJsonConfig* GetAIChatAIJsonConfigForTooling()
{
	return g_aiJsonConfig;
}

UINT GetAIChatToolExecMessageForTooling()
{
	return g_msgAIChatToolExec;
}

bool RequestConfirmationForTooling(
	const std::string& title,
	const std::string& content,
	const std::string& primaryText,
	const std::string& secondaryText,
	bool& outAccepted,
	bool& outSecondaryAccepted)
{
	return RequestConfirmationFromMainThread(title, content, primaryText, secondaryText, outAccepted, outSecondaryAccepted);
}

bool EnsureChatHostWindowCreated();
void FocusChatInputControl();

std::string GetChatTabCaption()
{
	return LocalFromWide(L"AutoLinker AI 对话");
}

void LogChatTab(const std::string& text);
void RefreshChatDialog(HWND hWnd);
void RequestClearChatHistoryAsync();
void HandleChatSubmitUi(HWND hWnd, ChatDialogContext* ctx, const std::string& text);
void HandleChatClearUi(HWND hWnd, ChatDialogContext* ctx);

std::string GetLeftWorkAreaTabCaption()
{
	return "AI";
}

std::wstring GetLeftWorkAreaTabCaptionWide()
{
	return L"AI";
}

std::string GetLeftWorkAreaTabToolTip()
{
	return "AutoLinker AI 对话";
}

int EnsureLeftWorkAreaTabImageIndex(HWND tabHwnd)
{
	if (tabHwnd == nullptr || !IsWindow(tabHwnd)) {
		return -1;
	}

	HIMAGELIST imageList = TabCtrl_GetImageList(tabHwnd);
	if (imageList == nullptr) {
		return -1;
	}

	if (g_leftWorkAreaHost.imageIndex >= 0) {
		return g_leftWorkAreaHost.imageIndex;
	}

	const int addedIndex = ImageList_AddIcon(imageList, GetAppIconSmall());
	if (addedIndex < 0) {
		LogChatTab("add left work area tab icon failed");
		return -1;
	}

	g_leftWorkAreaHost.imageIndex = addedIndex;
	return addedIndex;
}

void LogChatTab(const std::string& text)
{
	OutputStringToELog("[AI Chat][Tab] " + text);
}

bool IsLeftWorkAreaTabControl(HWND tabHwnd)
{
	if (tabHwnd == nullptr || !IsWindow(tabHwnd)) {
		return false;
	}

	const int itemCount = static_cast<int>(SendMessageA(tabHwnd, TCM_GETITEMCOUNT, 0, 0));
	if (itemCount < 3) {
		return false;
	}

	return ReadTabItemTextLocal(tabHwnd, 0) == LocalFromWide(L"支持库") &&
		ReadTabItemTextLocal(tabHwnd, 1) == LocalFromWide(L"程序") &&
		ReadTabItemTextLocal(tabHwnd, 2) == LocalFromWide(L"属性");
}

HWND FindLeftWorkAreaTabControl()
{
	const auto tabs = CollectChildWindowsByClass(g_mainWindow, WC_TABCONTROLA);
	for (HWND tabHwnd : tabs) {
		if (IsLeftWorkAreaTabControl(tabHwnd)) {
			return tabHwnd;
		}
	}
	return nullptr;
}

void RestoreLeftWorkAreaNativeChildren()
{
	for (HWND childHwnd : g_leftWorkAreaHost.hiddenNativeChildren) {
		if (childHwnd != nullptr && IsWindow(childHwnd)) {
			ShowWindow(childHwnd, SW_SHOW);
		}
	}
	g_leftWorkAreaHost.hiddenNativeChildren.clear();
}

RECT GetLeftWorkAreaPageRectInHost()
{
	RECT rc = {};
	if (g_leftWorkAreaHost.tabHwnd == nullptr || g_leftWorkAreaHost.hostHwnd == nullptr) {
		return rc;
	}

	rc = CalcTabPageRectLocal(g_leftWorkAreaHost.tabHwnd);
	MapWindowPoints(g_leftWorkAreaHost.tabHwnd, g_leftWorkAreaHost.hostHwnd, reinterpret_cast<LPPOINT>(&rc), 2);
	return rc;
}

void LayoutLeftWorkAreaChatPage()
{
	if (g_chatDialog == nullptr || !IsWindow(g_chatDialog) ||
		g_leftWorkAreaHost.hostHwnd == nullptr || !IsWindow(g_leftWorkAreaHost.hostHwnd) ||
		g_leftWorkAreaHost.tabHwnd == nullptr || !IsWindow(g_leftWorkAreaHost.tabHwnd)) {
		return;
	}

	RECT rc = GetLeftWorkAreaPageRectInHost();
	const int width = (std::max)(120, static_cast<int>(rc.right - rc.left));
	const int height = (std::max)(120, static_cast<int>(rc.bottom - rc.top) - kLeftWorkAreaPageBottomInset);
	SetWindowPos(
		g_chatDialog,
		HWND_TOP,
		rc.left,
		rc.top,
		width,
		height,
		SWP_NOACTIVATE);
	PostMessageA(g_chatDialog, WM_AUTOLINKER_AI_CHAT_DEFER_LAYOUT, 0, 0);
}

void ShowLeftWorkAreaChatPage(bool focusInput)
{
	if (g_leftWorkAreaHost.hostHwnd == nullptr || !IsWindow(g_leftWorkAreaHost.hostHwnd) ||
		g_leftWorkAreaHost.tabHwnd == nullptr || !IsWindow(g_leftWorkAreaHost.tabHwnd) ||
		g_chatDialog == nullptr || !IsWindow(g_chatDialog)) {
		return;
	}

	if (GetParent(g_chatDialog) != g_leftWorkAreaHost.hostHwnd) {
		SetParent(g_chatDialog, g_leftWorkAreaHost.hostHwnd);
	}

	if (!g_leftWorkAreaHost.pageVisible) {
		for (HWND childHwnd = GetWindow(g_leftWorkAreaHost.hostHwnd, GW_CHILD);
			childHwnd != nullptr;
			childHwnd = GetWindow(childHwnd, GW_HWNDNEXT)) {
			if (childHwnd == g_leftWorkAreaHost.tabHwnd || childHwnd == g_chatDialog) {
				continue;
			}
			if (!IsWindowVisible(childHwnd)) {
				continue;
			}
			g_leftWorkAreaHost.hiddenNativeChildren.push_back(childHwnd);
			ShowWindow(childHwnd, SW_HIDE);
		}
	}

	g_leftWorkAreaHost.pageVisible = true;
	LayoutLeftWorkAreaChatPage();
	ShowWindow(g_chatDialog, SW_SHOW);
	if (focusInput) {
		FocusChatInputControl();
	}
}

void HideLeftWorkAreaChatPage(bool restoreNativeChildren)
{
	if (g_chatDialog != nullptr && IsWindow(g_chatDialog)) {
		ShowWindow(g_chatDialog, SW_HIDE);
	}
	g_leftWorkAreaHost.pageVisible = false;
	if (restoreNativeChildren) {
		RestoreLeftWorkAreaNativeChildren();
	}
}

LRESULT CALLBACK LeftWorkAreaHostSubclassProc(
	HWND hWnd,
	UINT uMsg,
	WPARAM wParam,
	LPARAM lParam,
	UINT_PTR uIdSubclass,
	DWORD_PTR dwRefData)
{
	(void)dwRefData;
	switch (uMsg)
	{
	case WM_NOTIFY: {
		const auto* hdr = reinterpret_cast<NMHDR*>(lParam);
		if (hdr != nullptr &&
			(hdr->code == TTN_GETDISPINFOA || hdr->code == TTN_NEEDTEXTA)) {
			auto* info = reinterpret_cast<NMTTDISPINFOA*>(lParam);
			if (info != nullptr && static_cast<int>(info->hdr.idFrom) == g_leftWorkAreaHost.tabIndex) {
				static thread_local std::string tooltipText;
				tooltipText = GetLeftWorkAreaTabToolTip();
				info->lpszText = const_cast<LPSTR>(tooltipText.c_str());
				return TRUE;
			}
		}
		if (hdr != nullptr &&
			hdr->hwndFrom == g_leftWorkAreaHost.tabHwnd &&
			hdr->code == TCN_SELCHANGE) {
			const int curSel = static_cast<int>(SendMessageA(g_leftWorkAreaHost.tabHwnd, TCM_GETCURSEL, 0, 0));
			if (curSel == g_leftWorkAreaHost.tabIndex) {
				ShowLeftWorkAreaChatPage(true);
				return 0;
			}

			const bool leavingChatPage = g_leftWorkAreaHost.pageVisible;
			if (leavingChatPage) {
				HideLeftWorkAreaChatPage(true);
			}
			return DefSubclassProc(hWnd, uMsg, wParam, lParam);
		}
		break;
	}

	case WM_SIZE:
	{
		// 先让原始过程运行以更新子控件（含标签页控件）尺寸，
		// 再读取标签页控件的最新尺寸来重新布局聊天页。
		const LRESULT defResult = DefSubclassProc(hWnd, uMsg, wParam, lParam);
		if (g_leftWorkAreaHost.pageVisible) {
			LayoutLeftWorkAreaChatPage();
		}
		return defResult;
	}

	case WM_NCDESTROY:
		g_leftWorkAreaHost.tabHwnd = nullptr;
		g_leftWorkAreaHost.hostHwnd = nullptr;
		g_leftWorkAreaHost.tabIndex = -1;
		g_leftWorkAreaHost.imageIndex = -1;
		g_leftWorkAreaHost.pageVisible = false;
		g_leftWorkAreaHost.hiddenNativeChildren.clear();
		g_leftWorkAreaHost.subclassInstalled = false;
		RemoveWindowSubclass(hWnd, LeftWorkAreaHostSubclassProc, uIdSubclass);
		break;

	default:
		break;
	}
	return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

bool EnsureLeftWorkAreaChatTabAdded()
{
	if (!EnsureChatHostWindowCreated()) {
		return false;
	}

	HWND tabHwnd = g_leftWorkAreaHost.tabHwnd;
	if (tabHwnd == nullptr || !IsWindow(tabHwnd) || !IsLeftWorkAreaTabControl(tabHwnd)) {
		tabHwnd = FindLeftWorkAreaTabControl();
	}
	if (tabHwnd == nullptr || !IsWindow(tabHwnd)) {
		LogChatTab("left work area tab not found");
		return false;
	}

	HWND hostHwnd = GetParent(tabHwnd);
	if (hostHwnd == nullptr || !IsWindow(hostHwnd)) {
		LogChatTab("left work area host invalid");
		return false;
	}

	const std::string caption = GetLeftWorkAreaTabCaption();
	int tabIndex = -1;
	const int itemCount = static_cast<int>(SendMessageA(tabHwnd, TCM_GETITEMCOUNT, 0, 0));
	for (int index = 0; index < itemCount; ++index) {
		if (ReadTabItemTextLocal(tabHwnd, index) == caption) {
			tabIndex = index;
			break;
		}
	}

	if (tabIndex < 0) {
		std::wstring captionWide = GetLeftWorkAreaTabCaptionWide();
		TCITEMW item = {};
		item.mask = TCIF_TEXT;
		item.pszText = captionWide.data();
		const int imageIndex = EnsureLeftWorkAreaTabImageIndex(tabHwnd);
		if (imageIndex >= 0) {
			item.mask |= TCIF_IMAGE;
			item.iImage = imageIndex;
		}
		const LRESULT insertIndex = SendMessageW(tabHwnd, TCM_INSERTITEMW, static_cast<WPARAM>(itemCount), reinterpret_cast<LPARAM>(&item));
		if (insertIndex < 0) {
			LogChatTab("insert left work area tab failed");
			return false;
		}
		tabIndex = static_cast<int>(insertIndex);
	}

	g_leftWorkAreaHost.tabHwnd = tabHwnd;
	g_leftWorkAreaHost.hostHwnd = hostHwnd;
	g_leftWorkAreaHost.tabIndex = tabIndex;
	if (g_leftWorkAreaHost.imageIndex < 0) {
		g_leftWorkAreaHost.imageIndex = EnsureLeftWorkAreaTabImageIndex(tabHwnd);
	}
	if (g_leftWorkAreaHost.imageIndex >= 0) {
		TCITEMW item = {};
		item.mask = TCIF_IMAGE;
		item.iImage = g_leftWorkAreaHost.imageIndex;
		SendMessageW(tabHwnd, TCM_SETITEMW, static_cast<WPARAM>(tabIndex), reinterpret_cast<LPARAM>(&item));
	}
	if (!g_leftWorkAreaHost.subclassInstalled) {
		if (SetWindowSubclass(hostHwnd, LeftWorkAreaHostSubclassProc, kLeftWorkAreaHostSubclassId, 0) == FALSE) {
			LogChatTab("subclass left work area host failed");
			return false;
		}
		g_leftWorkAreaHost.subclassInstalled = true;
	}

	if (GetParent(g_chatDialog) != hostHwnd) {
		SetParent(g_chatDialog, hostHwnd);
	}
	ShowWindow(g_chatDialog, SW_HIDE);
	g_chatHostMode = ChatHostMode::LeftWorkArea;
	g_chatTabAdded = true;
	return true;
}

bool EnsureChatHostWindowCreated()
{
	if (g_mainWindow == nullptr || !IsWindow(g_mainWindow)) {
		return false;
	}
	if (g_chatDialog != nullptr && IsWindow(g_chatDialog)) {
		return true;
	}

	ComCtl6ActivationScope themeScope;
	INITCOMMONCONTROLSEX icex = {};
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_STANDARD_CLASSES | ICC_WIN95_CLASSES | ICC_LINK_CLASS;
	InitCommonControlsEx(&icex);

	WNDCLASSEXA wc = {};
	wc.cbSize = sizeof(wc);
	wc.lpfnWndProc = AIChatDialogProc;
	wc.hInstance = GetModuleHandleA(nullptr);
	wc.lpszClassName = "AutoLinkerAIChatDialogWindow";
	wc.hIcon = GetAppIconLarge();
	wc.hIconSm = GetAppIconSmall();
	wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_BTNFACE + 1);
	RegisterClassExA(&wc);

	g_chatDialog = CreateWindowExA(
		WS_EX_CONTROLPARENT,
		wc.lpszClassName,
		"",
		WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
		0, 0, 860, 680,
		g_mainWindow,
		nullptr,
		wc.hInstance,
		nullptr);
	if (g_chatDialog == nullptr) {
		LogChatTab("create host failed");
		return false;
	}

	PostMessageA(g_chatDialog, WM_AUTOLINKER_AI_CHAT_DEFER_LAYOUT, 0, 0);
	return true;
}

bool EnsureChatTabAddedInternal()
{
	if (!EnsureChatHostWindowCreated()) {
		return false;
	}
	if (g_chatTabAdded) {
		return true;
	}

	if (EnsureLeftWorkAreaChatTabAdded()) {
		return true;
	}

	const std::string caption = GetChatTabCaption();
	const bool ok = IDEFacade::Instance().AddOutputTab(g_chatDialog, caption, "AutoLinker AI Chat", GetAppIconSmall());
	if (!ok) {
		LogChatTab("add tab failed");
		return false;
	}

	g_chatHostMode = ChatHostMode::OutputTab;
	g_chatTabAdded = true;
	LogChatTab("add tab ok");
	PostMessageA(g_chatDialog, WM_AUTOLINKER_AI_CHAT_DEFER_LAYOUT, 0, 0);
	return true;
}

void FocusChatInputControl()
{
	if (g_chatDialog == nullptr || !IsWindow(g_chatDialog)) {
		return;
	}
	auto* ctx = reinterpret_cast<ChatDialogContext*>(GetWindowLongPtrA(g_chatDialog, GWLP_USERDATA));
	if (ctx != nullptr && ctx->webViewDesired) {
		FocusWebViewInput(ctx);
		return;
	}
	if (ctx != nullptr && ctx->hInput != nullptr) {
		SetFocus(ctx->hInput);
	}
}

namespace AIChatFeature {
void Initialize(HWND mainWindow, ConfigManager* configManager, AIJsonConfig* aiJsonConfig)
{
	g_mainWindow = mainWindow;
	g_configManager = configManager;
	g_aiJsonConfig = aiJsonConfig;
	g_msgAIChatDone = RegisterWindowMessageA("AutoLinker.AIChat.Done");
	g_msgAIChatToolDialog = RegisterWindowMessageA("AutoLinker.AIChat.ToolDialog");
	g_msgAIChatToolExec = RegisterWindowMessageA("AutoLinker.AIChat.ToolExec");
	EnsureTabCreated();
}

void Shutdown()
{
	if (g_leftWorkAreaHost.hostHwnd != nullptr && g_leftWorkAreaHost.subclassInstalled && IsWindow(g_leftWorkAreaHost.hostHwnd)) {
		RemoveWindowSubclass(g_leftWorkAreaHost.hostHwnd, LeftWorkAreaHostSubclassProc, kLeftWorkAreaHostSubclassId);
	}
	HideLeftWorkAreaChatPage(true);
	if (g_leftWorkAreaHost.tabHwnd != nullptr &&
		IsWindow(g_leftWorkAreaHost.tabHwnd) &&
		g_leftWorkAreaHost.tabIndex >= 0 &&
		ReadTabItemTextLocal(g_leftWorkAreaHost.tabHwnd, g_leftWorkAreaHost.tabIndex) == GetLeftWorkAreaTabCaption()) {
		SendMessageA(g_leftWorkAreaHost.tabHwnd, TCM_DELETEITEM, static_cast<WPARAM>(g_leftWorkAreaHost.tabIndex), 0);
	}
	g_leftWorkAreaHost = {};
	g_chatHostMode = ChatHostMode::None;
	if (g_chatDialog != nullptr && IsWindow(g_chatDialog)) {
		DestroyWindow(g_chatDialog);
		g_chatDialog = nullptr;
	}
	g_chatTabAdded = false;
}

void EnsureTabCreated()
{
	if (g_mainWindow == nullptr || !IsWindow(g_mainWindow) || g_configManager == nullptr) {
		OutputStringToELog("[AI Chat] Not initialized yet, please retry.");
		return;
	}

	if (!EnsureChatTabAddedInternal()) {
		return;
	}
}

void ActivateTab()
{
	EnsureTabCreated();
	if (g_chatDialog == nullptr || !IsWindow(g_chatDialog)) {
		return;
	}

	if (g_chatHostMode == ChatHostMode::LeftWorkArea &&
		g_leftWorkAreaHost.tabHwnd != nullptr &&
		IsWindow(g_leftWorkAreaHost.tabHwnd) &&
		g_leftWorkAreaHost.tabIndex >= 0) {
		SendMessageA(g_leftWorkAreaHost.tabHwnd, TCM_SETCURSEL, static_cast<WPARAM>(g_leftWorkAreaHost.tabIndex), 0);
		ShowLeftWorkAreaChatPage(true);
	} else {
		ShowWindow(g_chatDialog, SW_SHOW);
		PostMessageA(g_chatDialog, WM_AUTOLINKER_AI_CHAT_DEFER_LAYOUT, 0, 0);
		FocusChatInputControl();
	}
	LogChatTab("activate");
}

void OpenDialog()
{
	ActivateTab();
}

void OnCurrentSourceFilePathChanged(const std::string& previousPath, const std::string& currentPath)
{
	(void)previousPath;
	(void)currentPath;
	WorkspaceMirror::ResetAndCleanup();
	RebindChatSessionToCurrentSourceIfNeeded();
}

bool HandleMainWindowMessage(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	(void)wParam;
	if (g_msgAIChatDone != 0 && uMsg == g_msgAIChatDone) {
		HandleChatTaskDone(lParam);
		return true;
	}
	if (g_msgAIChatToolDialog != 0 && uMsg == g_msgAIChatToolDialog) {
		return HandleToolDialogRequest(lParam);
	}
	if (g_msgAIChatToolExec != 0 && uMsg == g_msgAIChatToolExec) {
		return HandleToolExecRequest(lParam);
	}
	return false;
}

void SetUpdateAvailable(const std::string& latestVersion)
{
	(void)latestVersion;
	g_updateAvailable.store(true);
	if (g_chatDialog != nullptr && IsWindow(g_chatDialog)) {
		PostMessageA(g_chatDialog, WM_AUTOLINKER_AI_CHAT_UPDATE_TAG, 0, 0);
	}
}

bool ExecutePublicTool(const std::string& toolName, const std::string& argumentsJson, std::string& outResultJsonUtf8, bool& outOk)
{
	outResultJsonUtf8.clear();
	outOk = false;
	if (TrimAsciiCopy(toolName).empty()) {
		return false;
	}

	bool toolOk = false;
	const std::string resultJsonLocal = ExecuteToolCall(toolName, argumentsJson, toolOk, false);
	outResultJsonUtf8 = LocalToUtf8Text(resultJsonLocal);
	outOk = toolOk;
	return true;
}
} // namespace AIChatFeature
