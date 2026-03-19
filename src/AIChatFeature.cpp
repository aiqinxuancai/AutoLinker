#include "AIChatFeature.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <cctype>
#include <chrono>
#include <CommCtrl.h>
#include <condition_variable>
#include <cstring>
#include <filesystem>
#include <fstream>
#include <memory>
#include <mutex>
#include <new>
#include <process.h>
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
#include "AIService.h"
#include "ConfigManager.h"
#include "AIChatTooling.h"
#include "AIChatToolingInternal.h"
#include "Global.h"
#include "IDEFacade.h"
#include "PathHelper.h"
#include "resource.h"
#include "..\\elib\\lib2.h"
#if defined(_M_IX86)
#include "direct_global_search_debug.hpp"
#include "native_module_public_info.hpp"
#endif

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "advapi32.lib")
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
constexpr UINT_PTR kHistoryWebViewFlushTimerId = 0xA17;

constexpr int IDC_AI_CHAT_HISTORY = 2101;
constexpr int IDC_AI_CHAT_INPUT = 2102;
constexpr int IDC_AI_CHAT_SEND = 32553;
constexpr int IDC_AI_CHAT_CLEAR_HISTORY = 32554;

constexpr int IDC_CODE_EDIT = 2201;
constexpr int IDC_CODE_OK = 1;
constexpr int IDC_CODE_CANCEL = 2;
constexpr int IDC_CODE_COPY = 2204;
constexpr UINT_PTR kEditSubclassId = 1;
constexpr UINT_PTR kActionControlSubclassId = 2;
constexpr UINT_PTR kLeftWorkAreaHostSubclassId = 3;
constexpr DWORD_PTR kEditFlagNone = 0;
constexpr DWORD_PTR kEditFlagSubmitOnEnter = 1;
constexpr UINT_PTR kActionSubmit = 1;
constexpr UINT_PTR kActionClear = 2;

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
};

struct AIChatSessionState {
	std::mutex mutex;
	std::vector<SessionMessage> messages;
	std::string rollingSummary;
	std::string streamingAssistantPreview;
	bool requestInFlight = false;
	unsigned long long activeRequestId = 0;
	unsigned long long nextRequestId = 1;
};

struct ChatDialogContext {
	HWND hHistory = nullptr;
	HWND hHistoryHost = nullptr;
	HWND hInput = nullptr;
	HWND hSend = nullptr;
	HWND hClearHistory = nullptr;
	int inputRowsVisible = 1;
	bool webViewDesired = false;
	bool webViewReady = false;
	bool webViewContentReady = false;
	bool webViewFlushScheduled = false;
	std::string pendingHistoryHtml;
	Microsoft::WRL::ComPtr<ICoreWebView2Environment> webViewEnvironment;
	Microsoft::WRL::ComPtr<ICoreWebView2Controller> webViewController;
	Microsoft::WRL::ComPtr<ICoreWebView2> webView;
};

struct CodeEditDialogContext {
	std::string title;
	std::string text;
	bool accepted = false;
	HWND hEdit = nullptr;
	HWND hCopy = nullptr;
	HWND hOk = nullptr;
	HWND hCancel = nullptr;
};

struct CodeEditToolRequest {
	std::string title;
	std::string hint;
	std::string initialCode;
	std::string resultCode;
	bool accepted = false;
	bool done = false;
	std::mutex mutex;
	std::condition_variable cv;
};

struct AIChatAsyncRequest {
	unsigned long long requestId = 0;
	AISettings settings = {};
	std::vector<AIChatMessage> contextMessages;
};

struct AIChatAsyncResult {
	unsigned long long requestId = 0;
	AIChatResult chatResult = {};
};

struct ChatHistoryGarbage {
	std::vector<SessionMessage> messages;
	std::string rollingSummary;
};

HWND g_mainWindow = nullptr;
ConfigManager* g_configManager = nullptr;
HWND g_chatDialog = nullptr;
bool g_chatTabAdded = false;
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

std::string ReadTabItemTextA(HWND tabHwnd, int index)
{
	if (tabHwnd == nullptr || !IsWindow(tabHwnd) || index < 0) {
		return std::string();
	}

	char textBuf[256] = {};
	TCITEMA item = {};
	item.mask = TCIF_TEXT;
	item.pszText = textBuf;
	item.cchTextMax = static_cast<int>(sizeof(textBuf));
	if (SendMessageA(tabHwnd, TCM_GETITEMA, static_cast<WPARAM>(index), reinterpret_cast<LPARAM>(&item)) == FALSE) {
		return std::string();
	}
	return textBuf;
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
	OutputStringToELog(std::format(
		"[AI Chat][WebView2] runtime available={}",
		g_webView2RuntimeAvailable ? 1 : 0));
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
		if (ctrlDown && (wParam == 'A' || wParam == 'a')) {
			SendMessageA(hWnd, EM_SETSEL, 0, -1);
			return 0;
		}

		if ((dwRefData & kEditFlagSubmitOnEnter) != 0 && wParam == VK_RETURN) {
			if (ctrlDown) {
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
		OutputStringToELog("[AI Chat][UI] click action: clear");
		PostMessageA(hParent, WM_AUTOLINKER_AI_CHAT_CLEAR, 0, 0);
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
		if (ctx->hInput != nullptr) {
			ShowWindow(ctx->hInput, SW_HIDE);
		}
		if (ctx->hSend != nullptr) {
			ShowWindow(ctx->hSend, SW_HIDE);
		}
		return;
	}

	const int margin = 0;
	const int bottomMargin = 4;
	const int gap = 6;
	const int actionRowHeight = 22;
	const int inputHeightSingle = 30;
	const int inputHeightDouble = 54;
	const int sendWidth = 92;
	const int clearHistoryWidth = 98;
	const int inputRowsVisible = ctx->inputRowsVisible >= 2 ? 2 : 1;
	const int inputHeight = inputRowsVisible >= 2 ? inputHeightDouble : inputHeightSingle;

	const int contentWidth = (std::max)(120, clientWidth - margin * 2);
	const int inputWidth = (std::max)(80, contentWidth - sendWidth - gap);
	const int sendX = margin + inputWidth + gap;
	const int inputY = clientHeight - bottomMargin - inputHeight;
	const int actionRowY = inputY - gap - actionRowHeight;
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
		ShowWindow(ctx->hClearHistory, SW_SHOW);
		MoveWindow(ctx->hClearHistory, margin, actionRowY, clearHistoryWidth, actionRowHeight, TRUE);
	}
	if (ctx->hInput != nullptr) {
		ShowWindow(ctx->hInput, SW_SHOW);
		MoveWindow(ctx->hInput, margin, inputY, inputWidth, inputHeight, TRUE);
	}
	if (ctx->hSend != nullptr) {
		ShowWindow(ctx->hSend, SW_SHOW);
		MoveWindow(ctx->hSend, sendX, inputY, sendWidth, inputHeight, TRUE);
	}
}

void LayoutCodeEditDialog(HWND hWnd, CodeEditDialogContext* ctx)
{
	if (hWnd == nullptr || ctx == nullptr) {
		return;
	}

	RECT rc = {};
	GetClientRect(hWnd, &rc);
	const int clientWidth = rc.right - rc.left;
	const int clientHeight = rc.bottom - rc.top;

	const int margin = 14;
	const int gap = 8;
	const int buttonHeight = 30;
	const int copyWidth = 116;
	const int okWidth = 84;
	const int cancelWidth = 56;

	const int buttonY = clientHeight - margin - buttonHeight;
	const int cancelX = clientWidth - margin - cancelWidth;
	const int okX = cancelX - gap - okWidth;
	const int copyX = okX - gap - copyWidth;

	const int editX = margin;
	const int editY = margin;
	const int editWidth = (std::max)(120, clientWidth - margin * 2);
	const int editHeight = (std::max)(100, buttonY - gap - editY);

	if (ctx->hEdit != nullptr) {
		MoveWindow(ctx->hEdit, editX, editY, editWidth, editHeight, TRUE);
	}
	if (ctx->hCopy != nullptr) {
		MoveWindow(ctx->hCopy, copyX, buttonY, copyWidth, buttonHeight, TRUE);
	}
	if (ctx->hOk != nullptr) {
		MoveWindow(ctx->hOk, okX, buttonY, okWidth, buttonHeight, TRUE);
	}
	if (ctx->hCancel != nullptr) {
		MoveWindow(ctx->hCancel, cancelX, buttonY, cancelWidth, buttonHeight, TRUE);
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

void UpdateWebViewComposerState(ChatDialogContext* ctx, bool busy)
{
	if (ctx == nullptr) {
		return;
	}
	std::wstring script = L"window.autolinkerSetBusy(";
	script += busy ? L"true" : L"false";
	script += L");";
	ExecuteWebViewScript(ctx, script);
}

void ClearWebViewInput(ChatDialogContext* ctx)
{
	ExecuteWebViewScript(ctx, L"window.autolinkerClearInput();");
}

void FocusWebViewInput(ChatDialogContext* ctx)
{
	ExecuteWebViewScript(ctx, L"window.autolinkerFocusInput();");
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
	const HRESULT hr = CreateCoreWebView2EnvironmentWithOptions(
		nullptr,
		nullptr,
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
												else if (action == "clear") {
													HandleChatClearUi(hWnd, msgCtx);
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
												{
													std::lock_guard<std::mutex> guard(g_session.mutex);
													inFlight = g_session.requestInFlight;
												}
												UpdateWebViewComposerState(navCtx, inFlight);
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
							OutputStringToELog(std::format("[AI Chat][WebView2] controller ready={}", innerCtx->webViewReady ? 1 : 0));
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

bool SetClipboardTextSimple(const std::string& text)
{
	if (text.empty()) {
		return false;
	}
	if (!OpenClipboard(nullptr)) {
		return false;
	}

	bool ok = false;
	if (EmptyClipboard()) {
		const size_t bytes = text.size() + 1;
		HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, bytes);
		if (hMem != nullptr) {
			void* ptr = GlobalLock(hMem);
			if (ptr != nullptr) {
				std::memcpy(ptr, text.c_str(), bytes);
				GlobalUnlock(hMem);
				ok = (SetClipboardData(CF_TEXT, hMem) != nullptr);
			}
			if (!ok) {
				GlobalFree(hMem);
			}
		}
	}
	CloseClipboard();
	return ok;
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

	for (const std::string& line : lines) {
		const std::string trimmed = TrimAsciiCopy(line);
		if (line.rfind("```", 0) == 0) {
			closeParagraph();
			closeList();
			if (!inCodeBlock) {
				html += "<pre><code>";
				inCodeBlock = true;
			}
			else {
				html += "</code></pre>";
				inCodeBlock = false;
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
	if (inCodeBlock) {
		html += "</code></pre>";
	}
	return html;
}

void CompactHistoryLocked(AIChatSessionState& state)
{
	size_t contextChars = 0;
	for (const auto& message : state.messages) {
		if (message.includeInContext) {
			contextChars += message.content.size();
		}
	}

	if (state.messages.size() <= 40 && contextChars <= 24000) {
		return;
	}

	const size_t keepCount = (std::min)(state.messages.size(), static_cast<size_t>(24));
	const size_t cutCount = state.messages.size() - keepCount;
	if (cutCount == 0) {
		return;
	}

	std::string summaryAppend;
	summaryAppend.reserve(2048);
	for (size_t i = 0; i < cutCount; ++i) {
		const auto& msg = state.messages[i];
		if (!msg.includeInContext) {
			continue;
		}

		std::string line = TrimAsciiCopy(msg.content);
		if (line.size() > 120) {
			line.resize(120);
			line += "...";
		}
		summaryAppend += "[";
		summaryAppend += RoleLabel(msg.role);
		summaryAppend += "] ";
		summaryAppend += line;
		summaryAppend += "\n";
		if (summaryAppend.size() > 4000) {
			break;
		}
	}

	state.messages.erase(state.messages.begin(), state.messages.begin() + static_cast<std::ptrdiff_t>(cutCount));
	if (!summaryAppend.empty()) {
		if (!state.rollingSummary.empty()) {
			state.rollingSummary += "\n";
		}
		state.rollingSummary += summaryAppend;
		if (state.rollingSummary.size() > 12000) {
			state.rollingSummary.erase(0, state.rollingSummary.size() - 12000);
		}
	}
}

std::vector<AIChatMessage> BuildContextMessagesLocked(const AIChatSessionState& state)
{
	std::vector<AIChatMessage> out;
	out.reserve(40);

	if (!TrimAsciiCopy(state.rollingSummary).empty()) {
		out.push_back(AIChatMessage{
			"system",
			LocalFromWide(L"\u5386\u53f2\u6458\u8981\uff08\u81ea\u52a8\u538b\u7f29\uff09\uff1a\n") + state.rollingSummary
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
		out.push_back(AIChatMessage{ role, msg.content });
	}
	return out;
}

std::string BuildHistoryHtmlLocked(const AIChatSessionState& state)
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

	std::string body;
	body.reserve(4096);
	for (const auto& msg : state.messages) {
		appendMessageCard(
			body,
			msg.role,
			RoleLabel(msg.role),
			msg.content,
			msg.role == SessionRole::Assistant);
	}

	if (state.requestInFlight) {
		const std::string preview = TrimAsciiCopy(state.streamingAssistantPreview);
		if (!preview.empty()) {
			appendMessageCard(body, SessionRole::Assistant, "AI", state.streamingAssistantPreview, true);
			appendMessageCard(body, SessionRole::System, LocalFromWide(L"系统"), LocalFromWide(L"AI 正在生成（流式）..."), false);
		}
		else {
			appendMessageCard(body, SessionRole::System, LocalFromWide(L"系统"), LocalFromWide(L"等待 AI 返回..."), false);
		}
	}

	if (TrimAsciiCopy(body).empty()) {
		appendMessageCard(body, SessionRole::System, LocalFromWide(L"系统"), LocalFromWide(L"等待开始对话..."), false);
	}

	std::string html;
	return body;
}

std::string BuildHistoryWebViewShellHtml()
{
	std::string html;
	html.reserve(2600);
	html += "<!doctype html><html><head><meta charset=\"utf-8\"><style>";
	html += "html,body{margin:0;padding:0;background:#ffffff;color:#222;font:13px/1.6 'Microsoft YaHei UI','Segoe UI',sans-serif;}";
	html += "body{padding:0;}";
	html += ".layout{display:flex;flex-direction:column;height:100vh;box-sizing:border-box;}";
	html += ".history{flex:1 1 auto;overflow:auto;padding:10px 10px 8px 10px;box-sizing:border-box;}";
	html += "#chat-root{min-height:100%;}";
	html += ".composer{flex:0 0 auto;border-top:1px solid #d9d9d9;background:#fbfbfb;padding:8px 10px 10px 10px;box-sizing:border-box;}";
	html += ".actions{display:flex;justify-content:space-between;align-items:flex-end;gap:8px;}";
	html += ".left-actions,.right-actions{display:flex;align-items:center;gap:8px;}";
	html += ".right-actions{margin-left:auto;}";
	html += ".input-wrap{display:flex;flex-direction:column;gap:8px;}";
	html += "#chat-input{width:100%;min-height:58px;max-height:160px;resize:vertical;box-sizing:border-box;border:1px solid #cfcfcf;border-radius:6px;padding:8px 10px;font:13px/1.5 'Microsoft YaHei UI','Segoe UI',sans-serif;outline:none;background:#fff;}";
	html += "#chat-input:focus{border-color:#7aa7e0;box-shadow:0 0 0 2px rgba(64,127,214,.12);}";
	html += ".btn{border:1px solid #c8c8c8;background:#fff;border-radius:6px;padding:6px 12px;font:12px 'Microsoft YaHei UI','Segoe UI',sans-serif;cursor:pointer;}";
	html += ".btn.primary{background:#0b63c9;border-color:#0b63c9;color:#fff;}";
	html += ".btn:disabled{opacity:.55;cursor:default;}";
	html += ".hint{font-size:12px;color:#6b6b6b;}";
	html += ".msg{border:1px solid #d8d8d8;border-radius:6px;padding:8px 10px;margin:0 0 10px 0;box-sizing:border-box;overflow:hidden;}";
	html += ".msg.user{background:#f7fbff;border-color:#c8ddf5;}";
	html += ".msg.assistant{background:#ffffff;border-color:#d8d8d8;}";
	html += ".msg.tool{background:#fbfbfb;border-color:#dddddd;}";
	html += ".msg.system{background:#fff9e8;border-color:#ecd9a1;}";
	html += ".role{font-size:12px;font-weight:700;color:#5a5a5a;margin-bottom:6px;}";
	html += ".body{overflow-wrap:anywhere;word-break:break-word;}";
	html += ".body p,.body ul,.body pre,.body h1,.body h2,.body h3,.body h4,.body h5,.body h6{margin:0 0 8px 0;}";
	html += ".body ul{padding-left:20px;}";
	html += ".body code{font-family:Consolas,'Courier New',monospace;background:#f2f2f2;border-radius:4px;padding:1px 4px;overflow-wrap:anywhere;word-break:break-word;}";
	html += ".body pre{overflow:auto;background:#f6f8fa;border-radius:6px;padding:10px;white-space:pre-wrap;word-break:break-word;overflow-wrap:anywhere;}";
	html += ".body pre code{background:transparent;padding:0;border-radius:0;}";
	html += ".body a{color:#0b63c9;text-decoration:none;}";
	html += ".body a:hover{text-decoration:underline;}";
	html += "</style></head><body><div class=\"layout\"><div id=\"history-scroll\" class=\"history\"><div id=\"chat-root\"></div></div><div class=\"composer\"><div class=\"input-wrap\"><textarea id=\"chat-input\" placeholder=\"输入消息，Enter 发送，Ctrl+Enter 换行\"></textarea><div class=\"actions\"><div class=\"left-actions\"><button id=\"clear-btn\" class=\"btn\" type=\"button\">清空历史会话</button></div><div class=\"right-actions\"><button id=\"send-btn\" class=\"btn primary\" type=\"button\">发送</button></div></div></div></div></div><script>";
	html += "window.autolinkerSetChatHtml=function(html){var root=document.getElementById('chat-root');var scroll=document.getElementById('history-scroll');if(root){root.innerHTML=html||'';if(scroll){scroll.scrollTop=scroll.scrollHeight;}}};";
	html += "window.autolinkerSetBusy=function(busy){var send=document.getElementById('send-btn');var clear=document.getElementById('clear-btn');var input=document.getElementById('chat-input');if(send){send.disabled=!!busy;}if(clear){clear.disabled=!!busy;}if(input){input.disabled=!!busy;}};";
	html += "window.autolinkerClearInput=function(){var input=document.getElementById('chat-input');if(input){input.value='';input.focus();}};";
	html += "window.autolinkerFocusInput=function(){var input=document.getElementById('chat-input');if(input){input.focus();}};";
	html += "(function(){var input=document.getElementById('chat-input');var send=document.getElementById('send-btn');var clear=document.getElementById('clear-btn');";
	html += "function post(obj){if(window.chrome&&window.chrome.webview){window.chrome.webview.postMessage(JSON.stringify(obj));}}";
	html += "function doSend(){if(!input||input.disabled){return;} post({action:'submit',text:input.value||''});}";
	html += "if(send){send.addEventListener('click',doSend);} if(clear){clear.addEventListener('click',function(){post({action:'clear'});});}";
	html += "if(input){input.addEventListener('keydown',function(e){if(e.key==='Enter'&&!e.ctrlKey){e.preventDefault();doSend();}}); setTimeout(function(){input.focus();},0);}})();";
	html += "</script></body></html>";
	return html;
}

std::string BuildHistoryTextLocked(const AIChatSessionState& state)
{
	std::string text;
	for (const auto& msg : state.messages) {
		text += "[";
		text += RoleLabel(msg.role);
		text += "]\r\n";
		text += msg.content;
		text += "\r\n\r\n";
	}
	if (state.requestInFlight) {
		const std::string preview = TrimAsciiCopy(state.streamingAssistantPreview);
		if (!preview.empty()) {
			text += "[AI]\r\n";
			text += state.streamingAssistantPreview;
			text += "\r\n\r\n";
			text += "[" + LocalFromWide(L"\u7cfb\u7edf") + "]\r\n"
				+ LocalFromWide(L"AI \u6b63\u5728\u751f\u6210\uff08\u6d41\u5f0f\uff09...") + "\r\n";
		}
		else {
			text += "[" + LocalFromWide(L"\u7cfb\u7edf") + "]\r\n"
				+ LocalFromWide(L"\u7b49\u5f85 AI \u8fd4\u56de...") + "\r\n";
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

void RecoverInFlightIfNeeded(const std::string& reason)
{
	std::lock_guard<std::mutex> guard(g_session.mutex);
	if (!g_session.requestInFlight) {
		return;
	}
	g_session.requestInFlight = false;
	g_session.activeRequestId = 0;
	g_session.streamingAssistantPreview.clear();
	g_session.messages.push_back(SessionMessage{
		SessionRole::System,
		"Chat request auto-recovered: " + reason,
		false
	});
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
	if (g_configManager == nullptr) {
		return false;
	}

	AIService::LoadSettings(*g_configManager, settings);
	std::string missing;
	if (AIService::HasRequiredSettings(settings, missing)) {
		return true;
	}

    OutputStringToELog("[AI Chat] AI settings missing, opening config dialog");
	if (!ShowAIConfigDialog(g_mainWindow, settings)) {
        OutputStringToELog("[AI Chat] AI config cancelled");
		return false;
	}
	AIService::SaveSettings(*g_configManager, settings);
    OutputStringToELog("[AI Chat] AI config saved");
	return true;
}

bool RequestCodeEditFromMainThread(
	const std::string& title,
	const std::string& hint,
	const std::string& initialCode,
	std::string& outCode)
{
	outCode.clear();
	if (g_mainWindow == nullptr || !IsWindow(g_mainWindow)) {
		return false;
	}

	CodeEditToolRequest request = {};
	request.title = title;
	request.hint = hint;
	request.initialCode = initialCode;
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
	if (!request.accepted) {
		return false;
	}

	outCode = request.resultCode;
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
		result->chatResult = AIService::ExecuteChatWithTools(
			request->contextMessages,
			request->settings,
			[](const std::string& toolName, const std::string& argumentsJson, bool& outOk) -> std::string {
				return ExecuteToolCall(toolName, argumentsJson, outOk);
			},
			[requestId = request->requestId](const std::string& deltaText) {
				AppendStreamingAssistantDelta(requestId, deltaText);
			});
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

		g_session.messages.push_back(SessionMessage{ SessionRole::User, trimmed, true });
		CompactHistoryLocked(g_session);

		request->requestId = g_session.nextRequestId++;
		request->settings = settings;
		request->contextMessages = BuildContextMessagesLocked(g_session);
		g_session.streamingAssistantPreview.clear();
		g_session.requestInFlight = true;
		g_session.activeRequestId = request->requestId;
	}

	const uintptr_t threadId = _beginthread(RunAIChatWorker, 0, request.get());
	if (threadId == static_cast<uintptr_t>(-1L)) {
		std::lock_guard<std::mutex> guard(g_session.mutex);
		g_session.requestInFlight = false;
		g_session.activeRequestId = 0;
		g_session.streamingAssistantPreview.clear();
        g_session.messages.push_back(SessionMessage{ SessionRole::System, "Failed to start background chat task.", false });
		PostRefreshDialog();
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

	const std::string trimmed = TrimAsciiCopy(text);
	if (trimmed.empty()) {
		if (ctx->webViewDesired) {
			FocusWebViewInput(ctx);
		}
		else if (ctx->hInput != nullptr) {
			SetFocus(ctx->hInput);
		}
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
	if (!ctx->webViewDesired && ctx->hHistory != nullptr) {
		SetWindowTextA(ctx->hHistory, "");
	}
	RequestClearChatHistoryAsync();
	UpdateWebViewComposerState(ctx, false);
	if (ctx->webViewDesired) {
		FocusWebViewInput(ctx);
	}
	else if (ctx->hInput != nullptr) {
		SetFocus(ctx->hInput);
	}
	(void)hWnd;
}

void ClearChatHistory()
{
	std::vector<SessionMessage> oldMessages;
	std::string oldSummary;
	{
		std::lock_guard<std::mutex> guard(g_session.mutex);
		oldMessages.swap(g_session.messages);
		oldSummary.swap(g_session.rollingSummary);
		g_session.streamingAssistantPreview.clear();
	}
}

void RequestClearChatHistoryAsync()
{
	if (g_clearHistoryInProgress.exchange(true)) {
		return;
	}

	const uintptr_t tid = _beginthread(
		[](void*) {
			ClearChatHistory();
			g_clearHistoryInProgress = false;
			PostRefreshDialog();
		},
		0,
		nullptr);
	if (tid == static_cast<uintptr_t>(-1L)) {
		// Thread creation failed: fall back to synchronous clear.
		ClearChatHistory();
		g_clearHistoryInProgress = false;
		PostRefreshDialog();
	}
}

void HandleChatTaskDone(LPARAM lParam)
{
	std::unique_ptr<AIChatAsyncResult> result(reinterpret_cast<AIChatAsyncResult*>(lParam));
	if (!result) {
		return;
	}

	std::lock_guard<std::mutex> guard(g_session.mutex);
	if (!g_session.requestInFlight || g_session.activeRequestId != result->requestId) {
		return;
	}

	g_session.requestInFlight = false;
	g_session.activeRequestId = 0;
	g_session.streamingAssistantPreview.clear();

	for (const auto& evt : result->chatResult.toolEvents) {
		std::string line =
			LocalFromWide(L"\u8c03\u7528 ") + evt.name +
			LocalFromWide(L"\uff0c\u8fd4\u56de\uff1a") + evt.resultJson;
		if (line.size() > 1200) {
			line.resize(1200);
			line += "...";
		}
		g_session.messages.push_back(SessionMessage{ SessionRole::Tool, line, false });
	}

	if (result->chatResult.ok) {
		g_session.messages.push_back(SessionMessage{
			SessionRole::Assistant,
			NormalizeCodeForEIDE(result->chatResult.content),
			true
		});
	}
	else {
		const std::string err = result->chatResult.error.empty()
			? LocalFromWide(L"AI\u5bf9\u8bdd\u5931\u8d25")
			: (LocalFromWide(L"AI\u5bf9\u8bdd\u5931\u8d25: ") + result->chatResult.error);
		g_session.messages.push_back(SessionMessage{ SessionRole::System, err, false });
		OutputStringToELog("[" + LocalFromWide(L"AI\u5bf9\u8bdd") + "]" + err);
	}

	CompactHistoryLocked(g_session);
	PostRefreshDialog();
}

void RefreshChatDialog(HWND hWnd)
{
	auto* ctx = reinterpret_cast<ChatDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
	if (ctx == nullptr) {
		return;
	}

	std::string history;
	std::string historyHtml;
	bool inFlight = false;
	{
		std::lock_guard<std::mutex> guard(g_session.mutex);
		if (g_session.requestInFlight && g_session.activeRequestId == 0) {
			g_session.requestInFlight = false;
			g_session.streamingAssistantPreview.clear();
		}
		history = BuildHistoryTextLocked(g_session);
		if (ctx->webViewDesired) {
			historyHtml = BuildHistoryHtmlLocked(g_session);
		}
		inFlight = g_session.requestInFlight;
	}

	SetWindowTextA(ctx->hHistory, history.c_str());
	ScrollEditToBottom(ctx->hHistory);
	if (ctx->webViewDesired) {
		UpdateHistoryWebViewHtml(ctx, historyHtml);
	}
	EnableWindow(ctx->hSend, inFlight ? FALSE : TRUE);
	UpdateWebViewComposerState(ctx, inFlight);
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
		ctx->hClearHistory = CreateWindowW(L"STATIC", L"\u6e05\u7a7a\u5386\u53f2\u4f1a\u8bdd",
			WS_CHILD | WS_VISIBLE | SS_NOTIFY | SS_CENTER | SS_CENTERIMAGE,
			14, 442, 106, 26, hWnd, reinterpret_cast<HMENU>(IDC_AI_CHAT_CLEAR_HISTORY), nullptr, nullptr);
		ctx->hInput = CreateWindowExA(WS_EX_CLIENTEDGE, "EDIT", "",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_MULTILINE | ES_AUTOVSCROLL | WS_VSCROLL,
			14, 476, 652, 32, hWnd, reinterpret_cast<HMENU>(IDC_AI_CHAT_INPUT), nullptr, nullptr);
		ctx->hSend = CreateWindowW(L"STATIC", L"\u53d1\u9001",
			WS_CHILD | WS_VISIBLE | SS_NOTIFY | SS_CENTER | SS_CENTERIMAGE,
			674, 476, 92, 32, hWnd, reinterpret_cast<HMENU>(IDC_AI_CHAT_SEND), nullptr, nullptr);

		SetDefaultFont(ctx->hHistory);
		SetDefaultFont(ctx->hClearHistory);
		SetDefaultFont(ctx->hInput);
		SetDefaultFont(ctx->hSend);
		InstallEditHotkeys(ctx->hHistory, kEditFlagNone);
		InstallEditHotkeys(ctx->hInput, kEditFlagSubmitOnEnter);
		InstallChatActionControl(ctx->hSend, kActionSubmit);
		InstallChatActionControl(ctx->hClearHistory, kActionClear);
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

	case WM_CTLCOLORSTATIC:
		if (ctx != nullptr) {
			HWND hStatic = reinterpret_cast<HWND>(lParam);
			if (hStatic != ctx->hClearHistory && hStatic != ctx->hSend) {
				break;
			}
			HDC hdc = reinterpret_cast<HDC>(wParam);
			SetTextColor(hdc, RGB(0, 102, 204));
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

	case WM_AUTOLINKER_AI_CHAT_REFRESH:
		RefreshChatDialog(hWnd);
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

LRESULT CALLBACK CodeEditDialogProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	auto* ctx = reinterpret_cast<CodeEditDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
	switch (uMsg)
	{
	case WM_NCCREATE: {
		const auto* create = reinterpret_cast<CREATESTRUCTA*>(lParam);
		SetWindowLongPtrA(hWnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create->lpCreateParams));
		return TRUE;
	}

	case WM_CREATE: {
		if (ctx == nullptr) {
			return -1;
		}

		ctx->hEdit = CreateWindowExA(WS_EX_CLIENTEDGE, "EDIT", "",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_MULTILINE | ES_AUTOVSCROLL | WS_VSCROLL,
			14, 14, 752, 450, hWnd, reinterpret_cast<HMENU>(IDC_CODE_EDIT), nullptr, nullptr);
		const std::string normalized = NormalizeCodeForEIDE(ctx->text);
		SetWindowTextA(ctx->hEdit, normalized.c_str());

		ctx->hCopy = CreateWindowW(L"BUTTON", L"\u590D\u5236\u5230\u526A\u8D34\u677F",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP, 494, 474, 116, 30, hWnd,
			reinterpret_cast<HMENU>(IDC_CODE_COPY), nullptr, nullptr);
		ctx->hOk = CreateWindowW(L"BUTTON", L"\u63D0\u4F9B\u7ED9AI",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON, 618, 474, 84, 30, hWnd,
			reinterpret_cast<HMENU>(IDC_CODE_OK), nullptr, nullptr);
		ctx->hCancel = CreateWindowW(L"BUTTON", L"\u53D6\u6D88",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP, 710, 474, 56, 30, hWnd,
			reinterpret_cast<HMENU>(IDC_CODE_CANCEL), nullptr, nullptr);

		SetDefaultFont(ctx->hEdit);
		SetDefaultFont(ctx->hCopy);
		SetDefaultFont(ctx->hOk);
		SetDefaultFont(ctx->hCancel);
		InstallEditHotkeys(ctx->hEdit, kEditFlagNone);
		LayoutCodeEditDialog(hWnd, ctx);
		SetFocus(ctx->hEdit);
		return 0;
	}

	case WM_GETMINMAXINFO: {
		auto* mmi = reinterpret_cast<MINMAXINFO*>(lParam);
		ApplyMinTrackSize(hWnd, mmi, 560, 360);
		return 0;
	}

	case WM_SIZE:
		if (ctx != nullptr) {
			LayoutCodeEditDialog(hWnd, ctx);
		}
		return 0;

	case WM_COMMAND: {
		if (ctx == nullptr) {
			return 0;
		}
		const int id = LOWORD(wParam);
		if (id == IDC_CODE_COPY) {
			const std::string text = GetEditTextA(ctx->hEdit);
			if (!SetClipboardTextSimple(text)) {
                MessageBoxA(hWnd, "Copy failed.", "AI Code Edit", MB_ICONWARNING | MB_OK);
			}
			return 0;
		}
		if (id == IDC_CODE_OK) {
			ctx->text = GetEditTextA(ctx->hEdit);
			ctx->accepted = true;
			DestroyWindow(hWnd);
			return 0;
		}
		if (id == IDC_CODE_CANCEL) {
			ctx->accepted = false;
			DestroyWindow(hWnd);
			return 0;
		}
		return 0;
	}

	case WM_CLOSE:
		DestroyWindow(hWnd);
		return 0;

	default:
		break;
	}
	return DefWindowProcA(hWnd, uMsg, wParam, lParam);
}

bool HandleToolDialogRequest(LPARAM lParam)
{
	auto* request = reinterpret_cast<CodeEditToolRequest*>(lParam);
	if (request == nullptr) {
		return true;
	}

    std::string title = request->title.empty() ? "AI Tool: Code Edit" : request->title;
	if (!TrimAsciiCopy(request->hint).empty()) {
		title += " - " + request->hint;
	}

	std::string edited = request->initialCode;
	const bool accepted = ShowAICodeEditDialog(g_mainWindow, title, request->initialCode, edited);

	{
		std::lock_guard<std::mutex> guard(request->mutex);
		request->accepted = accepted;
		request->resultCode = edited;
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

UINT GetAIChatToolExecMessageForTooling()
{
	return g_msgAIChatToolExec;
}

bool RequestCodeEditForTooling(
	const std::string& title,
	const std::string& hint,
	const std::string& initialCode,
	std::string& outCode)
{
	return RequestCodeEditFromMainThread(title, hint, initialCode, outCode);
}

bool EnsureChatHostWindowCreated();
void FocusChatInputControl();

bool ShowAICodeEditDialog(HWND owner, const std::string& title, const std::string& initialCode, std::string& ioCode)
{
	ComCtl6ActivationScope themeScope;
	INITCOMMONCONTROLSEX icex = {};
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_STANDARD_CLASSES | ICC_WIN95_CLASSES | ICC_LINK_CLASS;
	InitCommonControlsEx(&icex);

	WNDCLASSEXA wc = {};
	wc.cbSize = sizeof(wc);
	wc.lpfnWndProc = CodeEditDialogProc;
	wc.hInstance = GetModuleHandleA(nullptr);
	wc.lpszClassName = "AutoLinkerAICodeEditDialogWindow";
	wc.hIcon = GetAppIconLarge();
	wc.hIconSm = GetAppIconSmall();
	wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_BTNFACE + 1);
	RegisterClassExA(&wc);

	CodeEditDialogContext ctx = {};
	ctx.title = title.empty() ? "AutoLinker AI Code Edit" : title;
	ctx.text = initialCode;

	HWND hDialog = CreateWindowExA(
		WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
		wc.lpszClassName,
		ctx.title.c_str(),
		WS_CAPTION | WS_SYSMENU | WS_VISIBLE | WS_SIZEBOX,
		CW_USEDEFAULT, CW_USEDEFAULT, 800, 560,
		owner,
		nullptr,
		wc.hInstance,
		&ctx);
	if (hDialog == nullptr) {
		return false;
	}

	ApplyWindowIcon(hDialog);
	EnsureWindowTitle(hDialog, ctx.title);
	RunModalWindow(owner, hDialog);
	if (ctx.accepted) {
		ioCode = ctx.text;
	}
	return ctx.accepted;
}

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
	LogChatTab("add left work area tab icon ok");
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

	return ReadTabItemTextA(tabHwnd, 0) == LocalFromWide(L"支持库") &&
		ReadTabItemTextA(tabHwnd, 1) == LocalFromWide(L"程序") &&
		ReadTabItemTextA(tabHwnd, 2) == LocalFromWide(L"属性");
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

	GetClientRect(g_leftWorkAreaHost.tabHwnd, &rc);
	TabCtrl_AdjustRect(g_leftWorkAreaHost.tabHwnd, FALSE, &rc);
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
	const int height = (std::max)(120, static_cast<int>(rc.bottom - rc.top));
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
	case WM_WINDOWPOSCHANGED:
		if (g_leftWorkAreaHost.pageVisible) {
			LayoutLeftWorkAreaChatPage();
		}
		break;

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
		if (ReadTabItemTextA(tabHwnd, index) == caption) {
			tabIndex = index;
			break;
		}
	}

	if (tabIndex < 0) {
		TCITEMA item = {};
		item.mask = TCIF_TEXT;
		item.pszText = const_cast<LPSTR>(caption.c_str());
		const int imageIndex = EnsureLeftWorkAreaTabImageIndex(tabHwnd);
		if (imageIndex >= 0) {
			item.mask |= TCIF_IMAGE;
			item.iImage = imageIndex;
		}
		const LRESULT insertIndex = SendMessageA(tabHwnd, TCM_INSERTITEMA, static_cast<WPARAM>(itemCount), reinterpret_cast<LPARAM>(&item));
		if (insertIndex < 0) {
			LogChatTab("insert left work area tab failed");
			return false;
		}
		tabIndex = static_cast<int>(insertIndex);
		LogChatTab("insert left work area tab ok");
	}

	g_leftWorkAreaHost.tabHwnd = tabHwnd;
	g_leftWorkAreaHost.hostHwnd = hostHwnd;
	g_leftWorkAreaHost.tabIndex = tabIndex;
	if (g_leftWorkAreaHost.imageIndex < 0) {
		g_leftWorkAreaHost.imageIndex = EnsureLeftWorkAreaTabImageIndex(tabHwnd);
	}
	if (g_leftWorkAreaHost.imageIndex >= 0) {
		TCITEMA item = {};
		item.mask = TCIF_IMAGE;
		item.iImage = g_leftWorkAreaHost.imageIndex;
		SendMessageA(tabHwnd, TCM_SETITEMA, static_cast<WPARAM>(tabIndex), reinterpret_cast<LPARAM>(&item));
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
	LogChatTab("left work area host ok");
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

	LogChatTab("create host ok");
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
void Initialize(HWND mainWindow, ConfigManager* configManager)
{
	g_mainWindow = mainWindow;
	g_configManager = configManager;
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
		ReadTabItemTextA(g_leftWorkAreaHost.tabHwnd, g_leftWorkAreaHost.tabIndex) == GetLeftWorkAreaTabCaption()) {
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

bool ExecutePublicTool(const std::string& toolName, const std::string& argumentsJson, std::string& outResultJsonUtf8, bool& outOk)
{
	outResultJsonUtf8.clear();
	outOk = false;
	if (TrimAsciiCopy(toolName).empty()) {
		return false;
	}

	bool toolOk = false;
	const std::string resultJsonLocal = ExecuteToolCall(toolName, argumentsJson, toolOk);
	outResultJsonUtf8 = LocalToUtf8Text(resultJsonLocal);
	outOk = toolOk;
	return true;
}
} // namespace AIChatFeature
