#include "AIChatFeature.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <cctype>
#include <chrono>
#include <CommCtrl.h>
#include <condition_variable>
#include <cstring>
#include <memory>
#include <mutex>
#include <new>
#include <process.h>
#include <string>
#include <vector>

#include "..\\thirdparty\\json.hpp"

#include "AIConfigDialog.h"
#include "AIService.h"
#include "ConfigManager.h"
#include "Global.h"
#include "IDEFacade.h"
#include "resource.h"
#if defined(_M_IX86)
#include "direct_global_search_debug.hpp"
#endif

#pragma comment(lib, "comctl32.lib")
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
	HWND hInput = nullptr;
	HWND hSend = nullptr;
	HWND hClearHistory = nullptr;
	int inputRowsVisible = 1;
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

struct ToolExecutionRequest {
	std::string toolName;
	std::string argumentsJson;
	std::string resultJson;
	bool ok = false;
	bool done = false;
	std::mutex mutex;
	std::condition_variable cv;
};

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
	if (ctx->hClearHistory != nullptr) {
		MoveWindow(ctx->hClearHistory, margin, actionRowY, clearHistoryWidth, actionRowHeight, TRUE);
	}
	if (ctx->hInput != nullptr) {
		MoveWindow(ctx->hInput, margin, inputY, inputWidth, inputHeight, TRUE);
	}
	if (ctx->hSend != nullptr) {
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

std::uintptr_t GetCurrentProcessImageBaseForAI()
{
	HMODULE module = GetModuleHandleW(nullptr);
	return reinterpret_cast<std::uintptr_t>(module);
}

#if defined(_M_IX86)
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
	for (HWND treeHwnd : CollectTreeViewWindowsForAI(g_mainWindow)) {
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
	if (g_mainWindow == nullptr || !IsWindow(g_mainWindow)) {
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

bool TryGetProgramItemCodeByNameForAI(
	const std::string& name,
	const std::string& kindFilter,
	ProgramTreeItemInfo& outItem,
	std::string& outCode,
	std::string& outTrace,
	std::string& outError)
{
	outCode.clear();
	outTrace.clear();
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
	e571::RawSearchContextPageDumpDebugResult dumpResult;
	if (!e571::DebugDumpCodePageByProgramTreeItemData(
			outItem.itemData,
			GetCurrentProcessImageBaseForAI(),
			&outCode,
			&dumpResult)) {
		outTrace = dumpResult.trace;
		outError = "get page code failed";
		return false;
	}
	outTrace = dumpResult.trace;
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
		std::string pageCode;
		if (!IDEFacade::Instance().GetCurrentPageCode(pageCode)) {
			return R"({"ok":false,"error":"GetCurrentPageCode failed"})";
		}

		std::string pageName;
		std::string pageType;
		std::string pageNameTrace;
		const bool nameOk = IDEFacade::Instance().GetCurrentPageName(pageName, &pageType, &pageNameTrace);
		nlohmann::json r;
		r["ok"] = true;
		r["code"] = LocalToUtf8Text(pageCode);
		r["page_name_ok"] = nameOk;
		r["page_name"] = LocalToUtf8Text(pageName);
		r["page_type"] = LocalToUtf8Text(pageType);
		r["page_name_trace"] = pageNameTrace;
		outOk = true;
		return Utf8ToLocalText(r.dump());
	}

	if (toolName == "get_current_page_info") {
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
		r["page_name"] = LocalToUtf8Text(pageName);
		r["page_type"] = LocalToUtf8Text(pageType);
		r["page_name_trace"] = pageNameTrace;
		outOk = true;
		return Utf8ToLocalText(r.dump());
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

	if (toolName == "get_program_item_code") {
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

		ProgramTreeItemInfo item;
		std::string code;
		std::string trace;
		std::string error;
		if (!TryGetProgramItemCodeByNameForAI(name, kind, item, code, trace, error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "get program item code failed" : error;
			if (!trace.empty()) {
				r["trace"] = trace;
			}
			return Utf8ToLocalText(r.dump());
		}

		nlohmann::json r;
		r["ok"] = true;
		r["name"] = LocalToUtf8Text(item.name);
		r["type_key"] = item.typeKey;
		r["type_name"] = LocalToUtf8Text(item.typeName);
		r["item_data"] = item.itemData;
		r["trace"] = trace;
		r["code_kind"] = "pseudo_reference";
		r["warning"] = LocalToUtf8Text("该代码来自程序树按名称抓取，与IDE正常编辑页结构可能不同，仅可作为伪代码参考。");
		r["code"] = LocalToUtf8Text(code);
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

		std::string openTrace;
		if (!e571::DebugOpenProgramTreeItemByData(
				matched.front().itemData,
				GetCurrentProcessImageBaseForAI(),
				&openTrace)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = "open program item page failed";
			r["trace"] = openTrace;
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
		r["requested_name"] = LocalToUtf8Text(name);
		r["type_key"] = matched.front().typeKey;
		r["type_name"] = LocalToUtf8Text(matched.front().typeName);
		r["item_data"] = matched.front().itemData;
		r["trace"] = openTrace;
		r["current_page_ok"] = currentPageOk;
		r["current_page_name"] = LocalToUtf8Text(currentPageName);
		r["current_page_type"] = LocalToUtf8Text(currentPageType);
		r["current_page_trace"] = currentPageTrace;
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

	nlohmann::json r;
	r["ok"] = false;
	r["error"] = "unknown tool: " + toolName;
	return Utf8ToLocalText(r.dump());
}
#else
std::string ExecuteToolCallOnMainThread(const std::string& toolName, const std::string& argumentsJson, bool& outOk)
{
	(void)toolName;
	(void)argumentsJson;
	outOk = false;
	return R"({"ok":false,"error":"x86 only"})";
}
#endif

bool RequestToolExecutionFromMainThread(
	const std::string& toolName,
	const std::string& argumentsJson,
	std::string& outResultJson,
	bool& outOk)
{
	outResultJson.clear();
	outOk = false;
	if (g_mainWindow == nullptr || !IsWindow(g_mainWindow) || g_msgAIChatToolExec == 0) {
		return false;
	}

	ToolExecutionRequest request;
	request.toolName = toolName;
	request.argumentsJson = argumentsJson;
	if (PostMessage(g_mainWindow, g_msgAIChatToolExec, 0, reinterpret_cast<LPARAM>(&request)) == FALSE) {
		return false;
	}

	std::unique_lock<std::mutex> lock(request.mutex);
	if (!request.cv.wait_for(lock, std::chrono::minutes(20), [&request]() { return request.done; })) {
		return false;
	}

	outResultJson = request.resultJson;
	outOk = request.ok;
	return true;
}

std::string ExecuteToolCall(const std::string& toolName, const std::string& argumentsJson, bool& outOk)
{
	outOk = false;

	if (toolName == "request_code_edit") {
		std::string title = LocalFromWide(L"AI\u4ee3\u7801\u7f16\u8f91");
		std::string hint;
		std::string initialCode;
		try {
			const nlohmann::json args = nlohmann::json::parse(argumentsJson);
			if (args.contains("title") && args["title"].is_string()) {
				title = Utf8ToLocalText(args["title"].get<std::string>());
			}
			if (args.contains("hint") && args["hint"].is_string()) {
				hint = Utf8ToLocalText(args["hint"].get<std::string>());
			}
			if (args.contains("initial_code") && args["initial_code"].is_string()) {
				initialCode = Utf8ToLocalText(args["initial_code"].get<std::string>());
			}
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return Utf8ToLocalText(r.dump());
		}

		std::string editedCode;
		if (!RequestCodeEditFromMainThread(title, hint, initialCode, editedCode)) {
			return R"({"ok":false,"cancelled":true})";
		}

		nlohmann::json r;
		r["ok"] = true;
		r["code"] = LocalToUtf8Text(editedCode);
		outOk = true;
		return Utf8ToLocalText(r.dump());
	}

	if (toolName == "get_current_page_code" ||
		toolName == "get_current_page_info" ||
		toolName == "list_program_items" ||
		toolName == "get_program_item_code" ||
		toolName == "switch_to_program_item_page" ||
		toolName == "search_project_keyword" ||
		toolName == "jump_to_search_result") {
		std::string resultJson;
		if (!RequestToolExecutionFromMainThread(toolName, argumentsJson, resultJson, outOk)) {
			return R"({"ok":false,"error":"main thread tool execution failed"})";
		}
		return resultJson;
	}

	nlohmann::json r;
	r["ok"] = false;
	r["error"] = "unknown tool: " + toolName;
	return Utf8ToLocalText(r.dump());
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
	bool inFlight = false;
	{
		std::lock_guard<std::mutex> guard(g_session.mutex);
		if (g_session.requestInFlight && g_session.activeRequestId == 0) {
			g_session.requestInFlight = false;
			g_session.streamingAssistantPreview.clear();
		}
		history = BuildHistoryTextLocked(g_session);
		inFlight = g_session.requestInFlight;
	}

	SetWindowTextA(ctx->hHistory, history.c_str());
	ScrollEditToBottom(ctx->hHistory);
	EnableWindow(ctx->hSend, inFlight ? FALSE : TRUE);
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
		PostMessageA(hWnd, WM_AUTOLINKER_AI_CHAT_DEFER_LAYOUT, 0, 0);

		RefreshChatDialog(hWnd);
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

	case WM_AUTOLINKER_AI_CHAT_SUBMIT:
		if (ctx != nullptr) {
			OutputStringToELog("[AI Chat][UI] submit begin");
			const std::string text = GetEditTextA(ctx->hInput);
			if (!TrimAsciiCopy(text).empty()) {
				SetWindowTextA(ctx->hInput, "");
				ctx->inputRowsVisible = 1;
				LayoutAIChatDialog(hWnd, ctx);
				EnableWindow(ctx->hSend, FALSE);
				if (!StartChatRequest(text)) {
					RefreshChatDialog(hWnd);
				}
			}
			OutputStringToELog("[AI Chat][UI] submit end");
		}
		return 0;

	case WM_AUTOLINKER_AI_CHAT_CLEAR:
		if (ctx != nullptr && ctx->hHistory != nullptr) {
			SetWindowTextA(ctx->hHistory, "");
		}
		RequestClearChatHistoryAsync();
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

std::string GetLeftWorkAreaTabCaption()
{
	return "AI";
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
