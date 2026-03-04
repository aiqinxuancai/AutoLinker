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

constexpr int IDC_AI_CHAT_HISTORY = 2101;
constexpr int IDC_AI_CHAT_INPUT = 2102;
constexpr int IDC_AI_CHAT_SEND = 2103;
constexpr int IDC_AI_CHAT_CLEAR_HISTORY = 2104;

constexpr int IDC_CODE_EDIT = 2201;
constexpr int IDC_CODE_OK = 1;
constexpr int IDC_CODE_CANCEL = 2;
constexpr int IDC_CODE_COPY = 2204;
constexpr UINT_PTR kEditSubclassId = 1;
constexpr DWORD_PTR kEditFlagNone = 0;
constexpr DWORD_PTR kEditFlagSubmitOnEnter = 1;

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
	bool requestInFlight = false;
	unsigned long long activeRequestId = 0;
	unsigned long long nextRequestId = 1;
};

struct ChatDialogContext {
	HWND hHistory = nullptr;
	HWND hInput = nullptr;
	HWND hSend = nullptr;
	HWND hClearHistory = nullptr;
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

HWND g_mainWindow = nullptr;
ConfigManager* g_configManager = nullptr;
HWND g_chatDialog = nullptr;
AIChatSessionState g_session;
UINT g_msgAIChatDone = 0;
UINT g_msgAIChatToolDialog = 0;

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
			}
			else {
				HWND hParent = GetParent(hWnd);
				if (hParent != nullptr) {
					HWND hSend = GetDlgItem(hParent, IDC_AI_CHAT_SEND);
					if (hSend != nullptr && IsWindowEnabled(hSend)) {
						PostMessageA(
							hParent,
							WM_COMMAND,
							MAKEWPARAM(IDC_AI_CHAT_SEND, BN_CLICKED),
							reinterpret_cast<LPARAM>(hSend));
					}
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

	const int margin = 14;
	const int gap = 8;
	const int actionRowHeight = 26;
	const int inputHeight = 84;
	const int sendWidth = 92;
	const int clearHistoryWidth = 116;

	const int contentWidth = (std::max)(120, clientWidth - margin * 2);
	const int inputWidth = (std::max)(80, contentWidth - sendWidth - gap);
	const int sendX = margin + inputWidth + gap;
	const int inputY = clientHeight - margin - inputHeight;
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
		text += "[" + LocalFromWide(L"\u7cfb\u7edf") + "]\r\n"
			+ LocalFromWide(L"\u7b49\u5f85 AI \u8fd4\u56de...") + "\r\n";
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
	g_session.messages.push_back(SessionMessage{
		SessionRole::System,
		"Chat request auto-recovered: " + reason,
		false
	});
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

std::string ExecuteToolCall(const std::string& toolName, const std::string& argumentsJson, bool& outOk)
{
	outOk = false;
	if (toolName == "get_current_page_code") {
		std::string pageCode;
		if (!IDEFacade::Instance().GetCurrentPageCode(pageCode)) {
			return R"({"ok":false,"error":"GetCurrentPageCode failed"})";
		}
		nlohmann::json r;
		r["ok"] = true;
		r["code"] = LocalToUtf8Text(pageCode);
		outOk = true;
		return Utf8ToLocalText(r.dump());
	}

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
		g_session.requestInFlight = true;
		g_session.activeRequestId = request->requestId;
	}

	const uintptr_t threadId = _beginthread(RunAIChatWorker, 0, request.get());
	if (threadId == static_cast<uintptr_t>(-1L)) {
		std::lock_guard<std::mutex> guard(g_session.mutex);
		g_session.requestInFlight = false;
		g_session.activeRequestId = 0;
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
	std::lock_guard<std::mutex> guard(g_session.mutex);
	g_session.messages.clear();
	g_session.rollingSummary.clear();
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
		ctx->hClearHistory = CreateWindowW(L"BUTTON", L"\u6e05\u7a7a\u5386\u53f2\u5bf9\u8bdd",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP,
			14, 442, 106, 26, hWnd, reinterpret_cast<HMENU>(IDC_AI_CHAT_CLEAR_HISTORY), nullptr, nullptr);
		ctx->hInput = CreateWindowExA(WS_EX_CLIENTEDGE, "EDIT", "",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_MULTILINE | ES_AUTOVSCROLL | WS_VSCROLL,
			14, 476, 652, 72, hWnd, reinterpret_cast<HMENU>(IDC_AI_CHAT_INPUT), nullptr, nullptr);
		ctx->hSend = CreateWindowW(L"BUTTON", L"\u53D1\u9001",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
			674, 476, 92, 72, hWnd, reinterpret_cast<HMENU>(IDC_AI_CHAT_SEND), nullptr, nullptr);

		SetDefaultFont(ctx->hHistory);
		SetDefaultFont(ctx->hClearHistory);
		SetDefaultFont(ctx->hInput);
		SetDefaultFont(ctx->hSend);
		InstallEditHotkeys(ctx->hHistory, kEditFlagNone);
		InstallEditHotkeys(ctx->hInput, kEditFlagSubmitOnEnter);
		LayoutAIChatDialog(hWnd, ctx);

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
			LayoutAIChatDialog(hWnd, ctx);
			ScrollEditToBottom(ctx->hHistory);
		}
		return 0;

	case WM_COMMAND: {
		if (ctx == nullptr) {
			return 0;
		}
		const int id = LOWORD(wParam);
		if (id == IDC_AI_CHAT_SEND) {
			const std::string text = GetEditTextA(ctx->hInput);
			if (TrimAsciiCopy(text).empty()) {
				return 0;
			}

			SetWindowTextA(ctx->hInput, "");
			EnableWindow(ctx->hSend, FALSE);
			if (!StartChatRequest(text)) {
				RefreshChatDialog(hWnd);
			}
			return 0;
		}
		if (id == IDC_AI_CHAT_CLEAR_HISTORY) {
			const int answer = MessageBoxW(
				hWnd,
				L"\u786e\u5b9a\u6e05\u7a7a\u6240\u6709 AI \u5386\u53f2\u5bf9\u8bdd\u5417\uff1f",
				L"AutoLinker AI \u5bf9\u8bdd",
				MB_ICONQUESTION | MB_YESNO | MB_DEFBUTTON2);
			if (answer == IDYES) {
				ClearChatHistory();
				RefreshChatDialog(hWnd);
			}
			return 0;
		}
		return 0;
	}

	case WM_AUTOLINKER_AI_CHAT_REFRESH:
		RefreshChatDialog(hWnd);
		return 0;

	case WM_CLOSE:
		DestroyWindow(hWnd);
		return 0;

	case WM_DESTROY:
		if (g_chatDialog == hWnd) {
			g_chatDialog = nullptr;
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
} // namespace

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

namespace AIChatFeature {
void Initialize(HWND mainWindow, ConfigManager* configManager)
{
	g_mainWindow = mainWindow;
	g_configManager = configManager;
	g_msgAIChatDone = RegisterWindowMessageA("AutoLinker.AIChat.Done");
	g_msgAIChatToolDialog = RegisterWindowMessageA("AutoLinker.AIChat.ToolDialog");
}

void Shutdown()
{
	if (g_chatDialog != nullptr && IsWindow(g_chatDialog)) {
		DestroyWindow(g_chatDialog);
		g_chatDialog = nullptr;
	}
}

void OpenDialog()
{
	if (g_mainWindow == nullptr || !IsWindow(g_mainWindow) || g_configManager == nullptr) {
		OutputStringToELog("[AI Chat] Not initialized yet, please retry.");
		return;
	}

	if (g_chatDialog != nullptr && IsWindow(g_chatDialog)) {
		ShowWindow(g_chatDialog, SW_SHOW);
		SetForegroundWindow(g_chatDialog);
		return;
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
		LocalFromWide(L"AutoLinker AI \u5bf9\u8bdd").c_str(),
		WS_CAPTION | WS_SYSMENU | WS_SIZEBOX | WS_MINIMIZEBOX | WS_VISIBLE,
		CW_USEDEFAULT, CW_USEDEFAULT, 860, 680,
		g_mainWindow,
		nullptr,
		wc.hInstance,
		nullptr);
	ApplyWindowIcon(g_chatDialog);
	EnsureWindowTitle(g_chatDialog, LocalFromWide(L"AutoLinker AI \u5bf9\u8bdd"));
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
	return false;
}
} // namespace AIChatFeature
