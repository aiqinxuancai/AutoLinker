#include "AIConfigDialog.h"
#include "resource.h"

#include <algorithm>
#include <array>
#include <CommCtrl.h>
#include <format>
#include <new>
#include <Shellapi.h>
#include <wrl.h>

#include "..\\thirdparty\\json.hpp"
#include "..\\thirdparty\\WebView2.h"

#include "AIService.h"
#include "Global.h"
#include "ResourceTextLoader.h"

#pragma comment(lib, "comctl32.lib")
#if defined _M_IX86
#pragma comment(linker, "\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_X64
#pragma comment(linker, "\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
#pragma comment(linker, "\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif

namespace {
constexpr int IDC_CFG_PROTOCOL = 1000;
constexpr int IDC_CFG_BASE_URL = 1001;
constexpr int IDC_CFG_API_KEY = 1002;
constexpr int IDC_CFG_MODEL = 1003;
constexpr int IDC_CFG_EXTRA_PROMPT = 1004;
constexpr int IDC_CFG_GET_KEY_LINK = 1005;
constexpr int IDC_CFG_FILL_RIGHT_CODES = 1006; // 保留以备兼容，实际已被 IDC_CFG_PLATFORM_PRESET 取代
constexpr int IDC_CFG_TAVILY_API_KEY = 1007;
constexpr int IDC_CFG_PLATFORM_PRESET = 1008;
constexpr int IDC_CFG_TEST_CONNECTION = 1009;
constexpr int IDC_CFG_SAVE = 1;
constexpr int IDC_CFG_CANCEL = 2;

constexpr int IDC_PREVIEW_EDIT = 1101;
constexpr int IDC_PREVIEW_OK = 1;
constexpr int IDC_PREVIEW_SECONDARY = 1102;
constexpr int IDC_PREVIEW_CANCEL = 2;

constexpr int IDC_INPUT_EDIT = 1201;
constexpr int IDC_INPUT_OK = 1;
constexpr int IDC_INPUT_CANCEL = 2;
constexpr UINT_PTR kAIConfigWebViewInitTimerId = 0xAC01;
constexpr UINT kAIConfigWebViewInitTimeoutMs = 12000;
constexpr UINT WM_AUTOLINKER_AI_CONFIG_TEST_DONE = WM_APP + 301;
constexpr UINT_PTR kAIPreviewWebViewInitTimerId = 0xAC02;
constexpr UINT kAIPreviewWebViewInitTimeoutMs = 12000;

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

std::string GetEditTextA(HWND hEdit)
{
	if (hEdit == nullptr) {
		return std::string();
	}

	int len = GetWindowTextLengthA(hEdit);
	if (len <= 0) {
		return std::string();
	}

	std::string text(static_cast<size_t>(len + 1), '\0');
	GetWindowTextA(hEdit, &text[0], len + 1);
	text.resize(static_cast<size_t>(len));
	return text;
}

std::string NormalizeMultilineForEdit(const std::string& text)
{
	if (text.empty()) {
		return std::string();
	}

	std::string expanded = text;
	const bool hasRealLineBreak = expanded.find('\n') != std::string::npos || expanded.find('\r') != std::string::npos;
	if (!hasRealLineBreak) {
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

void SetDefaultFont(HWND hWnd)
{
	if (hWnd == nullptr) {
		return;
	}
	SendMessageA(hWnd, WM_SETFONT, reinterpret_cast<WPARAM>(GetDialogFont()), TRUE);
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

std::wstring Utf8ToWide(const std::string& text)
{
	if (text.empty()) {
		return std::wstring();
	}

	const int wideLen = MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0);
	if (wideLen <= 0) {
		return std::wstring();
	}

	std::wstring wide(static_cast<size_t>(wideLen), L'\0');
	if (MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		wide.data(),
		wideLen) <= 0) {
		return std::wstring();
	}
	return wide;
}

std::wstring WideFromLocal(const std::string& text)
{
	if (text.empty()) {
		return std::wstring();
	}

	const int wideLen = MultiByteToWideChar(
		CP_ACP,
		0,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0);
	if (wideLen <= 0) {
		return std::wstring();
	}

	std::wstring wide(static_cast<size_t>(wideLen), L'\0');
	if (MultiByteToWideChar(
		CP_ACP,
		0,
		text.data(),
		static_cast<int>(text.size()),
		wide.data(),
		wideLen) <= 0) {
		return std::wstring();
	}
	return wide;
}

std::string WideToUtf8(const std::wstring& text)
{
	if (text.empty()) {
		return std::string();
	}

	const int utf8Len = WideCharToMultiByte(
		CP_UTF8,
		0,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0,
		nullptr,
		nullptr);
	if (utf8Len <= 0) {
		return std::string();
	}

	std::string utf8(static_cast<size_t>(utf8Len), '\0');
	if (WideCharToMultiByte(
		CP_UTF8,
		0,
		text.data(),
		static_cast<int>(text.size()),
		utf8.data(),
		utf8Len,
		nullptr,
		nullptr) <= 0) {
		return std::string();
	}
	return utf8;
}

std::string JsonStringLiteral(const std::string& utf8Text)
{
	return nlohmann::json(utf8Text).dump();
}

bool IsWebView2RuntimeAvailableWithTag(const char* logTag)
{
	LPWSTR version = nullptr;
	const HRESULT hr = GetAvailableCoreWebView2BrowserVersionString(nullptr, &version);
	const bool available = SUCCEEDED(hr);
	OutputStringToELog(std::format("[{}][WebView2] runtime available={}", logTag == nullptr ? "AI Config" : logTag, available ? 1 : 0));
	if (version != nullptr) {
		CoTaskMemFree(version);
	}
	return available;
}

bool IsWebView2RuntimeAvailable()
{
	return IsWebView2RuntimeAvailableWithTag("AI Config");
}

struct AIConfigDialogRunResult {
	bool accepted = false;
	bool fallbackRequested = false;
};

struct AIPreviewWebViewRunResult {
	AIPreviewAction action = AIPreviewAction::Cancel;
	bool fallbackRequested = false;
};

struct AIConfigDialogContext {
	AISettings* settings = nullptr;
	bool accepted = false;
	bool useNativeLink = false;
	bool testInFlight = false;
	HWND hProtocol = nullptr;
	HWND hBaseUrl = nullptr;
	HWND hApiKey = nullptr;
	HWND hModel = nullptr;
	HWND hTavilyApiKey = nullptr;
	HWND hExtraPrompt = nullptr;
	HWND hGetKeyLink = nullptr;
	HWND hPlatformCombo = nullptr;
	HWND hTestConnection = nullptr;
};

struct AIConfigWebViewDialogContext {
	AISettings* settings = nullptr;
	bool accepted = false;
	bool fallbackRequested = false;
	bool webViewReady = false;
	bool testInFlight = false;
	HWND hHost = nullptr;
	HWND hLoading = nullptr;
	Microsoft::WRL::ComPtr<ICoreWebView2Environment> webViewEnvironment;
	Microsoft::WRL::ComPtr<ICoreWebView2Controller> webViewController;
	Microsoft::WRL::ComPtr<ICoreWebView2> webView;
};

struct AIConfigConnectionTestRequest {
	HWND dialogHwnd = nullptr;
	AISettings settings;
	bool forWebView = false;
};

struct AIConfigConnectionTestResult {
	bool forWebView = false;
	AIResult result;
};

struct AIPreviewWebViewDialogContext {
	std::string title;
	std::string content;
	std::string primaryText;
	std::string secondaryText;
	AIPreviewAction action = AIPreviewAction::Cancel;
	bool fallbackRequested = false;
	bool webViewReady = false;
	HWND hHost = nullptr;
	HWND hLoading = nullptr;
	Microsoft::WRL::ComPtr<ICoreWebView2Environment> webViewEnvironment;
	Microsoft::WRL::ComPtr<ICoreWebView2Controller> webViewController;
	Microsoft::WRL::ComPtr<ICoreWebView2> webView;
};

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

void PopulateProtocolCombo(HWND hCombo, AIProtocolType selected)
{
	if (hCombo == nullptr) {
		return;
	}

	SendMessageA(hCombo, CB_RESETCONTENT, 0, 0);
	const int idxOpenAI = static_cast<int>(SendMessageA(hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>("OpenAI Chat")));
	const int idxOpenAIResponses = static_cast<int>(SendMessageA(hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>("OpenAI Responses")));
	const int idxGemini = static_cast<int>(SendMessageA(hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>("Gemini")));
	const int idxClaude = static_cast<int>(SendMessageA(hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>("Claude")));
	SendMessageA(hCombo, CB_SETITEMDATA, idxOpenAI, static_cast<LPARAM>(AIProtocolType::OpenAI));
	SendMessageA(hCombo, CB_SETITEMDATA, idxOpenAIResponses, static_cast<LPARAM>(AIProtocolType::OpenAIResponses));
	SendMessageA(hCombo, CB_SETITEMDATA, idxGemini, static_cast<LPARAM>(AIProtocolType::Gemini));
	SendMessageA(hCombo, CB_SETITEMDATA, idxClaude, static_cast<LPARAM>(AIProtocolType::Claude));

	int selectedIndex = idxOpenAI;
	if (selected == AIProtocolType::OpenAIResponses) {
		selectedIndex = idxOpenAIResponses;
	}
	else if (selected == AIProtocolType::Gemini) {
		selectedIndex = idxGemini;
	}
	else if (selected == AIProtocolType::Claude) {
		selectedIndex = idxClaude;
	}
	SendMessageA(hCombo, CB_SETCURSEL, selectedIndex, 0);
}

AIProtocolType GetSelectedProtocol(HWND hCombo)
{
	if (hCombo == nullptr) {
		return AIProtocolType::OpenAI;
	}

	const int selected = static_cast<int>(SendMessageA(hCombo, CB_GETCURSEL, 0, 0));
	if (selected == CB_ERR) {
		return AIProtocolType::OpenAI;
	}

	const LRESULT data = SendMessageA(hCombo, CB_GETITEMDATA, selected, 0);
	switch (static_cast<AIProtocolType>(data)) {
	case AIProtocolType::OpenAIResponses:
		return AIProtocolType::OpenAIResponses;
	case AIProtocolType::Gemini:
		return AIProtocolType::Gemini;
	case AIProtocolType::Claude:
		return AIProtocolType::Claude;
	case AIProtocolType::OpenAI:
	default:
		return AIProtocolType::OpenAI;
	}
}
HFONT GetLinkFont()
{
	static HFONT s_linkFont = nullptr;
	if (s_linkFont != nullptr) {
		return s_linkFont;
	}

	LOGFONTA lf = {};
	HFONT base = GetDialogFont();
	if (base != nullptr && GetObjectA(base, sizeof(lf), &lf) == sizeof(lf)) {
		lf.lfUnderline = TRUE;
		s_linkFont = CreateFontIndirectA(&lf);
	}
	if (s_linkFont == nullptr) {
		s_linkFont = reinterpret_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
	}
	return s_linkFont;
}

AISettings ReadAISettingsFromNativeDialog(AIConfigDialogContext* ctx)
{
	AISettings next = {};
	if (ctx == nullptr || ctx->settings == nullptr) {
		return next;
	}

	next = *ctx->settings;
	next.protocolType = GetSelectedProtocol(ctx->hProtocol);
	next.baseUrl = GetEditTextA(ctx->hBaseUrl);
	next.apiKey = GetEditTextA(ctx->hApiKey);
	next.model = GetEditTextA(ctx->hModel);
	next.tavilyApiKey = GetEditTextA(ctx->hTavilyApiKey);
	next.extraSystemPrompt = GetEditTextA(ctx->hExtraPrompt);
	return next;
}

bool ValidateAISettingsForConnection(HWND hWnd, const AISettings& settings)
{
	if (AIService::Trim(settings.baseUrl).empty() ||
		AIService::Trim(settings.apiKey).empty() ||
		AIService::Trim(settings.model).empty()) {
		MessageBoxA(hWnd, "baseUrl / apiKey / model cannot be empty.", "AI Config", MB_ICONWARNING | MB_OK);
		return false;
	}
	return true;
}

std::string BuildAIConnectionTestMessage(const AIResult& result)
{
	if (result.ok) {
		std::string message = "连通性测试成功。";
		if (result.httpStatus > 0) {
			message += "\nHTTP: " + std::to_string(result.httpStatus);
		}
		const std::string trimmed = AIService::Trim(result.content);
		if (!trimmed.empty()) {
			message += "\n模型返回：";
			message += trimmed;
		}
		return message;
	}

	std::string message = "连通性测试失败。";
	if (result.httpStatus > 0) {
		message += "\nHTTP: " + std::to_string(result.httpStatus);
	}
	if (!result.error.empty()) {
		message += "\n错误：";
		message += result.error;
	}
	return message;
}

void SetAIConfigNativeTestBusy(AIConfigDialogContext* ctx, bool busy)
{
	if (ctx == nullptr) {
		return;
	}
	ctx->testInFlight = busy;
	if (ctx->hTestConnection != nullptr) {
		SetWindowTextA(ctx->hTestConnection, busy ? "测试中..." : "测试连通性");
		EnableWindow(ctx->hTestConnection, busy ? FALSE : TRUE);
	}
}

DWORD WINAPI AIConfigConnectionTestWorkerProc(LPVOID param)
{
	AIConfigConnectionTestRequest* request = reinterpret_cast<AIConfigConnectionTestRequest*>(param);
	if (request == nullptr) {
		return 0;
	}

	AIConfigConnectionTestResult* result = new (std::nothrow) AIConfigConnectionTestResult();
	if (result == nullptr) {
		delete request;
		return 0;
	}

	result->forWebView = request->forWebView;
	try {
		result->result = AIService::TestConnection(request->settings);
	}
	catch (const std::exception& ex) {
		result->result.ok = false;
		result->result.error = std::string("connection test exception: ") + ex.what();
	}
	catch (...) {
		result->result.ok = false;
		result->result.error = "connection test unknown exception";
	}

	const HWND hWnd = request->dialogHwnd;
	delete request;
	if (hWnd == nullptr || !PostMessageA(hWnd, WM_AUTOLINKER_AI_CONFIG_TEST_DONE, 0, reinterpret_cast<LPARAM>(result))) {
		delete result;
	}
	return 0;
}

bool StartAIConfigConnectionTest(HWND hWnd, const AISettings& settings, bool forWebView)
{
	AIConfigConnectionTestRequest* request = new (std::nothrow) AIConfigConnectionTestRequest();
	if (request == nullptr) {
		MessageBoxA(hWnd, "无法启动连通性测试线程。", "AI Config", MB_ICONERROR | MB_OK);
		return false;
	}

	request->dialogHwnd = hWnd;
	request->settings = settings;
	request->forWebView = forWebView;

	const HANDLE workerHandle = CreateThread(nullptr, 0, AIConfigConnectionTestWorkerProc, request, 0, nullptr);
	if (workerHandle == nullptr) {
		delete request;
		MessageBoxA(hWnd, "无法启动连通性测试线程。", "AI Config", MB_ICONERROR | MB_OK);
		return false;
	}

	CloseHandle(workerHandle);
	return true;
}

std::string BuildAIConfigWebViewSettingsJson(const AISettings& settings)
{
	nlohmann::json initialSettings;
	initialSettings["protocolType"] = AIService::ProtocolTypeToString(settings.protocolType);
	initialSettings["baseUrl"] = LocalToUtf8Text(settings.baseUrl);
	initialSettings["apiKey"] = LocalToUtf8Text(settings.apiKey);
	initialSettings["model"] = LocalToUtf8Text(settings.model);
	initialSettings["extraPrompt"] = LocalToUtf8Text(settings.extraSystemPrompt);
	initialSettings["tavilyApiKey"] = LocalToUtf8Text(settings.tavilyApiKey);
	return initialSettings.dump();
}

std::string BuildAIConfigWebViewShellHtml()
{
	std::string html = LoadUtf8HtmlResourceText(IDR_HTML_AI_CONFIG_DIALOG);
	if (!html.empty()) {
		return html;
	}
	return "<!doctype html><html><head><meta charset=\"utf-8\"></head><body>Config shell resource missing.</body></html>";
}

void LayoutAIConfigWebViewDialog(HWND hWnd, AIConfigWebViewDialogContext* ctx)
{
	if (hWnd == nullptr || ctx == nullptr) {
		return;
	}

	RECT rc = {};
	GetClientRect(hWnd, &rc);
	const int hostWidth = static_cast<int>((std::max)(0L, rc.right));
	const int hostHeight = static_cast<int>((std::max)(0L, rc.bottom));
	const int loadingLeft = 16;
	const int loadingTop = 14;
	const int loadingWidth = static_cast<int>((std::max)(0L, rc.right - loadingLeft * 2L));
	if (ctx->hHost != nullptr) {
		MoveWindow(ctx->hHost, 0, 0, hostWidth, hostHeight, TRUE);
	}
	if (ctx->hLoading != nullptr) {
		MoveWindow(ctx->hLoading, loadingLeft, loadingTop, loadingWidth, 24, TRUE);
	}
	if (ctx->webViewController != nullptr) {
		RECT bounds = {};
		bounds.left = 0;
		bounds.top = 0;
		bounds.right = static_cast<LONG>(hostWidth);
		bounds.bottom = static_cast<LONG>(hostHeight);
		ctx->webViewController->put_Bounds(bounds);
	}
}

LRESULT CALLBACK AIConfigDialogProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	auto* ctx = reinterpret_cast<AIConfigDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
	switch (uMsg)
	{
	case WM_NCCREATE: {
		const auto* create = reinterpret_cast<CREATESTRUCTA*>(lParam);
		SetWindowLongPtrA(hWnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create->lpCreateParams));
		return TRUE;
	}

	case WM_CREATE: {
		if (ctx == nullptr || ctx->settings == nullptr) {
			return -1;
		}

		HWND hProtocolLabel = CreateWindowA("STATIC", "Protocol:", WS_CHILD | WS_VISIBLE,
			16, 16, 100, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hProtocol = CreateWindowExA(0, "COMBOBOX", "",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST,
			120, 14, 180, 220, hWnd, reinterpret_cast<HMENU>(IDC_CFG_PROTOCOL), nullptr, nullptr);
		PopulateProtocolCombo(ctx->hProtocol, ctx->settings->protocolType);

		HWND hBaseUrlLabel = CreateWindowA("STATIC", "Base URL:", WS_CHILD | WS_VISIBLE,
			16, 50, 100, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hBaseUrl = CreateWindowExA(0, "EDIT", ctx->settings->baseUrl.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL,
			120, 48, 500, 24, hWnd, reinterpret_cast<HMENU>(IDC_CFG_BASE_URL), nullptr, nullptr);

		HWND hApiKeyLabel = CreateWindowA("STATIC", "API Key:", WS_CHILD | WS_VISIBLE,
			16, 84, 100, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hApiKey = CreateWindowExA(0, "EDIT", ctx->settings->apiKey.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL | ES_PASSWORD,
			120, 82, 500, 24, hWnd, reinterpret_cast<HMENU>(IDC_CFG_API_KEY), nullptr, nullptr);

		HWND hModelLabel = CreateWindowA("STATIC", "Model:", WS_CHILD | WS_VISIBLE,
			16, 118, 100, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hModel = CreateWindowExA(0, "EDIT", ctx->settings->model.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL,
			120, 116, 500, 24, hWnd, reinterpret_cast<HMENU>(IDC_CFG_MODEL), nullptr, nullptr);

		const std::wstring getKeyLinkText =
			L"<a href=\"https://right.codes/register?aff=3dc87885\">从转发平台获取Key</a>";
		HWND hGetKeyLink = CreateWindowExW(0, L"SysLink",
			getKeyLinkText.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP,
			120, 146, 240, 22, hWnd, reinterpret_cast<HMENU>(IDC_CFG_GET_KEY_LINK), nullptr, nullptr);
		if (hGetKeyLink != nullptr) {
			ctx->useNativeLink = true;
			ctx->hGetKeyLink = hGetKeyLink;
		}
		else {
			ctx->useNativeLink = false;
			ctx->hGetKeyLink = CreateWindowW(
				L"STATIC",
				L"\u4ECE\u8F6C\u53D1\u5E73\u53F0\u83B7\u53D6Key",
				WS_CHILD | WS_VISIBLE | WS_TABSTOP | SS_NOTIFY,
				120, 146, 240, 22, hWnd, reinterpret_cast<HMENU>(IDC_CFG_GET_KEY_LINK), nullptr, nullptr);
			hGetKeyLink = ctx->hGetKeyLink;
		}

		ctx->hPlatformCombo = CreateWindowExW(0, L"COMBOBOX", nullptr,
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST | WS_VSCROLL,
			430, 142, 190, 280, hWnd, reinterpret_cast<HMENU>(IDC_CFG_PLATFORM_PRESET), nullptr, nullptr);
		// 填入平台列表
		static const wchar_t* kPlatformNames[] = {
			L"（平台预设）", L"Right", L"DeepSeek", L"\u667A\u8C31", L"\u5343\u95EE",
			L"Kimi", L"\u8C46\u5305", L"MiniMax", L"aihubmix",
			L"\u7845\u57FA\u6D41\u52A8", L"OpenAI", L"Claude", L"Gemini"
		};
		for (const wchar_t* name : kPlatformNames) {
			SendMessageW(ctx->hPlatformCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(name));
		}
		SendMessageW(ctx->hPlatformCombo, CB_SETCURSEL, 0, 0);

		HWND hTavilyApiKeyLabel = CreateWindowA("STATIC", "Tavily API Key:", WS_CHILD | WS_VISIBLE,
			16, 182, 100, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hTavilyApiKey = CreateWindowExA(0, "EDIT", ctx->settings->tavilyApiKey.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL | ES_PASSWORD,
			120, 180, 500, 24, hWnd, reinterpret_cast<HMENU>(IDC_CFG_TAVILY_API_KEY), nullptr, nullptr);

		HWND hExtraPromptLabel = CreateWindowA("STATIC", "System Prompt:", WS_CHILD | WS_VISIBLE,
			16, 216, 140, 20, hWnd, nullptr, nullptr, nullptr);
		ctx->hExtraPrompt = CreateWindowExA(0, "EDIT", ctx->settings->extraSystemPrompt.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_MULTILINE | ES_AUTOVSCROLL | WS_VSCROLL,
			120, 216, 500, 150, hWnd, reinterpret_cast<HMENU>(IDC_CFG_EXTRA_PROMPT), nullptr, nullptr);

		ctx->hTestConnection = CreateWindowW(L"BUTTON", L"\u6D4B\u8BD5\u8FDE\u901A\u6027", WS_CHILD | WS_VISIBLE | WS_TABSTOP,
			300, 382, 110, 28, hWnd, reinterpret_cast<HMENU>(IDC_CFG_TEST_CONNECTION), nullptr, nullptr);
		HWND hSave = CreateWindowW(L"BUTTON", L"\u4FDD\u5B58\u5E76\u7EE7\u7EED", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
			420, 382, 100, 28, hWnd, reinterpret_cast<HMENU>(IDC_CFG_SAVE), nullptr, nullptr);
		HWND hCancel = CreateWindowW(L"BUTTON", L"\u53D6\u6D88", WS_CHILD | WS_VISIBLE | WS_TABSTOP,
			530, 382, 90, 28, hWnd, reinterpret_cast<HMENU>(IDC_CFG_CANCEL), nullptr, nullptr);

		std::array<HWND, 17> controls = {
			hProtocolLabel,
			ctx->hProtocol,
			hBaseUrlLabel,
			hApiKeyLabel,
			hModelLabel,
			hTavilyApiKeyLabel,
			ctx->hGetKeyLink,
			ctx->hPlatformCombo,
			hExtraPromptLabel,
			ctx->hBaseUrl,
			ctx->hApiKey,
			ctx->hModel,
			ctx->hTavilyApiKey,
			ctx->hExtraPrompt,
			ctx->hTestConnection,
			hSave,
			hCancel
		};
		for (HWND hControl : controls) {
			SetDefaultFont(hControl);
		}
		if (ctx->hGetKeyLink != nullptr && !ctx->useNativeLink) {
			SendMessageA(ctx->hGetKeyLink, WM_SETFONT, reinterpret_cast<WPARAM>(GetLinkFont()), TRUE);
		}
		SetFocus(ctx->hBaseUrl);
		return 0;
	}

	case WM_COMMAND: {
		const int id = LOWORD(wParam);
		if (ctx == nullptr || ctx->settings == nullptr) {
			return 0;
		}

		if (id == IDC_CFG_SAVE) {
			AISettings next = ReadAISettingsFromNativeDialog(ctx);
			if (!ValidateAISettingsForConnection(hWnd, next)) {
				return 0;
			}

			*ctx->settings = next;
			ctx->accepted = true;
			DestroyWindow(hWnd);
			return 0;
		}

		if (id == IDC_CFG_TEST_CONNECTION) {
			if (ctx->testInFlight) {
				return 0;
			}

			const AISettings next = ReadAISettingsFromNativeDialog(ctx);
			if (!ValidateAISettingsForConnection(hWnd, next)) {
				return 0;
			}
			if (!StartAIConfigConnectionTest(hWnd, next, false)) {
				return 0;
			}
			SetAIConfigNativeTestBusy(ctx, true);
			return 0;
		}

		if (id == IDC_CFG_PLATFORM_PRESET && HIWORD(wParam) == CBN_SELCHANGE) {
			// 与 WebView2 中 PLATFORMS 数组保持相同顺序（索引 0 = 自定义）
			struct PlatformPreset {
				const char* baseUrl;
				AIProtocolType protocol;
			};
			static const PlatformPreset kPresets[] = {
				{ nullptr, AIProtocolType::OpenAI },                                            // 0: 自定义
				{ "https://right.codes/codex",                         AIProtocolType::OpenAI },  // Right
				{ "https://api.deepseek.com/v1",                       AIProtocolType::OpenAI },  // DeepSeek
				{ "https://open.bigmodel.cn/api/paas/v4",              AIProtocolType::OpenAI },  // 智谱
				{ "https://dashscope.aliyuncs.com/compatible-mode/v1", AIProtocolType::OpenAI },  // 千问
				{ "https://api.moonshot.cn/v1",                        AIProtocolType::OpenAI },  // Kimi
				{ "https://ark.cn-beijing.volces.com/api/v3",          AIProtocolType::OpenAI },  // 豆包
				{ "https://api.minimax.chat/v1",                       AIProtocolType::OpenAI },  // MiniMax
				{ "https://aihubmix.com/v1",                           AIProtocolType::OpenAI },  // aihubmix
				{ "https://api.siliconflow.cn/v1",                     AIProtocolType::OpenAI },  // 硅基流动
				{ "https://api.openai.com/v1",                         AIProtocolType::OpenAIResponses },  // OpenAI
				{ "https://api.anthropic.com",                         AIProtocolType::Claude  },  // Claude
				{ "https://generativelanguage.googleapis.com",         AIProtocolType::Gemini  },  // Gemini
			};
			const int sel = static_cast<int>(SendMessageW(ctx->hPlatformCombo, CB_GETCURSEL, 0, 0));
			if (sel > 0 && sel < static_cast<int>(std::size(kPresets))) {
				const auto& p = kPresets[sel];
				PopulateProtocolCombo(ctx->hProtocol, p.protocol);
				SetWindowTextA(ctx->hBaseUrl, p.baseUrl);
				SetFocus(ctx->hApiKey);
			}
			return 0;
		}
		if (id == IDC_CFG_GET_KEY_LINK && HIWORD(wParam) == STN_CLICKED) {
			ShellExecuteA(hWnd, "open", "https://right.codes/register?aff=3dc87885", nullptr, nullptr, SW_SHOWNORMAL);
			return 0;
		}

		if (id == IDC_CFG_CANCEL) {
			ctx->accepted = false;
			DestroyWindow(hWnd);
			return 0;
		}
		return 0;
	}

	case WM_NOTIFY: {
		const NMHDR* hdr = reinterpret_cast<const NMHDR*>(lParam);
		if (hdr != nullptr && ctx != nullptr && ctx->useNativeLink && hdr->idFrom == IDC_CFG_GET_KEY_LINK &&
			(hdr->code == NM_CLICK || hdr->code == NM_RETURN)) {
			ShellExecuteA(hWnd, "open", "https://right.codes/register?aff=3dc87885", nullptr, nullptr, SW_SHOWNORMAL);
			return 0;
		}
		break;
	}

	case WM_CTLCOLORSTATIC: {
		if (ctx != nullptr && !ctx->useNativeLink && reinterpret_cast<HWND>(lParam) == ctx->hGetKeyLink) {
			HDC hdc = reinterpret_cast<HDC>(wParam);
			SetTextColor(hdc, RGB(0, 102, 204));
			SetBkMode(hdc, TRANSPARENT);
			return reinterpret_cast<LRESULT>(GetSysColorBrush(COLOR_BTNFACE));
		}
		break;
	}

	case WM_AUTOLINKER_AI_CONFIG_TEST_DONE: {
		AIConfigConnectionTestResult* result = reinterpret_cast<AIConfigConnectionTestResult*>(lParam);
		if (ctx == nullptr || result == nullptr) {
			delete result;
			return 0;
		}

		SetAIConfigNativeTestBusy(ctx, false);
		const std::string message = BuildAIConnectionTestMessage(result->result);
		MessageBoxA(hWnd, message.c_str(), "AI Config", (result->result.ok ? MB_ICONINFORMATION : MB_ICONERROR) | MB_OK);
		delete result;
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

AISettings ReadAISettingsFromWebPayload(const AISettings& current, const nlohmann::json& data)
{
	AISettings next = current;
	next.protocolType = AIService::ParseProtocolType(data.value("protocol_type", AIService::ProtocolTypeToString(next.protocolType)));
	next.baseUrl = Utf8ToLocalText(data.value("base_url", ""));
	next.apiKey = Utf8ToLocalText(data.value("api_key", ""));
	next.model = Utf8ToLocalText(data.value("model", ""));
	next.extraSystemPrompt = Utf8ToLocalText(data.value("extra_system_prompt", ""));
	next.tavilyApiKey = Utf8ToLocalText(data.value("tavily_api_key", ""));
	return next;
}

bool TryApplyAISettingsFromWebPayload(HWND hWnd, AIConfigWebViewDialogContext* ctx, const nlohmann::json& data)
{
	if (ctx == nullptr || ctx->settings == nullptr) {
		return false;
	}

	AISettings next = ReadAISettingsFromWebPayload(*ctx->settings, data);

	if (!ValidateAISettingsForConnection(hWnd, next)) {
		return false;
	}

	*ctx->settings = next;
	ctx->accepted = true;
	DestroyWindow(hWnd);
	return true;
}

void ExecuteAIConfigWebViewScript(AIConfigWebViewDialogContext* ctx, const std::wstring& script)
{
	if (ctx == nullptr || !ctx->webViewReady || ctx->webView == nullptr || script.empty()) {
		return;
	}
	ctx->webView->ExecuteScript(script.c_str(), nullptr);
}

void SetAIConfigWebViewTestBusy(AIConfigWebViewDialogContext* ctx, bool busy)
{
	if (ctx == nullptr) {
		return;
	}
	ctx->testInFlight = busy;
	ExecuteAIConfigWebViewScript(ctx, busy
		? L"window.autolinkerSetTestBusy(true);"
		: L"window.autolinkerSetTestBusy(false);");
}

void ShowAIConfigWebViewTestResult(AIConfigWebViewDialogContext* ctx, const AIResult& result)
{
	if (ctx == nullptr) {
		return;
	}

	nlohmann::json payload;
	payload["ok"] = result.ok;
	payload["message"] = LocalToUtf8Text(BuildAIConnectionTestMessage(result));

	const std::wstring payloadJsonWide = Utf8ToWide(payload.dump());
	if (payloadJsonWide.empty()) {
		return;
	}

	std::wstring script = L"window.autolinkerShowTestResult(JSON.parse('";
	script += EscapeJsSingleQuotedWide(payloadJsonWide);
	script += L"'));";
	ExecuteAIConfigWebViewScript(ctx, script);
}

void ApplyAIConfigWebViewSettings(AIConfigWebViewDialogContext* ctx)
{
	if (ctx == nullptr || ctx->settings == nullptr || !ctx->webViewReady) {
		return;
	}

	const std::string settingsJsonUtf8 = BuildAIConfigWebViewSettingsJson(*ctx->settings);
	const std::wstring settingsJsonWide = Utf8ToWide(settingsJsonUtf8);
	if (settingsJsonWide.empty()) {
		OutputStringToELog("[AI Config][WebView2] settings json conversion failed");
		return;
	}

	std::wstring script = L"window.autolinkerApplySettings(JSON.parse('";
	script += EscapeJsSingleQuotedWide(settingsJsonWide);
	script += L"'));window.autolinkerFocusPrimary();";
	OutputStringToELog(std::format("[AI Config][WebView2] apply settings jsonChars={}", settingsJsonWide.size()));
	ExecuteAIConfigWebViewScript(ctx, script);
}

void StartAIConfigWebView(HWND hWnd, AIConfigWebViewDialogContext* ctx)
{
	if (hWnd == nullptr || ctx == nullptr || ctx->hHost == nullptr) {
		return;
	}

	const std::wstring webViewUserDataFolder = GetWebView2UserDataFolderPath();
	const HRESULT hr = CreateCoreWebView2EnvironmentWithOptions(
		nullptr,
		webViewUserDataFolder.empty() ? nullptr : webViewUserDataFolder.c_str(),
		nullptr,
		Microsoft::WRL::Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
			[hWnd](HRESULT envResult, ICoreWebView2Environment* environment) -> HRESULT {
				auto* innerCtx = reinterpret_cast<AIConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
				if (innerCtx == nullptr || !IsWindow(hWnd)) {
					return S_OK;
				}
				if (FAILED(envResult) || environment == nullptr) {
					OutputStringToELog(std::format("[AI Config][WebView2] create environment failed hr=0x{:08X}", static_cast<unsigned int>(envResult)));
					innerCtx->fallbackRequested = true;
					DestroyWindow(hWnd);
					return S_OK;
				}

				innerCtx->webViewEnvironment = environment;
				return environment->CreateCoreWebView2Controller(
					innerCtx->hHost,
					Microsoft::WRL::Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
						[hWnd](HRESULT controllerResult, ICoreWebView2Controller* controller) -> HRESULT {
							auto* readyCtx = reinterpret_cast<AIConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
							if (readyCtx == nullptr || !IsWindow(hWnd)) {
								return S_OK;
							}
							if (FAILED(controllerResult) || controller == nullptr) {
								OutputStringToELog(std::format("[AI Config][WebView2] create controller failed hr=0x{:08X}", static_cast<unsigned int>(controllerResult)));
								readyCtx->fallbackRequested = true;
								DestroyWindow(hWnd);
								return S_OK;
							}

							readyCtx->webViewController = controller;
							readyCtx->webViewController->get_CoreWebView2(&readyCtx->webView);
							if (readyCtx->webView == nullptr) {
								OutputStringToELog("[AI Config][WebView2] get_CoreWebView2 returned null");
								readyCtx->fallbackRequested = true;
								DestroyWindow(hWnd);
								return S_OK;
							}

							Microsoft::WRL::ComPtr<ICoreWebView2Settings> webSettings;
							if (SUCCEEDED(readyCtx->webView->get_Settings(&webSettings)) && webSettings != nullptr) {
								webSettings->put_AreDevToolsEnabled(FALSE);
								webSettings->put_IsStatusBarEnabled(FALSE);
								webSettings->put_IsZoomControlEnabled(FALSE);
							}

							readyCtx->webView->add_WebMessageReceived(
								Microsoft::WRL::Callback<ICoreWebView2WebMessageReceivedEventHandler>(
									[hWnd](ICoreWebView2*, ICoreWebView2WebMessageReceivedEventArgs* args) -> HRESULT {
										auto* messageCtx = reinterpret_cast<AIConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
										if (messageCtx == nullptr || args == nullptr || !IsWindow(hWnd)) {
											return S_OK;
										}

										LPWSTR rawMessage = nullptr;
										if (FAILED(args->TryGetWebMessageAsString(&rawMessage)) || rawMessage == nullptr) {
											return S_OK;
										}

										const std::string utf8Message = WideToUtf8(rawMessage);
										CoTaskMemFree(rawMessage);
										try {
											const nlohmann::json payload = nlohmann::json::parse(utf8Message);
											const std::string action = payload.value("action", "");
											if (action == "save" && payload.contains("data") && payload["data"].is_object()) {
												TryApplyAISettingsFromWebPayload(hWnd, messageCtx, payload["data"]);
											}
											else if (action == "test_connection" && payload.contains("data") && payload["data"].is_object()) {
												if (!messageCtx->testInFlight) {
													const AISettings next = ReadAISettingsFromWebPayload(*messageCtx->settings, payload["data"]);
													if (ValidateAISettingsForConnection(hWnd, next) &&
														StartAIConfigConnectionTest(hWnd, next, true)) {
														SetAIConfigWebViewTestBusy(messageCtx, true);
													}
												}
											}
											else if (action == "cancel") {
												DestroyWindow(hWnd);
											}
											else if (action == "open_url") {
												const std::string openUrl = payload.value("url", "");
												if (!openUrl.empty() &&
													(openUrl.rfind("https://", 0) == 0 || openUrl.rfind("http://", 0) == 0)) {
													ShellExecuteA(hWnd, "open", openUrl.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
												}
											}
										}
										catch (...) {
										}
										return S_OK;
									}).Get(),
								nullptr);

							readyCtx->webView->add_NavigationCompleted(
								Microsoft::WRL::Callback<ICoreWebView2NavigationCompletedEventHandler>(
									[hWnd](ICoreWebView2*, ICoreWebView2NavigationCompletedEventArgs* args) -> HRESULT {
										auto* navCtx = reinterpret_cast<AIConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
										if (navCtx == nullptr || args == nullptr || !IsWindow(hWnd)) {
											return S_OK;
										}

										BOOL isSuccess = FALSE;
										args->get_IsSuccess(&isSuccess);
										if (isSuccess == TRUE) {
											navCtx->webViewReady = true;
											KillTimer(hWnd, kAIConfigWebViewInitTimerId);
											if (navCtx->hLoading != nullptr) {
												ShowWindow(navCtx->hLoading, SW_HIDE);
											}
											OutputStringToELog("[AI Config][WebView2] navigation completed successfully");
											LayoutAIConfigWebViewDialog(hWnd, navCtx);
											ApplyAIConfigWebViewSettings(navCtx);
											return S_OK;
										}

										COREWEBVIEW2_WEB_ERROR_STATUS webErrorStatus = COREWEBVIEW2_WEB_ERROR_STATUS_UNKNOWN;
										args->get_WebErrorStatus(&webErrorStatus);
										if (webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_OPERATION_CANCELED ||
											webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_ABORTED ||
											webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_RESET) {
											OutputStringToELog(std::format(
												"[AI Config][WebView2] navigation superseded errorStatus={}",
												static_cast<int>(webErrorStatus)));
											return S_OK;
										}

										OutputStringToELog(std::format(
											"[AI Config][WebView2] navigation failed errorStatus={}",
											static_cast<int>(webErrorStatus)));
										navCtx->fallbackRequested = true;
										DestroyWindow(hWnd);
										return S_OK;
									}).Get(),
								nullptr);

							LayoutAIConfigWebViewDialog(hWnd, readyCtx);
							const std::wstring shellHtml = Utf8ToWide(BuildAIConfigWebViewShellHtml());
							if (shellHtml.empty()) {
								OutputStringToELog("[AI Config][WebView2] shell html conversion failed");
								readyCtx->fallbackRequested = true;
								DestroyWindow(hWnd);
								return S_OK;
							}
							OutputStringToELog(std::format("[AI Config][WebView2] controller ready=1 shellHtmlChars={}", shellHtml.size()));
							readyCtx->webView->NavigateToString(shellHtml.c_str());
							return S_OK;
						}).Get());
			}).Get());

	if (FAILED(hr)) {
		OutputStringToELog(std::format("[AI Config][WebView2] bootstrap failed hr=0x{:08X}", static_cast<unsigned int>(hr)));
		ctx->fallbackRequested = true;
		DestroyWindow(hWnd);
	}
}

LRESULT CALLBACK AIConfigWebViewDialogProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	auto* ctx = reinterpret_cast<AIConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
	switch (uMsg)
	{
	case WM_NCCREATE: {
		const auto* create = reinterpret_cast<CREATESTRUCTA*>(lParam);
		SetWindowLongPtrA(hWnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create->lpCreateParams));
		return TRUE;
	}

	case WM_CREATE: {
		if (ctx == nullptr || ctx->settings == nullptr) {
			return -1;
		}

		ctx->hHost = CreateWindowExA(
			0,
			"STATIC",
			"",
			WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
			0,
			0,
			0,
			0,
			hWnd,
			nullptr,
			nullptr,
			nullptr);
		ctx->hLoading = CreateWindowExW(
			0,
			L"STATIC",
			L"正在初始化 WebView2 设置页...",
			WS_CHILD | WS_VISIBLE,
			0,
			0,
			0,
			0,
			hWnd,
			nullptr,
			nullptr,
			nullptr);
		SetDefaultFont(ctx->hLoading);
		LayoutAIConfigWebViewDialog(hWnd, ctx);
		SetTimer(hWnd, kAIConfigWebViewInitTimerId, kAIConfigWebViewInitTimeoutMs, nullptr);
		StartAIConfigWebView(hWnd, ctx);
		return 0;
	}

	case WM_GETMINMAXINFO: {
		auto* mmi = reinterpret_cast<MINMAXINFO*>(lParam);
		if (mmi != nullptr) {
			mmi->ptMinTrackSize.x = (std::max)(mmi->ptMinTrackSize.x, 820L);
			mmi->ptMinTrackSize.y = (std::max)(mmi->ptMinTrackSize.y, 620L);
		}
		return 0;
	}

	case WM_SIZE:
		if (ctx != nullptr) {
			LayoutAIConfigWebViewDialog(hWnd, ctx);
		}
		return 0;

	case WM_TIMER:
		if (ctx != nullptr && wParam == kAIConfigWebViewInitTimerId) {
			if (!ctx->webViewReady) {
				OutputStringToELog("[AI Config][WebView2] initialization timed out, fallback to native dialog");
				ctx->fallbackRequested = true;
				DestroyWindow(hWnd);
				return 0;
			}
			KillTimer(hWnd, kAIConfigWebViewInitTimerId);
			return 0;
		}
		break;

	case WM_AUTOLINKER_AI_CONFIG_TEST_DONE: {
		AIConfigConnectionTestResult* result = reinterpret_cast<AIConfigConnectionTestResult*>(lParam);
		if (ctx == nullptr || result == nullptr) {
			delete result;
			return 0;
		}

		SetAIConfigWebViewTestBusy(ctx, false);
		ShowAIConfigWebViewTestResult(ctx, result->result);
		delete result;
		return 0;
	}

	case WM_CLOSE:
		DestroyWindow(hWnd);
		return 0;

	case WM_DESTROY:
		KillTimer(hWnd, kAIConfigWebViewInitTimerId);
		if (ctx != nullptr) {
			ctx->webView = nullptr;
			ctx->webViewController = nullptr;
			ctx->webViewEnvironment = nullptr;
		}
		return 0;

	default:
		break;
	}

	return DefWindowProcA(hWnd, uMsg, wParam, lParam);
}

std::string BuildAIPreviewWebViewPayloadJson(const AIPreviewWebViewDialogContext& ctx)
{
	nlohmann::json payload;
	payload["title"] = LocalToUtf8Text(ctx.title);
	payload["content"] = LocalToUtf8Text(ctx.content);
	payload["primaryText"] = LocalToUtf8Text(ctx.primaryText);
	payload["secondaryText"] = LocalToUtf8Text(ctx.secondaryText);
	return payload.dump();
}

std::string BuildAIPreviewWebViewShellHtml()
{
	std::string html = LoadUtf8HtmlResourceText(IDR_HTML_AI_PREVIEW_DIALOG);
	if (!html.empty()) {
		return html;
	}
	return "<!doctype html><html><head><meta charset=\"utf-8\"></head><body>Preview shell resource missing.</body></html>";
}

void LayoutAIPreviewWebViewDialog(HWND hWnd, AIPreviewWebViewDialogContext* ctx)
{
	if (hWnd == nullptr || ctx == nullptr) {
		return;
	}

	RECT rc = {};
	GetClientRect(hWnd, &rc);
	const int hostWidth = static_cast<int>((std::max)(0L, rc.right));
	const int hostHeight = static_cast<int>((std::max)(0L, rc.bottom));
	const int loadingLeft = 16;
	const int loadingTop = 14;
	const int loadingWidth = static_cast<int>((std::max)(0L, rc.right - loadingLeft * 2L));
	if (ctx->hHost != nullptr) {
		MoveWindow(ctx->hHost, 0, 0, hostWidth, hostHeight, TRUE);
	}
	if (ctx->hLoading != nullptr) {
		MoveWindow(ctx->hLoading, loadingLeft, loadingTop, loadingWidth, 24, TRUE);
	}
	if (ctx->webViewController != nullptr) {
		RECT bounds = {};
		bounds.left = 0;
		bounds.top = 0;
		bounds.right = static_cast<LONG>(hostWidth);
		bounds.bottom = static_cast<LONG>(hostHeight);
		ctx->webViewController->put_Bounds(bounds);
	}
}

void ExecuteAIPreviewWebViewScript(AIPreviewWebViewDialogContext* ctx, const std::wstring& script)
{
	if (ctx == nullptr || !ctx->webViewReady || ctx->webView == nullptr || script.empty()) {
		return;
	}
	ctx->webView->ExecuteScript(script.c_str(), nullptr);
}

void ApplyAIPreviewWebViewPayload(AIPreviewWebViewDialogContext* ctx)
{
	if (ctx == nullptr || !ctx->webViewReady) {
		return;
	}

	const std::string payloadJsonUtf8 = BuildAIPreviewWebViewPayloadJson(*ctx);
	const std::wstring payloadJsonWide = Utf8ToWide(payloadJsonUtf8);
	if (payloadJsonWide.empty()) {
		OutputStringToELog("[AI Preview][WebView2] payload json conversion failed");
		return;
	}

	std::wstring script = L"window.autolinkerApplyPreview(JSON.parse('";
	script += EscapeJsSingleQuotedWide(payloadJsonWide);
	script += L"'));window.autolinkerFocusPrimary();";
	OutputStringToELog(std::format("[AI Preview][WebView2] apply payload jsonChars={}", payloadJsonWide.size()));
	ExecuteAIPreviewWebViewScript(ctx, script);
}

void StartAIPreviewWebView(HWND hWnd, AIPreviewWebViewDialogContext* ctx)
{
	if (hWnd == nullptr || ctx == nullptr || ctx->hHost == nullptr) {
		return;
	}

	const std::wstring webViewUserDataFolder = GetWebView2UserDataFolderPath();
	const HRESULT hr = CreateCoreWebView2EnvironmentWithOptions(
		nullptr,
		webViewUserDataFolder.empty() ? nullptr : webViewUserDataFolder.c_str(),
		nullptr,
		Microsoft::WRL::Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
			[hWnd](HRESULT envResult, ICoreWebView2Environment* environment) -> HRESULT {
				auto* innerCtx = reinterpret_cast<AIPreviewWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
				if (innerCtx == nullptr || !IsWindow(hWnd)) {
					return S_OK;
				}
				if (FAILED(envResult) || environment == nullptr) {
					OutputStringToELog(std::format("[AI Preview][WebView2] create environment failed hr=0x{:08X}", static_cast<unsigned int>(envResult)));
					innerCtx->fallbackRequested = true;
					DestroyWindow(hWnd);
					return S_OK;
				}

				innerCtx->webViewEnvironment = environment;
				return environment->CreateCoreWebView2Controller(
					innerCtx->hHost,
					Microsoft::WRL::Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
						[hWnd](HRESULT controllerResult, ICoreWebView2Controller* controller) -> HRESULT {
							auto* readyCtx = reinterpret_cast<AIPreviewWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
							if (readyCtx == nullptr || !IsWindow(hWnd)) {
								return S_OK;
							}
							if (FAILED(controllerResult) || controller == nullptr) {
								OutputStringToELog(std::format("[AI Preview][WebView2] create controller failed hr=0x{:08X}", static_cast<unsigned int>(controllerResult)));
								readyCtx->fallbackRequested = true;
								DestroyWindow(hWnd);
								return S_OK;
							}

							readyCtx->webViewController = controller;
							readyCtx->webViewController->get_CoreWebView2(&readyCtx->webView);
							if (readyCtx->webView == nullptr) {
								OutputStringToELog("[AI Preview][WebView2] get_CoreWebView2 returned null");
								readyCtx->fallbackRequested = true;
								DestroyWindow(hWnd);
								return S_OK;
							}

							Microsoft::WRL::ComPtr<ICoreWebView2Settings> webSettings;
							if (SUCCEEDED(readyCtx->webView->get_Settings(&webSettings)) && webSettings != nullptr) {
								webSettings->put_AreDevToolsEnabled(FALSE);
								webSettings->put_AreDefaultContextMenusEnabled(FALSE);
								webSettings->put_IsStatusBarEnabled(FALSE);
								webSettings->put_IsZoomControlEnabled(FALSE);
							}

							readyCtx->webView->add_WebMessageReceived(
								Microsoft::WRL::Callback<ICoreWebView2WebMessageReceivedEventHandler>(
									[hWnd](ICoreWebView2*, ICoreWebView2WebMessageReceivedEventArgs* args) -> HRESULT {
										auto* messageCtx = reinterpret_cast<AIPreviewWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
										if (messageCtx == nullptr || args == nullptr || !IsWindow(hWnd)) {
											return S_OK;
										}

										LPWSTR rawMessage = nullptr;
										if (FAILED(args->TryGetWebMessageAsString(&rawMessage)) || rawMessage == nullptr) {
											return S_OK;
										}

										const std::string utf8Message = WideToUtf8(rawMessage);
										CoTaskMemFree(rawMessage);
										try {
											const nlohmann::json payload = nlohmann::json::parse(utf8Message);
											const std::string action = payload.value("action", "");
											if (action == "primary") {
												messageCtx->action = AIPreviewAction::PrimaryConfirm;
												DestroyWindow(hWnd);
											}
											else if (action == "secondary") {
												messageCtx->action = AIPreviewAction::SecondaryConfirm;
												DestroyWindow(hWnd);
											}
											else if (action == "cancel") {
												messageCtx->action = AIPreviewAction::Cancel;
												DestroyWindow(hWnd);
											}
										}
										catch (...) {
										}
										return S_OK;
									}).Get(),
								nullptr);

							readyCtx->webView->add_NavigationCompleted(
								Microsoft::WRL::Callback<ICoreWebView2NavigationCompletedEventHandler>(
									[hWnd](ICoreWebView2*, ICoreWebView2NavigationCompletedEventArgs* args) -> HRESULT {
										auto* navCtx = reinterpret_cast<AIPreviewWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
										if (navCtx == nullptr || args == nullptr || !IsWindow(hWnd)) {
											return S_OK;
										}

										BOOL isSuccess = FALSE;
										args->get_IsSuccess(&isSuccess);
										if (isSuccess == TRUE) {
											navCtx->webViewReady = true;
											KillTimer(hWnd, kAIPreviewWebViewInitTimerId);
											if (navCtx->hLoading != nullptr) {
												ShowWindow(navCtx->hLoading, SW_HIDE);
											}
											OutputStringToELog("[AI Preview][WebView2] navigation completed successfully");
											LayoutAIPreviewWebViewDialog(hWnd, navCtx);
											ApplyAIPreviewWebViewPayload(navCtx);
											return S_OK;
										}

										COREWEBVIEW2_WEB_ERROR_STATUS webErrorStatus = COREWEBVIEW2_WEB_ERROR_STATUS_UNKNOWN;
										args->get_WebErrorStatus(&webErrorStatus);
										if (webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_OPERATION_CANCELED ||
											webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_ABORTED ||
											webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_RESET) {
											OutputStringToELog(std::format(
												"[AI Preview][WebView2] navigation superseded errorStatus={}",
												static_cast<int>(webErrorStatus)));
											return S_OK;
										}

										OutputStringToELog(std::format(
											"[AI Preview][WebView2] navigation failed errorStatus={}",
											static_cast<int>(webErrorStatus)));
										navCtx->fallbackRequested = true;
										DestroyWindow(hWnd);
										return S_OK;
									}).Get(),
								nullptr);

							LayoutAIPreviewWebViewDialog(hWnd, readyCtx);
							const std::wstring shellHtml = Utf8ToWide(BuildAIPreviewWebViewShellHtml());
							if (shellHtml.empty()) {
								OutputStringToELog("[AI Preview][WebView2] shell html conversion failed");
								readyCtx->fallbackRequested = true;
								DestroyWindow(hWnd);
								return S_OK;
							}
							OutputStringToELog(std::format("[AI Preview][WebView2] controller ready=1 shellHtmlChars={}", shellHtml.size()));
							readyCtx->webView->NavigateToString(shellHtml.c_str());
							return S_OK;
						}).Get());
			}).Get());

	if (FAILED(hr)) {
		OutputStringToELog(std::format("[AI Preview][WebView2] bootstrap failed hr=0x{:08X}", static_cast<unsigned int>(hr)));
		ctx->fallbackRequested = true;
		DestroyWindow(hWnd);
	}
}

LRESULT CALLBACK AIPreviewWebViewDialogProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	auto* ctx = reinterpret_cast<AIPreviewWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
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

		ctx->hHost = CreateWindowExA(
			0,
			"STATIC",
			"",
			WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
			0,
			0,
			0,
			0,
			hWnd,
			nullptr,
			nullptr,
			nullptr);
		ctx->hLoading = CreateWindowExW(
			0,
			L"STATIC",
			L"正在初始化 WebView2 确认窗口...",
			WS_CHILD | WS_VISIBLE,
			0,
			0,
			0,
			0,
			hWnd,
			nullptr,
			nullptr,
			nullptr);
		SetDefaultFont(ctx->hLoading);
		LayoutAIPreviewWebViewDialog(hWnd, ctx);
		SetTimer(hWnd, kAIPreviewWebViewInitTimerId, kAIPreviewWebViewInitTimeoutMs, nullptr);
		StartAIPreviewWebView(hWnd, ctx);
		return 0;
	}

	case WM_GETMINMAXINFO: {
		auto* mmi = reinterpret_cast<MINMAXINFO*>(lParam);
		if (mmi != nullptr) {
			mmi->ptMinTrackSize.x = (std::max)(mmi->ptMinTrackSize.x, 760L);
			mmi->ptMinTrackSize.y = (std::max)(mmi->ptMinTrackSize.y, 520L);
		}
		return 0;
	}

	case WM_SIZE:
		if (ctx != nullptr) {
			LayoutAIPreviewWebViewDialog(hWnd, ctx);
		}
		return 0;

	case WM_TIMER:
		if (ctx != nullptr && wParam == kAIPreviewWebViewInitTimerId) {
			if (!ctx->webViewReady) {
				OutputStringToELog("[AI Preview][WebView2] initialization timed out, fallback to native dialog");
				ctx->fallbackRequested = true;
				DestroyWindow(hWnd);
				return 0;
			}
			KillTimer(hWnd, kAIPreviewWebViewInitTimerId);
			return 0;
		}
		break;

	case WM_CLOSE:
		if (ctx != nullptr) {
			ctx->action = AIPreviewAction::Cancel;
		}
		DestroyWindow(hWnd);
		return 0;

	case WM_DESTROY:
		KillTimer(hWnd, kAIPreviewWebViewInitTimerId);
		if (ctx != nullptr) {
			ctx->webView = nullptr;
			ctx->webViewController = nullptr;
			ctx->webViewEnvironment = nullptr;
		}
		return 0;

	default:
		break;
	}

	return DefWindowProcA(hWnd, uMsg, wParam, lParam);
}

struct AIPreviewDialogContext {
	std::string title;
	std::string content;
	std::string primaryText;
	std::string secondaryText;
	AIPreviewAction action = AIPreviewAction::Cancel;
	HWND hEdit = nullptr;
	HWND hPrimary = nullptr;
	HWND hSecondary = nullptr;
	HWND hCancel = nullptr;
};

struct AIInputDialogContext {
	std::string title;
	std::string hint;
	std::string text;
	bool accepted = false;
	HWND hEdit = nullptr;
};

void LayoutAIPreviewDialog(HWND hWnd, AIPreviewDialogContext* ctx)
{
	if (hWnd == nullptr || ctx == nullptr) {
		return;
	}

	RECT rc = {};
	GetClientRect(hWnd, &rc);
	const int margin = 14;
	const int gap = 8;
	const int buttonW = 120;
	const int buttonH = 30;

	int right = static_cast<int>(rc.right) - margin;
	const int buttonTop = (std::max)(margin, static_cast<int>(rc.bottom) - margin - buttonH);

	if (ctx->hCancel != nullptr) {
		MoveWindow(ctx->hCancel, right - buttonW, buttonTop, buttonW, buttonH, TRUE);
		right -= (buttonW + gap);
	}
	if (ctx->hSecondary != nullptr) {
		MoveWindow(ctx->hSecondary, right - buttonW, buttonTop, buttonW, buttonH, TRUE);
		right -= (buttonW + gap);
	}
	if (ctx->hPrimary != nullptr) {
		MoveWindow(ctx->hPrimary, right - buttonW, buttonTop, buttonW, buttonH, TRUE);
	}

	const int editLeft = margin;
	const int editTop = margin;
	const int editWidth = (std::max)(120, static_cast<int>(rc.right) - margin * 2);
	const int editHeight = (std::max)(120, buttonTop - margin - editTop);
	if (ctx->hEdit != nullptr) {
		MoveWindow(ctx->hEdit, editLeft, editTop, editWidth, editHeight, TRUE);
	}
}

LRESULT CALLBACK AIPreviewDialogProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	auto* ctx = reinterpret_cast<AIPreviewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
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
			WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY | WS_VSCROLL,
			14, 14, 752, 450, hWnd, reinterpret_cast<HMENU>(IDC_PREVIEW_EDIT), nullptr, nullptr);
		const std::string displayText = NormalizeMultilineForEdit(ctx->content);
		SetWindowTextA(ctx->hEdit, displayText.c_str());

		ctx->hPrimary = CreateWindowA("BUTTON", ctx->primaryText.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
			590, 474, 120, 30, hWnd,
			reinterpret_cast<HMENU>(IDC_PREVIEW_OK), nullptr, nullptr);

		if (!ctx->secondaryText.empty()) {
			ctx->hSecondary = CreateWindowA("BUTTON", ctx->secondaryText.c_str(),
				WS_CHILD | WS_VISIBLE | WS_TABSTOP,
				460, 474, 120, 30, hWnd,
				reinterpret_cast<HMENU>(IDC_PREVIEW_SECONDARY), nullptr, nullptr);
		}

		ctx->hCancel = CreateWindowW(L"BUTTON", L"\u53D6\u6D88",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP,
			682, 474, 120, 30, hWnd,
			reinterpret_cast<HMENU>(IDC_PREVIEW_CANCEL), nullptr, nullptr);

		SetDefaultFont(ctx->hEdit);
		SetDefaultFont(ctx->hPrimary);
		SetDefaultFont(ctx->hSecondary);
		SetDefaultFont(ctx->hCancel);
		LayoutAIPreviewDialog(hWnd, ctx);
		SetFocus(ctx->hPrimary != nullptr ? ctx->hPrimary : ctx->hEdit);
		return 0;
	}

	case WM_SIZE:
		if (ctx != nullptr) {
			LayoutAIPreviewDialog(hWnd, ctx);
		}
		return 0;

	case WM_GETMINMAXINFO: {
		auto* mm = reinterpret_cast<MINMAXINFO*>(lParam);
		if (mm != nullptr) {
			mm->ptMinTrackSize.x = 680;
			mm->ptMinTrackSize.y = 420;
		}
		return 0;
	}

	case WM_COMMAND: {
		const int id = LOWORD(wParam);
		if (ctx == nullptr) {
			return 0;
		}
		if (id == IDC_PREVIEW_OK) {
			ctx->action = AIPreviewAction::PrimaryConfirm;
			DestroyWindow(hWnd);
			return 0;
		}
		if (id == IDC_PREVIEW_SECONDARY) {
			ctx->action = AIPreviewAction::SecondaryConfirm;
			DestroyWindow(hWnd);
			return 0;
		}
		if (id == IDC_PREVIEW_CANCEL) {
			ctx->action = AIPreviewAction::Cancel;
			DestroyWindow(hWnd);
			return 0;
		}
		return 0;
	}

	case WM_CLOSE:
		if (ctx != nullptr) {
			ctx->action = AIPreviewAction::Cancel;
		}
		DestroyWindow(hWnd);
		return 0;
	default:
		break;
	}
	return DefWindowProcA(hWnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK AIInputDialogProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	auto* ctx = reinterpret_cast<AIInputDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
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

		HWND hHint = CreateWindowA("STATIC", ctx->hint.c_str(),
			WS_CHILD | WS_VISIBLE, 14, 14, 612, 24, hWnd, nullptr, nullptr, nullptr);

		ctx->hEdit = CreateWindowExA(WS_EX_CLIENTEDGE, "EDIT", ctx->text.c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_MULTILINE | ES_AUTOVSCROLL | WS_VSCROLL,
			14, 44, 612, 200, hWnd, reinterpret_cast<HMENU>(IDC_INPUT_EDIT), nullptr, nullptr);

		HWND hOk = CreateWindowW(L"BUTTON", L"\u786E\u5B9A",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP, 448, 254, 84, 30, hWnd,
			reinterpret_cast<HMENU>(IDC_INPUT_OK), nullptr, nullptr);
		HWND hCancel = CreateWindowW(L"BUTTON", L"\u53D6\u6D88",
			WS_CHILD | WS_VISIBLE | WS_TABSTOP, 542, 254, 84, 30, hWnd,
			reinterpret_cast<HMENU>(IDC_INPUT_CANCEL), nullptr, nullptr);

		SetDefaultFont(hHint);
		SetDefaultFont(ctx->hEdit);
		SetDefaultFont(hOk);
		SetDefaultFont(hCancel);
		SetFocus(ctx->hEdit);
		return 0;
	}

	case WM_COMMAND: {
		if (ctx == nullptr) {
			return 0;
		}
		const int id = LOWORD(wParam);
		if (id == IDC_INPUT_OK) {
			ctx->text = GetEditTextA(ctx->hEdit);
			ctx->accepted = true;
			DestroyWindow(hWnd);
			return 0;
		}
		if (id == IDC_INPUT_CANCEL) {
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
} // namespace

bool ShowAIConfigDialogNative(HWND owner, AISettings& ioSettings)
{
	ComCtl6ActivationScope themeScope;
	INITCOMMONCONTROLSEX icex = {};
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_STANDARD_CLASSES | ICC_WIN95_CLASSES | ICC_LINK_CLASS;
	InitCommonControlsEx(&icex);

	WNDCLASSEXA wc = {};
	wc.cbSize = sizeof(wc);
	wc.lpfnWndProc = AIConfigDialogProc;
	wc.hInstance = GetModuleHandleA(nullptr);
	wc.lpszClassName = "AutoLinkerAIConfigDialogWindow";
	wc.hIcon = GetAppIconLarge();
	wc.hIconSm = GetAppIconSmall();
	wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_BTNFACE + 1);
	RegisterClassExA(&wc);

	AIConfigDialogContext ctx = {};
	ctx.settings = &ioSettings;

	HWND hDialog = CreateWindowExA(
		WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
		wc.lpszClassName,
		"AutoLinker AI Config",
		WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
		CW_USEDEFAULT, CW_USEDEFAULT, 650, 464,
		owner,
		nullptr,
		wc.hInstance,
		&ctx);

	if (hDialog == nullptr) {
		return false;
	}

	ApplyWindowIcon(hDialog);
	EnsureWindowTitle(hDialog, "AutoLinker AI Config");
	RunModalWindow(owner, hDialog);
	return ctx.accepted;
}

AIConfigDialogRunResult ShowAIConfigDialogWebView(HWND owner, AISettings& ioSettings)
{
	AIConfigDialogRunResult result = {};
	if (!IsWebView2RuntimeAvailable()) {
		OutputStringToELog("[AI Config][WebView2] runtime unavailable, fallback to native dialog");
		result.fallbackRequested = true;
		return result;
	}

	ComCtl6ActivationScope themeScope;
	INITCOMMONCONTROLSEX icex = {};
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_STANDARD_CLASSES | ICC_WIN95_CLASSES | ICC_LINK_CLASS;
	InitCommonControlsEx(&icex);

	WNDCLASSEXA wc = {};
	wc.cbSize = sizeof(wc);
	wc.lpfnWndProc = AIConfigWebViewDialogProc;
	wc.hInstance = GetModuleHandleA(nullptr);
	wc.lpszClassName = "AutoLinkerAIConfigWebViewDialogWindow";
	wc.hIcon = GetAppIconLarge();
	wc.hIconSm = GetAppIconSmall();
	wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_BTNFACE + 1);
	RegisterClassExA(&wc);

	AIConfigWebViewDialogContext ctx = {};
	ctx.settings = &ioSettings;

	HWND hDialog = CreateWindowExA(
		WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
		wc.lpszClassName,
		"AutoLinker AI Config",
		WS_CAPTION | WS_SYSMENU | WS_VISIBLE | WS_THICKFRAME,
		CW_USEDEFAULT, CW_USEDEFAULT, 1060, 860,
		owner,
		nullptr,
		wc.hInstance,
		&ctx);

	if (hDialog == nullptr) {
		OutputStringToELog("[AI Config][WebView2] CreateWindowExA failed, fallback to native dialog");
		result.fallbackRequested = true;
		return result;
	}

	ApplyWindowIcon(hDialog);
	EnsureWindowTitle(hDialog, "AutoLinker AI Config");
	RunModalWindow(owner, hDialog);
	result.accepted = ctx.accepted;
	result.fallbackRequested = ctx.fallbackRequested;
	return result;
}

bool ShowAIConfigDialog(HWND owner, AISettings& ioSettings)
{
	const AIConfigDialogRunResult webViewResult = ShowAIConfigDialogWebView(owner, ioSettings);
	if (webViewResult.fallbackRequested) {
		OutputStringToELog("[AI Config][WebView2] fallback to native dialog");
		return ShowAIConfigDialogNative(owner, ioSettings);
	}
	return webViewResult.accepted;
}

AIPreviewAction ShowAIPreviewDialogNative(
	HWND owner,
	const std::string& title,
	const std::string& content,
	const std::string& primaryText,
	const std::string& secondaryText)
{
	ComCtl6ActivationScope themeScope;
	INITCOMMONCONTROLSEX icex = {};
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_STANDARD_CLASSES | ICC_WIN95_CLASSES | ICC_LINK_CLASS;
	InitCommonControlsEx(&icex);

	WNDCLASSEXA wc = {};
	wc.cbSize = sizeof(wc);
	wc.lpfnWndProc = AIPreviewDialogProc;
	wc.hInstance = GetModuleHandleA(nullptr);
	wc.lpszClassName = "AutoLinkerAIPreviewDialogWindow";
	wc.hIcon = GetAppIconLarge();
	wc.hIconSm = GetAppIconSmall();
	wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_BTNFACE + 1);
	RegisterClassExA(&wc);

	AIPreviewDialogContext ctx = {};
	ctx.title = title.empty() ? "AutoLinker AI Preview" : title;
	ctx.content = content;
	ctx.primaryText = primaryText.empty() ? "\u786E\u5B9A" : primaryText;
	ctx.secondaryText = secondaryText;
	ctx.action = AIPreviewAction::Cancel;

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
		return AIPreviewAction::Cancel;
	}

	ApplyWindowIcon(hDialog);
	EnsureWindowTitle(hDialog, ctx.title);
	RunModalWindow(owner, hDialog);
	return ctx.action;
}

AIPreviewWebViewRunResult ShowAIPreviewDialogWebView(
	HWND owner,
	const std::string& title,
	const std::string& content,
	const std::string& primaryText,
	const std::string& secondaryText)
{
	AIPreviewWebViewRunResult result = {};
	if (!IsWebView2RuntimeAvailableWithTag("AI Preview")) {
		OutputStringToELog("[AI Preview][WebView2] runtime unavailable, fallback to native dialog");
		result.fallbackRequested = true;
		return result;
	}

	ComCtl6ActivationScope themeScope;
	INITCOMMONCONTROLSEX icex = {};
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_STANDARD_CLASSES | ICC_WIN95_CLASSES | ICC_LINK_CLASS;
	InitCommonControlsEx(&icex);

	WNDCLASSEXA wc = {};
	wc.cbSize = sizeof(wc);
	wc.lpfnWndProc = AIPreviewWebViewDialogProc;
	wc.hInstance = GetModuleHandleA(nullptr);
	wc.lpszClassName = "AutoLinkerAIPreviewWebViewDialogWindow";
	wc.hIcon = GetAppIconLarge();
	wc.hIconSm = GetAppIconSmall();
	wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_BTNFACE + 1);
	RegisterClassExA(&wc);

	AIPreviewWebViewDialogContext ctx = {};
	ctx.title = title.empty() ? "AutoLinker AI Preview" : title;
	ctx.content = content;
	ctx.primaryText = primaryText.empty() ? "\u786E\u5B9A" : primaryText;
	ctx.secondaryText = secondaryText;
	ctx.action = AIPreviewAction::Cancel;

	HWND hDialog = CreateWindowExA(
		WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
		wc.lpszClassName,
		ctx.title.c_str(),
		WS_CAPTION | WS_SYSMENU | WS_VISIBLE | WS_THICKFRAME,
		CW_USEDEFAULT, CW_USEDEFAULT, 960, 700,
		owner,
		nullptr,
		wc.hInstance,
		&ctx);

	if (hDialog == nullptr) {
		OutputStringToELog("[AI Preview][WebView2] CreateWindowExA failed, fallback to native dialog");
		result.fallbackRequested = true;
		return result;
	}

	ApplyWindowIcon(hDialog);
	EnsureWindowTitle(hDialog, ctx.title);
	RunModalWindow(owner, hDialog);
	result.action = ctx.action;
	result.fallbackRequested = ctx.fallbackRequested;
	return result;
}

AIPreviewAction ShowAIPreviewDialogEx(
	HWND owner,
	const std::string& title,
	const std::string& content,
	const std::string& primaryText,
	const std::string& secondaryText)
{
	const AIPreviewWebViewRunResult webViewResult = ShowAIPreviewDialogWebView(
		owner,
		title,
		content,
		primaryText,
		secondaryText);
	if (webViewResult.fallbackRequested) {
		OutputStringToELog("[AI Preview][WebView2] fallback to native dialog");
		return ShowAIPreviewDialogNative(owner, title, content, primaryText, secondaryText);
	}
	return webViewResult.action;
}

bool ShowAIPreviewDialog(HWND owner, const std::string& title, const std::string& content, const std::string& confirmText)
{
	return ShowAIPreviewDialogEx(owner, title, content, confirmText, "") == AIPreviewAction::PrimaryConfirm;
}

bool ShowAITextInputDialog(HWND owner, const std::string& title, const std::string& hint, std::string& ioText)
{
	ComCtl6ActivationScope themeScope;
	INITCOMMONCONTROLSEX icex = {};
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_STANDARD_CLASSES | ICC_WIN95_CLASSES | ICC_LINK_CLASS;
	InitCommonControlsEx(&icex);

	WNDCLASSEXA wc = {};
	wc.cbSize = sizeof(wc);
	wc.lpfnWndProc = AIInputDialogProc;
	wc.hInstance = GetModuleHandleA(nullptr);
	wc.lpszClassName = "AutoLinkerAIInputDialogWindow";
	wc.hIcon = GetAppIconLarge();
	wc.hIconSm = GetAppIconSmall();
	wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_BTNFACE + 1);
	RegisterClassExA(&wc);

	AIInputDialogContext ctx = {};
	ctx.title = title.empty() ? "AutoLinker AI Input" : title;
	ctx.hint = hint;
	ctx.text = ioText;

	HWND hDialog = CreateWindowExA(
		WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
		wc.lpszClassName,
		ctx.title.c_str(),
		WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
		CW_USEDEFAULT, CW_USEDEFAULT, 650, 330,
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
		ioText = ctx.text;
	}
	return ctx.accepted;
}
