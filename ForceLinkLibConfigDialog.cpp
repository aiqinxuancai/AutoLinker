#include "ForceLinkLibConfigDialog.h"

#include "ForceLinkLibManager.h"
#include "Global.h"
#include "resource.h"
#include "ResourceTextLoader.h"

#include <Windows.h>
#include <CommCtrl.h>
#include <cctype>
#include <filesystem>
#include <format>
#include <string>
#include <utility>
#include <vector>
#include <wrl.h>

#include "..\\thirdparty\\json.hpp"
#include "..\\thirdparty\\WebView2.h"

#pragma comment(lib, "comctl32.lib")

namespace {

constexpr UINT_PTR kForceLinkLibConfigWebViewInitTimerId = 0xAC05;
constexpr UINT kForceLinkLibConfigWebViewInitTimeoutMs = 12000;

struct ForceLinkLibConfigWebViewDialogContext {
	bool webViewReady = false;
	HWND hHost = nullptr;
	HWND hLoading = nullptr;
	std::vector<ForceLinkLibRule> rules;
	bool appendForce = false;
	std::filesystem::path configPath;
	Microsoft::WRL::ComPtr<ICoreWebView2Environment> webViewEnvironment;
	Microsoft::WRL::ComPtr<ICoreWebView2Controller> webViewController;
	Microsoft::WRL::ComPtr<ICoreWebView2> webView;
};

HMODULE GetCurrentModuleHandleLocal()
{
	HMODULE module = nullptr;
	GetModuleHandleExW(
		GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
		reinterpret_cast<LPCWSTR>(&GetCurrentModuleHandleLocal),
		&module);
	return module;
}

HICON LoadAppIconHandleLocal(int cx, int cy)
{
	const HMODULE module = GetCurrentModuleHandleLocal();
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

HICON GetAppIconLargeLocal()
{
	static HICON s_largeIcon = LoadAppIconHandleLocal(GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON));
	return s_largeIcon;
}

HICON GetAppIconSmallLocal()
{
	static HICON s_smallIcon = LoadAppIconHandleLocal(GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON));
	return s_smallIcon;
}

void ApplyWindowIconLocal(HWND hWnd)
{
	if (hWnd == nullptr) {
		return;
	}
	SendMessageA(hWnd, WM_SETICON, ICON_BIG, reinterpret_cast<LPARAM>(GetAppIconLargeLocal()));
	SendMessageA(hWnd, WM_SETICON, ICON_SMALL, reinterpret_cast<LPARAM>(GetAppIconSmallLocal()));
}

HFONT GetDialogFontLocal()
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

void SetDefaultFontLocal(HWND hWnd)
{
	if (hWnd != nullptr) {
		SendMessageA(hWnd, WM_SETFONT, reinterpret_cast<WPARAM>(GetDialogFontLocal()), TRUE);
	}
}

std::wstring Utf8ToWideLocal(const std::string& text)
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

std::string WideToUtf8Local(const std::wstring& text)
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

std::string ConvertCodePageLocal(const std::string& text, UINT fromCodePage, UINT toCodePage, DWORD fromFlags = 0)
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

std::string Utf8ToLocalTextLocal(const std::string& text)
{
	return ConvertCodePageLocal(text, CP_UTF8, CP_ACP, MB_ERR_INVALID_CHARS);
}

std::string LocalToUtf8TextLocal(const std::string& text)
{
	return ConvertCodePageLocal(text, CP_ACP, CP_UTF8, 0);
}

std::wstring EscapeJsSingleQuotedWideLocal(const std::wstring& text)
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

std::string TrimAsciiLocal(const std::string& text)
{
	size_t begin = 0;
	size_t end = text.size();
	while (begin < end) {
		const unsigned char ch = static_cast<unsigned char>(text[begin]);
		if (ch != ' ' && ch != '\t' && ch != '\r' && ch != '\n') {
			break;
		}
		++begin;
	}
	while (end > begin) {
		const unsigned char ch = static_cast<unsigned char>(text[end - 1]);
		if (ch != ' ' && ch != '\t' && ch != '\r' && ch != '\n') {
			break;
		}
		--end;
	}
	return text.substr(begin, end - begin);
}

bool IsWebView2RuntimeAvailableLocal()
{
	LPWSTR version = nullptr;
	const HRESULT hr = GetAvailableCoreWebView2BrowserVersionString(nullptr, &version);
	if (version != nullptr) {
		CoTaskMemFree(version);
	}
	return SUCCEEDED(hr);
}

std::vector<ForceLinkLibRule> LoadForceLinkLibRulesFromDisk(std::filesystem::path& outConfigPath, bool& outAppendForce)
{
	ForceLinkLibManager manager;
	outConfigPath = manager.getConfigFilePath();
	outAppendForce = manager.getAppendForce();
	return manager.getRules();
}

std::string BuildForceLinkLibWebViewPayloadJson(const ForceLinkLibConfigWebViewDialogContext* ctx)
{
	nlohmann::json payload;
	payload["configPath"] = ctx != nullptr ? WideToUtf8Local(ctx->configPath.wstring()) : std::string();
	payload["appendForce"] = ctx != nullptr ? ctx->appendForce : false;
	payload["rules"] = nlohmann::json::array();
	if (ctx == nullptr) {
		return payload.dump();
	}
	for (const auto& rule : ctx->rules) {
		nlohmann::json item;
		item["enabled"] = rule.enabled;
		item["linkerName"] = LocalToUtf8TextLocal(rule.linkerName);
		item["libPath"] = LocalToUtf8TextLocal(rule.libPath);
		payload["rules"].push_back(std::move(item));
	}
	return payload.dump();
}

std::string BuildForceLinkLibWebViewShellHtml()
{
	std::string html = LoadUtf8HtmlResourceText(IDR_HTML_FORCE_LINK_LIB_CONFIG_DIALOG);
	if (!html.empty()) {
		return html;
	}
	return "<!doctype html><html><head><meta charset=\"utf-8\"></head><body>Force link lib config shell resource missing.</body></html>";
}

void ExecuteForceLinkLibWebViewScript(ForceLinkLibConfigWebViewDialogContext* ctx, const std::wstring& script)
{
	if (ctx == nullptr || !ctx->webViewReady || ctx->webView == nullptr || script.empty()) {
		return;
	}
	ctx->webView->ExecuteScript(script.c_str(), nullptr);
}

void ApplyForceLinkLibWebViewData(ForceLinkLibConfigWebViewDialogContext* ctx)
{
	if (ctx == nullptr) {
		return;
	}
	const std::wstring payloadWide = Utf8ToWideLocal(BuildForceLinkLibWebViewPayloadJson(ctx));
	std::wstring script = L"window.autolinkerApplyForceLinkLibConfig(JSON.parse('";
	script += EscapeJsSingleQuotedWideLocal(payloadWide);
	script += L"'));";
	ExecuteForceLinkLibWebViewScript(ctx, script);
}

void NotifyForceLinkLibSaveResult(ForceLinkLibConfigWebViewDialogContext* ctx, bool ok, const std::string& message)
{
	nlohmann::json result;
	result["ok"] = ok;
	result["message"] = LocalToUtf8TextLocal(message);
	std::wstring script = L"window.autolinkerForceLinkLibSaveResult(JSON.parse('";
	script += EscapeJsSingleQuotedWideLocal(Utf8ToWideLocal(result.dump()));
	script += L"'));";
	ExecuteForceLinkLibWebViewScript(ctx, script);
}

bool BuildRulesFromWebPayload(const nlohmann::json& data, std::vector<ForceLinkLibRule>& outRules, std::string& errorMessage)
{
	outRules.clear();
	if (!data.contains("rules") || !data["rules"].is_array()) {
		return true;
	}

	for (const auto& item : data["rules"]) {
		if (!item.is_object()) {
			continue;
		}

		ForceLinkLibRule rule;
		rule.enabled = item.value("enabled", true);
		rule.linkerName = TrimAsciiLocal(Utf8ToLocalTextLocal(item.value("linkerName", std::string())));
		rule.libPath = TrimAsciiLocal(Utf8ToLocalTextLocal(item.value("libPath", std::string())));
		if (rule.libPath.empty()) {
			continue;
		}
		if (rule.linkerName.find('=') != std::string::npos ||
			rule.linkerName.find('\r') != std::string::npos ||
			rule.linkerName.find('\n') != std::string::npos) {
			errorMessage = "链接器匹配文本不能包含等号或换行。";
			return false;
		}
		if (rule.libPath.find('\r') != std::string::npos ||
			rule.libPath.find('\n') != std::string::npos) {
			errorMessage = "Lib 路径不能包含换行。";
			return false;
		}
		outRules.push_back(std::move(rule));
	}
	return true;
}

void HandleForceLinkLibWebViewSave(ForceLinkLibConfigWebViewDialogContext* ctx, const nlohmann::json& data)
{
	if (ctx == nullptr) {
		return;
	}

	std::vector<ForceLinkLibRule> nextRules;
	std::string errorMessage;
	const bool nextAppendForce = data.value("appendForce", false);
	if (!BuildRulesFromWebPayload(data, nextRules, errorMessage)) {
		NotifyForceLinkLibSaveResult(ctx, false, errorMessage.empty() ? "配置内容无效。" : errorMessage);
		return;
	}

	ForceLinkLibManager manager;
	if (!manager.replaceRules(nextRules, nextAppendForce, &errorMessage)) {
		NotifyForceLinkLibSaveResult(ctx, false, errorMessage.empty() ? "保存 ForceLinkLib.ini 失败。" : errorMessage);
		return;
	}

	ctx->configPath = manager.getConfigFilePath();
	ctx->appendForce = manager.getAppendForce();
	ctx->rules = manager.getRules();
	NotifyForceLinkLibSaveResult(ctx, true, std::format("已保存 {} 条核心库函数重写强制链接规则。", ctx->rules.size()));
}

void LayoutForceLinkLibConfigWebViewDialog(HWND hWnd, ForceLinkLibConfigWebViewDialogContext* ctx)
{
	if (hWnd == nullptr || ctx == nullptr) {
		return;
	}
	RECT rc = {};
	GetClientRect(hWnd, &rc);
	const int hostWidth = static_cast<int>((std::max)(0L, rc.right));
	const int hostHeight = static_cast<int>((std::max)(0L, rc.bottom));
	if (ctx->hHost != nullptr) {
		MoveWindow(ctx->hHost, 0, 0, hostWidth, hostHeight, TRUE);
	}
	if (ctx->hLoading != nullptr) {
		MoveWindow(ctx->hLoading, 16, 14, (std::max)(0, hostWidth - 32), 24, TRUE);
	}
	if (ctx->webViewController != nullptr) {
		RECT bounds = { 0, 0, static_cast<LONG>(hostWidth), static_cast<LONG>(hostHeight) };
		ctx->webViewController->put_Bounds(bounds);
	}
}

HRESULT OnForceLinkLibConfigControllerCreated(HWND hWnd, HRESULT controllerResult, ICoreWebView2Controller* controller)
{
	auto* readyCtx = reinterpret_cast<ForceLinkLibConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
	if (readyCtx == nullptr || !IsWindow(hWnd)) {
		return S_OK;
	}
	if (FAILED(controllerResult) || controller == nullptr) {
		OutputStringToELog(std::format("[ForceLinkLib Config][WebView2] create controller failed hr=0x{:08X}", static_cast<unsigned int>(controllerResult)));
		DestroyWindow(hWnd);
		return S_OK;
	}

	readyCtx->webViewController = controller;
	readyCtx->webViewController->get_CoreWebView2(&readyCtx->webView);
	if (readyCtx->webView == nullptr) {
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
				auto* messageCtx = reinterpret_cast<ForceLinkLibConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
				if (messageCtx == nullptr || args == nullptr || !IsWindow(hWnd)) {
					return S_OK;
				}
				LPWSTR rawMessage = nullptr;
				if (FAILED(args->TryGetWebMessageAsString(&rawMessage)) || rawMessage == nullptr) {
					return S_OK;
				}
				const std::string utf8Message = WideToUtf8Local(rawMessage);
				CoTaskMemFree(rawMessage);
				try {
					const nlohmann::json payload = nlohmann::json::parse(utf8Message);
					const std::string action = payload.value("action", "");
					if (action == "save" && payload.contains("data") && payload["data"].is_object()) {
						HandleForceLinkLibWebViewSave(messageCtx, payload["data"]);
					}
					else if (action == "cancel") {
						DestroyWindow(hWnd);
					}
				}
				catch (...) {
					NotifyForceLinkLibSaveResult(messageCtx, false, "页面消息解析失败。");
				}
				return S_OK;
			}).Get(),
		nullptr);

	readyCtx->webView->add_NavigationCompleted(
		Microsoft::WRL::Callback<ICoreWebView2NavigationCompletedEventHandler>(
			[hWnd](ICoreWebView2*, ICoreWebView2NavigationCompletedEventArgs* args) -> HRESULT {
				auto* navCtx = reinterpret_cast<ForceLinkLibConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
				if (navCtx == nullptr || args == nullptr || !IsWindow(hWnd)) {
					return S_OK;
				}
				BOOL isSuccess = FALSE;
				args->get_IsSuccess(&isSuccess);
				if (isSuccess == TRUE) {
					navCtx->webViewReady = true;
					KillTimer(hWnd, kForceLinkLibConfigWebViewInitTimerId);
					if (navCtx->hLoading != nullptr) {
						ShowWindow(navCtx->hLoading, SW_HIDE);
					}
					LayoutForceLinkLibConfigWebViewDialog(hWnd, navCtx);
					ApplyForceLinkLibWebViewData(navCtx);
					return S_OK;
				}
				COREWEBVIEW2_WEB_ERROR_STATUS webErrorStatus = COREWEBVIEW2_WEB_ERROR_STATUS_UNKNOWN;
				args->get_WebErrorStatus(&webErrorStatus);
				if (webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_OPERATION_CANCELED ||
					webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_ABORTED ||
					webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_RESET) {
					return S_OK;
				}
				DestroyWindow(hWnd);
				return S_OK;
			}).Get(),
		nullptr);

	LayoutForceLinkLibConfigWebViewDialog(hWnd, readyCtx);
	const std::wstring shellHtml = Utf8ToWideLocal(BuildForceLinkLibWebViewShellHtml());
	if (shellHtml.empty()) {
		DestroyWindow(hWnd);
		return S_OK;
	}
	readyCtx->webView->NavigateToString(shellHtml.c_str());
	return S_OK;
}

void StartForceLinkLibConfigWebView(HWND hWnd, ForceLinkLibConfigWebViewDialogContext* ctx)
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
				auto* innerCtx = reinterpret_cast<ForceLinkLibConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
				if (innerCtx == nullptr || !IsWindow(hWnd)) {
					return S_OK;
				}
				if (FAILED(envResult) || environment == nullptr) {
					OutputStringToELog(std::format("[ForceLinkLib Config][WebView2] create environment failed hr=0x{:08X}", static_cast<unsigned int>(envResult)));
					DestroyWindow(hWnd);
					return S_OK;
				}
				innerCtx->webViewEnvironment = environment;
				return environment->CreateCoreWebView2Controller(
					innerCtx->hHost,
					Microsoft::WRL::Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
						[hWnd](HRESULT controllerResult, ICoreWebView2Controller* controller) -> HRESULT {
							return OnForceLinkLibConfigControllerCreated(hWnd, controllerResult, controller);
						}).Get());
			}).Get());

	if (FAILED(hr)) {
		OutputStringToELog(std::format("[ForceLinkLib Config][WebView2] bootstrap failed hr=0x{:08X}", static_cast<unsigned int>(hr)));
		DestroyWindow(hWnd);
	}
}

void CenterWindowOnOwnerOrScreenLocal(HWND hDialog, HWND owner)
{
	if (hDialog == nullptr) {
		return;
	}

	RECT dlgRc = {};
	if (GetWindowRect(hDialog, &dlgRc) == FALSE) {
		return;
	}
	const int dlgW = dlgRc.right - dlgRc.left;
	const int dlgH = dlgRc.bottom - dlgRc.top;

	RECT refRc = {};
	bool haveRef = false;
	if (owner != nullptr && IsWindow(owner) && !IsIconic(owner)) {
		haveRef = GetWindowRect(owner, &refRc) != FALSE;
	}
	if (!haveRef) {
		HMONITOR hMon = MonitorFromWindow(hDialog, MONITOR_DEFAULTTOPRIMARY);
		MONITORINFO mi = {};
		mi.cbSize = sizeof(mi);
		if (GetMonitorInfoA(hMon, &mi)) {
			refRc = mi.rcWork;
			haveRef = true;
		}
	}
	if (!haveRef) {
		return;
	}

	int x = refRc.left + ((refRc.right - refRc.left) - dlgW) / 2;
	int y = refRc.top + ((refRc.bottom - refRc.top) - dlgH) / 2;

	HMONITOR hMon = MonitorFromPoint(POINT{ x + dlgW / 2, y + dlgH / 2 }, MONITOR_DEFAULTTONEAREST);
	MONITORINFO mi = {};
	mi.cbSize = sizeof(mi);
	if (GetMonitorInfoA(hMon, &mi)) {
		const RECT& work = mi.rcWork;
		if (x + dlgW > work.right) {
			x = work.right - dlgW;
		}
		if (y + dlgH > work.bottom) {
			y = work.bottom - dlgH;
		}
		if (x < work.left) {
			x = work.left;
		}
		if (y < work.top) {
			y = work.top;
		}
	}

	SetWindowPos(hDialog, nullptr, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
}

bool RunModalWindowLocal(HWND owner, HWND hDialog)
{
	if (hDialog == nullptr) {
		return false;
	}

	if (owner != nullptr && IsWindow(owner)) {
		EnableWindow(owner, FALSE);
	}

	CenterWindowOnOwnerOrScreenLocal(hDialog, owner);
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

LRESULT CALLBACK ForceLinkLibConfigWebViewDialogProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	auto* ctx = reinterpret_cast<ForceLinkLibConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
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
			0, "STATIC", "",
			WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
			0, 0, 0, 0, hWnd, nullptr, nullptr, nullptr);
		ctx->hLoading = CreateWindowExW(
			0, L"STATIC", L"正在初始化 WebView2 核心库函数重写设置页...",
			WS_CHILD | WS_VISIBLE,
			0, 0, 0, 0, hWnd, nullptr, nullptr, nullptr);
		SetDefaultFontLocal(ctx->hLoading);
		LayoutForceLinkLibConfigWebViewDialog(hWnd, ctx);
		SetTimer(hWnd, kForceLinkLibConfigWebViewInitTimerId, kForceLinkLibConfigWebViewInitTimeoutMs, nullptr);
		StartForceLinkLibConfigWebView(hWnd, ctx);
		return 0;
	}

	case WM_GETMINMAXINFO: {
		auto* mmi = reinterpret_cast<MINMAXINFO*>(lParam);
		if (mmi != nullptr) {
			mmi->ptMinTrackSize.x = (std::max)(mmi->ptMinTrackSize.x, 800L);
			mmi->ptMinTrackSize.y = (std::max)(mmi->ptMinTrackSize.y, 560L);
		}
		return 0;
	}

	case WM_SIZE:
		if (ctx != nullptr) {
			LayoutForceLinkLibConfigWebViewDialog(hWnd, ctx);
		}
		return 0;

	case WM_TIMER:
		if (ctx != nullptr && wParam == kForceLinkLibConfigWebViewInitTimerId) {
			if (!ctx->webViewReady) {
				OutputStringToELog("[ForceLinkLib Config][WebView2] initialization timed out");
				DestroyWindow(hWnd);
				return 0;
			}
			KillTimer(hWnd, kForceLinkLibConfigWebViewInitTimerId);
			return 0;
		}
		break;

	case WM_CLOSE:
		DestroyWindow(hWnd);
		return 0;

	case WM_DESTROY:
		KillTimer(hWnd, kForceLinkLibConfigWebViewInitTimerId);
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

} // namespace

void ShowForceLinkLibConfigDialog(HWND owner)
{
	if (!IsWebView2RuntimeAvailableLocal()) {
		MessageBoxW(owner,
			L"未检测到 WebView2 运行时，无法打开核心库函数重写设置页。\n请安装 Microsoft Edge WebView2 Runtime 后重试。",
			L"AutoLinker 核心库函数重写设置",
			MB_OK | MB_ICONWARNING);
		return;
	}

	INITCOMMONCONTROLSEX icex = {};
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_STANDARD_CLASSES | ICC_WIN95_CLASSES | ICC_LINK_CLASS;
	InitCommonControlsEx(&icex);

	WNDCLASSEXA wc = {};
	wc.cbSize = sizeof(wc);
	wc.lpfnWndProc = ForceLinkLibConfigWebViewDialogProc;
	wc.hInstance = GetModuleHandleA(nullptr);
	wc.lpszClassName = "AutoLinkerForceLinkLibConfigWebViewDialogWindow";
	wc.hIcon = GetAppIconLargeLocal();
	wc.hIconSm = GetAppIconSmallLocal();
	wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_BTNFACE + 1);
	RegisterClassExA(&wc);

	ForceLinkLibConfigWebViewDialogContext ctx = {};
	ctx.rules = LoadForceLinkLibRulesFromDisk(ctx.configPath, ctx.appendForce);

	HWND hDialog = CreateWindowExA(
		WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
		wc.lpszClassName,
		"AutoLinker Force Link Lib Config",
		WS_CAPTION | WS_SYSMENU | WS_VISIBLE | WS_THICKFRAME,
		CW_USEDEFAULT, CW_USEDEFAULT, 960, 680,
		owner,
		nullptr,
		wc.hInstance,
		&ctx);

	if (hDialog == nullptr) {
		OutputStringToELog("[ForceLinkLib Config][WebView2] CreateWindowExA failed");
		return;
	}

	ApplyWindowIconLocal(hDialog);
	SetWindowTextW(hDialog, L"AutoLinker 核心库函数重写设置");
	RunModalWindowLocal(owner, hDialog);
}
