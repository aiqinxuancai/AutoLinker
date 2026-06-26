#include "EcSwitchConfigDialog.h"

#include "Global.h"
#include "resource.h"
#include "ResourceTextLoader.h"

#include <Windows.h>
#include <CommCtrl.h>
#include <cctype>
#include <cstring>
#include <filesystem>
#include <format>
#include <map>
#include <set>
#include <string>
#include <vector>
#include <wrl.h>

#include "..\\thirdparty\\json.hpp"
#include "..\\thirdparty\\WebView2.h"

#pragma comment(lib, "comctl32.lib")

namespace {

constexpr UINT_PTR kEcSwitchConfigWebViewInitTimerId = 0xAC04;
constexpr UINT kEcSwitchConfigWebViewInitTimeoutMs = 12000;

struct EcSwitchRule {
	std::string dynamicName; // 调试时使用的动态版 ec 文件名
	std::string staticName;  // 编译时使用的静态版 ec 文件名
};

struct EcSwitchConfigWebViewDialogContext {
	bool webViewReady = false;
	HWND hHost = nullptr;
	HWND hLoading = nullptr;
	std::vector<EcSwitchRule> rules;
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

std::string ToLowerAsciiLocal(std::string text)
{
	std::transform(text.begin(), text.end(), text.begin(), [](unsigned char ch) {
		return static_cast<char>(std::tolower(ch));
	});
	return text;
}

bool EndsWithEcExtension(const std::string& name)
{
	if (name.size() < 3) {
		return false;
	}
	const std::string ext = ToLowerAsciiLocal(name.substr(name.size() - 3));
	return ext == ".ec";
}

std::string NormalizeEcFileName(const std::string& rawName)
{
	std::string name = TrimAsciiLocal(rawName);
	if (!name.empty() && !EndsWithEcExtension(name)) {
		name += ".ec";
	}
	return name;
}

bool IsValidEcFileName(const std::string& name, std::string& errorMessage)
{
	if (name.empty()) {
		errorMessage = "模块文件名不能为空。";
		return false;
	}
	if (!EndsWithEcExtension(name)) {
		errorMessage = "模块文件名必须以 .ec 结尾。";
		return false;
	}
	for (char ch : name) {
		switch (ch) {
		case '\\': case '/': case ':': case '*':
		case '?': case '"': case '<': case '>': case '|':
		case '=': case '\r': case '\n':
			errorMessage = "模块文件名不能包含路径、等号或 Windows 文件名非法字符。";
			return false;
		default:
			if (static_cast<unsigned char>(ch) < 32) {
				errorMessage = "模块文件名不能包含控制字符。";
				return false;
			}
			break;
		}
	}
	return true;
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

std::vector<EcSwitchRule> LoadEcSwitchRulesFromManager()
{
	std::vector<EcSwitchRule> rules;
	for (const auto& [dynamicName, staticName] : g_modelManager.getAllValues()) {
		if (dynamicName.empty() || staticName.empty()) {
			continue;
		}
		rules.push_back(EcSwitchRule{ dynamicName, staticName });
	}
	std::sort(rules.begin(), rules.end(), [](const EcSwitchRule& left, const EcSwitchRule& right) {
		return ToLowerAsciiLocal(left.dynamicName) < ToLowerAsciiLocal(right.dynamicName);
	});
	return rules;
}

std::string BuildEcSwitchWebViewPayloadJson(const std::vector<EcSwitchRule>& rules)
{
	nlohmann::json payload;
	payload["configPath"] = WideToUtf8Local(g_modelManager.getConfigFilePath().wstring());
	payload["rules"] = nlohmann::json::array();
	for (const auto& rule : rules) {
		nlohmann::json item;
		item["dynamicName"] = LocalToUtf8TextLocal(rule.dynamicName);
		item["staticName"] = LocalToUtf8TextLocal(rule.staticName);
		payload["rules"].push_back(std::move(item));
	}
	return payload.dump();
}

std::string BuildEcSwitchWebViewShellHtml()
{
	std::string html = LoadUtf8HtmlResourceText(IDR_HTML_EC_SWITCH_CONFIG_DIALOG);
	if (!html.empty()) {
		return html;
	}
	return "<!doctype html><html><head><meta charset=\"utf-8\"></head><body>EC switch config shell resource missing.</body></html>";
}

void ExecuteEcSwitchWebViewScript(EcSwitchConfigWebViewDialogContext* ctx, const std::wstring& script)
{
	if (ctx == nullptr || !ctx->webViewReady || ctx->webView == nullptr || script.empty()) {
		return;
	}
	ctx->webView->ExecuteScript(script.c_str(), nullptr);
}

void ApplyEcSwitchWebViewData(EcSwitchConfigWebViewDialogContext* ctx)
{
	if (ctx == nullptr) {
		return;
	}
	const std::wstring payloadWide = Utf8ToWideLocal(BuildEcSwitchWebViewPayloadJson(ctx->rules));
	std::wstring script = L"window.autolinkerApplyEcSwitchConfig(JSON.parse('";
	script += EscapeJsSingleQuotedWideLocal(payloadWide);
	script += L"'));";
	ExecuteEcSwitchWebViewScript(ctx, script);
}

void NotifyEcSwitchSaveResult(EcSwitchConfigWebViewDialogContext* ctx, bool ok, const std::string& message)
{
	nlohmann::json result;
	result["ok"] = ok;
	result["message"] = LocalToUtf8TextLocal(message);
	std::wstring script = L"window.autolinkerEcSwitchSaveResult(JSON.parse('";
	script += EscapeJsSingleQuotedWideLocal(Utf8ToWideLocal(result.dump()));
	script += L"'));";
	ExecuteEcSwitchWebViewScript(ctx, script);
}

bool BuildRulesFromWebPayload(const nlohmann::json& data, std::map<std::string, std::string>& outValues, std::string& errorMessage)
{
	outValues.clear();
	std::set<std::string> dynamicNames;
	std::set<std::string> staticNames;

	if (!data.contains("rules") || !data["rules"].is_array()) {
		return true;
	}

	for (const auto& item : data["rules"]) {
		if (!item.is_object()) {
			continue;
		}
		const std::string dynamicName = NormalizeEcFileName(Utf8ToLocalTextLocal(item.value("dynamicName", std::string())));
		const std::string staticName = NormalizeEcFileName(Utf8ToLocalTextLocal(item.value("staticName", std::string())));

		if (dynamicName.empty() && staticName.empty()) {
			continue;
		}
		if (!IsValidEcFileName(dynamicName, errorMessage)) {
			return false;
		}
		if (!IsValidEcFileName(staticName, errorMessage)) {
			return false;
		}
		if (_stricmp(dynamicName.c_str(), staticName.c_str()) == 0) {
			errorMessage = "动态版和静态版不能使用同一个 ec 文件名：" + dynamicName;
			return false;
		}

		const std::string dynamicKey = ToLowerAsciiLocal(dynamicName);
		if (!dynamicNames.insert(dynamicKey).second) {
			errorMessage = "存在重复的动态调试版模块：" + dynamicName;
			return false;
		}
		const std::string staticKey = ToLowerAsciiLocal(staticName);
		if (!staticNames.insert(staticKey).second) {
			errorMessage = "存在重复的静态编译版模块：" + staticName;
			return false;
		}
		outValues[dynamicName] = staticName;
	}
	return true;
}

void HandleEcSwitchWebViewSave(EcSwitchConfigWebViewDialogContext* ctx, const nlohmann::json& data)
{
	if (ctx == nullptr) {
		return;
	}

	std::map<std::string, std::string> nextValues;
	std::string errorMessage;
	if (!BuildRulesFromWebPayload(data, nextValues, errorMessage)) {
		NotifyEcSwitchSaveResult(ctx, false, errorMessage.empty() ? "配置内容无效。" : errorMessage);
		return;
	}

	if (!g_modelManager.replaceAllValues(nextValues, &errorMessage)) {
		NotifyEcSwitchSaveResult(ctx, false, errorMessage.empty() ? "保存 ModelManager.ini 失败。" : errorMessage);
		return;
	}

	ctx->rules = LoadEcSwitchRulesFromManager();
	NotifyEcSwitchSaveResult(ctx, true, std::format("已保存 {} 条 EC 模块自动切换规则。", ctx->rules.size()));
}

void LayoutEcSwitchConfigWebViewDialog(HWND hWnd, EcSwitchConfigWebViewDialogContext* ctx)
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

HRESULT OnEcSwitchConfigControllerCreated(HWND hWnd, HRESULT controllerResult, ICoreWebView2Controller* controller)
{
	auto* readyCtx = reinterpret_cast<EcSwitchConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
	if (readyCtx == nullptr || !IsWindow(hWnd)) {
		return S_OK;
	}
	if (FAILED(controllerResult) || controller == nullptr) {
		OutputStringToELog(std::format("[EC Switch Config][WebView2] create controller failed hr=0x{:08X}", static_cast<unsigned int>(controllerResult)));
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
				auto* messageCtx = reinterpret_cast<EcSwitchConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
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
						HandleEcSwitchWebViewSave(messageCtx, payload["data"]);
					}
					else if (action == "cancel") {
						DestroyWindow(hWnd);
					}
				}
				catch (...) {
					NotifyEcSwitchSaveResult(messageCtx, false, "页面消息解析失败。");
				}
				return S_OK;
			}).Get(),
		nullptr);

	readyCtx->webView->add_NavigationCompleted(
		Microsoft::WRL::Callback<ICoreWebView2NavigationCompletedEventHandler>(
			[hWnd](ICoreWebView2*, ICoreWebView2NavigationCompletedEventArgs* args) -> HRESULT {
				auto* navCtx = reinterpret_cast<EcSwitchConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
				if (navCtx == nullptr || args == nullptr || !IsWindow(hWnd)) {
					return S_OK;
				}
				BOOL isSuccess = FALSE;
				args->get_IsSuccess(&isSuccess);
				if (isSuccess == TRUE) {
					navCtx->webViewReady = true;
					KillTimer(hWnd, kEcSwitchConfigWebViewInitTimerId);
					if (navCtx->hLoading != nullptr) {
						ShowWindow(navCtx->hLoading, SW_HIDE);
					}
					LayoutEcSwitchConfigWebViewDialog(hWnd, navCtx);
					ApplyEcSwitchWebViewData(navCtx);
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

	LayoutEcSwitchConfigWebViewDialog(hWnd, readyCtx);
	const std::wstring shellHtml = Utf8ToWideLocal(BuildEcSwitchWebViewShellHtml());
	if (shellHtml.empty()) {
		DestroyWindow(hWnd);
		return S_OK;
	}
	readyCtx->webView->NavigateToString(shellHtml.c_str());
	return S_OK;
}

void StartEcSwitchConfigWebView(HWND hWnd, EcSwitchConfigWebViewDialogContext* ctx)
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
				auto* innerCtx = reinterpret_cast<EcSwitchConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
				if (innerCtx == nullptr || !IsWindow(hWnd)) {
					return S_OK;
				}
				if (FAILED(envResult) || environment == nullptr) {
					OutputStringToELog(std::format("[EC Switch Config][WebView2] create environment failed hr=0x{:08X}", static_cast<unsigned int>(envResult)));
					DestroyWindow(hWnd);
					return S_OK;
				}
				innerCtx->webViewEnvironment = environment;
				return environment->CreateCoreWebView2Controller(
					innerCtx->hHost,
					Microsoft::WRL::Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
						[hWnd](HRESULT controllerResult, ICoreWebView2Controller* controller) -> HRESULT {
							return OnEcSwitchConfigControllerCreated(hWnd, controllerResult, controller);
						}).Get());
			}).Get());

	if (FAILED(hr)) {
		OutputStringToELog(std::format("[EC Switch Config][WebView2] bootstrap failed hr=0x{:08X}", static_cast<unsigned int>(hr)));
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

LRESULT CALLBACK EcSwitchConfigWebViewDialogProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	auto* ctx = reinterpret_cast<EcSwitchConfigWebViewDialogContext*>(GetWindowLongPtrA(hWnd, GWLP_USERDATA));
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
			0, L"STATIC", L"正在初始化 WebView2 EC 模块自动切换设置页...",
			WS_CHILD | WS_VISIBLE,
			0, 0, 0, 0, hWnd, nullptr, nullptr, nullptr);
		SetDefaultFontLocal(ctx->hLoading);
		LayoutEcSwitchConfigWebViewDialog(hWnd, ctx);
		SetTimer(hWnd, kEcSwitchConfigWebViewInitTimerId, kEcSwitchConfigWebViewInitTimeoutMs, nullptr);
		StartEcSwitchConfigWebView(hWnd, ctx);
		return 0;
	}

	case WM_GETMINMAXINFO: {
		auto* mmi = reinterpret_cast<MINMAXINFO*>(lParam);
		if (mmi != nullptr) {
			mmi->ptMinTrackSize.x = (std::max)(mmi->ptMinTrackSize.x, 760L);
			mmi->ptMinTrackSize.y = (std::max)(mmi->ptMinTrackSize.y, 560L);
		}
		return 0;
	}

	case WM_SIZE:
		if (ctx != nullptr) {
			LayoutEcSwitchConfigWebViewDialog(hWnd, ctx);
		}
		return 0;

	case WM_TIMER:
		if (ctx != nullptr && wParam == kEcSwitchConfigWebViewInitTimerId) {
			if (!ctx->webViewReady) {
				OutputStringToELog("[EC Switch Config][WebView2] initialization timed out");
				DestroyWindow(hWnd);
				return 0;
			}
			KillTimer(hWnd, kEcSwitchConfigWebViewInitTimerId);
			return 0;
		}
		break;

	case WM_CLOSE:
		DestroyWindow(hWnd);
		return 0;

	case WM_DESTROY:
		KillTimer(hWnd, kEcSwitchConfigWebViewInitTimerId);
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

void ShowEcSwitchConfigDialog(HWND owner)
{
	if (!IsWebView2RuntimeAvailableLocal()) {
		MessageBoxW(owner,
			L"未检测到 WebView2 运行时，无法打开 EC 模块自动切换设置页。\n请安装 Microsoft Edge WebView2 Runtime 后重试。",
			L"AutoLinker EC 模块自动切换设置",
			MB_OK | MB_ICONWARNING);
		return;
	}

	INITCOMMONCONTROLSEX icex = {};
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_STANDARD_CLASSES | ICC_WIN95_CLASSES | ICC_LINK_CLASS;
	InitCommonControlsEx(&icex);

	WNDCLASSEXA wc = {};
	wc.cbSize = sizeof(wc);
	wc.lpfnWndProc = EcSwitchConfigWebViewDialogProc;
	wc.hInstance = GetModuleHandleA(nullptr);
	wc.lpszClassName = "AutoLinkerEcSwitchConfigWebViewDialogWindow";
	wc.hIcon = GetAppIconLargeLocal();
	wc.hIconSm = GetAppIconSmallLocal();
	wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_BTNFACE + 1);
	RegisterClassExA(&wc);

	EcSwitchConfigWebViewDialogContext ctx = {};
	ctx.rules = LoadEcSwitchRulesFromManager();

	HWND hDialog = CreateWindowExA(
		WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
		wc.lpszClassName,
		"AutoLinker EC Switch Config",
		WS_CAPTION | WS_SYSMENU | WS_VISIBLE | WS_THICKFRAME,
		CW_USEDEFAULT, CW_USEDEFAULT, 900, 680,
		owner,
		nullptr,
		wc.hInstance,
		&ctx);

	if (hDialog == nullptr) {
		OutputStringToELog("[EC Switch Config][WebView2] CreateWindowExA failed");
		return;
	}

	ApplyWindowIconLocal(hDialog);
	SetWindowTextW(hDialog, L"AutoLinker EC 模块自动切换设置");
	RunModalWindowLocal(owner, hDialog);
}
