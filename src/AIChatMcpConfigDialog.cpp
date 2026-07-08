#include "AIChatMcpConfigDialog.h"

#include "resource.h"

#include <Windows.h>
#include <algorithm>
#include <filesystem>
#include <fstream>
#include <format>
#include <string>
#include <wrl.h>

#include "..\\thirdparty\\json.hpp"
#include "..\\thirdparty\\WebView2.h"

#include "AIChatMcpConfig.h"
#include "Global.h"
#include "ResourceTextLoader.h"

namespace {

constexpr int kEditId = 1101;
constexpr int kSaveId = 1102;
constexpr int kCancelId = 1103;
constexpr int kTemplateId = 1104;
constexpr UINT_PTR kMcpWebViewInitTimerId = 0xAD01;
constexpr UINT kMcpWebViewInitTimeoutMs = 12000;

struct McpConfigDialogContext {
	HWND hEdit = nullptr;
	HWND hSave = nullptr;
	HWND hCancel = nullptr;
	HWND hTemplate = nullptr;
	HWND owner = nullptr;
	bool done = false;
	bool saved = false;
};

struct McpWebViewDialogContext {
	HWND hHost = nullptr;
	HWND hLoading = nullptr;
	bool done = false;
	bool saved = false;
	bool fallbackRequested = false;
	bool webViewReady = false;
	Microsoft::WRL::ComPtr<ICoreWebView2Environment> webViewEnvironment;
	Microsoft::WRL::ComPtr<ICoreWebView2Controller> webViewController;
	Microsoft::WRL::ComPtr<ICoreWebView2> webView;
};

struct McpWebViewRunResult {
	bool saved = false;
	bool fallbackRequested = false;
};

std::string ReadCurrentConfigText()
{
	const std::filesystem::path path = AIChatMcpConfigStore::GetConfigPath();
	if (!std::filesystem::exists(path)) {
		return AIChatMcpConfigStore::BuildDefaultConfigJson();
	}
	std::ifstream input(path, std::ios::binary);
	if (!input.is_open()) {
		return AIChatMcpConfigStore::BuildDefaultConfigJson();
	}
	std::string text{ std::istreambuf_iterator<char>(input), std::istreambuf_iterator<char>() };
	return text.empty() ? AIChatMcpConfigStore::BuildDefaultConfigJson() : text;
}

std::wstring Utf8ToWide(const std::string& text)
{
	if (text.empty()) {
		return std::wstring();
	}
	const int size = MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (size <= 0) {
		return std::wstring(text.begin(), text.end());
	}
	std::wstring wide(static_cast<size_t>(size), L'\0');
	if (MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), wide.data(), size) <= 0) {
		return std::wstring(text.begin(), text.end());
	}
	return wide;
}

std::string NarrowAsciiFallback(const std::wstring& text)
{
	std::string out;
	out.reserve(text.size());
	for (wchar_t ch : text) {
		out.push_back(ch >= 0 && ch <= 0x7f ? static_cast<char>(ch) : '?');
	}
	return out;
}

std::string WideToUtf8(const std::wstring& text)
{
	if (text.empty()) {
		return std::string();
	}
	const int size = WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
	if (size <= 0) {
		return NarrowAsciiFallback(text);
	}
	std::string utf8(static_cast<size_t>(size), '\0');
	if (WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), utf8.data(), size, nullptr, nullptr) <= 0) {
		return NarrowAsciiFallback(text);
	}
	return utf8;
}

std::string WideToUtf8(const wchar_t* text)
{
	return text == nullptr ? std::string() : WideToUtf8(std::wstring(text));
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

std::string GetEditTextUtf8(HWND edit)
{
	const int len = GetWindowTextLengthW(edit);
	if (len <= 0) {
		return std::string();
	}
	std::wstring text(static_cast<size_t>(len) + 1, L'\0');
	const int copied = GetWindowTextW(edit, text.data(), len + 1);
	if (copied <= 0) {
		return std::string();
	}
	text.resize(static_cast<size_t>(copied));
	return WideToUtf8(text);
}

void SetDefaultFont(HWND hWnd)
{
	HFONT font = reinterpret_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
	if (font != nullptr) {
		SendMessageW(hWnd, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
	}
}

bool IsMcpWebViewRuntimeAvailable()
{
	LPWSTR version = nullptr;
	const HRESULT hr = GetAvailableCoreWebView2BrowserVersionString(nullptr, &version);
	const bool available = SUCCEEDED(hr);
	if (version != nullptr) {
		CoTaskMemFree(version);
	}
	return available;
}

std::string BuildMcpWebViewShellHtml()
{
	std::string html = LoadUtf8HtmlResourceText(IDR_HTML_AI_CHAT_MCP_CONFIG_DIALOG);
	if (!html.empty()) {
		return html;
	}
	return "<!doctype html><html><head><meta charset=\"utf-8\"></head><body>MCP config shell resource missing.</body></html>";
}

std::string BuildMcpWebViewConfigJson()
{
	AIChatMcpConfig config;
	std::string error;
	if (!AIChatMcpConfigStore::ParseConfigJson(ReadCurrentConfigText(), config, error)) {
		OutputStringToELog(std::format("[AI Chat MCP][WebView2] config parse failed, using default: {}", error));
		if (!AIChatMcpConfigStore::ParseConfigJson(AIChatMcpConfigStore::BuildDefaultConfigJson(), config, error)) {
			config = {};
			config.version = 1;
		}
	}
	return AIChatMcpConfigStore::SerializeConfigJson(config, false);
}

void ExecuteMcpWebViewScript(McpWebViewDialogContext* ctx, const std::wstring& script)
{
	if (ctx == nullptr || !ctx->webViewReady || ctx->webView == nullptr || script.empty()) {
		return;
	}
	ctx->webView->ExecuteScript(script.c_str(), nullptr);
}

void ApplyMcpWebViewConfig(McpWebViewDialogContext* ctx)
{
	if (ctx == nullptr || !ctx->webViewReady) {
		return;
	}

	const std::wstring configJsonWide = Utf8ToWide(BuildMcpWebViewConfigJson());
	if (configJsonWide.empty()) {
		OutputStringToELog("[AI Chat MCP][WebView2] config json conversion failed");
		return;
	}

	std::wstring script = L"window.autolinkerApplyMcpConfig(JSON.parse('";
	script += EscapeJsSingleQuotedWide(configJsonWide);
	script += L"'));";
	ExecuteMcpWebViewScript(ctx, script);
}

void LayoutMcpWebViewDialog(HWND hWnd, McpWebViewDialogContext* ctx)
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
		MoveWindow(ctx->hLoading, 16, 14, (std::max)(120, hostWidth - 32), 24, TRUE);
	}
	if (ctx->webViewController != nullptr) {
		RECT bounds = {};
		bounds.right = static_cast<LONG>(hostWidth);
		bounds.bottom = static_cast<LONG>(hostHeight);
		ctx->webViewController->put_Bounds(bounds);
	}
}

bool TrySaveWebConfig(HWND hWnd, McpWebViewDialogContext* ctx, const nlohmann::json& data)
{
	if (ctx == nullptr || !data.is_object()) {
		return false;
	}

	std::string error;
	const std::string configText = data.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace);
	if (!AIChatMcpConfigStore::SaveJsonText(configText, &error)) {
		MessageBoxW(
			hWnd,
			Utf8ToWide(error.empty() ? std::string("保存失败。") : error).c_str(),
			L"MCP 配置无效",
			MB_ICONERROR | MB_OK);
		return false;
	}

	ctx->saved = true;
	ctx->done = true;
	DestroyWindow(hWnd);
	return true;
}

void StartMcpWebView(HWND hWnd, McpWebViewDialogContext* ctx)
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
				auto* innerCtx = reinterpret_cast<McpWebViewDialogContext*>(GetWindowLongPtrW(hWnd, GWLP_USERDATA));
				if (innerCtx == nullptr || !IsWindow(hWnd)) {
					return S_OK;
				}
				if (FAILED(envResult) || environment == nullptr) {
					OutputStringToELog(std::format("[AI Chat MCP][WebView2] create environment failed hr=0x{:08X}", static_cast<unsigned int>(envResult)));
					innerCtx->fallbackRequested = true;
					DestroyWindow(hWnd);
					return S_OK;
				}

				innerCtx->webViewEnvironment = environment;
				return environment->CreateCoreWebView2Controller(
					innerCtx->hHost,
					Microsoft::WRL::Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
						[hWnd](HRESULT controllerResult, ICoreWebView2Controller* controller) -> HRESULT {
							auto* readyCtx = reinterpret_cast<McpWebViewDialogContext*>(GetWindowLongPtrW(hWnd, GWLP_USERDATA));
							if (readyCtx == nullptr || !IsWindow(hWnd)) {
								return S_OK;
							}
							if (FAILED(controllerResult) || controller == nullptr) {
								OutputStringToELog(std::format("[AI Chat MCP][WebView2] create controller failed hr=0x{:08X}", static_cast<unsigned int>(controllerResult)));
								readyCtx->fallbackRequested = true;
								DestroyWindow(hWnd);
								return S_OK;
							}

							readyCtx->webViewController = controller;
							readyCtx->webViewController->get_CoreWebView2(&readyCtx->webView);
							if (readyCtx->webView == nullptr) {
								OutputStringToELog("[AI Chat MCP][WebView2] get_CoreWebView2 returned null");
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
										auto* messageCtx = reinterpret_cast<McpWebViewDialogContext*>(GetWindowLongPtrW(hWnd, GWLP_USERDATA));
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
												TrySaveWebConfig(hWnd, messageCtx, payload["data"]);
											}
											else if (action == "cancel") {
												messageCtx->done = true;
												DestroyWindow(hWnd);
											}
										}
										catch (const std::exception& ex) {
											OutputStringToELog(std::format("[AI Chat MCP][WebView2] web message parse failed: {}", ex.what()));
										}
										catch (...) {
											OutputStringToELog("[AI Chat MCP][WebView2] web message parse failed");
										}
										return S_OK;
									}).Get(),
								nullptr);

							readyCtx->webView->add_NavigationCompleted(
								Microsoft::WRL::Callback<ICoreWebView2NavigationCompletedEventHandler>(
									[hWnd](ICoreWebView2*, ICoreWebView2NavigationCompletedEventArgs* args) -> HRESULT {
										auto* navCtx = reinterpret_cast<McpWebViewDialogContext*>(GetWindowLongPtrW(hWnd, GWLP_USERDATA));
										if (navCtx == nullptr || args == nullptr || !IsWindow(hWnd)) {
											return S_OK;
										}

										BOOL isSuccess = FALSE;
										args->get_IsSuccess(&isSuccess);
										if (isSuccess == TRUE) {
											navCtx->webViewReady = true;
											KillTimer(hWnd, kMcpWebViewInitTimerId);
											if (navCtx->hLoading != nullptr) {
												ShowWindow(navCtx->hLoading, SW_HIDE);
											}
											LayoutMcpWebViewDialog(hWnd, navCtx);
											ApplyMcpWebViewConfig(navCtx);
											return S_OK;
										}

										COREWEBVIEW2_WEB_ERROR_STATUS webErrorStatus = COREWEBVIEW2_WEB_ERROR_STATUS_UNKNOWN;
										args->get_WebErrorStatus(&webErrorStatus);
										if (webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_OPERATION_CANCELED ||
											webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_ABORTED ||
											webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_RESET) {
											OutputStringToELog(std::format(
												"[AI Chat MCP][WebView2] navigation superseded errorStatus={}",
												static_cast<int>(webErrorStatus)));
											return S_OK;
										}

										OutputStringToELog(std::format(
											"[AI Chat MCP][WebView2] navigation failed errorStatus={}",
											static_cast<int>(webErrorStatus)));
										navCtx->fallbackRequested = true;
										DestroyWindow(hWnd);
										return S_OK;
									}).Get(),
								nullptr);

							LayoutMcpWebViewDialog(hWnd, readyCtx);
							const std::wstring shellHtml = Utf8ToWide(BuildMcpWebViewShellHtml());
							if (shellHtml.empty()) {
								OutputStringToELog("[AI Chat MCP][WebView2] shell html conversion failed");
								readyCtx->fallbackRequested = true;
								DestroyWindow(hWnd);
								return S_OK;
							}
							readyCtx->webView->NavigateToString(shellHtml.c_str());
							return S_OK;
						}).Get());
			}).Get());

	if (FAILED(hr)) {
		OutputStringToELog(std::format("[AI Chat MCP][WebView2] bootstrap failed hr=0x{:08X}", static_cast<unsigned int>(hr)));
		ctx->fallbackRequested = true;
		DestroyWindow(hWnd);
	}
}

LRESULT CALLBACK McpWebViewDialogProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	auto* ctx = reinterpret_cast<McpWebViewDialogContext*>(GetWindowLongPtrW(hWnd, GWLP_USERDATA));
	switch (message)
	{
	case WM_NCCREATE: {
		const auto* cs = reinterpret_cast<CREATESTRUCTW*>(lParam);
		SetWindowLongPtrW(hWnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(cs->lpCreateParams));
		return TRUE;
	}
	case WM_CREATE:
		ctx = reinterpret_cast<McpWebViewDialogContext*>(GetWindowLongPtrW(hWnd, GWLP_USERDATA));
		if (ctx == nullptr) {
			return -1;
		}
		ctx->hHost = CreateWindowExW(
			0,
			L"STATIC",
			L"",
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
			L"正在初始化 MCP 设置页...",
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
		LayoutMcpWebViewDialog(hWnd, ctx);
		SetTimer(hWnd, kMcpWebViewInitTimerId, kMcpWebViewInitTimeoutMs, nullptr);
		StartMcpWebView(hWnd, ctx);
		return 0;
	case WM_GETMINMAXINFO: {
		auto* mmi = reinterpret_cast<MINMAXINFO*>(lParam);
		if (mmi != nullptr) {
			mmi->ptMinTrackSize.x = (std::max)(mmi->ptMinTrackSize.x, 720L);
			mmi->ptMinTrackSize.y = (std::max)(mmi->ptMinTrackSize.y, 560L);
		}
		return 0;
	}
	case WM_SIZE:
		LayoutMcpWebViewDialog(hWnd, ctx);
		return 0;
	case WM_TIMER:
		if (ctx != nullptr && wParam == kMcpWebViewInitTimerId) {
			if (!ctx->webViewReady) {
				OutputStringToELog("[AI Chat MCP][WebView2] initialization timed out, fallback to native dialog");
				ctx->fallbackRequested = true;
				DestroyWindow(hWnd);
				return 0;
			}
			KillTimer(hWnd, kMcpWebViewInitTimerId);
			return 0;
		}
		break;
	case WM_CLOSE:
		if (ctx != nullptr) {
			ctx->done = true;
		}
		DestroyWindow(hWnd);
		return 0;
	case WM_DESTROY:
		KillTimer(hWnd, kMcpWebViewInitTimerId);
		if (ctx != nullptr) {
			ctx->done = true;
			ctx->webView = nullptr;
			ctx->webViewController = nullptr;
			ctx->webViewEnvironment = nullptr;
		}
		return 0;
	default:
		break;
	}
	return DefWindowProcW(hWnd, message, wParam, lParam);
}

void LayoutMcpConfigDialog(HWND hWnd, McpConfigDialogContext* ctx)
{
	if (ctx == nullptr) {
		return;
	}
	RECT rc = {};
	GetClientRect(hWnd, &rc);
	const int width = rc.right - rc.left;
	const int height = rc.bottom - rc.top;
	const int margin = 12;
	const int buttonWidth = 92;
	const int buttonHeight = 28;
	const int gap = 8;
	const int footerY = height - margin - buttonHeight;
	MoveWindow(ctx->hEdit, margin, margin, (std::max)(120, width - margin * 2), (std::max)(80, footerY - margin - gap), TRUE);
	MoveWindow(ctx->hTemplate, margin, footerY, 120, buttonHeight, TRUE);
	MoveWindow(ctx->hCancel, width - margin - buttonWidth, footerY, buttonWidth, buttonHeight, TRUE);
	MoveWindow(ctx->hSave, width - margin - buttonWidth * 2 - gap, footerY, buttonWidth, buttonHeight, TRUE);
}

bool TrySaveConfig(HWND hWnd, McpConfigDialogContext* ctx)
{
	if (ctx == nullptr || ctx->hEdit == nullptr) {
		return false;
	}
	std::string error;
	if (!AIChatMcpConfigStore::SaveJsonText(GetEditTextUtf8(ctx->hEdit), &error)) {
		MessageBoxW(
			hWnd,
			Utf8ToWide(error.empty() ? std::string("保存失败。") : error).c_str(),
			L"MCP 配置无效",
			MB_ICONERROR | MB_OK);
		return false;
	}
	ctx->saved = true;
	ctx->done = true;
	DestroyWindow(hWnd);
	return true;
}

LRESULT CALLBACK McpConfigDialogProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	auto* ctx = reinterpret_cast<McpConfigDialogContext*>(GetWindowLongPtrW(hWnd, GWLP_USERDATA));
	switch (message)
	{
	case WM_NCCREATE: {
		const auto* cs = reinterpret_cast<CREATESTRUCTW*>(lParam);
		SetWindowLongPtrW(hWnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(cs->lpCreateParams));
		return TRUE;
	}
	case WM_CREATE:
		ctx = reinterpret_cast<McpConfigDialogContext*>(GetWindowLongPtrW(hWnd, GWLP_USERDATA));
		if (ctx == nullptr) {
			return -1;
		}
		ctx->hEdit = CreateWindowExW(
			WS_EX_CLIENTEDGE,
			L"EDIT",
			Utf8ToWide(ReadCurrentConfigText()).c_str(),
			WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_MULTILINE | ES_AUTOVSCROLL | ES_AUTOHSCROLL | WS_VSCROLL | WS_HSCROLL,
			0, 0, 100, 100,
			hWnd,
			reinterpret_cast<HMENU>(kEditId),
			nullptr,
			nullptr);
		ctx->hTemplate = CreateWindowW(L"BUTTON", L"重置示例", WS_CHILD | WS_VISIBLE | WS_TABSTOP, 0, 0, 80, 24, hWnd, reinterpret_cast<HMENU>(kTemplateId), nullptr, nullptr);
		ctx->hSave = CreateWindowW(L"BUTTON", L"保存", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON, 0, 0, 80, 24, hWnd, reinterpret_cast<HMENU>(kSaveId), nullptr, nullptr);
		ctx->hCancel = CreateWindowW(L"BUTTON", L"取消", WS_CHILD | WS_VISIBLE | WS_TABSTOP, 0, 0, 80, 24, hWnd, reinterpret_cast<HMENU>(kCancelId), nullptr, nullptr);
		SetDefaultFont(ctx->hEdit);
		SetDefaultFont(ctx->hTemplate);
		SetDefaultFont(ctx->hSave);
		SetDefaultFont(ctx->hCancel);
		LayoutMcpConfigDialog(hWnd, ctx);
		SetFocus(ctx->hEdit);
		return 0;
	case WM_SIZE:
		LayoutMcpConfigDialog(hWnd, ctx);
		return 0;
	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
		case kSaveId:
			TrySaveConfig(hWnd, ctx);
			return 0;
		case kTemplateId:
			if (ctx != nullptr && ctx->hEdit != nullptr &&
				MessageBoxW(hWnd, L"用默认示例覆盖当前编辑内容？", L"重置示例", MB_ICONQUESTION | MB_YESNO) == IDYES) {
				SetWindowTextW(ctx->hEdit, Utf8ToWide(AIChatMcpConfigStore::BuildDefaultConfigJson()).c_str());
			}
			return 0;
		case kCancelId:
			if (ctx != nullptr) {
				ctx->done = true;
			}
			DestroyWindow(hWnd);
			return 0;
		default:
			break;
		}
		break;
	case WM_CLOSE:
		if (ctx != nullptr) {
			ctx->done = true;
		}
		DestroyWindow(hWnd);
		return 0;
	case WM_DESTROY:
		if (ctx != nullptr) {
			ctx->done = true;
		}
		return 0;
	default:
		break;
	}
	return DefWindowProcW(hWnd, message, wParam, lParam);
}

template <typename DoneGetter>
void RunMcpModalWindow(HWND owner, HWND hWnd, DoneGetter doneGetter)
{
	if (owner != nullptr && IsWindow(owner)) {
		EnableWindow(owner, FALSE);
	}
	if (hWnd != nullptr && IsWindow(hWnd)) {
		ShowWindow(hWnd, SW_SHOW);
		UpdateWindow(hWnd);
	}

	MSG msg = {};
	while (!doneGetter() && GetMessageW(&msg, nullptr, 0, 0) > 0) {
		if (hWnd == nullptr || !IsWindow(hWnd) || !IsDialogMessageW(hWnd, &msg)) {
			TranslateMessage(&msg);
			DispatchMessageW(&msg);
		}
	}

	if (owner != nullptr && IsWindow(owner)) {
		EnableWindow(owner, TRUE);
		SetForegroundWindow(owner);
	}
}

bool ShowAIChatMcpConfigDialogNative(HWND owner)
{
	const wchar_t* className = L"AutoLinkerAIChatMcpConfigDialogWindow";
	WNDCLASSW wc = {};
	wc.lpfnWndProc = McpConfigDialogProc;
	wc.hInstance = GetModuleHandleW(nullptr);
	wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
	wc.lpszClassName = className;
	RegisterClassW(&wc);

	McpConfigDialogContext ctx = {};
	ctx.owner = owner;
	HWND hWnd = CreateWindowExW(
		WS_EX_DLGMODALFRAME,
		className,
		L"外部 MCP 服务器",
		WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_THICKFRAME,
		CW_USEDEFAULT,
		CW_USEDEFAULT,
		760,
		560,
		owner,
		nullptr,
		GetModuleHandleW(nullptr),
		&ctx);
	if (hWnd == nullptr) {
		return false;
	}

	RunMcpModalWindow(owner, hWnd, [&ctx]() { return ctx.done; });
	return ctx.saved;
}

McpWebViewRunResult ShowAIChatMcpConfigDialogWebView(HWND owner)
{
	McpWebViewRunResult result = {};
	if (!IsMcpWebViewRuntimeAvailable()) {
		OutputStringToELog("[AI Chat MCP][WebView2] runtime unavailable, fallback to native dialog");
		result.fallbackRequested = true;
		return result;
	}

	const wchar_t* className = L"AutoLinkerAIChatMcpConfigWebViewDialogWindow";
	WNDCLASSW wc = {};
	wc.lpfnWndProc = McpWebViewDialogProc;
	wc.hInstance = GetModuleHandleW(nullptr);
	wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
	wc.lpszClassName = className;
	RegisterClassW(&wc);

	McpWebViewDialogContext ctx = {};
	HWND hWnd = CreateWindowExW(
		WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
		className,
		L"外部 MCP 服务器",
		WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_THICKFRAME,
		CW_USEDEFAULT,
		CW_USEDEFAULT,
		820,
		680,
		owner,
		nullptr,
		GetModuleHandleW(nullptr),
		&ctx);
	if (hWnd == nullptr) {
		OutputStringToELog("[AI Chat MCP][WebView2] CreateWindowExW failed, fallback to native dialog");
		result.fallbackRequested = true;
		return result;
	}

	RunMcpModalWindow(owner, hWnd, [&ctx]() { return ctx.done; });
	result.saved = ctx.saved;
	result.fallbackRequested = ctx.fallbackRequested;
	return result;
}

} // namespace

bool ShowAIChatMcpConfigDialog(HWND owner)
{
	const McpWebViewRunResult webViewResult = ShowAIChatMcpConfigDialogWebView(owner);
	if (webViewResult.fallbackRequested) {
		OutputStringToELog("[AI Chat MCP][WebView2] fallback to native dialog");
		return ShowAIChatMcpConfigDialogNative(owner);
	}
	return webViewResult.saved;
}
