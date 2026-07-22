#include "AutoLinkerSettingsDialog.h"

#include <CommCtrl.h>
#include <Shellapi.h>

#include <algorithm>
#include <array>
#include <filesystem>
#include <format>
#include <string>
#include <string_view>
#include <vector>
#include <wrl.h>

#include "..\\thirdparty\\json.hpp"
#include "..\\thirdparty\\WebView2.h"

#include "AIChatMcpConfigDialog.h"
#include "AIConfigDialog.h"
#include "AutoLinkerInternal.h"
#include "AutoLinkerUpdateManager.h"
#include "AutoLinkerVersion.h"
#include "EcSwitchConfigDialog.h"
#include "EPackagerIntegration.h"
#include "ForceLinkLibConfigDialog.h"
#include "Global.h"
#include "IdeCompileOutputCapture.h"
#include "IdeLogViewer.h"
#include "Logger.h"
#include "ProjectAgentsConfigDialog.h"
#include "resource.h"
#include "ResourceTextLoader.h"

#pragma comment(lib, "comctl32.lib")

namespace {

constexpr wchar_t kSettingsWindowClass[] = L"AutoLinker.UnifiedSettings.Window.v1";
constexpr wchar_t kNativePageWindowClass[] = L"AutoLinker.UnifiedSettings.NativePage.v1";
constexpr wchar_t kWebViewPageWindowClass[] = L"AutoLinker.UnifiedSettings.WebViewPage.v1";
constexpr char kLastPageConfigKey[] = "ui.settings.last_page";
constexpr char kDebugOutputOptimizationConfigKey[] = "debug.output_optimization.enabled";
constexpr char kCompileOutputCaptureHookConfigKey[] = "ide.compile_output_capture_hook.enabled";
constexpr char kDebugOutputOptimizationEnabledMessage[] =
	"开启后输出调试文本等调试输出函数的性能大幅提升，减少调试与界面更新在同一线程造成的延时";

constexpr int kNavigationWidth = 218;
constexpr int kNavigationButtonHeight = 40;
constexpr int kNavigationButtonGap = 4;
constexpr int kSettingsInitialWidth = 1200;
constexpr int kSettingsInitialHeight = 760;
constexpr UINT kNavigationCommandBase = 0x5200;
constexpr UINT_PTR kSettingsWebViewInitTimerId = 0xAC08;
constexpr UINT kSettingsWebViewInitTimeoutMs = 12000;

enum class SettingsWebViewPageKind {
	LogOptimization,
	About
};

struct NativeSettingsPageContext {
	std::wstring unavailableMessage;
	HFONT normalFont = nullptr;
};

struct SettingsWebViewPageContext {
	SettingsWebViewPageKind kind = SettingsWebViewPageKind::LogOptimization;
	HWND hostWindow = nullptr;
	HWND loadingLabel = nullptr;
	bool webViewReady = false;
	Microsoft::WRL::ComPtr<ICoreWebView2Environment> environment;
	Microsoft::WRL::ComPtr<ICoreWebView2Controller> controller;
	Microsoft::WRL::ComPtr<ICoreWebView2> webView;
};

struct SettingsWindowContext {
	HWND owner = nullptr;
	AutoLinkerSettingsPageId currentPage = AutoLinkerSettingsPageId::AiService;
	std::array<HWND, static_cast<size_t>(AutoLinkerSettingsPageId::Count)> pages = {};
	std::array<HWND, static_cast<size_t>(AutoLinkerSettingsPageId::Count)> navigationButtons = {};
	AutoLinkerSettingsResult result;
	HFONT normalFont = nullptr;
	HFONT titleFont = nullptr;
};

constexpr std::array<const wchar_t*, static_cast<size_t>(AutoLinkerSettingsPageId::Count)> kPageTitles = {
	L"AI 接口",
	L"MCP 服务",
	L"AI 对话配色",
	L"当前项目 AGENTS.md",
	L"链接器",
	L"EC 模块切换",
	L"核心库函数重写",
	L"日志优化",
	L"关于"
};

HMODULE GetCurrentModuleHandle()
{
	HMODULE module = nullptr;
	GetModuleHandleExW(
		GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
		reinterpret_cast<LPCWSTR>(&GetCurrentModuleHandle),
		&module);
	return module;
}

HFONT CreateUiFont(int pointSize, int weight = FW_NORMAL)
{
	HDC dc = GetDC(nullptr);
	const int dpi = dc != nullptr ? GetDeviceCaps(dc, LOGPIXELSY) : 96;
	if (dc != nullptr) {
		ReleaseDC(nullptr, dc);
	}
	return CreateFontW(
		-MulDiv(pointSize, dpi, 72), 0, 0, 0, weight, FALSE, FALSE, FALSE,
		DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
		CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Microsoft YaHei UI");
}

void ApplyFont(HWND window, HFONT font)
{
	if (window != nullptr && font != nullptr) {
		SendMessageW(window, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
	}
}

HWND CreateLabel(
	HWND parent,
	const wchar_t* text,
	DWORD style,
	int x,
	int y,
	int width,
	int height,
	HFONT font)
{
	HWND label = CreateWindowExW(
		0, L"STATIC", text, WS_CHILD | WS_VISIBLE | style,
		x, y, width, height, parent, nullptr, GetCurrentModuleHandle(), nullptr);
	ApplyFont(label, font);
	return label;
}

HWND CreateButton(
	HWND parent,
	UINT id,
	const wchar_t* text,
	DWORD style,
	int x,
	int y,
	int width,
	int height,
	HFONT font)
{
	HWND button = CreateWindowExW(
		0, L"BUTTON", text, WS_CHILD | WS_VISIBLE | WS_TABSTOP | style,
		x, y, width, height, parent,
		reinterpret_cast<HMENU>(static_cast<UINT_PTR>(id)),
		GetCurrentModuleHandle(), nullptr);
	ApplyFont(button, font);
	return button;
}

void WriteBoolConfig(const char* key, bool value)
{
	g_configManager.setValue(key, value ? "1" : "0");
}

std::wstring WideFromUtf8(const std::string& text)
{
	if (text.empty()) {
		return {};
	}
	const int length = MultiByteToWideChar(
		CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (length <= 0) {
		return std::wstring(text.begin(), text.end());
	}
	std::wstring result(static_cast<size_t>(length), L'\0');
	MultiByteToWideChar(
		CP_UTF8, 0, text.data(), static_cast<int>(text.size()), result.data(), length);
	return result;
}

std::string Utf8FromWide(const std::wstring& text)
{
	if (text.empty()) {
		return {};
	}
	const int length = WideCharToMultiByte(
		CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
	if (length <= 0) {
		return {};
	}
	std::string result(static_cast<size_t>(length), '\0');
	WideCharToMultiByte(
		CP_UTF8, 0, text.data(), static_cast<int>(text.size()), result.data(), length, nullptr, nullptr);
	return result;
}

std::wstring EscapeJsSingleQuotedWide(const std::wstring& text)
{
	std::wstring escaped;
	escaped.reserve(text.size() + 16);
	for (const wchar_t ch : text) {
		switch (ch) {
		case L'\\': escaped += L"\\\\"; break;
		case L'\'': escaped += L"\\'"; break;
		case L'\r': escaped += L"\\r"; break;
		case L'\n': escaped += L"\\n"; break;
		case 0x2028: escaped += L"\\u2028"; break;
		case 0x2029: escaped += L"\\u2029"; break;
		default: escaped.push_back(ch); break;
		}
	}
	return escaped;
}

bool IsSettingsWebViewAvailable()
{
	LPWSTR version = nullptr;
	const HRESULT result = GetAvailableCoreWebView2BrowserVersionString(nullptr, &version);
	if (version != nullptr) {
		CoTaskMemFree(version);
	}
	return SUCCEEDED(result);
}

std::string Base64Encode(const unsigned char* data, size_t size)
{
	static constexpr char kAlphabet[] =
		"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
	if (data == nullptr || size == 0) {
		return {};
	}
	std::string output;
	output.reserve(((size + 2) / 3) * 4);
	for (size_t offset = 0; offset < size; offset += 3) {
		const unsigned int first = data[offset];
		const unsigned int second = offset + 1 < size ? data[offset + 1] : 0;
		const unsigned int third = offset + 2 < size ? data[offset + 2] : 0;
		const unsigned int value = (first << 16) | (second << 8) | third;
		output.push_back(kAlphabet[(value >> 18) & 0x3F]);
		output.push_back(kAlphabet[(value >> 12) & 0x3F]);
		output.push_back(offset + 1 < size ? kAlphabet[(value >> 6) & 0x3F] : '=');
		output.push_back(offset + 2 < size ? kAlphabet[value & 0x3F] : '=');
	}
	return output;
}

std::string LoadAboutIconDataUrl()
{
	static const std::string dataUrl = []() {
		const HMODULE module = GetCurrentModuleHandle();
		if (module == nullptr) {
			return std::string();
		}
		const HRSRC resource = FindResourceW(
			module,
			MAKEINTRESOURCEW(IDR_PNG_APP_ICON),
			MAKEINTRESOURCEW(10));
		if (resource == nullptr) {
			return std::string();
		}
		const DWORD size = SizeofResource(module, resource);
		const HGLOBAL loaded = LoadResource(module, resource);
		const auto* data = loaded != nullptr
			? reinterpret_cast<const unsigned char*>(LockResource(loaded))
			: nullptr;
		if (data == nullptr || size == 0) {
			return std::string();
		}
		return "data:image/png;base64," + Base64Encode(data, size);
	}();
	return dataUrl;
}

const char* ComponentStateName(ComponentUpdateState state)
{
	switch (state) {
	case ComponentUpdateState::Checking: return "checking";
	case ComponentUpdateState::UpToDate: return "up_to_date";
	case ComponentUpdateState::UpdateAvailable: return "update_available";
	case ComponentUpdateState::Downloading: return "downloading";
	case ComponentUpdateState::Installing: return "installing";
	case ComponentUpdateState::ReadyToRestart: return "ready_to_restart";
	case ComponentUpdateState::Completed: return "completed";
	case ComponentUpdateState::Error: return "error";
	default: return "idle";
	}
}

nlohmann::json BuildComponentStatusJson(const ComponentUpdateStatus& status)
{
	return {
		{"state", ComponentStateName(status.state)},
		{"currentVersion", status.currentVersion},
		{"latestVersion", status.latestVersion},
		{"message", status.message},
		{"progressPercent", status.progressPercent}
	};
}

std::string DumpWebViewJson(const nlohmann::json& payload)
{
	return payload.dump(
		-1,
		' ',
		false,
		nlohmann::json::error_handler_t::replace);
}

std::string BuildLogOptimizationPayload(
	const std::string& notice = {},
	bool noticeIsError = false)
{
	const bool hookEnabled = IdeCompileOutputCapture::IsCaptureHookEnabled();
	const bool hookAvailable = IdeCompileOutputCapture::IsHookAvailable();
	const bool debugEnabled = IdeCompileOutputCapture::IsDebugOutputOptimizationEnabled();
	std::string hookStatus;
	if (!hookEnabled) {
		hookStatus = "当前已旁路；下次启动不会安装该 Hook。";
	}
	else if (hookAvailable) {
		hookStatus = "当前 Hook 已安装并启用。";
	}
	else {
		hookStatus = "设置已保存；重启 IDE 后尝试安装 Hook。";
	}
	std::string debugStatus;
	if (!hookEnabled) {
		debugStatus = "需要先启用编译阶段内部日志 Hook";
	}
	else if (!debugEnabled) {
		debugStatus = "当前未启用";
	}
	else if (hookAvailable) {
		debugStatus = "当前快速路径已启用";
	}
	else {
		debugStatus = "设置已保存；重启 IDE 后启用快速路径";
	}
	nlohmann::json payload = {
		{"logCenterOpen", IdeLogViewer::IsOpen()},
		{"compileHookEnabled", hookEnabled},
		{"hookAvailable", hookAvailable},
		{"debugOptimizationEnabled", debugEnabled},
		{"hookStatus", hookStatus},
		{"debugStatus", debugStatus},
		{"notice", notice},
		{"noticeIsError", noticeIsError}
	};
	return DumpWebViewJson(payload);
}

std::string BuildAboutPayload()
{
	const std::filesystem::path ePackagerPath =
		std::filesystem::path(GetBasePath()) / "tools" / "e-packager.exe";
	std::error_code existsError;
	nlohmann::json payload = {
		{"version", AUTOLINKER_VERSION},
		{"iconDataUrl", LoadAboutIconDataUrl()},
		{"ePackagerInstalled", std::filesystem::exists(ePackagerPath, existsError)},
		{"autoLinker", BuildComponentStatusJson(AutoLinkerUpdateManager::GetStatus())},
		{"ePackager", BuildComponentStatusJson(EPackagerIntegration::GetUpdateStatus())}
	};
	return DumpWebViewJson(payload);
}

void ExecuteSettingsWebViewScript(
	SettingsWebViewPageContext* context,
	const wchar_t* functionName,
	const std::string& payload)
{
	if (context == nullptr || !context->webViewReady || context->webView == nullptr ||
		functionName == nullptr) {
		return;
	}
	std::wstring script = L"window.";
	script += functionName;
	script += L"(JSON.parse('";
	script += EscapeJsSingleQuotedWide(WideFromUtf8(payload));
	script += L"'));";
	context->webView->ExecuteScript(script.c_str(), nullptr);
}

void ApplySettingsWebViewData(
	SettingsWebViewPageContext* context,
	const std::string& notice = {},
	bool noticeIsError = false) noexcept
{
	if (context == nullptr) {
		return;
	}
	try {
		if (context->kind == SettingsWebViewPageKind::LogOptimization) {
			ExecuteSettingsWebViewScript(
				context,
				L"autolinkerApplyLogOptimization",
				BuildLogOptimizationPayload(notice, noticeIsError));
		}
		else {
			ExecuteSettingsWebViewScript(context, L"autolinkerApplyAbout", BuildAboutPayload());
		}
	}
	catch (const std::exception& exception) {
		Logger::Instance().Write(
			"AutoLinkerSettings",
			std::format("WebView data refresh failed: {}", exception.what()));
	}
	catch (...) {
		Logger::Instance().Write("AutoLinkerSettings", "WebView data refresh failed: unknown exception");
	}
}

bool OpenAboutLink(HWND owner, std::string_view linkId)
{
	const wchar_t* url = nullptr;
	if (linkId == "github") {
		url = L"https://github.com/aiqinxuancai/AutoLinker";
	}
	else if (linkId == "releases") {
		url = L"https://github.com/aiqinxuancai/AutoLinker/releases";
	}
	else if (linkId == "config") {
		url = L"https://github.com/aiqinxuancai/AutoLinker/blob/main/CONFIG.md";
	}
	else if (linkId == "awesome_agent") {
		url = L"https://github.com/aiqinxuancai/Awesome-E-Agent";
	}
	else if (linkId == "e_packager") {
		url = L"https://github.com/aiqinxuancai/e-packager";
	}
	if (url == nullptr) {
		return false;
	}
	ShellExecuteW(owner, L"open", url, nullptr, nullptr, SW_SHOWNORMAL);
	return true;
}

void HandleLogOptimizationAction(
	SettingsWebViewPageContext* context,
	const std::string& action,
	bool enabled)
{
	std::string notice;
	bool noticeIsError = false;
	if (action == "set_log_center") {
		if (enabled) {
			if (!IdeLogViewer::Initialize(g_hwnd)) {
				notice = "未找到可用的 IDE 原生日志控件，日志中心没有打开。";
				noticeIsError = true;
			}
		}
		else {
			IdeLogViewer::Close();
		}
	}
	else if (action == "set_compile_hook") {
		WriteBoolConfig(kCompileOutputCaptureHookConfigKey, enabled);
		IdeCompileOutputCapture::SetCaptureHookEnabled(enabled);
		if (!enabled) {
			WriteBoolConfig(kDebugOutputOptimizationConfigKey, false);
		}
		else if (!IdeCompileOutputCapture::IsHookAvailable()) {
			notice = "设置已保存，重启 IDE 后才会安装编译阶段内部日志 Hook。";
		}
	}
	else if (action == "set_debug_optimization") {
		if (!IdeCompileOutputCapture::IsCaptureHookEnabled()) {
			notice = "请先启用编译阶段内部日志 Hook。";
			noticeIsError = true;
		}
		else {
			WriteBoolConfig(kDebugOutputOptimizationConfigKey, enabled);
			IdeCompileOutputCapture::SetDebugOutputOptimizationEnabled(enabled);
			if (enabled) {
				OutputStringToELog(kDebugOutputOptimizationEnabledMessage);
				if (!IdeCompileOutputCapture::IsHookAvailable()) {
					notice = "设置已保存，重启 IDE 后启用调试输出快速路径。";
				}
			}
		}
	}
	ApplySettingsWebViewData(context, notice, noticeIsError);
}

void HandleSettingsWebViewMessage(
	HWND page,
	SettingsWebViewPageContext* context,
	const std::string& message)
{
	if (context == nullptr) {
		return;
	}
	const nlohmann::json payload = nlohmann::json::parse(message);
	const std::string action = payload.value("action", "");
	if (action == "ready") {
		ApplySettingsWebViewData(context);
		return;
	}
	if (context->kind == SettingsWebViewPageKind::LogOptimization) {
		HandleLogOptimizationAction(context, action, payload.value("enabled", false));
		return;
	}
	if (action == "open_link") {
		const auto data = payload.value("data", nlohmann::json::object());
		OpenAboutLink(page, data.value("id", ""));
	}
	else if (action == "check_autolinker") {
		AutoLinkerUpdateManager::CheckForUpdatesInBackground();
		ApplySettingsWebViewData(context);
	}
	else if (action == "update_autolinker") {
		AutoLinkerUpdateManager::RunUpdateInBackground();
		ApplySettingsWebViewData(context);
	}
	else if (action == "check_e_packager") {
		EPackagerIntegration::CheckForToolUpdatesInBackground();
		ApplySettingsWebViewData(context);
	}
	else if (action == "update_e_packager") {
		EPackagerIntegration::RunToolUpdateInBackground();
		ApplySettingsWebViewData(context);
	}
}

void LayoutSettingsWebViewPage(HWND page, SettingsWebViewPageContext* context)
{
	if (context == nullptr) {
		return;
	}
	RECT rc = {};
	GetClientRect(page, &rc);
	const int width = (std::max)(0L, rc.right);
	const int height = (std::max)(0L, rc.bottom);
	if (context->hostWindow != nullptr) {
		MoveWindow(context->hostWindow, 0, 0, width, height, TRUE);
	}
	if (context->loadingLabel != nullptr) {
		MoveWindow(context->loadingLabel, 20, 18, (std::max)(0, width - 40), 28, TRUE);
	}
	if (context->controller != nullptr) {
		RECT bounds = {0, 0, static_cast<LONG>(width), static_cast<LONG>(height)};
		context->controller->put_Bounds(bounds);
	}
}

void ShowSettingsWebViewFailure(SettingsWebViewPageContext* context, const wchar_t* message)
{
	if (context == nullptr || context->loadingLabel == nullptr) {
		return;
	}
	SetWindowTextW(context->loadingLabel, message);
	ShowWindow(context->loadingLabel, SW_SHOW);
}

std::string LoadSettingsWebViewHtml(SettingsWebViewPageKind kind)
{
	return LoadUtf8HtmlResourceText(
		kind == SettingsWebViewPageKind::LogOptimization
			? IDR_HTML_LOG_OPTIMIZATION_SETTINGS
			: IDR_HTML_ABOUT_SETTINGS);
}

HRESULT OnSettingsWebViewControllerCreated(
	HWND page,
	HRESULT controllerResult,
	ICoreWebView2Controller* controller)
{
	auto* context = reinterpret_cast<SettingsWebViewPageContext*>(
		GetWindowLongPtrW(page, GWLP_USERDATA));
	if (context == nullptr || !IsWindow(page)) {
		return S_OK;
	}
	if (FAILED(controllerResult) || controller == nullptr) {
		ShowSettingsWebViewFailure(context, L"WebView2 页面初始化失败，请确认运行时安装完整。");
		return S_OK;
	}
	context->controller = controller;
	context->controller->get_CoreWebView2(&context->webView);
	if (context->webView == nullptr) {
		ShowSettingsWebViewFailure(context, L"无法创建 WebView2 页面。");
		return S_OK;
	}

	Microsoft::WRL::ComPtr<ICoreWebView2Settings> settings;
	if (SUCCEEDED(context->webView->get_Settings(&settings)) && settings != nullptr) {
		settings->put_AreDevToolsEnabled(FALSE);
		settings->put_AreDefaultContextMenusEnabled(FALSE);
		settings->put_IsStatusBarEnabled(FALSE);
		settings->put_IsZoomControlEnabled(FALSE);
	}
	context->webView->add_WebMessageReceived(
		Microsoft::WRL::Callback<ICoreWebView2WebMessageReceivedEventHandler>(
			[page](ICoreWebView2*, ICoreWebView2WebMessageReceivedEventArgs* args) -> HRESULT {
				auto* messageContext = reinterpret_cast<SettingsWebViewPageContext*>(
					GetWindowLongPtrW(page, GWLP_USERDATA));
				if (messageContext == nullptr || args == nullptr || !IsWindow(page)) {
					return S_OK;
				}
				LPWSTR rawMessage = nullptr;
				if (FAILED(args->TryGetWebMessageAsString(&rawMessage)) || rawMessage == nullptr) {
					return S_OK;
				}
				const std::string message = Utf8FromWide(rawMessage);
				CoTaskMemFree(rawMessage);
				try {
					HandleSettingsWebViewMessage(page, messageContext, message);
				}
				catch (const std::exception& exception) {
					Logger::Instance().Write(
						"AutoLinkerSettings",
						std::format("WebView message failed: {}", exception.what()));
				}
				catch (...) {
					Logger::Instance().Write(
						"AutoLinkerSettings",
						"WebView message failed: unknown exception");
				}
				return S_OK;
			}).Get(),
		nullptr);
	context->webView->add_NavigationCompleted(
		Microsoft::WRL::Callback<ICoreWebView2NavigationCompletedEventHandler>(
			[page](ICoreWebView2*, ICoreWebView2NavigationCompletedEventArgs* args) -> HRESULT {
				auto* navigationContext = reinterpret_cast<SettingsWebViewPageContext*>(
					GetWindowLongPtrW(page, GWLP_USERDATA));
				if (navigationContext == nullptr || args == nullptr || !IsWindow(page)) {
					return S_OK;
				}
				BOOL succeeded = FALSE;
				args->get_IsSuccess(&succeeded);
				if (succeeded == TRUE) {
					navigationContext->webViewReady = true;
					KillTimer(page, kSettingsWebViewInitTimerId);
					ShowWindow(navigationContext->loadingLabel, SW_HIDE);
					LayoutSettingsWebViewPage(page, navigationContext);
					ApplySettingsWebViewData(navigationContext);
				}
				else {
					ShowSettingsWebViewFailure(navigationContext, L"WebView2 页面加载失败。");
				}
				return S_OK;
			}).Get(),
		nullptr);

	LayoutSettingsWebViewPage(page, context);
	const std::string html = LoadSettingsWebViewHtml(context->kind);
	if (html.empty()) {
		ShowSettingsWebViewFailure(context, L"设置页 HTML 资源缺失。");
		return S_OK;
	}
	context->webView->NavigateToString(WideFromUtf8(html).c_str());
	return S_OK;
}

void StartSettingsWebView(HWND page, SettingsWebViewPageContext* context)
{
	if (page == nullptr || context == nullptr || context->hostWindow == nullptr) {
		return;
	}
	const std::wstring userDataFolder = GetWebView2UserDataFolderPath();
	const HRESULT result = CreateCoreWebView2EnvironmentWithOptions(
		nullptr,
		userDataFolder.empty() ? nullptr : userDataFolder.c_str(),
		nullptr,
		Microsoft::WRL::Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
			[page](HRESULT environmentResult, ICoreWebView2Environment* environment) -> HRESULT {
				auto* innerContext = reinterpret_cast<SettingsWebViewPageContext*>(
					GetWindowLongPtrW(page, GWLP_USERDATA));
				if (innerContext == nullptr || !IsWindow(page)) {
					return S_OK;
				}
				if (FAILED(environmentResult) || environment == nullptr) {
					ShowSettingsWebViewFailure(innerContext, L"无法启动 WebView2 运行时。");
					return S_OK;
				}
				innerContext->environment = environment;
				return environment->CreateCoreWebView2Controller(
					innerContext->hostWindow,
					Microsoft::WRL::Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
						[page](HRESULT controllerResult, ICoreWebView2Controller* controller) -> HRESULT {
							return OnSettingsWebViewControllerCreated(page, controllerResult, controller);
						}).Get());
			}).Get());
	if (FAILED(result)) {
		ShowSettingsWebViewFailure(context, L"WebView2 初始化请求失败。");
	}
}

void LayoutNativePage(HWND page, NativeSettingsPageContext* context)
{
	if (context == nullptr) {
		return;
	}
	RECT rc = {};
	GetClientRect(page, &rc);
	HWND label = GetDlgItem(page, 1);
	if (label != nullptr) {
		MoveWindow(label, 40, 40, (std::max)(0L, rc.right - 80), 90, TRUE);
	}
}

LRESULT CALLBACK NativeSettingsPageProc(HWND page, UINT message, WPARAM wParam, LPARAM lParam)
{
	auto* context = reinterpret_cast<NativeSettingsPageContext*>(GetWindowLongPtrW(page, GWLP_USERDATA));
	switch (message) {
	case WM_NCCREATE: {
		const auto* create = reinterpret_cast<CREATESTRUCTW*>(lParam);
		SetWindowLongPtrW(page, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create->lpCreateParams));
		return TRUE;
	}
	case WM_CREATE:
		if (context == nullptr) {
			return -1;
		}
		context->normalFont = CreateUiFont(10);
		ApplyFont(
			CreateWindowExW(
				0, L"STATIC", context->unavailableMessage.c_str(),
				WS_CHILD | WS_VISIBLE | SS_CENTER,
				0, 0, 0, 0, page,
				reinterpret_cast<HMENU>(1), GetCurrentModuleHandle(), nullptr),
			context->normalFont);
		return 0;
	case WM_SIZE:
		LayoutNativePage(page, context);
		return 0;
	case WM_ERASEBKGND: {
		RECT rc = {};
		GetClientRect(page, &rc);
		FillRect(reinterpret_cast<HDC>(wParam), &rc, reinterpret_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));
		return 1;
	}
	case WM_NCDESTROY:
		SetWindowLongPtrW(page, GWLP_USERDATA, 0);
		if (context != nullptr) {
			if (context->normalFont != nullptr) DeleteObject(context->normalFont);
			delete context;
		}
		break;
	default:
		break;
	}
	return DefWindowProcW(page, message, wParam, lParam);
}

LRESULT CALLBACK SettingsWebViewPageProc(HWND page, UINT message, WPARAM wParam, LPARAM lParam)
{
	auto* context = reinterpret_cast<SettingsWebViewPageContext*>(
		GetWindowLongPtrW(page, GWLP_USERDATA));
	switch (message) {
	case WM_NCCREATE: {
		const auto* create = reinterpret_cast<CREATESTRUCTW*>(lParam);
		SetWindowLongPtrW(page, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create->lpCreateParams));
		return TRUE;
	}
	case WM_CREATE:
		if (context == nullptr) {
			return -1;
		}
		context->hostWindow = CreateWindowExW(
			0, L"STATIC", L"",
			WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
			0, 0, 0, 0, page, nullptr, GetCurrentModuleHandle(), nullptr);
		context->loadingLabel = CreateLabel(
			page,
			context->kind == SettingsWebViewPageKind::LogOptimization
				? L"正在初始化 WebView2 日志优化页面..."
				: L"正在初始化 WebView2 关于页面...",
			SS_LEFT, 20, 18, 600, 28,
			reinterpret_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT)));
		if (context->kind == SettingsWebViewPageKind::About) {
			AutoLinkerUpdateManager::SetStatusNotificationWindow(page);
			EPackagerIntegration::SetUpdateStatusNotificationWindow(page);
		}
		LayoutSettingsWebViewPage(page, context);
		SetTimer(page, kSettingsWebViewInitTimerId, kSettingsWebViewInitTimeoutMs, nullptr);
		StartSettingsWebView(page, context);
		return 0;
	case WM_SIZE:
		LayoutSettingsWebViewPage(page, context);
		return 0;
	case WM_SHOWWINDOW:
		if (wParam != FALSE && context != nullptr && context->webViewReady) {
			ApplySettingsWebViewData(context);
		}
		break;
	case WM_TIMER:
		if (context != nullptr && wParam == kSettingsWebViewInitTimerId) {
			KillTimer(page, kSettingsWebViewInitTimerId);
			if (!context->webViewReady) {
				ShowSettingsWebViewFailure(context, L"WebView2 初始化超时，请重新打开设置窗口。");
			}
			return 0;
		}
		break;
	case WM_AUTOLINKER_COMPONENT_UPDATE_STATUS:
		if (context != nullptr && context->kind == SettingsWebViewPageKind::About) {
			ApplySettingsWebViewData(context);
		}
		return 0;
	case WM_DESTROY:
		KillTimer(page, kSettingsWebViewInitTimerId);
		if (context != nullptr) {
			if (context->controller != nullptr) {
				context->controller->Close();
			}
			context->webView = nullptr;
			context->controller = nullptr;
			context->environment = nullptr;
		}
		return 0;
	case WM_NCDESTROY: {
		if (context != nullptr && context->kind == SettingsWebViewPageKind::About) {
			AutoLinkerUpdateManager::SetStatusNotificationWindow(nullptr);
			EPackagerIntegration::SetUpdateStatusNotificationWindow(nullptr);
		}
		SetWindowLongPtrW(page, GWLP_USERDATA, 0);
		const LRESULT result = DefWindowProcW(page, message, wParam, lParam);
		delete context;
		return result;
	}
	case WM_ERASEBKGND: {
		RECT rc = {};
		GetClientRect(page, &rc);
		FillRect(reinterpret_cast<HDC>(wParam), &rc, reinterpret_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));
		return 1;
	}
	default:
		break;
	}
	return DefWindowProcW(page, message, wParam, lParam);
}

HWND CreateNativePage(HWND parent, const std::wstring& unavailableMessage)
{
	auto* context = new NativeSettingsPageContext();
	context->unavailableMessage = unavailableMessage;
	HWND page = CreateWindowExW(
		WS_EX_CONTROLPARENT,
		kNativePageWindowClass,
		L"",
		WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
		0, 0, 0, 0,
		parent, nullptr, GetCurrentModuleHandle(), context);
	if (page == nullptr) {
		delete context;
	}
	return page;
}

HWND CreateSettingsWebViewPage(HWND parent, SettingsWebViewPageKind kind)
{
	if (parent == nullptr || !IsWindow(parent) || !IsSettingsWebViewAvailable()) {
		return nullptr;
	}
	auto* context = new SettingsWebViewPageContext();
	context->kind = kind;
	HWND page = CreateWindowExW(
		WS_EX_CONTROLPARENT,
		kWebViewPageWindowClass,
		L"",
		WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
		0, 0, 0, 0,
		parent, nullptr, GetCurrentModuleHandle(), context);
	if (page == nullptr) {
		delete context;
	}
	return page;
}

AutoLinkerSettingsPageId NormalizePageId(AutoLinkerSettingsPageId requested)
{
	if (requested != AutoLinkerSettingsPageId::LastUsed) {
		const int value = static_cast<int>(requested);
		if (value >= 0 && value < static_cast<int>(AutoLinkerSettingsPageId::Count)) {
			return requested;
		}
	}
	try {
		const int stored = std::stoi(g_configManager.getValue(kLastPageConfigKey));
		if (stored >= 0 && stored < static_cast<int>(AutoLinkerSettingsPageId::Count)) {
			return static_cast<AutoLinkerSettingsPageId>(stored);
		}
	}
	catch (...) {
	}
	return AutoLinkerSettingsPageId::AiService;
}

HWND CreateSettingsPage(HWND window, AutoLinkerSettingsPageId pageId)
{
	switch (pageId) {
	case AutoLinkerSettingsPageId::AiService:
		return CreateAIConfigSettingsPage(window);
	case AutoLinkerSettingsPageId::Mcp:
		return CreateAIChatMcpConfigSettingsPage(window);
	case AutoLinkerSettingsPageId::ChatTheme:
		return CreateAIChatThemeConfigSettingsPage(window);
	case AutoLinkerSettingsPageId::ProjectAgents:
		return CreateProjectAgentsConfigSettingsPage(window);
	case AutoLinkerSettingsPageId::Linker:
		return CreateLinkerConfigSettingsPage(window);
	case AutoLinkerSettingsPageId::EcSwitch:
		return CreateEcSwitchConfigSettingsPage(window);
	case AutoLinkerSettingsPageId::ForceLinkLib:
		return CreateForceLinkLibConfigSettingsPage(window);
	case AutoLinkerSettingsPageId::LogOptimization:
		return CreateSettingsWebViewPage(window, SettingsWebViewPageKind::LogOptimization);
	case AutoLinkerSettingsPageId::About:
		return CreateSettingsWebViewPage(window, SettingsWebViewPageKind::About);
	default:
		return nullptr;
	}
}

void LayoutSettingsWindow(HWND window, SettingsWindowContext* context)
{
	if (context == nullptr) {
		return;
	}
	RECT rc = {};
	GetClientRect(window, &rc);
	const int width = (std::max)(0L, rc.right);
	const int height = (std::max)(0L, rc.bottom);
	const int navLeft = 14;
	const int navWidth = kNavigationWidth - 28;
	int y = 76;
	for (size_t index = 0; index + 1 < context->navigationButtons.size(); ++index) {
		MoveWindow(context->navigationButtons[index], navLeft, y, navWidth, kNavigationButtonHeight, TRUE);
		y += kNavigationButtonHeight + kNavigationButtonGap;
	}
	MoveWindow(
		context->navigationButtons.back(), navLeft,
		(std::max)(y + 10, height - kNavigationButtonHeight - 18),
		navWidth, kNavigationButtonHeight, TRUE);

	const int pageWidth = (std::max)(0, width - kNavigationWidth - 1);
	for (HWND page : context->pages) {
		if (page != nullptr && IsWindow(page)) {
			MoveWindow(page, kNavigationWidth + 1, 0, pageWidth, height, TRUE);
		}
	}
}

void SelectSettingsPage(HWND window, SettingsWindowContext* context, AutoLinkerSettingsPageId pageId)
{
	if (context == nullptr) {
		return;
	}
	const int pageIndex = static_cast<int>(pageId);
	if (pageIndex < 0 || pageIndex >= static_cast<int>(AutoLinkerSettingsPageId::Count)) {
		return;
	}
	for (HWND page : context->pages) {
		if (page != nullptr && IsWindow(page)) {
			ShowWindow(page, SW_HIDE);
		}
	}

	HWND& page = context->pages[static_cast<size_t>(pageIndex)];
	if (page != nullptr && !IsWindow(page)) {
		page = nullptr;
	}
	if (page == nullptr) {
		page = CreateSettingsPage(window, pageId);
		if (page == nullptr) {
			const std::wstring message = pageId == AutoLinkerSettingsPageId::ProjectAgents
				? L"当前没有打开可用的 .e 或 .ec 工程，暂时无法编辑项目 AGENTS.md。"
				: L"该设置页初始化失败，请确认已安装 Microsoft Edge WebView2 Runtime。";
			page = CreateNativePage(window, message);
		}
	}
	context->currentPage = pageId;
	g_configManager.setValue(kLastPageConfigKey, std::to_string(pageIndex));
	LayoutSettingsWindow(window, context);
	if (page != nullptr) {
		ShowWindow(page, SW_SHOW);
		SetWindowPos(page, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
	}
	for (HWND button : context->navigationButtons) {
		InvalidateRect(button, nullptr, TRUE);
	}
}

void DrawNavigationButton(const DRAWITEMSTRUCT* draw, SettingsWindowContext* context)
{
	if (draw == nullptr || context == nullptr) {
		return;
	}
	const int pageIndex = static_cast<int>(draw->CtlID - kNavigationCommandBase);
	const bool selected = pageIndex == static_cast<int>(context->currentPage);
	const bool pressed = (draw->itemState & ODS_SELECTED) != 0;
	const COLORREF background = selected
		? RGB(225, 236, 248)
		: (pressed ? RGB(235, 238, 242) : RGB(246, 247, 249));
	HBRUSH brush = CreateSolidBrush(background);
	FillRect(draw->hDC, &draw->rcItem, brush);
	DeleteObject(brush);
	if (selected) {
		RECT accent = draw->rcItem;
		accent.right = accent.left + 3;
		HBRUSH accentBrush = CreateSolidBrush(RGB(35, 103, 170));
		FillRect(draw->hDC, &accent, accentBrush);
		DeleteObject(accentBrush);
	}
	SetBkMode(draw->hDC, TRANSPARENT);
	SetTextColor(draw->hDC, selected ? RGB(25, 72, 116) : RGB(48, 55, 64));
	SelectObject(draw->hDC, context->normalFont);
	RECT textRect = draw->rcItem;
	textRect.left += 16;
	DrawTextW(
		draw->hDC,
		kPageTitles[static_cast<size_t>(pageIndex)],
		-1,
		&textRect,
		DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
}

void CenterWindow(HWND window, HWND owner)
{
	RECT windowRect = {};
	GetWindowRect(window, &windowRect);
	RECT reference = {};
	if (owner == nullptr || !IsWindow(owner) || !GetWindowRect(owner, &reference)) {
		MONITORINFO info = {};
		info.cbSize = sizeof(info);
		GetMonitorInfoW(MonitorFromWindow(window, MONITOR_DEFAULTTOPRIMARY), &info);
		reference = info.rcWork;
	}
	const int width = windowRect.right - windowRect.left;
	const int height = windowRect.bottom - windowRect.top;
	SetWindowPos(
		window, nullptr,
		reference.left + ((reference.right - reference.left) - width) / 2,
		reference.top + ((reference.bottom - reference.top) - height) / 2,
		0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
}

LRESULT CALLBACK SettingsWindowProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam)
{
	auto* context = reinterpret_cast<SettingsWindowContext*>(GetWindowLongPtrW(window, GWLP_USERDATA));
	switch (message) {
	case WM_NCCREATE: {
		const auto* create = reinterpret_cast<CREATESTRUCTW*>(lParam);
		SetWindowLongPtrW(window, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create->lpCreateParams));
		return TRUE;
	}
	case WM_CREATE:
		if (context == nullptr) {
			return -1;
		}
		context->normalFont = CreateUiFont(10);
		context->titleFont = CreateUiFont(15, FW_SEMIBOLD);
		CreateLabel(window, L"AutoLinker 设置", SS_LEFT, 18, 20, 186, 34, context->titleFont);
		for (size_t index = 0; index < context->navigationButtons.size(); ++index) {
			context->navigationButtons[index] = CreateButton(
				window,
				kNavigationCommandBase + static_cast<UINT>(index),
				kPageTitles[index],
				BS_OWNERDRAW,
				0, 0, 0, 0,
				context->normalFont);
		}
		SelectSettingsPage(window, context, context->currentPage);
		return 0;
	case WM_SIZE:
		LayoutSettingsWindow(window, context);
		return 0;
	case WM_GETMINMAXINFO: {
		auto* minMax = reinterpret_cast<MINMAXINFO*>(lParam);
		if (minMax != nullptr) {
			minMax->ptMinTrackSize.x = (std::max)(minMax->ptMinTrackSize.x, 1080L);
			minMax->ptMinTrackSize.y = (std::max)(minMax->ptMinTrackSize.y, 660L);
		}
		return 0;
	}
	case WM_COMMAND: {
		const UINT commandId = LOWORD(wParam);
		if (commandId >= kNavigationCommandBase &&
			commandId < kNavigationCommandBase + static_cast<UINT>(AutoLinkerSettingsPageId::Count)) {
			SelectSettingsPage(
				window,
				context,
				static_cast<AutoLinkerSettingsPageId>(commandId - kNavigationCommandBase));
			return 0;
		}
		break;
	}
	case WM_DRAWITEM:
		DrawNavigationButton(reinterpret_cast<DRAWITEMSTRUCT*>(lParam), context);
		return TRUE;
	case WM_AUTOLINKER_SETTINGS_PAGE_SAVED:
		if (context != nullptr) {
			if (wParam == static_cast<WPARAM>(AutoLinkerSettingsPageId::AiService)) {
				context->result.aiSettingsSaved = true;
			}
			else if (wParam == static_cast<WPARAM>(AutoLinkerSettingsPageId::Mcp)) {
				context->result.mcpSettingsSaved = true;
			}
		}
		return 0;
	case WM_CTLCOLORSTATIC:
		SetBkColor(reinterpret_cast<HDC>(wParam), RGB(246, 247, 249));
		SetDCBrushColor(reinterpret_cast<HDC>(wParam), RGB(246, 247, 249));
		return reinterpret_cast<LRESULT>(GetStockObject(DC_BRUSH));
	case WM_ERASEBKGND: {
		RECT rc = {};
		GetClientRect(window, &rc);
		RECT navRect = rc;
		navRect.right = kNavigationWidth;
		HBRUSH navBrush = CreateSolidBrush(RGB(246, 247, 249));
		FillRect(reinterpret_cast<HDC>(wParam), &navRect, navBrush);
		DeleteObject(navBrush);
		RECT divider = { kNavigationWidth, 0, kNavigationWidth + 1, rc.bottom };
		HBRUSH dividerBrush = CreateSolidBrush(RGB(218, 222, 227));
		FillRect(reinterpret_cast<HDC>(wParam), &divider, dividerBrush);
		DeleteObject(dividerBrush);
		return 1;
	}
	case WM_CLOSE:
		DestroyWindow(window);
		return 0;
	case WM_DESTROY:
		if (context != nullptr) {
			if (context->normalFont != nullptr) DeleteObject(context->normalFont);
			if (context->titleFont != nullptr) DeleteObject(context->titleFont);
			context->normalFont = nullptr;
			context->titleFont = nullptr;
		}
		return 0;
	default:
		break;
	}
	return DefWindowProcW(window, message, wParam, lParam);
}

void RegisterSettingsClasses()
{
	WNDCLASSEXW pageClass = {};
	pageClass.cbSize = sizeof(pageClass);
	pageClass.lpfnWndProc = NativeSettingsPageProc;
	pageClass.hInstance = GetCurrentModuleHandle();
	pageClass.hCursor = LoadCursor(nullptr, IDC_ARROW);
	pageClass.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
	pageClass.lpszClassName = kNativePageWindowClass;
	RegisterClassExW(&pageClass);

	WNDCLASSEXW webViewPageClass = {};
	webViewPageClass.cbSize = sizeof(webViewPageClass);
	webViewPageClass.lpfnWndProc = SettingsWebViewPageProc;
	webViewPageClass.hInstance = GetCurrentModuleHandle();
	webViewPageClass.hCursor = LoadCursor(nullptr, IDC_ARROW);
	webViewPageClass.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
	webViewPageClass.lpszClassName = kWebViewPageWindowClass;
	RegisterClassExW(&webViewPageClass);

	WNDCLASSEXW windowClass = {};
	windowClass.cbSize = sizeof(windowClass);
	windowClass.lpfnWndProc = SettingsWindowProc;
	windowClass.hInstance = GetCurrentModuleHandle();
	windowClass.hCursor = LoadCursor(nullptr, IDC_ARROW);
	windowClass.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
	windowClass.hIcon = reinterpret_cast<HICON>(LoadImageW(
		GetCurrentModuleHandle(), MAKEINTRESOURCEW(IDI_APP_ICON), IMAGE_ICON,
		GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR));
	windowClass.hIconSm = reinterpret_cast<HICON>(LoadImageW(
		GetCurrentModuleHandle(), MAKEINTRESOURCEW(IDI_APP_ICON), IMAGE_ICON,
		GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR));
	windowClass.lpszClassName = kSettingsWindowClass;
	RegisterClassExW(&windowClass);
}

} // namespace

AutoLinkerSettingsResult ShowAutoLinkerSettingsDialog(HWND owner, AutoLinkerSettingsPageId initialPage)
{
	INITCOMMONCONTROLSEX commonControls = {};
	commonControls.dwSize = sizeof(commonControls);
	commonControls.dwICC = ICC_STANDARD_CLASSES | ICC_WIN95_CLASSES;
	InitCommonControlsEx(&commonControls);
	RegisterSettingsClasses();

	SettingsWindowContext context = {};
	context.owner = owner;
	context.currentPage = NormalizePageId(initialPage);
	HWND window = CreateWindowExW(
		WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
		kSettingsWindowClass,
		L"AutoLinker 设置",
		WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_THICKFRAME,
		CW_USEDEFAULT, CW_USEDEFAULT, kSettingsInitialWidth, kSettingsInitialHeight,
		owner, nullptr, GetCurrentModuleHandle(), &context);
	if (window == nullptr) {
		OutputStringToELog("AutoLinker 统一设置窗口创建失败");
		return context.result;
	}

	if (owner != nullptr && IsWindow(owner)) {
		EnableWindow(owner, FALSE);
	}
	CenterWindow(window, owner);
	ShowWindow(window, SW_SHOW);
	UpdateWindow(window);

	MSG message = {};
	bool repostQuit = false;
	int quitCode = 0;
	while (IsWindow(window)) {
		const BOOL result = GetMessageW(&message, nullptr, 0, 0);
		if (result == 0) {
			repostQuit = true;
			quitCode = static_cast<int>(message.wParam);
			DestroyWindow(window);
			break;
		}
		if (result < 0) {
			DestroyWindow(window);
			break;
		}
		if (!IsDialogMessageW(window, &message)) {
			TranslateMessage(&message);
			DispatchMessageW(&message);
		}
	}
	if (owner != nullptr && IsWindow(owner)) {
		EnableWindow(owner, TRUE);
		SetForegroundWindow(owner);
	}
	if (repostQuit) {
		PostQuitMessage(quitCode);
	}
	return context.result;
}

std::string BuildAutoLinkerSettingsSelfTestJson()
{
	using nlohmann::json;
	const json pages = {
		"ai_service", "mcp", "chat_theme", "project_agents", "linker",
		"ec_switch", "force_link_lib", "log_optimization", "about"
	};
	const json links = {
		"https://github.com/aiqinxuancai/AutoLinker",
		"https://github.com/aiqinxuancai/AutoLinker/releases",
		"https://github.com/aiqinxuancai/AutoLinker/blob/main/CONFIG.md",
		"https://github.com/aiqinxuancai/Awesome-E-Agent",
		"https://github.com/aiqinxuancai/e-packager"
	};
	const bool hookDependencyValid =
		!IdeCompileOutputCapture::IsDebugOutputOptimizationEnabled() ||
		IdeCompileOutputCapture::IsCaptureHookEnabled();
	const std::string logOptimizationHtml = LoadUtf8HtmlResourceText(IDR_HTML_LOG_OPTIMIZATION_SETTINGS);
	const std::string aboutHtml = LoadUtf8HtmlResourceText(IDR_HTML_ABOUT_SETTINGS);
	const std::string aboutIconDataUrl = LoadAboutIconDataUrl();
	const bool webViewResourcesValid =
		logOptimizationHtml.find("autolinkerApplyLogOptimization") != std::string::npos &&
		aboutHtml.find("autolinkerApplyAbout") != std::string::npos &&
		!aboutIconDataUrl.empty();
	const bool ePackagerTwoStepUpdateValid =
		aboutHtml.find("check_e_packager") != std::string::npos &&
		aboutHtml.find("update_e_packager") != std::string::npos &&
		aboutHtml.find("checkPackagerBtn") != std::string::npos &&
		aboutHtml.find("applyPackagerBtn") != std::string::npos &&
		aboutHtml.find("检查并更新组件") == std::string::npos;
	bool invalidUtf8PayloadSafe = false;
	try {
		const json invalidPayload = {
			{"text", std::string(1, static_cast<char>(0xD0))}
		};
		invalidUtf8PayloadSafe = json::parse(DumpWebViewJson(invalidPayload)).is_object();
	}
	catch (...) {
	}
	json report = {
		{"name", "unified-settings-self-test"},
		{"ok", pages.size() == static_cast<size_t>(AutoLinkerSettingsPageId::Count) &&
			hookDependencyValid && links.size() == 5 && webViewResourcesValid &&
			invalidUtf8PayloadSafe && ePackagerTwoStepUpdateValid},
		{"page_count", pages.size()},
		{"pages", pages},
		{"about_fixed_links", links},
		{"compile_hook_default_enabled", false},
		{"debug_optimization_requires_compile_hook", hookDependencyValid},
		{"log_optimization_webview2", !logOptimizationHtml.empty()},
		{"about_webview2", !aboutHtml.empty()},
		{"about_icon_resource", !aboutIconDataUrl.empty()},
		{"e_packager_two_step_update", ePackagerTwoStepUpdateValid},
		{"webview_invalid_utf8_payload_safe", invalidUtf8PayloadSafe},
		{"log_center_config_key", "ui.log_center.open"},
		{"compile_hook_config_key", kCompileOutputCaptureHookConfigKey},
		{"debug_optimization_config_key", kDebugOutputOptimizationConfigKey}
	};
	return report.dump();
}
