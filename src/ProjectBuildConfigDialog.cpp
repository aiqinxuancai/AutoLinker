#include "ProjectBuildConfigDialog.h"

#include <CommCtrl.h>
#include <filesystem>
#include <format>
#include <string>
#include <vector>
#include <wrl.h>

#include "..\\thirdparty\\WebView2.h"
#include "..\\thirdparty\\json.hpp"
#include "AutoLinkerSettingsDialog.h"
#include "Global.h"
#include "ProjectBuildConfigManager.h"
#include "ResourceTextLoader.h"
#include "resource.h"

namespace {

constexpr wchar_t kWindowClass[] = L"AutoLinker.ProjectBuildConfig.Page.v1";
constexpr UINT_PTR kInitTimer = 0xAC0A;

struct Context {
	HWND host = nullptr;
	HWND loading = nullptr;
	std::filesystem::path sourcePath;
	ProjectBuildConfigFile configFile;
	bool ready = false;
	Microsoft::WRL::ComPtr<ICoreWebView2Environment> environment;
	Microsoft::WRL::ComPtr<ICoreWebView2Controller> controller;
	Microsoft::WRL::ComPtr<ICoreWebView2> webView;
};

std::string WideToUtf8(const std::wstring& text)
{
	if (text.empty()) return {};
	const int size = WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
	if (size <= 0) return {};
	std::string result(static_cast<size_t>(size), '\0');
	WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), result.data(), size, nullptr, nullptr);
	return result;
}

std::wstring Utf8ToWide(const std::string& text)
{
	if (text.empty()) return {};
	const int size = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (size <= 0) return {};
	std::wstring result(static_cast<size_t>(size), L'\0');
	MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), size);
	return result;
}

std::wstring EscapeScriptString(const std::wstring& text)
{
	std::wstring result;
	for (wchar_t ch : text) {
		switch (ch) {
		case L'\\': result += L"\\\\"; break;
		case L'\'': result += L"\\'"; break;
		case L'\r': result += L"\\r"; break;
		case L'\n': result += L"\\n"; break;
		default: result.push_back(ch); break;
		}
	}
	return result;
}

std::filesystem::path GetProgramDirectory()
{
	std::vector<wchar_t> buffer(MAX_PATH);
	for (;;) {
		const DWORD length = GetModuleFileNameW(nullptr, buffer.data(), static_cast<DWORD>(buffer.size()));
		if (length == 0) return {};
		if (length < buffer.size() - 1) {
			return std::filesystem::path(std::wstring(buffer.data(), length)).parent_path();
		}
		buffer.resize(buffer.size() * 2);
	}
}

nlohmann::json ConfigToJson(const ProjectBuildConfig& config)
{
	return {
		{"name", config.name},
		{"target", ProjectBuildTargetToString(config.target)},
		{"staticCompile", config.staticCompile},
		{"outputPath", config.outputPath},
		{"preBuildCommands", config.preBuildCommands},
		{"postBuildCommands", config.postBuildCommands}
	};
}

bool JsonToConfigFile(const nlohmann::json& data, ProjectBuildConfigFile& output, std::string& error)
{
	if (!data.contains("configurations") || !data["configurations"].is_array()) {
		error = "配置列表格式无效。";
		return false;
	}
	ProjectBuildConfigFile file;
	file.activeConfiguration = data.value("activeConfiguration", std::string());
	for (const auto& item : data["configurations"]) {
		if (!item.is_object()) continue;
		ProjectBuildConfig config;
		config.name = item.value("name", std::string());
		config.target = ProjectBuildTargetFromString(item.value("target", std::string("auto")));
		config.staticCompile = item.value("staticCompile", false);
		config.outputPath = item.value("outputPath", std::string());
		if (config.name.empty() || config.outputPath.empty()) {
			error = "配置名称和输出路径不能为空。";
			return false;
		}
		for (const auto& existing : file.configurations) {
			if (_stricmp(existing.name.c_str(), config.name.c_str()) == 0) {
				error = "配置名称不能重复。";
				return false;
			}
		}
		if (item.contains("preBuildCommands") && item["preBuildCommands"].is_array()) {
			for (const auto& command : item["preBuildCommands"]) if (command.is_string() && !command.get<std::string>().empty()) config.preBuildCommands.push_back(command.get<std::string>());
		}
		if (item.contains("postBuildCommands") && item["postBuildCommands"].is_array()) {
			for (const auto& command : item["postBuildCommands"]) if (command.is_string() && !command.get<std::string>().empty()) config.postBuildCommands.push_back(command.get<std::string>());
		}
		file.configurations.push_back(std::move(config));
	}
	if (file.configurations.empty()) {
		file.activeConfiguration.clear();
	}
	else if (file.activeConfiguration.empty()) {
		file.activeConfiguration = file.configurations.front().name;
	}
	output = std::move(file);
	return true;
}

nlohmann::json BuildPayload(Context* context)
{
	nlohmann::json configs = nlohmann::json::array();
	for (const auto& config : context->configFile.configurations) configs.push_back(ConfigToJson(config));
	return {
		{"sourcePath", WideToUtf8(context->sourcePath.wstring())},
		{"configPath", WideToUtf8(GetProjectBuildConfigPath(context->sourcePath).wstring())},
		{"variables", {
			{"programDir", WideToUtf8(GetProgramDirectory().wstring())},
			{"projectDir", WideToUtf8(context->sourcePath.parent_path().wstring())},
			{"projectName", WideToUtf8(context->sourcePath.stem().wstring())},
			{"sourcePath", WideToUtf8(context->sourcePath.wstring())}
		}},
		{"activeConfiguration", context->configFile.activeConfiguration},
		{"configurations", configs}
	};
}

void SendResult(Context* context, bool ok, const std::string& message)
{
	if (context == nullptr || !context->ready || context->webView == nullptr) return;
	const nlohmann::json result = { {"ok", ok}, {"message", message}, {"data", BuildPayload(context)} };
	std::wstring script = L"window.autolinkerProjectBuildSaveResult(JSON.parse('";
	script += EscapeScriptString(Utf8ToWide(result.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace)));
	script += L"'));";
	context->webView->ExecuteScript(script.c_str(), nullptr);
}

void HandleMessage(HWND page, Context* context, const std::string& message)
{
	try {
		const auto payload = nlohmann::json::parse(message);
		const std::string action = payload.value("action", std::string());
		if (action == "ready") {
			context->ready = true;
			SendResult(context, true, "");
			return;
		}
		if (action == "save" && payload.contains("data")) {
			ProjectBuildConfigFile file;
			std::string error;
			if (!JsonToConfigFile(payload["data"], file, error) || !SaveProjectBuildConfigFile(context->sourcePath, file, &error)) {
				SendResult(context, false, error.empty() ? "保存项目配置失败。" : error);
				return;
			}
			context->configFile = std::move(file);
			SendResult(context, true, "项目配置已保存。编译菜单已同步刷新。");
			PostMessageW(GetParent(page), WM_AUTOLINKER_SETTINGS_PAGE_SAVED, static_cast<WPARAM>(AutoLinkerSettingsPageId::ProjectBuild), 0);
		}
	}
	catch (const std::exception& ex) {
		SendResult(context, false, std::string("页面消息解析失败：") + ex.what());
	}
}

void Layout(HWND page, Context* context)
{
	RECT rc = {};
	GetClientRect(page, &rc);
	if (context->host != nullptr) MoveWindow(context->host, 0, 0, rc.right, rc.bottom, TRUE);
	if (context->loading != nullptr) MoveWindow(context->loading, 18, 16, (std::max)(0L, rc.right - 36), 28, TRUE);
	if (context->controller != nullptr) context->controller->put_Bounds(rc);
}

HRESULT ControllerCreated(HWND page, HRESULT hr, ICoreWebView2Controller* controller)
{
	auto* context = reinterpret_cast<Context*>(GetWindowLongPtrW(page, GWLP_USERDATA));
	if (context == nullptr || FAILED(hr) || controller == nullptr) return S_OK;
	context->controller = controller;
	controller->get_CoreWebView2(&context->webView);
	if (context->webView == nullptr) return S_OK;
	Microsoft::WRL::ComPtr<ICoreWebView2Settings> settings;
	if (SUCCEEDED(context->webView->get_Settings(&settings)) && settings != nullptr) {
		settings->put_AreDevToolsEnabled(FALSE);
		settings->put_IsStatusBarEnabled(FALSE);
		settings->put_IsZoomControlEnabled(FALSE);
	}
	context->webView->add_WebMessageReceived(
		Microsoft::WRL::Callback<ICoreWebView2WebMessageReceivedEventHandler>(
			[page](ICoreWebView2*, ICoreWebView2WebMessageReceivedEventArgs* args) -> HRESULT {
				auto* inner = reinterpret_cast<Context*>(GetWindowLongPtrW(page, GWLP_USERDATA));
				LPWSTR raw = nullptr;
				if (inner != nullptr && args != nullptr && SUCCEEDED(args->TryGetWebMessageAsString(&raw)) && raw != nullptr) {
					const std::string message = WideToUtf8(raw);
					CoTaskMemFree(raw);
					HandleMessage(page, inner, message);
				}
				return S_OK;
			}).Get(), nullptr);
	context->webView->add_NavigationCompleted(
		Microsoft::WRL::Callback<ICoreWebView2NavigationCompletedEventHandler>(
			[page](ICoreWebView2*, ICoreWebView2NavigationCompletedEventArgs* args) -> HRESULT {
				auto* inner = reinterpret_cast<Context*>(GetWindowLongPtrW(page, GWLP_USERDATA));
				BOOL success = FALSE;
				if (inner != nullptr && args != nullptr && SUCCEEDED(args->get_IsSuccess(&success)) && success) {
					inner->ready = true;
					KillTimer(page, kInitTimer);
					ShowWindow(inner->loading, SW_HIDE);
					Layout(page, inner);
				}
				return S_OK;
			}).Get(), nullptr);
	Layout(page, context);
	const std::string html = LoadUtf8HtmlResourceText(IDR_HTML_PROJECT_BUILD_CONFIG_DIALOG);
	context->webView->NavigateToString(Utf8ToWide(html).c_str());
	return S_OK;
}

void StartWebView(HWND page, Context* context)
{
	const std::wstring folder = GetWebView2UserDataFolderPath();
	CreateCoreWebView2EnvironmentWithOptions(nullptr, folder.empty() ? nullptr : folder.c_str(), nullptr,
		Microsoft::WRL::Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
			[page](HRESULT hr, ICoreWebView2Environment* environment) -> HRESULT {
				auto* context = reinterpret_cast<Context*>(GetWindowLongPtrW(page, GWLP_USERDATA));
				if (context == nullptr || FAILED(hr) || environment == nullptr) return S_OK;
				context->environment = environment;
				return environment->CreateCoreWebView2Controller(context->host,
					Microsoft::WRL::Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
						[page](HRESULT result, ICoreWebView2Controller* controller) -> HRESULT { return ControllerCreated(page, result, controller); }).Get());
			}).Get());
}

LRESULT CALLBACK PageProc(HWND page, UINT message, WPARAM wParam, LPARAM lParam)
{
	auto* context = reinterpret_cast<Context*>(GetWindowLongPtrW(page, GWLP_USERDATA));
	switch (message) {
	case WM_NCCREATE:
		SetWindowLongPtrW(page, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(reinterpret_cast<CREATESTRUCTW*>(lParam)->lpCreateParams));
		return TRUE;
	case WM_CREATE:
		context = reinterpret_cast<Context*>(GetWindowLongPtrW(page, GWLP_USERDATA));
		context->host = CreateWindowExW(0, L"STATIC", L"", WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN, 0, 0, 0, 0, page, nullptr, nullptr, nullptr);
		context->loading = CreateWindowExW(0, L"STATIC", L"正在加载项目配置...", WS_CHILD | WS_VISIBLE, 18, 16, 500, 28, page, nullptr, nullptr, nullptr);
		SetTimer(page, kInitTimer, 12000, nullptr);
		StartWebView(page, context);
		return 0;
	case WM_SIZE: if (context != nullptr) Layout(page, context); return 0;
	case WM_TIMER: if (wParam == kInitTimer && context != nullptr && !context->ready) SetWindowTextW(context->loading, L"WebView2 初始化超时，请重新打开设置。"); return 0;
	case WM_DESTROY:
		KillTimer(page, kInitTimer);
		if (context != nullptr && context->controller != nullptr) context->controller->Close();
		return 0;
	case WM_NCDESTROY:
		SetWindowLongPtrW(page, GWLP_USERDATA, 0);
		delete context;
		break;
	}
	return DefWindowProcW(page, message, wParam, lParam);
}

} // namespace

HWND CreateProjectBuildConfigSettingsPage(HWND parent)
{
	if (parent == nullptr || !IsWindow(parent)) return nullptr;
	LPWSTR browserVersion = nullptr;
	const HRESULT webViewCheck = GetAvailableCoreWebView2BrowserVersionString(nullptr, &browserVersion);
	if (browserVersion != nullptr) CoTaskMemFree(browserVersion);
	if (FAILED(webViewCheck)) return nullptr;
	UpdateCurrentOpenSourceFile();
	if (g_nowOpenSourceFilePath.empty()) return nullptr;
	std::filesystem::path sourcePath(g_nowOpenSourceFilePath);
	if (sourcePath.extension() != ".e" && sourcePath.extension() != ".ec") return nullptr;
	WNDCLASSW wc = {};
	wc.lpfnWndProc = PageProc;
	wc.hInstance = GetModuleHandleW(nullptr);
	wc.lpszClassName = kWindowClass;
	wc.hCursor = LoadCursorW(nullptr, MAKEINTRESOURCEW(IDC_ARROW));
	wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
	RegisterClassW(&wc);
	auto* context = new Context();
	context->sourcePath = sourcePath;
	std::string error;
	context->configFile = LoadProjectBuildConfigFile(sourcePath, &error);
	if (!error.empty()) {
		delete context;
		return nullptr;
	}
	HWND page = CreateWindowExW(WS_EX_CONTROLPARENT, kWindowClass, L"", WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
		0, 0, 0, 0, parent, nullptr, wc.hInstance, context);
	if (page == nullptr) delete context;
	return page;
}
