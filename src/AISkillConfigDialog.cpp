#include "AISkillConfigDialog.h"

#include <Windows.h>
#include <Shellapi.h>
#include <Shobjidl.h>

#include <algorithm>
#include <atomic>
#include <format>
#include <memory>
#include <mutex>
#include <string>
#include <thread>
#include <wrl.h>

#include "..\\thirdparty\\json.hpp"
#include "..\\thirdparty\\WebView2.h"

#include "AISkillManager.h"
#include "AutoLinkerInternal.h"
#include "AutoLinkerSettingsDialog.h"
#include "Global.h"
#include "ResourceTextLoader.h"
#include "resource.h"

namespace {

using nlohmann::json;

constexpr wchar_t kSkillPageClass[] = L"AutoLinker.AISkillSettings.Page.v1";
constexpr UINT_PTR kInitTimerId = 0xAD31;
constexpr UINT kInitTimeoutMs = 12000;
constexpr UINT kAsyncResultMessage = WM_APP + 0x3B2;

struct SkillPageContext {
	HWND host = nullptr;
	HWND loading = nullptr;
	bool ready = false;
	Microsoft::WRL::ComPtr<ICoreWebView2Environment> environment;
	Microsoft::WRL::ComPtr<ICoreWebView2Controller> controller;
	Microsoft::WRL::ComPtr<ICoreWebView2> webView;
};

struct AsyncResult {
	std::string action;
	std::string jsonText;
};

std::wstring Utf8ToWide(const std::string& text)
{
	if (text.empty()) {
		return {};
	}
	const int size = MultiByteToWideChar(
		CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (size <= 0) {
		return std::wstring(text.begin(), text.end());
	}
	std::wstring result(static_cast<size_t>(size), L'\0');
	if (MultiByteToWideChar(
			CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), size) <= 0) {
		return {};
	}
	return result;
}

std::string WideToUtf8(const std::wstring& text)
{
	if (text.empty()) {
		return {};
	}
	const int size = WideCharToMultiByte(
		CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
	if (size <= 0) {
		return {};
	}
	std::string result(static_cast<size_t>(size), '\0');
	if (WideCharToMultiByte(
			CP_UTF8, 0, text.data(), static_cast<int>(text.size()), result.data(), size, nullptr, nullptr) <= 0) {
		return {};
	}
	return result;
}

HMODULE GetCurrentModuleHandleLocal()
{
	HMODULE module = nullptr;
	GetModuleHandleExW(
		GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
		reinterpret_cast<LPCWSTR>(&GetCurrentModuleHandleLocal), &module);
	return module;
}

bool IsWebViewAvailable()
{
	LPWSTR version = nullptr;
	const HRESULT hr = GetAvailableCoreWebView2BrowserVersionString(nullptr, &version);
	if (version != nullptr) {
		CoTaskMemFree(version);
	}
	return SUCCEEDED(hr);
}

void SetDefaultFont(HWND window)
{
	SendMessageW(window, WM_SETFONT, reinterpret_cast<WPARAM>(GetStockObject(DEFAULT_GUI_FONT)), TRUE);
}

void LayoutPage(HWND window, SkillPageContext* context)
{
	if (window == nullptr || context == nullptr) {
		return;
	}
	RECT rect = {};
	GetClientRect(window, &rect);
	if (context->host != nullptr) {
		MoveWindow(context->host, 0, 0, (std::max)(0L, rect.right), (std::max)(0L, rect.bottom), TRUE);
	}
	if (context->loading != nullptr) {
		MoveWindow(context->loading, 18, 16, (std::max)(120L, rect.right - 36), 28, TRUE);
	}
	if (context->controller != nullptr) {
		context->controller->put_Bounds(rect);
	}
}

void SendWebPayload(SkillPageContext* context, const std::string& action, const std::string& resultText)
{
	if (context == nullptr || !context->ready || context->webView == nullptr) {
		return;
	}
	json result = json::parse(resultText, nullptr, false);
	if (result.is_discarded()) {
		result = json({{"ok", false}, {"error", "本机返回了无效 JSON"}});
	}
	const std::string payload = json({{"action", action}, {"result", std::move(result)}})
		.dump(-1, ' ', false, json::error_handler_t::replace);
	context->webView->PostWebMessageAsJson(Utf8ToWide(payload).c_str());
}

void SendState(SkillPageContext* context)
{
	SendWebPayload(context, "state", AISkillManager::BuildSettingsPayloadJson());
}

template <typename Work>
void RunAsync(HWND window, std::string action, Work work)
{
	std::thread([window, action = std::move(action), work = std::move(work)]() mutable {
		auto result = std::make_unique<AsyncResult>();
		result->action = std::move(action);
		try {
			result->jsonText = work();
		}
		catch (const std::exception& ex) {
			result->jsonText = json({{"ok", false}, {"error", ex.what()}}).dump();
		}
		catch (...) {
			result->jsonText = R"({"ok":false,"error":"后台操作发生未知错误"})";
		}
		if (!PostMessageW(window, kAsyncResultMessage, 0, reinterpret_cast<LPARAM>(result.get()))) {
			return;
		}
		result.release();
	}).detach();
}

bool Confirm(HWND owner, const std::wstring& text, const wchar_t* title)
{
	return MessageBoxW(
		owner, text.c_str(), title,
		MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON2 | MB_SETFOREGROUND) == IDYES;
}

std::string PickLocalSource(HWND owner, bool pickFolder)
{
	Microsoft::WRL::ComPtr<IFileOpenDialog> dialog;
	const HRESULT createResult = CoCreateInstance(
		CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&dialog));
	if (FAILED(createResult) || dialog == nullptr) {
		return json({{"ok", false}, {"error", std::format("无法创建本地选择窗口：0x{:08X}", static_cast<unsigned int>(createResult))}}).dump();
	}
	DWORD options = 0;
	if (SUCCEEDED(dialog->GetOptions(&options))) {
		options |= FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST;
		if (pickFolder) {
			options |= FOS_PICKFOLDERS;
		}
		else {
			options |= FOS_FILEMUSTEXIST;
		}
		dialog->SetOptions(options);
	}
	if (pickFolder) {
		dialog->SetTitle(L"选择包含 AI 技能的目录");
	}
	else {
		dialog->SetTitle(L"选择 AI 技能 ZIP 文件");
		const COMDLG_FILTERSPEC filters[] = {
			{L"ZIP 压缩包 (*.zip)", L"*.zip"},
			{L"所有文件 (*.*)", L"*.*"}
		};
		dialog->SetFileTypes(static_cast<UINT>(std::size(filters)), filters);
		dialog->SetFileTypeIndex(1);
	}
	const HRESULT showResult = dialog->Show(owner);
	if (showResult == HRESULT_FROM_WIN32(ERROR_CANCELLED)) {
		return R"({"ok":false,"cancelled":true})";
	}
	if (FAILED(showResult)) {
		return json({{"ok", false}, {"error", std::format("选择本地来源失败：0x{:08X}", static_cast<unsigned int>(showResult))}}).dump();
	}
	Microsoft::WRL::ComPtr<IShellItem> item;
	if (FAILED(dialog->GetResult(&item)) || item == nullptr) {
		return R"({"ok":false,"error":"无法读取选择结果"})";
	}
	PWSTR rawPath = nullptr;
	const HRESULT pathResult = item->GetDisplayName(SIGDN_FILESYSPATH, &rawPath);
	if (FAILED(pathResult) || rawPath == nullptr) {
		return R"({"ok":false,"error":"无法读取本地来源路径"})";
	}
	const std::string path = WideToUtf8(rawPath);
	CoTaskMemFree(rawPath);
	return json({{"ok", true}, {"path", path}}).dump();
}

const AISkillInfo* FindSkill(const std::vector<AISkillInfo>& skills, const std::string& skillFile)
{
	const std::wstring requested = Utf8ToWide(skillFile);
	for (const auto& skill : skills) {
		if (_wcsicmp(skill.skillFile.c_str(), requested.c_str()) == 0) {
			return &skill;
		}
	}
	return nullptr;
}

void HandleMessage(HWND window, SkillPageContext* context, const json& payload)
{
	const std::string action = payload.value("action", std::string());
	if (action == "load") {
		SendState(context);
		return;
	}
	if (action == "search") {
		const std::string query = payload.value("query", std::string());
		RunAsync(window, action, [query]() { return AISkillManager::SearchSkillsSh(query); });
		return;
	}
	if (action == "details") {
		const std::string source = payload.value("source", std::string());
		const std::string skillId = payload.value("skill_id", std::string());
		RunAsync(window, action, [source, skillId]() {
			return AISkillManager::GetSkillsShSkillDetails(source, skillId);
		});
		return;
	}
	if (action == "inspect") {
		const std::string source = payload.value("source", std::string());
		RunAsync(window, action, [source]() { return AISkillManager::InspectGitHubRepository(source); });
		return;
	}
	if (action == "browse_local_folder" || action == "browse_local_zip") {
		SendWebPayload(context, action, PickLocalSource(window, action == "browse_local_folder"));
		return;
	}
	if (action == "inspect_local") {
		const std::string source = payload.value("source", std::string());
		RunAsync(window, action, [source]() { return AISkillManager::InspectLocalSource(source); });
		return;
	}
	if (action == "install_local") {
		const std::string source = payload.value("source", std::string());
		const std::string selectedPath = payload.value("candidate_path", std::string());
		const bool replace = payload.value("allow_replace", false);
		const AISkillScope scope = payload.value("scope", std::string("global")) == "project"
			? AISkillScope::Repo : AISkillScope::User;
		const AISkillLocalInstallMode mode = payload.value("mode", std::string("copy")) == "reference"
			? AISkillLocalInstallMode::Reference : AISkillLocalInstallMode::Copy;
		RunAsync(window, action, [source, scope, selectedPath, mode, replace]() {
			return AISkillManager::InstallFromLocal(source, scope, selectedPath, mode, replace);
		});
		return;
	}
	if (action == "install") {
		const std::string source = payload.value("source", std::string());
		const std::string selectedPath = payload.value("repository_path", std::string());
		const std::string skillId = payload.value("skill_id", std::string());
		const bool replace = payload.value("allow_replace", false);
		const AISkillScope scope = payload.value("scope", std::string("global")) == "project"
			? AISkillScope::Repo : AISkillScope::User;
		RunAsync(window, action, [source, scope, selectedPath, skillId, replace]() {
			return AISkillManager::InstallFromGitHub(source, scope, selectedPath, skillId, replace);
		});
		return;
	}
	if (action == "update") {
		const std::string skillFile = payload.value("skill_file", std::string());
		RunAsync(window, action, [skillFile]() { return AISkillManager::UpdateInstalledSkill(skillFile); });
		return;
	}
	if (action == "toggle") {
		std::string error;
		const bool ok = AISkillManager::SetSkillEnabled(
			payload.value("skill_file", std::string()), payload.value("enabled", true), error);
		SendWebPayload(context, action, json({{"ok", ok}, {"error", error}}).dump());
		if (ok) {
			PostMessageW(
				GetParent(window), WM_AUTOLINKER_SETTINGS_PAGE_SAVED,
				static_cast<WPARAM>(AutoLinkerSettingsPageId::Skills), 0);
		}
		SendState(context);
		return;
	}
	if (action == "remove") {
		const std::string skillFile = payload.value("skill_file", std::string());
		const auto skills = AISkillManager::DiscoverSkills();
		const AISkillInfo* skill = FindSkill(skills, skillFile);
		if (skill == nullptr) {
			SendWebPayload(context, action, R"({"ok":false,"error":"未找到技能"})");
			return;
		}
		const std::wstring confirmText = skill->externalReference
			? L"将移除此技能的原路径引用，源目录和文件不会被删除：\r\n\r\n" + skill->directory.wstring()
			: L"将删除技能目录及其中全部文件：\r\n\r\n" + skill->directory.wstring();
		if (!Confirm(window, confirmText, skill->externalReference ? L"移除 AI 技能引用" : L"卸载 AI 技能")) {
			SendWebPayload(context, action, R"({"ok":false,"cancelled":true})");
			return;
		}
		RunAsync(window, action, [skillFile]() { return AISkillManager::RemoveInstalledSkill(skillFile); });
		return;
	}
	if (action == "open_folder") {
		const auto skills = AISkillManager::DiscoverSkills();
		const AISkillInfo* skill = FindSkill(skills, payload.value("skill_file", std::string()));
		if (skill != nullptr) {
			ShellExecuteW(window, L"open", skill->directory.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
		}
		return;
	}
}

void StartWebView(HWND window, SkillPageContext* context)
{
	const std::wstring userDataFolder = GetWebView2UserDataFolderPath();
	const HRESULT hr = CreateCoreWebView2EnvironmentWithOptions(
		nullptr, userDataFolder.empty() ? nullptr : userDataFolder.c_str(), nullptr,
		Microsoft::WRL::Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
			[window](HRESULT result, ICoreWebView2Environment* environment) -> HRESULT {
				auto* page = reinterpret_cast<SkillPageContext*>(GetWindowLongPtrW(window, GWLP_USERDATA));
				if (page == nullptr || FAILED(result) || environment == nullptr || !IsWindow(window)) {
					return S_OK;
				}
				page->environment = environment;
				return environment->CreateCoreWebView2Controller(
					page->host,
					Microsoft::WRL::Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
						[window](HRESULT controllerResult, ICoreWebView2Controller* controller) -> HRESULT {
							auto* ready = reinterpret_cast<SkillPageContext*>(GetWindowLongPtrW(window, GWLP_USERDATA));
							if (ready == nullptr || FAILED(controllerResult) || controller == nullptr || !IsWindow(window)) {
								return S_OK;
							}
							ready->controller = controller;
							ready->controller->get_CoreWebView2(&ready->webView);
							if (ready->webView == nullptr) {
								return S_OK;
							}
							Microsoft::WRL::ComPtr<ICoreWebView2Settings> settings;
							if (SUCCEEDED(ready->webView->get_Settings(&settings)) && settings != nullptr) {
								settings->put_AreDevToolsEnabled(FALSE);
								settings->put_AreDefaultContextMenusEnabled(FALSE);
								settings->put_IsStatusBarEnabled(FALSE);
								settings->put_IsZoomControlEnabled(FALSE);
							}
							ready->webView->add_WebMessageReceived(
								Microsoft::WRL::Callback<ICoreWebView2WebMessageReceivedEventHandler>(
									[window](ICoreWebView2*, ICoreWebView2WebMessageReceivedEventArgs* args) -> HRESULT {
										auto* messageContext = reinterpret_cast<SkillPageContext*>(
											GetWindowLongPtrW(window, GWLP_USERDATA));
										if (messageContext == nullptr || args == nullptr || !IsWindow(window)) {
											return S_OK;
										}
										LPWSTR raw = nullptr;
										if (FAILED(args->TryGetWebMessageAsString(&raw)) || raw == nullptr) {
											return S_OK;
										}
										const json payload = json::parse(WideToUtf8(raw), nullptr, false);
										CoTaskMemFree(raw);
										if (payload.is_object()) {
											HandleMessage(window, messageContext, payload);
										}
										return S_OK;
									}).Get(), nullptr);
							ready->webView->add_NavigationCompleted(
								Microsoft::WRL::Callback<ICoreWebView2NavigationCompletedEventHandler>(
									[window](ICoreWebView2*, ICoreWebView2NavigationCompletedEventArgs* args) -> HRESULT {
										auto* nav = reinterpret_cast<SkillPageContext*>(GetWindowLongPtrW(window, GWLP_USERDATA));
										BOOL success = FALSE;
										if (nav != nullptr && args != nullptr) {
											args->get_IsSuccess(&success);
										}
										if (nav != nullptr && success == TRUE) {
											nav->ready = true;
											KillTimer(window, kInitTimerId);
											ShowWindow(nav->loading, SW_HIDE);
											LayoutPage(window, nav);
											SendState(nav);
										}
										return S_OK;
									}).Get(), nullptr);
							LayoutPage(window, ready);
							const std::string html = LoadUtf8HtmlResourceText(IDR_HTML_AI_SKILL_CONFIG_DIALOG);
							ready->webView->NavigateToString(Utf8ToWide(html).c_str());
							return S_OK;
						}).Get());
			}).Get());
	if (FAILED(hr)) {
		SetWindowTextW(context->loading, L"无法初始化 WebView2 技能设置页。\r\n请安装或修复 Microsoft Edge WebView2 Runtime。");
	}
}

LRESULT CALLBACK SkillPageProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam)
{
	auto* context = reinterpret_cast<SkillPageContext*>(GetWindowLongPtrW(window, GWLP_USERDATA));
	switch (message) {
	case WM_NCCREATE: {
		const auto* create = reinterpret_cast<CREATESTRUCTW*>(lParam);
		SetWindowLongPtrW(window, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create->lpCreateParams));
		return TRUE;
	}
	case WM_CREATE:
		context = reinterpret_cast<SkillPageContext*>(GetWindowLongPtrW(window, GWLP_USERDATA));
		if (context == nullptr) {
			return -1;
		}
		context->host = CreateWindowExW(
			0, L"STATIC", L"", WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
			0, 0, 0, 0, window, nullptr, GetCurrentModuleHandleLocal(), nullptr);
		context->loading = CreateWindowExW(
			0, L"STATIC", L"正在初始化 SKILL 设置页...", WS_CHILD | WS_VISIBLE,
			0, 0, 0, 0, window, nullptr, GetCurrentModuleHandleLocal(), nullptr);
		SetDefaultFont(context->loading);
		LayoutPage(window, context);
		SetTimer(window, kInitTimerId, kInitTimeoutMs, nullptr);
		StartWebView(window, context);
		return 0;
	case WM_SIZE:
		LayoutPage(window, context);
		return 0;
	case WM_TIMER:
		if (wParam == kInitTimerId && context != nullptr && !context->ready) {
			KillTimer(window, kInitTimerId);
			SetWindowTextW(context->loading, L"SKILL 设置页初始化超时，请重新打开设置窗口。");
			return 0;
		}
		break;
	case kAsyncResultMessage: {
		std::unique_ptr<AsyncResult> result(reinterpret_cast<AsyncResult*>(lParam));
		if (context != nullptr && result != nullptr) {
			SendWebPayload(context, result->action, result->jsonText);
			if (result->action == "install" || result->action == "install_local" ||
				result->action == "update" || result->action == "remove") {
				const json parsed = json::parse(result->jsonText, nullptr, false);
				if (parsed.is_object() && parsed.value("ok", false)) {
					PostMessageW(
						GetParent(window), WM_AUTOLINKER_SETTINGS_PAGE_SAVED,
						static_cast<WPARAM>(AutoLinkerSettingsPageId::Skills), 0);
					SendState(context);
				}
			}
		}
		return 0;
	}
	case WM_DESTROY:
		KillTimer(window, kInitTimerId);
		if (context != nullptr) {
			context->ready = false;
			context->webView = nullptr;
			context->controller = nullptr;
			context->environment = nullptr;
		}
		return 0;
	case WM_NCDESTROY: {
		SetWindowLongPtrW(window, GWLP_USERDATA, 0);
		const LRESULT result = DefWindowProcW(window, message, wParam, lParam);
		delete context;
		return result;
	}
	default:
		break;
	}
	return DefWindowProcW(window, message, wParam, lParam);
}

void RegisterPageClass()
{
	static std::once_flag once;
	std::call_once(once, []() {
		WNDCLASSEXW windowClass = {};
		windowClass.cbSize = sizeof(windowClass);
		windowClass.lpfnWndProc = SkillPageProc;
		windowClass.hInstance = GetCurrentModuleHandleLocal();
		windowClass.hCursor = LoadCursor(nullptr, IDC_ARROW);
		windowClass.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
		windowClass.lpszClassName = kSkillPageClass;
		RegisterClassExW(&windowClass);
	});
}

} // namespace

HWND CreateAISkillConfigSettingsPage(HWND parent)
{
	if (parent == nullptr || !IsWindow(parent) || !IsWebViewAvailable()) {
		return nullptr;
	}
	UpdateCurrentOpenSourceFile();
	RegisterPageClass();
	auto* context = new SkillPageContext();
	HWND page = CreateWindowExW(
		WS_EX_CONTROLPARENT, kSkillPageClass, L"",
		WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
		0, 0, 0, 0, parent, nullptr, GetCurrentModuleHandleLocal(), context);
	if (page == nullptr) {
		delete context;
	}
	return page;
}
