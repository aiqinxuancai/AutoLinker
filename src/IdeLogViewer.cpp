#include "IdeLogViewer.h"

#include <Windows.h>
#include <CommDlg.h>

#include <algorithm>
#include <cstdint>
#include <ctime>
#include <filesystem>
#include <format>
#include <fstream>
#include <iterator>
#include <new>
#include <string>
#include <utility>
#include <vector>
#include <wrl.h>

#include "..\\thirdparty\\json.hpp"
#include "..\\thirdparty\\WebView2.h"
#include "Global.h"
#include "IDEFacade.h"
#include "IdeCompileOutputCapture.h"
#include "IdeLogStore.h"
#include "IdeOutputTabController.h"
#include "Logger.h"
#include "ResourceTextLoader.h"
#include "resource.h"

namespace IdeLogViewer {
namespace {

constexpr wchar_t kViewerWindowClass[] = L"AutoLinkerIdeLogViewerWindow";
constexpr UINT_PTR kFlushTimerId = 1;
constexpr UINT kFlushIntervalMs = 80;
constexpr std::size_t kMaxEntriesPerBatch = 500;
constexpr std::size_t kMaxTextBytesPerBatch = 512 * 1024;
constexpr ULONGLONG kControlFallbackPollMs = 400;

struct ViewerContext {
	HWND hostWindow = nullptr;
	HWND loadingLabel = nullptr;
	HWND fallbackEdit = nullptr;
	bool webViewContentReady = false;
	bool fallbackActive = false;
	std::uint64_t lastSequence = 0;
	std::uint64_t generation = 0;
	ULONGLONG lastControlPollTick = 0;
	std::string lastControlText;
	Microsoft::WRL::ComPtr<ICoreWebView2Environment> environment;
	Microsoft::WRL::ComPtr<ICoreWebView2Controller> controller;
	Microsoft::WRL::ComPtr<ICoreWebView2> webView;
};

HWND g_mainWindow = nullptr;
HWND g_viewerWindow = nullptr;
bool g_isOpen = false;
IdeOutputTabController::HiddenTabState g_hiddenTabState;

HMODULE GetCurrentModuleHandle()
{
	HMODULE module = nullptr;
	GetModuleHandleExW(
		GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
		reinterpret_cast<LPCWSTR>(&GetCurrentModuleHandle),
		&module);
	return module;
}

std::wstring Utf8ToWide(const std::string& text)
{
	if (text.empty()) {
		return {};
	}
	const int length = MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0);
	if (length <= 0) {
		return {};
	}
	std::wstring wide(static_cast<std::size_t>(length), L'\0');
	if (MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		wide.data(),
		length) <= 0) {
		return {};
	}
	return wide;
}

std::string WideToUtf8(const std::wstring& text)
{
	if (text.empty()) {
		return {};
	}
	const int length = WideCharToMultiByte(
		CP_UTF8,
		0,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0,
		nullptr,
		nullptr);
	if (length <= 0) {
		return {};
	}
	std::string utf8(static_cast<std::size_t>(length), '\0');
	if (WideCharToMultiByte(
		CP_UTF8,
		0,
		text.data(),
		static_cast<int>(text.size()),
		utf8.data(),
		length,
		nullptr,
		nullptr) <= 0) {
		return {};
	}
	return utf8;
}

std::string WideToLocal(const std::wstring& text)
{
	if (text.empty()) {
		return {};
	}
	const int length = WideCharToMultiByte(
		CP_ACP,
		0,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0,
		nullptr,
		nullptr);
	if (length <= 0) {
		return {};
	}
	std::string local(static_cast<std::size_t>(length), '\0');
	WideCharToMultiByte(
		CP_ACP,
		0,
		text.data(),
		static_cast<int>(text.size()),
		local.data(),
		length,
		nullptr,
		nullptr);
	return local;
}

std::string LocalToUtf8(const std::string& text)
{
	if (text.empty()) {
		return {};
	}
	const int wideLength = MultiByteToWideChar(
		CP_ACP,
		0,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0);
	if (wideLength <= 0) {
		return {};
	}
	std::wstring wide(static_cast<std::size_t>(wideLength), L'\0');
	if (MultiByteToWideChar(
		CP_ACP,
		0,
		text.data(),
		static_cast<int>(text.size()),
		wide.data(),
		wideLength) <= 0) {
		return {};
	}
	return WideToUtf8(wide);
}

std::string FormatEntryTime(std::uint64_t unixTimeMs)
{
	const std::time_t seconds = static_cast<std::time_t>(unixTimeMs / 1000ULL);
	std::tm local = {};
	if (localtime_s(&local, &seconds) != 0) {
		return {};
	}
	return std::format(
		"{:04}-{:02}-{:02} {:02}:{:02}:{:02}.{:03}",
		local.tm_year + 1900,
		local.tm_mon + 1,
		local.tm_mday,
		local.tm_hour,
		local.tm_min,
		local.tm_sec,
		unixTimeMs % 1000ULL);
}

void LayoutViewer(HWND window, ViewerContext* context)
{
	if (context == nullptr) {
		return;
	}
	RECT client = {};
	GetClientRect(window, &client);
	const int width = (std::max)(0L, client.right - client.left);
	const int height = (std::max)(0L, client.bottom - client.top);
	if (context->hostWindow != nullptr) {
		MoveWindow(context->hostWindow, 0, 0, width, height, TRUE);
	}
	if (context->loadingLabel != nullptr) {
		MoveWindow(context->loadingLabel, 16, 16, (std::max)(0, width - 32), 24, TRUE);
	}
	if (context->fallbackEdit != nullptr) {
		MoveWindow(context->fallbackEdit, 0, 0, width, height, TRUE);
	}
	if (context->controller != nullptr) {
		RECT bounds = {0, 0, width, height};
		context->controller->put_Bounds(bounds);
	}
}

void ShowFallback(ViewerContext* context, const std::wstring& reason)
{
	if (context == nullptr) {
		return;
	}
	context->fallbackActive = true;
	context->webViewContentReady = false;
	if (context->hostWindow != nullptr) {
		ShowWindow(context->hostWindow, SW_HIDE);
	}
	if (context->loadingLabel != nullptr) {
		ShowWindow(context->loadingLabel, SW_HIDE);
	}
	if (context->fallbackEdit != nullptr) {
		SetWindowTextW(context->fallbackEdit, reason.c_str());
		ShowWindow(context->fallbackEdit, SW_SHOW);
	}
}

nlohmann::json BuildEntryJson(const IdeLogStore::Entry& entry)
{
	const char* source = "ide";
	if (entry.source == IdeLogStore::EntrySource::ExistingOutputSnapshot) {
		source = "snapshot";
	}
	else if (entry.source == IdeLogStore::EntrySource::OutputControlFallback) {
		source = "control";
	}
	else if (entry.source == IdeLogStore::EntrySource::AutoLinkerDirect) {
		source = "autolinker";
	}
	return {
		{"seq", entry.sequence},
		{"time", entry.unixTimeMs},
		{"thread", entry.threadId},
		{"source", source},
		{"text", LocalToUtf8(entry.localText)},
		{"truncated", entry.truncated}
	};
}

bool PostJson(ViewerContext* context, const nlohmann::json& payload)
{
	if (context == nullptr || context->webView == nullptr || !context->webViewContentReady) {
		return false;
	}
	const std::wstring json = Utf8ToWide(payload.dump(
		-1,
		' ',
		false,
		nlohmann::json::error_handler_t::replace));
	return !json.empty() && SUCCEEDED(context->webView->PostWebMessageAsJson(json.c_str()));
}

void FlushWebView(ViewerContext* context)
{
	if (context == nullptr || !context->webViewContentReady || context->webView == nullptr) {
		return;
	}

	const IdeLogStore::Batch batch = IdeLogStore::ReadAfter(
		context->lastSequence,
		context->generation,
		kMaxEntriesPerBatch,
		kMaxTextBytesPerBatch);
	if (!batch.resetRequired && batch.entries.empty()) {
		return;
	}

	nlohmann::json entries = nlohmann::json::array();
	for (const IdeLogStore::Entry& entry : batch.entries) {
		entries.push_back(BuildEntryJson(entry));
	}
	const nlohmann::json payload = {
		{"type", "log-batch"},
		{"reset", batch.resetRequired},
		{"generation", batch.generation},
		{"latestSequence", batch.latestSequence},
		{"droppedEntries", batch.droppedEntries},
		{"entries", std::move(entries)}
	};
	if (!PostJson(context, payload)) {
		return;
	}

	context->generation = batch.generation;
	if (!batch.entries.empty()) {
		context->lastSequence = batch.entries.back().sequence;
	}
	else if (batch.resetRequired) {
		context->lastSequence = batch.latestSequence;
	}
}

void AppendFallbackBatch(ViewerContext* context)
{
	if (context == nullptr || !context->fallbackActive || context->fallbackEdit == nullptr) {
		return;
	}
	const IdeLogStore::Batch batch = IdeLogStore::ReadAfter(
		context->lastSequence,
		context->generation,
		100,
		kMaxTextBytesPerBatch);
	if (batch.resetRequired) {
		SetWindowTextW(context->fallbackEdit, L"");
	}
	for (const IdeLogStore::Entry& entry : batch.entries) {
		const std::wstring line = Utf8ToWide(
			FormatEntryTime(entry.unixTimeMs) + " [T" + std::to_string(entry.threadId) + "] " +
			LocalToUtf8(entry.localText) + "\r\n");
		SendMessageW(context->fallbackEdit, EM_SETSEL, static_cast<WPARAM>(-1), static_cast<LPARAM>(-1));
		SendMessageW(context->fallbackEdit, EM_REPLACESEL, FALSE, reinterpret_cast<LPARAM>(line.c_str()));
	}
	context->generation = batch.generation;
	if (!batch.entries.empty()) {
		context->lastSequence = batch.entries.back().sequence;
	}
	else if (batch.resetRequired) {
		context->lastSequence = batch.latestSequence;
	}
}

void PostExportResult(ViewerContext* context, bool ok, const std::wstring& path, const std::string& error)
{
	PostJson(context, {
		{"type", "export-result"},
		{"ok", ok},
		{"path", WideToUtf8(path)},
		{"error", error}
	});
}

void ExportLogs(HWND owner, ViewerContext* context)
{
	SYSTEMTIME now = {};
	GetLocalTime(&now);
	wchar_t fileBuffer[MAX_PATH] = {};
	const std::wstring defaultName = std::format(
		L"AutoLinker-logs-{:04}{:02}{:02}-{:02}{:02}{:02}.log",
		now.wYear,
		now.wMonth,
		now.wDay,
		now.wHour,
		now.wMinute,
		now.wSecond);
	wcsncpy_s(fileBuffer, defaultName.c_str(), _TRUNCATE);

	OPENFILENAMEW dialog = {};
	dialog.lStructSize = sizeof(dialog);
	dialog.hwndOwner = owner;
	dialog.lpstrFilter = L"日志文件 (*.log)\0*.log\0文本文件 (*.txt)\0*.txt\0所有文件 (*.*)\0*.*\0";
	dialog.lpstrFile = fileBuffer;
	dialog.nMaxFile = static_cast<DWORD>(std::size(fileBuffer));
	dialog.lpstrDefExt = L"log";
	dialog.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR;
	if (!GetSaveFileNameW(&dialog)) {
		return;
	}

	const std::vector<IdeLogStore::Entry> entries = IdeLogStore::SnapshotAll();
	std::ofstream file(std::filesystem::path(fileBuffer), std::ios::binary | std::ios::trunc);
	if (!file.is_open()) {
		PostExportResult(context, false, fileBuffer, "open_failed");
		return;
	}
	static constexpr char kUtf8Bom[] = "\xEF\xBB\xBF";
	file.write(kUtf8Bom, 3);
	for (const IdeLogStore::Entry& entry : entries) {
		const std::string prefix = std::format(
			"[{}] [T{}] ",
			FormatEntryTime(entry.unixTimeMs),
			entry.threadId);
		const std::string text = LocalToUtf8(entry.localText);
		file.write(prefix.data(), static_cast<std::streamsize>(prefix.size()));
		file.write(text.data(), static_cast<std::streamsize>(text.size()));
		if (text.size() < 2 || text.substr(text.size() - 2) != "\r\n") {
			file.write("\r\n", 2);
		}
	}
	file.flush();
	const bool ok = file.good();
	file.close();
	PostExportResult(context, ok, fileBuffer, ok ? "" : "write_failed");
	Logger::Instance().Write(
		"IdeLogViewer",
		ok ? "exported log file" : "failed to export log file");
}

void HandleWebMessage(HWND window, ViewerContext* context, ICoreWebView2WebMessageReceivedEventArgs* args)
{
	if (context == nullptr || args == nullptr) {
		return;
	}
	LPWSTR rawMessage = nullptr;
	if (FAILED(args->TryGetWebMessageAsString(&rawMessage)) || rawMessage == nullptr) {
		return;
	}
	const std::string message = WideToUtf8(rawMessage);
	CoTaskMemFree(rawMessage);
	const nlohmann::json payload = nlohmann::json::parse(message, nullptr, false);
	if (!payload.is_object()) {
		return;
	}
	const std::string action = payload.value("action", std::string());
	if (action == "ready") {
		context->webViewContentReady = true;
		context->lastSequence = 0;
		context->generation = 0;
		if (context->loadingLabel != nullptr) {
			ShowWindow(context->loadingLabel, SW_HIDE);
		}
		Logger::Instance().Write("IdeLogViewer", "webview ready");
		FlushWebView(context);
	}
	else if (action == "batch-applied") {
		Logger::Instance().Write(
			"IdeLogViewer",
			std::format("first webview log batch applied count={}", payload.value("count", 0)));
	}
	else if (action == "clear") {
		IdeLogStore::Clear();
		FlushWebView(context);
	}
	else if (action == "export") {
		ExportLogs(window, context);
	}
}

void StartWebView(HWND window, ViewerContext* context)
{
	if (context == nullptr || context->hostWindow == nullptr) {
		return;
	}
	LPWSTR version = nullptr;
	const HRESULT runtimeResult = GetAvailableCoreWebView2BrowserVersionString(nullptr, &version);
	if (version != nullptr) {
		CoTaskMemFree(version);
	}
	if (FAILED(runtimeResult)) {
		ShowFallback(context, L"未检测到 WebView2 Runtime，已切换到基础日志视图。\r\n");
		return;
	}

	using Microsoft::WRL::Callback;
	const std::wstring userDataFolder = GetWebView2UserDataFolderPath();
	const HRESULT createResult = CreateCoreWebView2EnvironmentWithOptions(
		nullptr,
		userDataFolder.empty() ? nullptr : userDataFolder.c_str(),
		nullptr,
		Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
			[window](HRESULT environmentResult, ICoreWebView2Environment* environment) -> HRESULT {
				if (!IsWindow(window)) {
					return S_OK;
				}
				auto* current = reinterpret_cast<ViewerContext*>(GetWindowLongPtrW(window, GWLP_USERDATA));
				if (current == nullptr || FAILED(environmentResult) || environment == nullptr) {
					if (current != nullptr) {
						ShowFallback(current, L"WebView2 环境初始化失败，已切换到基础日志视图。\r\n");
					}
					return S_OK;
				}
				current->environment = environment;
				return environment->CreateCoreWebView2Controller(
					current->hostWindow,
					Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
						[window](HRESULT controllerResult, ICoreWebView2Controller* controller) -> HRESULT {
							if (!IsWindow(window)) {
								return S_OK;
							}
							auto* inner = reinterpret_cast<ViewerContext*>(GetWindowLongPtrW(window, GWLP_USERDATA));
							if (inner == nullptr || FAILED(controllerResult) || controller == nullptr) {
								if (inner != nullptr) {
									ShowFallback(inner, L"WebView2 控件初始化失败，已切换到基础日志视图。\r\n");
								}
								return S_OK;
							}

							inner->controller = controller;
							inner->controller->get_CoreWebView2(&inner->webView);
							if (inner->webView == nullptr) {
								ShowFallback(inner, L"WebView2 核心不可用，已切换到基础日志视图。\r\n");
								return S_OK;
							}

							Microsoft::WRL::ComPtr<ICoreWebView2Settings> settings;
							if (SUCCEEDED(inner->webView->get_Settings(&settings)) && settings != nullptr) {
								settings->put_IsStatusBarEnabled(FALSE);
								settings->put_AreDevToolsEnabled(FALSE);
								settings->put_IsZoomControlEnabled(FALSE);
								settings->put_AreDefaultContextMenusEnabled(TRUE);
							}

							inner->webView->add_WebMessageReceived(
								Callback<ICoreWebView2WebMessageReceivedEventHandler>(
									[window](ICoreWebView2*, ICoreWebView2WebMessageReceivedEventArgs* args) -> HRESULT {
										if (IsWindow(window)) {
											auto* messageContext = reinterpret_cast<ViewerContext*>(
												GetWindowLongPtrW(window, GWLP_USERDATA));
											HandleWebMessage(window, messageContext, args);
										}
										return S_OK;
									}).Get(),
								nullptr);

							inner->webView->add_NavigationCompleted(
								Callback<ICoreWebView2NavigationCompletedEventHandler>(
									[window](ICoreWebView2*, ICoreWebView2NavigationCompletedEventArgs* args) -> HRESULT {
										if (!IsWindow(window) || args == nullptr) {
											return S_OK;
										}
										BOOL success = FALSE;
										args->get_IsSuccess(&success);
										if (success != TRUE) {
											auto* navigationContext = reinterpret_cast<ViewerContext*>(
												GetWindowLongPtrW(window, GWLP_USERDATA));
											ShowFallback(navigationContext, L"日志页面加载失败，已切换到基础日志视图。\r\n");
										}
										return S_OK;
									}).Get(),
								nullptr);

							inner->controller->put_IsVisible(g_isOpen ? TRUE : FALSE);
							LayoutViewer(window, inner);
							const std::string html = LoadUtf8HtmlResourceText(IDR_HTML_IDE_LOG_VIEWER);
							const std::wstring htmlWide = Utf8ToWide(html);
							if (htmlWide.empty()) {
								ShowFallback(inner, L"日志页面资源缺失，已切换到基础日志视图。\r\n");
								return S_OK;
							}
							inner->webView->NavigateToString(htmlWide.c_str());
							return S_OK;
						}).Get());
			}).Get());

	if (FAILED(createResult)) {
		ShowFallback(context, L"WebView2 启动失败，已切换到基础日志视图。\r\n");
	}
}

void PollOutputControlFallback(ViewerContext* context)
{
	if (context == nullptr || IdeCompileOutputCapture::IsHookAvailable()) {
		return;
	}
	const ULONGLONG now = GetTickCount64();
	if (context->lastControlPollTick != 0 &&
		now - context->lastControlPollTick < kControlFallbackPollMs) {
		return;
	}
	context->lastControlPollTick = now;

	std::string currentText;
	if (!IDEFacade::Instance().GetOutputWindowText(currentText)) {
		return;
	}
	std::string delta;
	if (currentText.size() > context->lastControlText.size() &&
		currentText.compare(0, context->lastControlText.size(), context->lastControlText) == 0) {
		delta = currentText.substr(context->lastControlText.size());
	}
	else if (currentText != context->lastControlText) {
		delta = currentText;
	}
	context->lastControlText = std::move(currentText);
	if (!delta.empty()) {
		IdeLogStore::AppendLocal(
			delta.data(),
			delta.size(),
			IdeLogStore::EntrySource::OutputControlFallback);
	}
}

void PauseViewer(HWND window)
{
	if (window == nullptr || !IsWindow(window)) {
		return;
	}
	KillTimer(window, kFlushTimerId);
	auto* context = reinterpret_cast<ViewerContext*>(GetWindowLongPtrW(window, GWLP_USERDATA));
	if (context != nullptr && context->controller != nullptr) {
		context->controller->put_IsVisible(FALSE);
	}
}

void ResumeViewer(HWND window)
{
	if (window == nullptr || !IsWindow(window)) {
		return;
	}
	auto* context = reinterpret_cast<ViewerContext*>(GetWindowLongPtrW(window, GWLP_USERDATA));
	if (context == nullptr) {
		return;
	}
	context->lastSequence = 0;
	context->generation = 0;
	context->lastControlPollTick = 0;
	if (!IdeCompileOutputCapture::IsHookAvailable()) {
		IDEFacade::Instance().GetOutputWindowText(context->lastControlText);
	}
	if (context->controller != nullptr) {
		context->controller->put_IsVisible(TRUE);
		LayoutViewer(window, context);
	}
	SetTimer(window, kFlushTimerId, kFlushIntervalMs, nullptr);
}

LRESULT CALLBACK ViewerWindowProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam)
{
	auto* context = reinterpret_cast<ViewerContext*>(GetWindowLongPtrW(window, GWLP_USERDATA));
	switch (message) {
	case WM_CREATE: {
		auto* created = new (std::nothrow) ViewerContext();
		if (created == nullptr) {
			return -1;
		}
		SetWindowLongPtrW(window, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(created));
		created->hostWindow = CreateWindowExW(
			0,
			L"STATIC",
			L"",
			WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
			0,
			0,
			0,
			0,
			window,
			nullptr,
			GetCurrentModuleHandle(),
			nullptr);
		created->loadingLabel = CreateWindowExW(
			0,
			L"STATIC",
			L"正在初始化现代日志视图...",
			WS_CHILD | WS_VISIBLE,
			0,
			0,
			0,
			0,
			window,
			nullptr,
			GetCurrentModuleHandle(),
			nullptr);
		created->fallbackEdit = CreateWindowExW(
			WS_EX_CLIENTEDGE,
			L"EDIT",
			L"",
			WS_CHILD | ES_MULTILINE | ES_READONLY | ES_AUTOVSCROLL | WS_VSCROLL | WS_HSCROLL,
			0,
			0,
			0,
			0,
			window,
			nullptr,
			GetCurrentModuleHandle(),
			nullptr);
		SendMessageW(created->loadingLabel, WM_SETFONT, reinterpret_cast<WPARAM>(GetStockObject(DEFAULT_GUI_FONT)), TRUE);
		SendMessageW(created->fallbackEdit, WM_SETFONT, reinterpret_cast<WPARAM>(GetStockObject(ANSI_FIXED_FONT)), TRUE);
		SetTimer(window, kFlushTimerId, kFlushIntervalMs, nullptr);
		if (!IdeCompileOutputCapture::IsHookAvailable()) {
			IDEFacade::Instance().GetOutputWindowText(created->lastControlText);
		}
		StartWebView(window, created);
		return 0;
	}
	case WM_SIZE:
		LayoutViewer(window, context);
		return 0;
	case WM_TIMER:
		if (wParam == kFlushTimerId) {
			PollOutputControlFallback(context);
			if (context != nullptr && context->fallbackActive) {
				AppendFallbackBatch(context);
			}
			else {
				FlushWebView(context);
			}
			return 0;
		}
		break;
	case WM_DESTROY:
		KillTimer(window, kFlushTimerId);
		if (context != nullptr && context->controller != nullptr) {
			context->controller->Close();
		}
		return 0;
	case WM_NCDESTROY:
		IdeLogStore::EndRecording();
		SetWindowLongPtrW(window, GWLP_USERDATA, 0);
		delete context;
		if (g_viewerWindow == window) {
			g_viewerWindow = nullptr;
			g_isOpen = false;
			g_hiddenTabState = {};
		}
		return DefWindowProcW(window, message, wParam, lParam);
	default:
		break;
	}
	return DefWindowProcW(window, message, wParam, lParam);
}

bool RegisterViewerWindowClass()
{
	WNDCLASSEXW windowClass = {};
	windowClass.cbSize = sizeof(windowClass);
	windowClass.lpfnWndProc = ViewerWindowProc;
	windowClass.hInstance = GetCurrentModuleHandle();
	windowClass.hCursor = LoadCursor(nullptr, IDC_ARROW);
	windowClass.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
	windowClass.lpszClassName = kViewerWindowClass;
	if (RegisterClassExW(&windowClass) != 0) {
		return true;
	}
	return GetLastError() == ERROR_CLASS_ALREADY_EXISTS;
}

} // namespace

bool Initialize(HWND mainWindow)
{
	if (mainWindow == nullptr || !IsWindow(mainWindow)) {
		return false;
	}
	if (g_viewerWindow != nullptr && IsWindow(g_viewerWindow)) {
		if (g_isOpen) {
			return true;
		}
		IdeLogStore::BeginRecording();
		const std::string caption = WideToLocal(L"日志中心");
		if (!IdeOutputTabController::RestoreHiddenTab(
			g_viewerWindow,
			caption,
			g_hiddenTabState)) {
			IdeLogStore::EndRecording();
			Logger::Instance().Write("IdeLogViewer", "failed to restore hidden log tab");
			return false;
		}
		ResumeViewer(g_viewerWindow);
		g_isOpen = true;
		Logger::Instance().Write("IdeLogViewer", "restored registered log viewer");
		OutputStringToELog("日志中心已开始记录");
		return true;
	}
	if (!RegisterViewerWindowClass()) {
		Logger::Instance().Write("IdeLogViewer", "failed to register viewer window class");
		return false;
	}

	IdeLogStore::BeginRecording();
	g_mainWindow = mainWindow;
	g_viewerWindow = CreateWindowExW(
		WS_EX_CONTROLPARENT,
		kViewerWindowClass,
		L"",
		WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
		0,
		0,
		960,
		640,
		mainWindow,
		nullptr,
		GetCurrentModuleHandle(),
		nullptr);
	if (g_viewerWindow == nullptr) {
		IdeLogStore::EndRecording();
		Logger::Instance().Write("IdeLogViewer", "failed to create viewer window");
		return false;
	}

	const std::string caption = WideToLocal(L"日志中心");
	const std::string toolTip = WideToLocal(L"AutoLinker 高性能日志查看器");
	if (!IDEFacade::Instance().AddOutputTab(g_viewerWindow, caption, toolTip, nullptr)) {
		DestroyWindow(g_viewerWindow);
		g_viewerWindow = nullptr;
		IdeLogStore::EndRecording();
		Logger::Instance().Write("IdeLogViewer", "FN_ADD_TAB failed");
		return false;
	}
	g_hiddenTabState = {};
	g_isOpen = true;
	Logger::Instance().Write("IdeLogViewer", "FN_ADD_TAB registered log viewer");
	OutputStringToELog("日志中心已开始记录");
	return true;
}

void Close()
{
	if (!g_isOpen || g_viewerWindow == nullptr || !IsWindow(g_viewerWindow)) {
		return;
	}
	const std::string caption = WideToLocal(L"日志中心");
	if (g_mainWindow == nullptr || !IsWindow(g_mainWindow) ||
		!IdeOutputTabController::HideTabByCaption(
			g_mainWindow,
			g_viewerWindow,
			caption,
			g_hiddenTabState)) {
		Logger::Instance().Write("IdeLogViewer", "failed to hide log tab; viewer remains open");
		return;
	}
	PauseViewer(g_viewerWindow);
	ShowWindow(g_viewerWindow, SW_HIDE);
	IdeLogStore::EndRecording();
	g_isOpen = false;
	Logger::Instance().Write("IdeLogViewer", "paused and hid registered log viewer");
}

void Shutdown()
{
	g_isOpen = false;
	IdeLogStore::EndRecording();
	if (g_viewerWindow != nullptr && IsWindow(g_viewerWindow)) {
		DestroyWindow(g_viewerWindow);
	}
	g_viewerWindow = nullptr;
	g_mainWindow = nullptr;
	g_hiddenTabState = {};
}

bool IsOpen()
{
	return g_isOpen && g_viewerWindow != nullptr && IsWindow(g_viewerWindow);
}

std::string BuildSelfTestJson()
{
	const std::string html = LoadUtf8HtmlResourceText(IDR_HTML_IDE_LOG_VIEWER);
	const bool resourceLoaded = !html.empty();
	const bool virtualListPresent = html.find("virtual-spacer") != std::string::npos &&
		html.find("renderVirtualRows") != std::string::npos;
	const bool plainSearchPresent = html.find("search-input") != std::string::npos;
	const bool regexSearchPresent = html.find("regex-toggle") != std::string::npos &&
		html.find("new RegExp") != std::string::npos;
	const bool webMessageBridgePresent = html.find("chrome.webview") != std::string::npos;
	const bool ok = resourceLoaded && virtualListPresent && plainSearchPresent &&
		regexSearchPresent && webMessageBridgePresent;
	return nlohmann::json({
		{"name", "ide-log-viewer-resource"},
		{"ok", ok},
		{"resource_loaded", resourceLoaded},
		{"virtual_list_present", virtualListPresent},
		{"plain_search_present", plainSearchPresent},
		{"regex_search_present", regexSearchPresent},
		{"web_message_bridge_present", webMessageBridgePresent}
	}).dump();
}

} // namespace IdeLogViewer
