#include <cstdlib>
#include <algorithm>
#include <chrono>
#include <filesystem>
#include <fstream>
#include <format>
#include <iostream>
#include <string>
#include <thread>
#include <vector>
#include <Windows.h>
#include <winhttp.h>

#include "..\\thirdparty\\json.hpp"

#include "..\\src\\AutoLinkerTestApi.h"

#pragma comment(lib, "winhttp.lib")

namespace {

void PrintUsage();

struct HeadlessLauncherOptions {
	std::string eExePath;
	std::string projectPath;
	std::string outputPath;
	std::string resultPath;
	std::string target = "auto";
	bool staticCompile = false;
	bool hideWindow = true;
	int timeoutSeconds = 120;
};

struct CapturedDialog {
	std::string kind = "info";
	std::string caption;
	std::string text;
	std::vector<std::string> listItems;
	std::vector<std::string> supportLibraries;
};

struct DialogEnumContext {
	DWORD processId = 0;
	std::vector<CapturedDialog>* dialogs = nullptr;
};

std::wstring AnsiToWide(const std::string& text)
{
	if (text.empty()) {
		return std::wstring();
	}
	const int size = MultiByteToWideChar(CP_ACP, 0, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (size <= 0) {
		return std::wstring();
	}
	std::wstring wide(static_cast<size_t>(size), L'\0');
	if (MultiByteToWideChar(CP_ACP, 0, text.data(), static_cast<int>(text.size()), wide.data(), size) <= 0) {
		return std::wstring();
	}
	return wide;
}

std::string WideToUtf8(const std::wstring& text)
{
	if (text.empty()) {
		return std::string();
	}
	const int size = WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
	if (size <= 0) {
		return std::string();
	}
	std::string utf8(static_cast<size_t>(size), '\0');
	if (WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), utf8.data(), size, nullptr, nullptr) <= 0) {
		return std::string();
	}
	return utf8;
}

std::wstring Utf8ToWide(const std::string& text)
{
	if (text.empty()) {
		return std::wstring();
	}
	const int size = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (size <= 0) {
		return AnsiToWide(text);
	}
	std::wstring wide(static_cast<size_t>(size), L'\0');
	if (MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), wide.data(), size) <= 0) {
		return AnsiToWide(text);
	}
	return wide;
}

std::wstring QuoteCommandLineArgWide(const std::wstring& arg)
{
	std::wstring quoted = L"\"";
	for (wchar_t ch : arg) {
		if (ch == L'"') {
			quoted += L"\\\"";
		}
		else {
			quoted.push_back(ch);
		}
	}
	quoted += L"\"";
	return quoted;
}

std::filesystem::path MakePathFromText(const std::string& text)
{
	const std::wstring wide = Utf8ToWide(text);
	if (!wide.empty()) {
		return std::filesystem::path(wide);
	}
	return std::filesystem::path(text);
}

std::string DefaultHeadlessResultPath(const std::string& outputPath)
{
	const std::filesystem::path path = MakePathFromText(outputPath);
	if (path.empty()) {
		return "autolinker-headless-result.json";
	}
	return WideToUtf8(path.wstring() + L".headless.json");
}

std::filesystem::path GetHeadlessRequestPath(const std::string& eExePath)
{
	const std::filesystem::path exePath = MakePathFromText(eExePath);
	const std::filesystem::path basePath = exePath.parent_path().empty()
		? std::filesystem::current_path()
		: exePath.parent_path();
	return basePath / "AutoLinker" / "Log" / "headless_compile_request.json";
}

std::wstring GetWindowTextWideLocal(HWND hWnd)
{
	if (hWnd == nullptr || !IsWindow(hWnd)) {
		return std::wstring();
	}
	const int len = GetWindowTextLengthW(hWnd);
	if (len <= 0) {
		return std::wstring();
	}
	std::wstring text(static_cast<size_t>(len) + 1, L'\0');
	const int copied = GetWindowTextW(hWnd, text.data(), len + 1);
	if (copied <= 0) {
		return std::wstring();
	}
	text.resize(static_cast<size_t>(copied));
	return text;
}

std::wstring GetWindowClassWideLocal(HWND hWnd)
{
	wchar_t className[128] = {};
	if (hWnd == nullptr || !IsWindow(hWnd)) {
		return std::wstring();
	}
	if (GetClassNameW(hWnd, className, static_cast<int>(sizeof(className) / sizeof(className[0]))) <= 0) {
		return std::wstring();
	}
	return className;
}

struct ChildTextContext {
	std::vector<std::wstring> texts;
	std::vector<std::wstring> listItems;
};

std::vector<std::wstring> ReadListBoxItems(HWND hWnd)
{
	std::vector<std::wstring> items;
	const LRESULT count = SendMessageW(hWnd, LB_GETCOUNT, 0, 0);
	if (count <= 0 || count == LB_ERR) {
		return items;
	}

	for (int i = 0; i < count; ++i) {
		const LRESULT len = SendMessageW(hWnd, LB_GETTEXTLEN, static_cast<WPARAM>(i), 0);
		if (len == LB_ERR) {
			continue;
		}
		std::wstring item(static_cast<size_t>(len) + 1, L'\0');
		const LRESULT copied = SendMessageW(hWnd, LB_GETTEXT, static_cast<WPARAM>(i), reinterpret_cast<LPARAM>(item.data()));
		if (copied == LB_ERR || copied <= 0) {
			continue;
		}
		item.resize(static_cast<size_t>(copied));
		if (std::find(items.begin(), items.end(), item) == items.end()) {
			items.push_back(std::move(item));
		}
	}
	return items;
}

BOOL CALLBACK EnumChildTextProc(HWND hWnd, LPARAM lParam)
{
	auto* ctx = reinterpret_cast<ChildTextContext*>(lParam);
	if (ctx == nullptr || hWnd == nullptr || !IsWindow(hWnd)) {
		return TRUE;
	}
	const std::wstring className = GetWindowClassWideLocal(hWnd);
	if (_wcsicmp(className.c_str(), L"ListBox") == 0) {
		for (auto& item : ReadListBoxItems(hWnd)) {
			if (!item.empty() && std::find(ctx->listItems.begin(), ctx->listItems.end(), item) == ctx->listItems.end()) {
				ctx->listItems.push_back(std::move(item));
			}
		}
		return TRUE;
	}
	if (_wcsicmp(className.c_str(), L"Static") != 0 &&
		_wcsicmp(className.c_str(), L"Edit") != 0) {
		return TRUE;
	}
	std::wstring text = GetWindowTextWideLocal(hWnd);
	if (!text.empty() && std::find(ctx->texts.begin(), ctx->texts.end(), text) == ctx->texts.end()) {
		ctx->texts.push_back(std::move(text));
	}
	return TRUE;
}

std::wstring CollectDialogText(HWND hWnd, std::vector<std::wstring>& listItems)
{
	ChildTextContext ctx;
	EnumChildWindows(hWnd, EnumChildTextProc, reinterpret_cast<LPARAM>(&ctx));
	listItems = ctx.listItems;
	std::wstring merged;
	for (const auto& text : ctx.texts) {
		if (text.empty()) {
			continue;
		}
		if (!merged.empty()) {
			merged += L"\r\n";
		}
		merged += text;
	}
	for (const auto& item : ctx.listItems) {
		if (item.empty() || std::find(ctx.texts.begin(), ctx.texts.end(), item) != ctx.texts.end()) {
			continue;
		}
		if (!merged.empty()) {
			merged += L"\r\n";
		}
		merged += item;
	}
	return merged;
}

bool HasEcomModuleListItem(const std::vector<std::wstring>& listItems)
{
	for (const auto& item : listItems) {
		if (item.find(L".ec") != std::wstring::npos || item.find(L".EC") != std::wstring::npos) {
			return true;
		}
	}
	return false;
}

bool TextLooksLikeMissingEcomModule(const std::wstring& text)
{
	return text.find(L"易模块文件已经无法找到") != std::wstring::npos ||
		text.find(L"模块文件已经无法找到") != std::wstring::npos ||
		(text.find(L"易模块文件") != std::wstring::npos && text.find(L"不存在") != std::wstring::npos) ||
		text.find(L".ec") != std::wstring::npos ||
		text.find(L".EC") != std::wstring::npos;
}

bool TextLooksLikeMissingSupportLibrary(const std::wstring& text)
{
	return text.find(L"不能载入支持库") != std::wstring::npos ||
		text.find(L"支持库不能被载入") != std::wstring::npos;
}

std::vector<std::wstring> ExtractMissingSupportLibraries(const std::wstring& text)
{
	std::vector<std::wstring> rows;
	size_t begin = 0;
	while (begin <= text.size()) {
		size_t end = text.find_first_of(L"\r\n", begin);
		if (end == std::wstring::npos) {
			end = text.size();
		}
		std::wstring line = text.substr(begin, end - begin);
		if (line.find(L"不能载入支持库") != std::wstring::npos &&
			std::find(rows.begin(), rows.end(), line) == rows.end()) {
			rows.push_back(std::move(line));
		}
		begin = end + 1;
		while (begin < text.size() && (text[begin] == L'\r' || text[begin] == L'\n')) {
			++begin;
		}
		if (end == text.size()) {
			break;
		}
	}
	if (rows.empty() && TextLooksLikeMissingSupportLibrary(text)) {
		rows.push_back(text);
	}
	return rows;
}

std::string ClassifyDialog(const std::wstring& text, const std::vector<std::wstring>& listItems)
{
	if (TextLooksLikeMissingSupportLibrary(text)) {
		return "missing_support_library";
	}
	if (TextLooksLikeMissingEcomModule(text) || HasEcomModuleListItem(listItems)) {
		return "missing_ecom_module";
	}
	return "info";
}

bool ContainsDialog(const std::vector<CapturedDialog>& dialogs, const CapturedDialog& dialog)
{
	return std::find_if(dialogs.begin(), dialogs.end(), [&dialog](const CapturedDialog& item) {
		return item.kind == dialog.kind && item.caption == dialog.caption && item.text == dialog.text;
	}) != dialogs.end();
}

bool HasBlockingDialog(const std::vector<CapturedDialog>& dialogs)
{
	return std::find_if(dialogs.begin(), dialogs.end(), [](const CapturedDialog& dialog) {
		return dialog.kind != "info";
	}) != dialogs.end();
}

BOOL CALLBACK EnumLauncherDialogProc(HWND hWnd, LPARAM lParam)
{
	auto* ctx = reinterpret_cast<DialogEnumContext*>(lParam);
	if (ctx == nullptr || ctx->dialogs == nullptr || hWnd == nullptr || !IsWindow(hWnd) || !IsWindowVisible(hWnd)) {
		return TRUE;
	}
	DWORD pid = 0;
	GetWindowThreadProcessId(hWnd, &pid);
	if (pid != ctx->processId || _wcsicmp(GetWindowClassWideLocal(hWnd).c_str(), L"#32770") != 0) {
		return TRUE;
	}
	CapturedDialog dialog;
	std::vector<std::wstring> listItems;
	const std::wstring text = CollectDialogText(hWnd, listItems);
	dialog.kind = ClassifyDialog(text, listItems);
	dialog.caption = WideToUtf8(GetWindowTextWideLocal(hWnd));
	dialog.text = WideToUtf8(text);
	for (const auto& item : listItems) {
		dialog.listItems.push_back(WideToUtf8(item));
	}
	for (const auto& item : ExtractMissingSupportLibraries(text)) {
		dialog.supportLibraries.push_back(WideToUtf8(item));
	}
	if (!ContainsDialog(*ctx->dialogs, dialog)) {
		ctx->dialogs->push_back(dialog);
	}
	PostMessageW(hWnd, WM_COMMAND, static_cast<WPARAM>(IDCANCEL), 0);
	PostMessageW(hWnd, WM_CLOSE, 0, 0);
	return TRUE;
}

void CaptureAndDismissProcessDialogs(DWORD processId, std::vector<CapturedDialog>& dialogs)
{
	DialogEnumContext ctx;
	ctx.processId = processId;
	ctx.dialogs = &dialogs;
	EnumWindows(EnumLauncherDialogProc, reinterpret_cast<LPARAM>(&ctx));
}

bool WriteTextFile(const std::filesystem::path& path, const std::string& text)
{
	std::error_code ec;
	const std::filesystem::path parent = path.parent_path();
	if (!parent.empty()) {
		std::filesystem::create_directories(parent, ec);
	}
	std::ofstream out(path, std::ios::binary | std::ios::trunc);
	if (!out.is_open()) {
		return false;
	}
	out.write(text.data(), static_cast<std::streamsize>(text.size()));
	return true;
}

std::string ReadTextFile(const std::filesystem::path& path)
{
	std::ifstream in(path, std::ios::binary);
	if (!in.is_open()) {
		return std::string();
	}
	return std::string(std::istreambuf_iterator<char>(in), std::istreambuf_iterator<char>());
}

std::wstring Utf8ToWideStrict(const std::string& text)
{
	if (text.empty()) {
		return std::wstring();
	}
	const int size = MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (size <= 0) {
		return std::wstring();
	}
	std::wstring wide(static_cast<size_t>(size), L'\0');
	if (MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), wide.data(), size) <= 0) {
		return std::wstring();
	}
	return wide;
}

struct SimpleHttpResult {
	bool ok = false;
	int httpStatus = 0;
	std::string bodyUtf8;
	std::string error;
};

bool IsValidUtf8(const std::string& text)
{
	if (text.empty()) {
		return true;
	}
	return MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0) > 0;
}

std::string LocalToUtf8(const std::string& text)
{
	if (text.empty() || IsValidUtf8(text)) {
		return text;
	}

	const int wideLen = MultiByteToWideChar(CP_ACP, 0, text.data(), static_cast<int>(text.size()), nullptr, 0);
	if (wideLen <= 0) {
		return text;
	}

	std::wstring wide(static_cast<size_t>(wideLen), L'\0');
	if (MultiByteToWideChar(CP_ACP, 0, text.data(), static_cast<int>(text.size()), wide.data(), wideLen) <= 0) {
		return text;
	}

	const int utf8Len = WideCharToMultiByte(CP_UTF8, 0, wide.data(), wideLen, nullptr, 0, nullptr, nullptr);
	if (utf8Len <= 0) {
		return text;
	}

	std::string utf8(static_cast<size_t>(utf8Len), '\0');
	if (WideCharToMultiByte(CP_UTF8, 0, wide.data(), wideLen, utf8.data(), utf8Len, nullptr, nullptr) <= 0) {
		return text;
	}
	return utf8;
}

size_t ClampUtf8PrefixBoundary(const std::string& text, size_t maxBytes)
{
	size_t end = (std::min)(maxBytes, text.size());
	while (end > 0 && end < text.size() &&
		(static_cast<unsigned char>(text[end]) & 0xC0) == 0x80) {
		--end;
	}
	return end;
}

std::string TruncateUtf8(const std::string& text, size_t maxBytes)
{
	if (text.size() <= maxBytes) {
		return text;
	}
	return text.substr(0, ClampUtf8PrefixBoundary(text, maxBytes));
}

std::string DumpJsonCompactSafe(const nlohmann::json& value)
{
	return value.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace);
}

std::string DumpJsonPrettySafe(const nlohmann::json& value)
{
	return value.dump(2, ' ', false, nlohmann::json::error_handler_t::replace);
}

void AppendTraceLine(const std::filesystem::path& path, const std::string& line)
{
	std::error_code ec;
	const std::filesystem::path parent = path.parent_path();
	if (!parent.empty()) {
		std::filesystem::create_directories(parent, ec);
	}

	std::ofstream out(path, std::ios::binary | std::ios::app);
	if (!out.is_open()) {
		return;
	}
	out << line << "\r\n";
}

bool ParseHttpUrl(const std::string& urlUtf8, std::wstring& host, std::wstring& pathAndQuery, INTERNET_PORT& port, bool& useHttps)
{
	useHttps = false;
	port = INTERNET_DEFAULT_HTTPS_PORT;
	const std::wstring url = Utf8ToWideStrict(urlUtf8);
	if (url.empty()) {
		return false;
	}
	URL_COMPONENTSW uc = {};
	uc.dwStructSize = sizeof(uc);
	uc.dwSchemeLength = static_cast<DWORD>(-1);
	uc.dwHostNameLength = static_cast<DWORD>(-1);
	uc.dwUrlPathLength = static_cast<DWORD>(-1);
	uc.dwExtraInfoLength = static_cast<DWORD>(-1);
	std::wstring mutableUrl = url;
	if (!WinHttpCrackUrl(mutableUrl.data(), static_cast<DWORD>(mutableUrl.size()), 0, &uc)) {
		return false;
	}
	useHttps = uc.nScheme == INTERNET_SCHEME_HTTPS;
	port = uc.nPort;
	if (uc.lpszHostName != nullptr && uc.dwHostNameLength > 0) {
		host.assign(uc.lpszHostName, uc.dwHostNameLength);
	}
	if (uc.lpszUrlPath != nullptr && uc.dwUrlPathLength > 0) {
		pathAndQuery.assign(uc.lpszUrlPath, uc.dwUrlPathLength);
	}
	if (uc.lpszExtraInfo != nullptr && uc.dwExtraInfoLength > 0) {
		pathAndQuery.append(uc.lpszExtraInfo, uc.dwExtraInfoLength);
	}
	if (pathAndQuery.empty()) {
		pathAndQuery = L"/";
	}
	return !host.empty();
}

SimpleHttpResult PerformJsonRequest(
	const std::string& methodUtf8,
	const std::string& urlUtf8,
	const std::string& bodyUtf8,
	const std::vector<std::pair<std::string, std::string>>& headers)
{
	SimpleHttpResult result;
	std::wstring host;
	std::wstring pathAndQuery;
	INTERNET_PORT port = INTERNET_DEFAULT_HTTPS_PORT;
	bool useHttps = false;
	if (!ParseHttpUrl(urlUtf8, host, pathAndQuery, port, useHttps)) {
		result.error = "parse url failed";
		return result;
	}

	const std::wstring methodWide = Utf8ToWideStrict(methodUtf8);
	HINTERNET hSession = WinHttpOpen(L"AutoLinkerTest/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
	if (hSession == nullptr) {
		result.error = "WinHttpOpen failed";
		return result;
	}
	WinHttpSetTimeouts(hSession, 30000, 30000, 180000, 180000);

	HINTERNET hConnect = WinHttpConnect(hSession, host.c_str(), port, 0);
	if (hConnect == nullptr) {
		result.error = "WinHttpConnect failed";
		WinHttpCloseHandle(hSession);
		return result;
	}
	HINTERNET hRequest = WinHttpOpenRequest(
		hConnect,
		methodWide.c_str(),
		pathAndQuery.c_str(),
		nullptr,
		WINHTTP_NO_REFERER,
		WINHTTP_DEFAULT_ACCEPT_TYPES,
		useHttps ? WINHTTP_FLAG_SECURE : 0);
	if (hRequest == nullptr) {
		result.error = "WinHttpOpenRequest failed";
		WinHttpCloseHandle(hConnect);
		WinHttpCloseHandle(hSession);
		return result;
	}

	std::wstring headerText;
	for (const auto& kv : headers) {
		headerText += Utf8ToWideStrict(kv.first + ": " + kv.second + "\r\n");
	}
	const BOOL sent = WinHttpSendRequest(
		hRequest,
		headerText.empty() ? WINHTTP_NO_ADDITIONAL_HEADERS : headerText.c_str(),
		headerText.empty() ? 0 : static_cast<DWORD>(headerText.size()),
		bodyUtf8.empty() ? WINHTTP_NO_REQUEST_DATA : const_cast<char*>(bodyUtf8.data()),
		static_cast<DWORD>(bodyUtf8.size()),
		static_cast<DWORD>(bodyUtf8.size()),
		0);
	if (!sent || !WinHttpReceiveResponse(hRequest, nullptr)) {
		result.error = "send/receive failed";
		WinHttpCloseHandle(hRequest);
		WinHttpCloseHandle(hConnect);
		WinHttpCloseHandle(hSession);
		return result;
	}

	DWORD statusCode = 0;
	DWORD size = sizeof(statusCode);
	WinHttpQueryHeaders(
		hRequest,
		WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
		WINHTTP_HEADER_NAME_BY_INDEX,
		&statusCode,
		&size,
		WINHTTP_NO_HEADER_INDEX);
	result.httpStatus = static_cast<int>(statusCode);

	std::string response;
	for (;;) {
		DWORD available = 0;
		if (!WinHttpQueryDataAvailable(hRequest, &available) || available == 0) {
			break;
		}
		std::string chunk(static_cast<size_t>(available), '\0');
		DWORD read = 0;
		if (!WinHttpReadData(hRequest, chunk.data(), available, &read) || read == 0) {
			break;
		}
		chunk.resize(static_cast<size_t>(read));
		response += chunk;
	}

	result.ok = result.httpStatus >= 200 && result.httpStatus < 300;
	result.bodyUtf8 = response;
	WinHttpCloseHandle(hRequest);
	WinHttpCloseHandle(hConnect);
	WinHttpCloseHandle(hSession);
	return result;
}

SimpleHttpResult PerformJsonPost(
	const std::string& urlUtf8,
	const std::string& bodyUtf8,
	const std::vector<std::pair<std::string, std::string>>& headers)
{
	return PerformJsonRequest("POST", urlUtf8, bodyUtf8, headers);
}

std::string PerformSimpleGetText(const std::string& urlUtf8, int maxBytes, std::string& outError)
{
	outError.clear();
	const SimpleHttpResult response = PerformJsonRequest("GET", urlUtf8, "", {});
	if (!response.ok) {
		outError = response.error.empty()
			? ("http_status=" + std::to_string(response.httpStatus))
			: response.error;
		return std::string();
	}
	if (static_cast<int>(response.bodyUtf8.size()) <= maxBytes) {
		return response.bodyUtf8;
	}
	return TruncateUtf8(response.bodyUtf8, static_cast<size_t>(maxBytes));
}

bool ExtractToolTargetUrl(const nlohmann::json& toolCall, std::string& outUrl)
{
	outUrl.clear();
	if (!toolCall.is_object() ||
		!toolCall.contains("function") ||
		!toolCall["function"].is_object() ||
		!toolCall["function"].contains("arguments") ||
		!toolCall["function"]["arguments"].is_string()) {
		return false;
	}

	try {
		const nlohmann::json args = nlohmann::json::parse(toolCall["function"]["arguments"].get<std::string>());
		if (!args.is_object() || !args.contains("url") || !args["url"].is_string()) {
			return false;
		}
		outUrl = args["url"].get<std::string>();
		return !outUrl.empty();
	}
	catch (...) {
		return false;
	}
}

void PrintHeadlessSummary(const nlohmann::json& result)
{
	const bool ok = result.value("ok", false);
	std::ostream& out = ok ? std::cout : std::cerr;
	out << (ok ? "[AutoLinkerLauncher] compile success" : "[AutoLinkerLauncher] compile failed") << std::endl;
	const std::string error = result.value("error", std::string());
	if (!ok && !error.empty()) {
		out << "error: " << error << std::endl;
	}
	if (!ok && result.contains("launcher_process_exit_code")) {
		const DWORD exitCode = result.value("launcher_process_exit_code", 0u);
		out << std::format("process_exit_code: {} (0x{:08X})", exitCode, exitCode) << std::endl;
	}
	if (result.contains("resolved_target")) {
		out << "target: " << result.value("resolved_target", std::string()) << std::endl;
	}
	if (result.contains("compile_result") && result["compile_result"].is_object()) {
		const auto& compileResult = result["compile_result"];
		out << "output: " << compileResult.value("output_path", std::string()) << std::endl;
		if (!ok) {
			out << "error_location: page=" << compileResult.value("caret_page_name", std::string())
				<< " type=" << compileResult.value("caret_page_type", std::string())
				<< " row=" << compileResult.value("caret_row", -1) << std::endl;
			const std::string line = compileResult.value("caret_line_text", std::string());
			if (!line.empty()) {
				out << "error_line: " << line << std::endl;
			}
		}
		const std::string outputText = compileResult.value("output_window_text", std::string());
		if (!outputText.empty()) {
			out << "[IDE output]" << std::endl << outputText << std::endl;
		}
	}
	if (result.contains("compile_dialogs") && result["compile_dialogs"].is_array()) {
		for (const auto& row : result["compile_dialogs"]) {
			if (row.value("type", std::string()) != "compile_output_target") {
				continue;
			}
			const std::string mode = row.value("mode", std::string());
			out << (mode == "pending"
				? "[AutoLinkerLauncher] compile output target dialog pending"
				: "[AutoLinkerLauncher] compile output target auto-suppressed")
				<< std::endl;
			const std::string target = row.value("target", std::string());
			if (!target.empty()) {
				out << "dialog_target: " << target << std::endl;
			}
			const std::string outputPath = row.value("output_path", std::string());
			if (!outputPath.empty()) {
				out << "dialog_output_path: " << outputPath << std::endl;
			}
		}
	}
	const auto printBoxes = [](const nlohmann::json& boxes, const char* label) {
		if (!boxes.is_array()) {
			return;
		}
		for (const auto& box : boxes) {
			const std::string kind = box.value("kind", std::string("info"));
			std::cerr << label;
			if (kind == "missing_ecom_module") {
				std::cerr << "[MissingModule] ";
			}
			else if (kind == "missing_support_library") {
				std::cerr << "[MissingSupportLibrary] ";
			}
			else if (kind == "info") {
				std::cerr << "[Info] ";
			}
			std::cerr << box.value("caption", std::string()) << std::endl;
			const std::string text = box.value("text", std::string());
			if (!text.empty()) {
				std::cerr << text << std::endl;
			}
			if (box.contains("support_libraries") && box["support_libraries"].is_array() && !box["support_libraries"].empty()) {
				std::cerr << "support_libraries:" << std::endl;
				for (const auto& item : box["support_libraries"]) {
					if (item.is_string()) {
						std::cerr << "  " << item.get<std::string>() << std::endl;
					}
				}
			}
			if (box.contains("list_items") && box["list_items"].is_array() && !box["list_items"].empty()) {
				std::cerr << "list_items:" << std::endl;
				for (const auto& item : box["list_items"]) {
					if (item.is_string()) {
						std::cerr << "  " << item.get<std::string>() << std::endl;
					}
				}
			}
		}
	};
	if (result.contains("launcher_message_boxes")) {
		printBoxes(result["launcher_message_boxes"], "[AutoLinkerLauncher][MessageBox] ");
	}
	if (result.contains("ide_message_boxes")) {
		printBoxes(result["ide_message_boxes"], "[AutoLinker][MessageBox] ");
	}
}

int RunHeadlessCompile(int argc, char* argv[])
{
	if (argc < 5) {
		PrintUsage();
		return EXIT_FAILURE;
	}

	HeadlessLauncherOptions options;
	options.eExePath = argv[2];
	options.projectPath = argv[3];
	options.outputPath = argv[4];
	options.resultPath = DefaultHeadlessResultPath(options.outputPath);

	for (int i = 5; i < argc; ++i) {
		const std::string arg = argv[i];
		if (arg == "--static") {
			options.staticCompile = true;
		}
		else if (arg == "--show-window") {
			options.hideWindow = false;
		}
		else if (arg == "--target" && i + 1 < argc) {
			options.target = argv[++i];
		}
		else if (arg == "--result" && i + 1 < argc) {
			options.resultPath = argv[++i];
		}
		else if (arg == "--timeout" && i + 1 < argc) {
			options.timeoutSeconds = (std::max)(1, std::atoi(argv[++i]));
		}
		else {
			std::cerr << "unknown headless-compile option: " << arg << std::endl;
			return EXIT_FAILURE;
		}
	}

	nlohmann::json request = {
		{"enabled", true},
		{"launcher_pid", GetCurrentProcessId()},
		{"launcher_name", "AutoLinkerTest.exe"},
		{"target", options.target},
		{"static_compile", options.staticCompile},
		{"output_path", options.outputPath},
		{"result_path", options.resultPath},
		{"startup_timeout_seconds", options.timeoutSeconds},
		{"hide_window", options.hideWindow},
		{"exit_after_compile", true}
	};

	const std::filesystem::path requestPath = GetHeadlessRequestPath(options.eExePath);
	if (!WriteTextFile(requestPath, request.dump(2))) {
		std::cerr << "write headless request failed: " << WideToUtf8(requestPath.wstring()) << std::endl;
		return EXIT_FAILURE;
	}

	STARTUPINFOW si = {};
	si.cb = sizeof(si);
	si.dwFlags = STARTF_USESHOWWINDOW;
	si.wShowWindow = options.hideWindow ? SW_HIDE : SW_SHOW;
	PROCESS_INFORMATION pi = {};

	const std::wstring exePath = Utf8ToWide(options.eExePath);
	const std::wstring projectPath = Utf8ToWide(options.projectPath);
	std::wstring commandLine = QuoteCommandLineArgWide(exePath) + L" " + QuoteCommandLineArgWide(projectPath);
	const std::filesystem::path projectFsPath = MakePathFromText(options.projectPath);
	const std::wstring workingDir = projectFsPath.parent_path().empty()
		? std::filesystem::current_path().wstring()
		: projectFsPath.parent_path().wstring();

	std::vector<wchar_t> mutableCommandLine(commandLine.begin(), commandLine.end());
	mutableCommandLine.push_back(L'\0');
	const BOOL created = CreateProcessW(
		exePath.c_str(),
		mutableCommandLine.data(),
		nullptr,
		nullptr,
		FALSE,
		0,
		nullptr,
		workingDir.c_str(),
		&si,
		&pi);

	if (created == FALSE) {
		std::error_code ec;
		std::filesystem::remove(requestPath, ec);
		std::cerr << "CreateProcessW failed, error=" << GetLastError() << std::endl;
		return EXIT_FAILURE;
	}

	std::vector<CapturedDialog> dialogs;
	const auto deadline = std::chrono::steady_clock::now() + std::chrono::seconds(options.timeoutSeconds + 30);
	bool timedOut = false;
	for (;;) {
		const DWORD waitResult = WaitForSingleObject(pi.hProcess, 200);
		CaptureAndDismissProcessDialogs(pi.dwProcessId, dialogs);
		if (waitResult == WAIT_OBJECT_0) {
			break;
		}
		if (std::chrono::steady_clock::now() >= deadline) {
			timedOut = true;
			TerminateProcess(pi.hProcess, 5);
			break;
		}
	}

	DWORD processExitCode = 0;
	GetExitCodeProcess(pi.hProcess, &processExitCode);
	CloseHandle(pi.hThread);
	CloseHandle(pi.hProcess);
	{
		std::error_code ec;
		std::filesystem::remove(requestPath, ec);
	}

	const std::filesystem::path resultPath = MakePathFromText(options.resultPath);
	nlohmann::json result;
	const std::string resultText = ReadTextFile(resultPath);
	const bool resultLoadedFromFile = !resultText.empty();
	if (!resultText.empty()) {
		try {
			result = nlohmann::json::parse(resultText);
		}
		catch (const std::exception& ex) {
			result = {
				{"ok", false},
				{"error", std::string("parse_headless_result_failed: ") + ex.what()},
				{"raw_result", resultText}
			};
		}
	}
	else {
		result = {
			{"ok", false},
			{"mode", "autolinker_headless_launcher"},
			{"error", timedOut ? "headless_process_timeout" : "headless_result_missing"},
			{"process_exit_code", processExitCode},
			{"request", request},
			{"request_file", WideToUtf8(requestPath.wstring())}
		};
	}

	if (!dialogs.empty()) {
		nlohmann::json rows = nlohmann::json::array();
		for (const auto& dialog : dialogs) {
			nlohmann::json row = {
				{"kind", dialog.kind},
				{"caption", dialog.caption},
				{"text", dialog.text}
			};
			if (!dialog.listItems.empty()) {
				row["list_items"] = dialog.listItems;
			}
			if (!dialog.supportLibraries.empty()) {
				row["support_libraries"] = dialog.supportLibraries;
			}
			rows.push_back(std::move(row));
		}
		result["launcher_message_boxes"] = std::move(rows);
		if (HasBlockingDialog(dialogs) && !result.value("ok", false) && !timedOut && (!resultLoadedFromFile || !result.contains("error"))) {
			result["error"] = "ide_startup_message_box_captured";
		}
	}
	result["launcher_process_exit_code"] = processExitCode;
	result["launcher_timed_out"] = timedOut;
	result["launcher_request_file"] = WideToUtf8(requestPath.wstring());
	WriteTextFile(resultPath, result.dump(2));
	PrintHeadlessSummary(result);
	std::cout << "result: " << options.resultPath << std::endl;

	if (result.value("ok", false)) {
		return EXIT_SUCCESS;
	}
	return result.value("exit_code", timedOut ? 5 : 1);
}

std::string JoinArguments(int startIndex, int argc, char* argv[])
{
	std::string text;
	for (int i = startIndex; i < argc; ++i) {
		if (!text.empty()) {
			text.push_back(' ');
		}
		text.append(argv[i]);
	}
	return text;
}

int PrintStringResult(const char* label, int result, const char* text)
{
	if (result >= 0) {
		std::cout << label << ": " << text << std::endl;
		return EXIT_SUCCESS;
	}

	if (result == AUTOLINKER_TEST_STRING_BUFFER_TOO_SMALL) {
		std::cerr << label << " failed: buffer too small" << std::endl;
	}
	else {
		std::cerr << label << " failed: invalid argument" << std::endl;
	}
	return EXIT_FAILURE;
}

int RunSmokeTest()
{
	char buffer[512] = {};
	int compareResult = 0;

	if (!AutoLinkerTest_CompareVersion("1.2.3", "1.2.0", &compareResult)) {
		std::cerr << "version-compare failed" << std::endl;
		return EXIT_FAILURE;
	}

	std::cout << "version-compare: " << compareResult << std::endl;

	int result = AutoLinkerTest_GetVersionText(buffer, static_cast<int>(sizeof(buffer)));
	if (PrintStringResult("version-text", result, buffer) != EXIT_SUCCESS) {
		return EXIT_FAILURE;
	}

	const char* linkerCommand = "link /out:\"D:\\demo\\AutoLinkerTest.exe\" \"D:\\demo\\main.obj\" \"D:\\deps\\static_lib\\krnln_static.lib\"";

	result = AutoLinkerTest_GetLinkerOutFileName(linkerCommand, buffer, static_cast<int>(sizeof(buffer)));
	if (PrintStringResult("linker-out", result, buffer) != EXIT_SUCCESS) {
		return EXIT_FAILURE;
	}

	result = AutoLinkerTest_GetLinkerKrnlnFileName(linkerCommand, buffer, static_cast<int>(sizeof(buffer)));
	if (PrintStringResult("linker-krnln", result, buffer) != EXIT_SUCCESS) {
		return EXIT_FAILURE;
	}

	result = AutoLinkerTest_ExtractBetweenDashes("before - middle - after", buffer, static_cast<int>(sizeof(buffer)));
	return PrintStringResult("extract-between-dashes", result, buffer);
}

int RunVersionCompare(const std::string& left, const std::string& right)
{
	int compareResult = 0;
	if (!AutoLinkerTest_CompareVersion(left.c_str(), right.c_str(), &compareResult)) {
		std::cerr << "version-compare failed" << std::endl;
		return EXIT_FAILURE;
	}

	std::cout << compareResult << std::endl;
	return EXIT_SUCCESS;
}

int RunDeepSeekModelIntegrationTest(int argc, char* argv[])
{
	if (argc < 4) {
		PrintUsage();
		return EXIT_FAILURE;
	}

	const char* apiKey = argv[2];
	const char* model = argv[3];
	const char* baseUrl = "https://api.deepseek.com";
	std::string outputPath;
	for (int i = 4; i < argc; ++i) {
		const std::string arg = argv[i];
		if (arg == "--out" && i + 1 < argc) {
			outputPath = argv[++i];
		}
		else if (std::string(baseUrl) == "https://api.deepseek.com") {
			baseUrl = argv[i];
		}
		else {
			PrintUsage();
			return EXIT_FAILURE;
		}
	}

	const std::filesystem::path tracePath = std::filesystem::current_path() / "temp" / "deepseek_run_trace.log";
	AppendTraceLine(tracePath, "run:start");

	try {
		auto postChat = [&](const nlohmann::json& messages) -> nlohmann::json {
			AppendTraceLine(tracePath, "post_chat:begin");
			nlohmann::json body = {
				{"model", LocalToUtf8(std::string(model))},
				{"stream", false},
				{"temperature", 0},
				{"messages", messages},
				{"thinking", {
					{"type", "enabled"}
				}},
				{"reasoning_effort", "high"},
				{"tools", nlohmann::json::array({
					{
						{"type", "function"},
						{"function", {
							{"name", "fetch_url"},
							{"description", "Fetch a URL and return plain text body."},
							{"parameters", {
								{"type", "object"},
								{"properties", {
									{"url", {{"type", "string"}}}
								}},
								{"required", nlohmann::json::array({"url"})},
								{"additionalProperties", false}
							}}
						}}
					},
					{
						{"type", "function"},
						{"function", {
							{"name", "extract_web_document"},
							{"description", "Fetch a URL and return a shortened readable text excerpt."},
							{"parameters", {
								{"type", "object"},
								{"properties", {
									{"url", {{"type", "string"}}}
								}},
								{"required", nlohmann::json::array({"url"})},
								{"additionalProperties", false}
							}}
						}}
					}
				})}
			};
			const std::string requestBody = DumpJsonCompactSafe(body);
			AppendTraceLine(tracePath, "post_chat:request_ready bytes=" + std::to_string(requestBody.size()));
			SimpleHttpResult response = PerformJsonPost(
				std::string(baseUrl) + "/chat/completions",
				requestBody,
				{
					{"Content-Type", "application/json"},
					{"Authorization", std::string("Bearer ") + apiKey}
				});
			AppendTraceLine(
				tracePath,
				"post_chat:response status=" + std::to_string(response.httpStatus) +
				" ok=" + std::to_string(response.ok ? 1 : 0) +
				" body_bytes=" + std::to_string(response.bodyUtf8.size()));
			nlohmann::json r = {
				{"http_status", response.httpStatus},
				{"ok", response.ok},
				{"error", response.error},
				{"body_utf8", response.bodyUtf8}
			};
			if (!response.bodyUtf8.empty()) {
				try {
					r["parsed"] = nlohmann::json::parse(response.bodyUtf8);
				}
				catch (const std::exception& ex) {
					r["parse_error"] = ex.what();
				}
			}
			return r;
		};

		nlohmann::json report = {
			{"provider", "deepseek"},
			{"model", model},
			{"base_url", baseUrl}
		};

		AppendTraceLine(tracePath, "step:test_connection");
		const nlohmann::json connectionMessages = nlohmann::json::array({
			{{"role", "system"}, {"content", "You are a connectivity test assistant. Reply with OK only."}},
			{{"role", "user"}, {"content", "Reply with OK only."}}
		});
		const nlohmann::json connectionResult = postChat(connectionMessages);
		report["test_connection"] = connectionResult;
		if (!connectionResult.value("ok", false)) {
			report["ok"] = false;
			const std::string text = DumpJsonPrettySafe(report);
			std::cout << text << std::endl;
			return EXIT_FAILURE;
		}

		AppendTraceLine(tracePath, "step:simple_task");
		const nlohmann::json simpleMessages = nlohmann::json::array({
			{{"role", "user"}, {"content", LocalToUtf8("只回答这四个汉字：测试通过")}}
		});
		const nlohmann::json simpleTaskResult = postChat(simpleMessages);
		report["simple_task"] = simpleTaskResult;
		if (!simpleTaskResult.value("ok", false)) {
			report["ok"] = false;
			const std::string text = DumpJsonPrettySafe(report);
			std::cout << text << std::endl;
			return EXIT_FAILURE;
		}

		AppendTraceLine(tracePath, "step:tool_chat");
		nlohmann::json toolMessages = nlohmann::json::array({
			{{"role", "user"}, {"content", LocalToUtf8("You must call fetch_url on https://api-docs.deepseek.com/quick_start/rate_limit and then call extract_web_document on https://api-docs.deepseek.com/guides/thinking_mode . After both tools finish, reply in one Chinese line exactly in this format: 限速页已读；思考页已读；工具数=N。")}}
		});

		nlohmann::json toolEvents = nlohmann::json::array();
		nlohmann::json finalAssistantMessage = nlohmann::json::object();
		bool toolChatOk = false;
		int toolRoundCount = 0;
		for (; toolRoundCount < 8; ++toolRoundCount) {
			AppendTraceLine(tracePath, "step:tool_chat_round" + std::to_string(toolRoundCount + 1));
			const nlohmann::json roundResult = postChat(toolMessages);
			report[std::format("tool_chat_round{}", toolRoundCount + 1)] = roundResult;
			if (!roundResult.value("ok", false)) {
				break;
			}

			nlohmann::json assistantMessage = nlohmann::json::object();
			if (roundResult.contains("parsed")) {
				const auto& parsed = roundResult["parsed"];
				if (parsed.contains("choices") && parsed["choices"].is_array() && !parsed["choices"].empty()) {
					assistantMessage = parsed["choices"][0].value("message", nlohmann::json::object());
				}
			}
			if (assistantMessage.contains("content") && assistantMessage["content"].is_null()) {
				assistantMessage["content"] = "";
			}
			if (assistantMessage.contains("reasoning_content") && assistantMessage["reasoning_content"].is_null()) {
				assistantMessage["reasoning_content"] = "";
			}

			finalAssistantMessage = assistantMessage;
			if (!assistantMessage.contains("tool_calls") || !assistantMessage["tool_calls"].is_array() || assistantMessage["tool_calls"].empty()) {
				toolChatOk = true;
				break;
			}

			toolMessages.push_back(assistantMessage);
			for (const auto& toolCall : assistantMessage["tool_calls"]) {
				if (!toolCall.is_object()) {
					continue;
				}

				const std::string toolName =
					toolCall.contains("function") && toolCall["function"].contains("name") && toolCall["function"]["name"].is_string()
					? toolCall["function"]["name"].get<std::string>()
					: std::string();
				const std::string callId = toolCall.value("id", std::string());
				std::string targetUrl;
				ExtractToolTargetUrl(toolCall, targetUrl);

				std::string toolResultText;
				std::string toolError;
				if (toolName == "fetch_url") {
					toolResultText = PerformSimpleGetText(targetUrl, 262144, toolError);
				}
				else if (toolName == "extract_web_document") {
					toolResultText = PerformSimpleGetText(targetUrl, 32768, toolError);
					if (toolResultText.size() > 4000) {
						toolResultText = TruncateUtf8(toolResultText, 4000);
					}
				}
				else {
					toolError = "unsupported tool";
				}

				nlohmann::json toolPayload = {
					{"ok", toolError.empty()},
					{"tool_name", toolName},
					{"url", targetUrl},
					{"text", toolResultText},
					{"error", toolError}
				};
				toolEvents.push_back(toolPayload);
				toolMessages.push_back({
					{"role", "tool"},
					{"tool_call_id", callId},
					{"name", toolName},
					{"content", DumpJsonCompactSafe(toolPayload)}
				});
			}
		}
		report["tool_events"] = toolEvents;
		report["tool_chat_ok"] = toolChatOk;
		report["tool_chat_round_count"] = toolRoundCount + (toolChatOk ? 1 : 0);

		AppendTraceLine(tracePath, "step:followup_chat");
		nlohmann::json followupMessages = toolMessages;
		if (!finalAssistantMessage.empty()) {
			followupMessages.push_back(finalAssistantMessage);
		}
		followupMessages.push_back({
			{"role", "user"},
			{"content", LocalToUtf8("只回答：上一轮你实际调用了几个工具？输出阿拉伯数字。")}
		});
		const nlohmann::json followupResult = postChat(followupMessages);
		report["followup_chat"] = followupResult;

		bool reasoningContentSeen = false;
		const auto scanReasoningSeen = [&reasoningContentSeen](const nlohmann::json& node) {
			if (!node.is_object()) {
				return;
			}
			if (node.contains("parsed")) {
				const auto& parsed = node["parsed"];
				if (parsed.contains("choices") && parsed["choices"].is_array() && !parsed["choices"].empty()) {
					const auto& message = parsed["choices"][0]["message"];
					if (message.contains("reasoning_content") && message["reasoning_content"].is_string() &&
						!message["reasoning_content"].get<std::string>().empty()) {
						reasoningContentSeen = true;
					}
				}
			}
		};
		for (int round = 1; round <= toolRoundCount + 1; ++round) {
			const std::string key = std::format("tool_chat_round{}", round);
			if (report.contains(key)) {
				scanReasoningSeen(report[key]);
			}
		}
		scanReasoningSeen(followupResult);
		report["reasoning_content_seen"] = reasoningContentSeen;

		report["ok"] =
			connectionResult.value("ok", false) &&
			simpleTaskResult.value("ok", false) &&
			toolChatOk &&
			followupResult.value("ok", false) &&
			toolEvents.size() >= 2;

		const std::string text = DumpJsonPrettySafe(report);
		if (!outputPath.empty()) {
			std::error_code ec;
			const std::filesystem::path outputFsPath = MakePathFromText(outputPath);
			const std::filesystem::path parent = outputFsPath.parent_path();
			if (!parent.empty()) {
				std::filesystem::create_directories(parent, ec);
			}
			std::ofstream out(outputFsPath, std::ios::binary | std::ios::trunc);
			if (!out.is_open()) {
				std::cerr << "deepseek-model-test failed: cannot open output file: " << outputPath << std::endl;
				return EXIT_FAILURE;
			}
			out.write(text.data(), static_cast<std::streamsize>(text.size()));
		}
		std::cout << text << std::endl;
		AppendTraceLine(tracePath, "run:done ok=" + std::to_string(report.value("ok", false) ? 1 : 0));
		return report.value("ok", false) ? EXIT_SUCCESS : EXIT_FAILURE;
	}
	catch (const std::exception& ex) {
		AppendTraceLine(tracePath, std::string("run:exception ") + ex.what());
		nlohmann::json errorReport = {
			{"provider", "deepseek"},
			{"model", model},
			{"base_url", baseUrl},
			{"ok", false},
			{"error", std::string("exception: ") + ex.what()}
		};
		const std::string text = DumpJsonPrettySafe(errorReport);
		std::cerr << text << std::endl;
		return EXIT_FAILURE;
	}
	catch (...) {
		AppendTraceLine(tracePath, "run:exception unknown");
		nlohmann::json errorReport = {
			{"provider", "deepseek"},
			{"model", model},
			{"base_url", baseUrl},
			{"ok", false},
			{"error", "unknown exception"}
		};
		const std::string text = DumpJsonPrettySafe(errorReport);
		std::cerr << text << std::endl;
		return EXIT_FAILURE;
	}
}

int RunOpenAIIntegrationCommand(
	int argc,
	char* argv[],
	bool useResponsesProtocol)
{
	if (argc < 4) {
		PrintUsage();
		return EXIT_FAILURE;
	}

	const char* apiKey = argv[2];
	const char* model = argv[3];
	const char* baseUrl = "https://api.openai.com/v1";
	std::string outputPath;
	std::vector<char> buffer(2 * 1024 * 1024);
	for (int i = 4; i < argc; ++i) {
		const std::string arg = argv[i];
		if (arg == "--out" && i + 1 < argc) {
			outputPath = argv[++i];
		}
		else if (std::string(baseUrl) == "https://api.openai.com/v1") {
			baseUrl = argv[i];
		}
		else {
			PrintUsage();
			return EXIT_FAILURE;
		}
	}

	const int result = useResponsesProtocol
		? AutoLinkerTest_RunOpenAIResponsesIntegrationTest(apiKey, model, baseUrl, buffer.data(), static_cast<int>(buffer.size()))
		: AutoLinkerTest_RunOpenAIChatIntegrationTest(apiKey, model, baseUrl, buffer.data(), static_cast<int>(buffer.size()));
	if (result < 0) {
		return PrintStringResult(useResponsesProtocol ? "openai-responses-test" : "openai-chat-test", result, buffer.data());
	}

	const std::string text(buffer.data(), static_cast<size_t>(result));
	if (!outputPath.empty()) {
		std::error_code ec;
		const std::filesystem::path outputFsPath = MakePathFromText(outputPath);
		const std::filesystem::path parent = outputFsPath.parent_path();
		if (!parent.empty()) {
			std::filesystem::create_directories(parent, ec);
		}
		std::ofstream out(outputFsPath, std::ios::binary | std::ios::trunc);
		if (!out.is_open()) {
			std::cerr << (useResponsesProtocol ? "openai-responses-test" : "openai-chat-test")
				<< " failed: cannot open output file: " << outputPath << std::endl;
			return EXIT_FAILURE;
		}
		out.write(text.data(), static_cast<std::streamsize>(text.size()));
	}

	std::cout << text << std::endl;
	try {
		const nlohmann::json parsed = nlohmann::json::parse(text);
		return parsed.value("ok", false) ? EXIT_SUCCESS : EXIT_FAILURE;
	}
	catch (...) {
		return EXIT_FAILURE;
	}
}

int RunGeminiIntegrationCommand(int argc, char* argv[])
{
	if (argc < 4) {
		PrintUsage();
		return EXIT_FAILURE;
	}

	const char* apiKey = argv[2];
	const char* model = argv[3];
	const char* baseUrl = "https://generativelanguage.googleapis.com";
	std::string outputPath;
	std::vector<char> buffer(2 * 1024 * 1024);
	for (int i = 4; i < argc; ++i) {
		const std::string arg = argv[i];
		if (arg == "--out" && i + 1 < argc) {
			outputPath = argv[++i];
		}
		else if (std::string(baseUrl) == "https://generativelanguage.googleapis.com") {
			baseUrl = argv[i];
		}
		else {
			PrintUsage();
			return EXIT_FAILURE;
		}
	}

	const int result = AutoLinkerTest_RunGeminiIntegrationTest(
		apiKey,
		model,
		baseUrl,
		buffer.data(),
		static_cast<int>(buffer.size()));
	if (result < 0) {
		return PrintStringResult("gemini-model-test", result, buffer.data());
	}

	const std::string text(buffer.data(), static_cast<size_t>(result));
	if (!outputPath.empty()) {
		std::error_code ec;
		const std::filesystem::path outputFsPath = MakePathFromText(outputPath);
		const std::filesystem::path parent = outputFsPath.parent_path();
		if (!parent.empty()) {
			std::filesystem::create_directories(parent, ec);
		}
		std::ofstream out(outputFsPath, std::ios::binary | std::ios::trunc);
		if (!out.is_open()) {
			std::cerr << "gemini-model-test failed: cannot open output file: " << outputPath << std::endl;
			return EXIT_FAILURE;
		}
		out.write(text.data(), static_cast<std::streamsize>(text.size()));
	}

	std::cout << text << std::endl;
	try {
		const nlohmann::json parsed = nlohmann::json::parse(text);
		return parsed.value("ok", false) ? EXIT_SUCCESS : EXIT_FAILURE;
	}
	catch (...) {
		return EXIT_FAILURE;
	}
}

int RunClaudeIntegrationCommand(int argc, char* argv[])
{
	if (argc < 4) {
		PrintUsage();
		return EXIT_FAILURE;
	}

	const char* apiKey = argv[2];
	const char* model = argv[3];
	const char* baseUrl = "https://api.anthropic.com";
	std::string outputPath;
	std::vector<char> buffer(2 * 1024 * 1024);
	for (int i = 4; i < argc; ++i) {
		const std::string arg = argv[i];
		if (arg == "--out" && i + 1 < argc) {
			outputPath = argv[++i];
		}
		else if (std::string(baseUrl) == "https://api.anthropic.com") {
			baseUrl = argv[i];
		}
		else {
			PrintUsage();
			return EXIT_FAILURE;
		}
	}

	const int result = AutoLinkerTest_RunClaudeIntegrationTest(
		apiKey,
		model,
		baseUrl,
		buffer.data(),
		static_cast<int>(buffer.size()));
	if (result < 0) {
		return PrintStringResult("claude-model-test", result, buffer.data());
	}

	const std::string text(buffer.data(), static_cast<size_t>(result));
	if (!outputPath.empty()) {
		std::error_code ec;
		const std::filesystem::path outputFsPath = MakePathFromText(outputPath);
		const std::filesystem::path parent = outputFsPath.parent_path();
		if (!parent.empty()) {
			std::filesystem::create_directories(parent, ec);
		}
		std::ofstream out(outputFsPath, std::ios::binary | std::ios::trunc);
		if (!out.is_open()) {
			std::cerr << "claude-model-test failed: cannot open output file: " << outputPath << std::endl;
			return EXIT_FAILURE;
		}
		out.write(text.data(), static_cast<std::streamsize>(text.size()));
	}

	std::cout << text << std::endl;
	try {
		const nlohmann::json parsed = nlohmann::json::parse(text);
		return parsed.value("ok", false) ? EXIT_SUCCESS : EXIT_FAILURE;
	}
	catch (...) {
		return EXIT_FAILURE;
	}
}

int RunStringCommand(const std::string& commandName, const std::string& input)
{
	char buffer[524288] = {};
	int result = AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;

	if (commandName == "linker-out") {
		result = AutoLinkerTest_GetLinkerOutFileName(input.c_str(), buffer, static_cast<int>(sizeof(buffer)));
	}
	else if (commandName == "linker-krnln") {
		result = AutoLinkerTest_GetLinkerKrnlnFileName(input.c_str(), buffer, static_cast<int>(sizeof(buffer)));
	}
	else if (commandName == "between-dashes") {
		result = AutoLinkerTest_ExtractBetweenDashes(input.c_str(), buffer, static_cast<int>(sizeof(buffer)));
	}
	else if (commandName == "version-text") {
		result = AutoLinkerTest_GetVersionText(buffer, static_cast<int>(sizeof(buffer)));
	}
	else if (commandName == "module-local-dump") {
		result = AutoLinkerTest_DumpLocalModulePublicInfo(input.c_str(), buffer, static_cast<int>(sizeof(buffer)));
	}

	if (result < 0) {
		return PrintStringResult(commandName.c_str(), result, buffer);
	}

	std::cout << buffer << std::endl;
	return EXIT_SUCCESS;
}

void PrintUsage()
{
	std::cout << "AutoLinkerTest commands:" << std::endl;
	std::cout << "  AutoLinkerTest" << std::endl;
	std::cout << "  AutoLinkerTest version-compare <left> <right>" << std::endl;
	std::cout << "  AutoLinkerTest linker-out <link-command>" << std::endl;
	std::cout << "  AutoLinkerTest linker-krnln <link-command>" << std::endl;
	std::cout << "  AutoLinkerTest between-dashes <text>" << std::endl;
	std::cout << "  AutoLinkerTest version-text" << std::endl;
	std::cout << "  AutoLinkerTest module-local-dump <module-path>" << std::endl;
	std::cout << "  AutoLinkerTest e2txt-generate <input-path> <output-path>" << std::endl;
	std::cout << "  AutoLinkerTest e2txt-restore <input-path> <output-path>" << std::endl;
	std::cout << "  AutoLinkerTest bundle-digest-compare <input.e> <input-dir>" << std::endl;
	std::cout << "  AutoLinkerTest deepseek-model-test <api-key> <model> [base-url] [--out result.json]" << std::endl;
	std::cout << "  AutoLinkerTest openai-chat-test <api-key> <model> [base-url] [--out result.json]" << std::endl;
	std::cout << "  AutoLinkerTest openai-responses-test <api-key> <model> [base-url] [--out result.json]" << std::endl;
	std::cout << "  AutoLinkerTest gemini-model-test <api-key> <model> [base-url] [--out result.json]" << std::endl;
	std::cout << "  AutoLinkerTest claude-model-test <api-key> <model> [base-url] [--out result.json]" << std::endl;
	std::cout << "  AutoLinkerTest headless-compile <e.exe> <input.e> <output> [--target auto|win_exe|win_console_exe|win_dll|ecom] [--static] [--result path] [--timeout seconds]" << std::endl;
}

}

int main(int argc, char* argv[])
{
	if (argc == 1) {
		return RunSmokeTest();
	}

	const std::string commandName = argv[1];
	if (commandName == "headless-compile") {
		return RunHeadlessCompile(argc, argv);
	}
	if (commandName == "deepseek-model-test") {
		return RunDeepSeekModelIntegrationTest(argc, argv);
	}
	if (commandName == "openai-chat-test") {
		return RunOpenAIIntegrationCommand(argc, argv, false);
	}
	if (commandName == "openai-responses-test") {
		return RunOpenAIIntegrationCommand(argc, argv, true);
	}
	if (commandName == "gemini-model-test") {
		return RunGeminiIntegrationCommand(argc, argv);
	}
	if (commandName == "claude-model-test") {
		return RunClaudeIntegrationCommand(argc, argv);
	}

	if (commandName == "version-compare") {
		if (argc != 4) {
			PrintUsage();
			return EXIT_FAILURE;
		}
		return RunVersionCompare(argv[2], argv[3]);
	}

	if (commandName == "linker-out" || commandName == "linker-krnln" || commandName == "between-dashes" || commandName == "module-local-dump") {
		if (argc < 3) {
			PrintUsage();
			return EXIT_FAILURE;
		}
		return RunStringCommand(commandName, JoinArguments(2, argc, argv));
	}

	if (commandName == "version-text") {
		if (argc != 2) {
			PrintUsage();
			return EXIT_FAILURE;
		}
		char buffer[524288] = {};
		const int result = AutoLinkerTest_GetVersionText(buffer, static_cast<int>(sizeof(buffer)));
		if (result < 0) {
			return PrintStringResult(commandName.c_str(), result, buffer);
		}
		std::cout << buffer << std::endl;
		return EXIT_SUCCESS;
	}

	if (commandName == "e2txt-generate") {
		if (argc != 4) {
			PrintUsage();
			return EXIT_FAILURE;
		}

		char buffer[524288] = {};
		const int result = AutoLinkerTest_GenerateE2Txt(argv[2], argv[3], buffer, static_cast<int>(sizeof(buffer)));
		if (result < 0) {
			return PrintStringResult(commandName.c_str(), result, buffer);
		}
		std::cout << buffer << std::endl;
		return EXIT_SUCCESS;
	}

	if (commandName == "e2txt-restore") {
		if (argc != 4) {
			PrintUsage();
			return EXIT_FAILURE;
		}

		char buffer[524288] = {};
		const int result = AutoLinkerTest_RestoreE2Txt(argv[2], argv[3], buffer, static_cast<int>(sizeof(buffer)));
		if (result < 0) {
			return PrintStringResult(commandName.c_str(), result, buffer);
		}
		std::cout << buffer << std::endl;
		return EXIT_SUCCESS;
	}

	if (commandName == "bundle-digest-compare") {
		if (argc != 4) {
			PrintUsage();
			return EXIT_FAILURE;
		}

		char buffer[524288] = {};
		const int result = AutoLinkerTest_CompareBundleDigest(argv[2], argv[3], buffer, static_cast<int>(sizeof(buffer)));
		if (result < 0) {
			return PrintStringResult(commandName.c_str(), result, buffer);
		}
		std::cout << buffer << std::endl;
		return EXIT_SUCCESS;
	}

	PrintUsage();
	return EXIT_FAILURE;
}
