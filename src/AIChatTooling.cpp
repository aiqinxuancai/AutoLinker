#include "AIChatTooling.h"
#include "AIChatToolingInternal.h"
#include "AIService.h"
#include "ConfigManager.h"
#include "PowerShellToolRunner.h"
#include "TavilyClient.h"
#include "WebDocumentClient.h"
#include "WebDocumentExtractor.h"

#include <Windows.h>

#include <algorithm>
#include <chrono>
#include <cctype>
#include <condition_variable>
#include <mutex>
#include <string>

#include "..\\thirdparty\\json.hpp"

namespace {std::string TrimAsciiCopy(const std::string& text)
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

std::string ToLowerAsciiCopyLocal(std::string text)
{
	for (char& ch : text) {
		ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
	}
	return text;
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
bool RequestToolExecutionFromMainThread(
	const std::string& toolName,
	const std::string& argumentsJson,
	std::string& outResultJson,
	bool& outOk)
{
	outResultJson.clear();
	outOk = false;
	const HWND mainWindow = GetAIChatMainWindowForTooling();
	const UINT toolExecMessage = GetAIChatToolExecMessageForTooling();
	if (mainWindow == nullptr || !IsWindow(mainWindow) || toolExecMessage == 0) {
		return false;
	}

	ToolExecutionRequest request;
	request.toolName = toolName;
	request.argumentsJson = argumentsJson;
	if (PostMessage(mainWindow, toolExecMessage, 0, reinterpret_cast<LPARAM>(&request)) == FALSE) {
		return false;
	}

	std::unique_lock<std::mutex> lock(request.mutex);
	if (!request.cv.wait_for(lock, std::chrono::minutes(20), [&request]() { return request.done; })) {
		return false;
	}

	outResultJson = request.resultJson;
	outOk = request.ok;
	return true;
}

} // namespace
std::string ExecuteToolCall(const std::string& toolName, const std::string& argumentsJson, bool& outOk)
{
	outOk = false;

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
		if (!RequestCodeEditForTooling(title, hint, initialCode, editedCode)) {
			return R"({"ok":false,"cancelled":true})";
		}

		nlohmann::json r;
		r["ok"] = true;
		r["code"] = LocalToUtf8Text(editedCode);
		outOk = true;
		return Utf8ToLocalText(r.dump());
	}

	if (toolName == "run_powershell_command") {
		std::string commandUtf8;
		std::string workingDirectoryUtf8;
		int timeoutSeconds = 60;
		try {
			const nlohmann::json args = nlohmann::json::parse(argumentsJson);
			if (args.contains("command") && args["command"].is_string()) {
				commandUtf8 = args["command"].get<std::string>();
			}
			if (args.contains("working_directory") && args["working_directory"].is_string()) {
				workingDirectoryUtf8 = args["working_directory"].get<std::string>();
			}
			if (args.contains("timeout_seconds") && args["timeout_seconds"].is_number_integer()) {
				timeoutSeconds = (std::clamp)(args["timeout_seconds"].get<int>(), 1, 600);
			}
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return Utf8ToLocalText(r.dump());
		}

		if (TrimAsciiCopy(commandUtf8).empty()) {
			return R"({"ok":false,"error":"command is required"})";
		}

		std::string confirmationText =
			std::string("即将执行 PowerShell 命令：\r\n\r\n") +
			Utf8ToLocalText(commandUtf8) +
			"\r\n\r\n工作目录：\r\n" +
			(TrimAsciiCopy(workingDirectoryUtf8).empty() ? std::string("(当前进程目录)") : Utf8ToLocalText(workingDirectoryUtf8)) +
			"\r\n\r\n超时：\r\n" +
			std::to_string(timeoutSeconds) +
			" 秒\r\n\r\n请确认该命令不会造成你不希望的本机副作用。";
		bool accepted = false;
		if (!RequestConfirmationForTooling(
				LocalFromWide(L"AI PowerShell 执行确认"),
				confirmationText,
				LocalFromWide(L"执行"),
				LocalFromWide(L"取消"),
				accepted) ||
			!accepted) {
			nlohmann::json r;
			r["ok"] = false;
			r["cancelled"] = true;
			r["error"] = "user cancelled powershell execution";
			return Utf8ToLocalText(r.dump());
		}

		const PowerShellRunResult runResult = PowerShellToolRunner::Run(commandUtf8, workingDirectoryUtf8, timeoutSeconds);
		nlohmann::json r;
		r["ok"] = runResult.ok;
		r["cancelled"] = false;
		r["command"] = commandUtf8;
		r["working_directory"] = runResult.effectiveWorkingDirectory;
		r["stdout"] = runResult.stdOut;
		r["stderr"] = runResult.stdErr;
		r["exit_code"] = runResult.exitCode;
		r["timed_out"] = runResult.timedOut;
		if (!runResult.error.empty()) {
			r["error"] = runResult.error;
		}
		outOk = runResult.ok;
		return Utf8ToLocalText(r.dump());
	}

	if (toolName == "search_web_tavily") {
		std::string queryUtf8;
		std::string topicUtf8;
		int maxResults = 5;
		try {
			const nlohmann::json args = nlohmann::json::parse(argumentsJson);
			if (args.contains("query") && args["query"].is_string()) {
				queryUtf8 = args["query"].get<std::string>();
			}
			if (args.contains("topic") && args["topic"].is_string()) {
				topicUtf8 = args["topic"].get<std::string>();
			}
			if (args.contains("max_results") && args["max_results"].is_number_integer()) {
				maxResults = (std::clamp)(args["max_results"].get<int>(), 1, 10);
			}
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return Utf8ToLocalText(r.dump());
		}

		AISettings settings = {};
		std::string tavilyApiKey;
		ConfigManager* configManager = GetAIChatConfigManagerForTooling();
		if (configManager != nullptr && AIService::LoadSettings(*configManager, settings)) {
			tavilyApiKey = LocalToUtf8Text(settings.tavilyApiKey);
		}

		const TavilySearchResult searchResult = TavilyClient::Search(tavilyApiKey, queryUtf8, maxResults, topicUtf8);
		if (!searchResult.ok) {
			nlohmann::json r;
			r["ok"] = false;
			r["http_status"] = searchResult.httpStatus;
			r["error"] = searchResult.error;
			return Utf8ToLocalText(r.dump());
		}

		outOk = true;
		return Utf8ToLocalText(searchResult.normalizedResultJsonUtf8);
	}

	if (toolName == "fetch_url") {
		std::string urlUtf8;
		int timeoutSeconds = 60;
		size_t maxBytes = 512 * 1024;
		try {
			const nlohmann::json args = nlohmann::json::parse(argumentsJson);
			if (args.contains("url") && args["url"].is_string()) {
				urlUtf8 = args["url"].get<std::string>();
			}
			if (args.contains("timeout_seconds") && args["timeout_seconds"].is_number_integer()) {
				timeoutSeconds = (std::clamp)(args["timeout_seconds"].get<int>(), 1, 300);
			}
			if (args.contains("max_bytes") && args["max_bytes"].is_number_integer()) {
				const int value = args["max_bytes"].get<int>();
				maxBytes = (std::clamp)(value, 4096, 2097152);
			}
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return Utf8ToLocalText(r.dump());
		}

		const HttpFetchResult fetchResult = WebDocumentClient::FetchTextUrl(urlUtf8, timeoutSeconds, maxBytes);
		nlohmann::json r;
		r["ok"] = fetchResult.ok;
		r["url"] = fetchResult.url;
		r["final_url"] = fetchResult.finalUrl;
		r["http_status"] = fetchResult.httpStatus;
		r["content_type"] = fetchResult.contentType;
		r["content_length"] = static_cast<unsigned long long>(fetchResult.contentLength);
		r["body_text"] = fetchResult.bodyText;
		r["body_truncated"] = fetchResult.bodyTruncated;
		if (!fetchResult.error.empty()) {
			r["error"] = fetchResult.error;
		}
		outOk = fetchResult.ok;
		return Utf8ToLocalText(r.dump());
	}

	if (toolName == "extract_web_document") {
		std::string urlUtf8;
		int timeoutSeconds = 60;
		size_t maxBytes = 512 * 1024;
		try {
			const nlohmann::json args = nlohmann::json::parse(argumentsJson);
			if (args.contains("url") && args["url"].is_string()) {
				urlUtf8 = args["url"].get<std::string>();
			}
			if (args.contains("timeout_seconds") && args["timeout_seconds"].is_number_integer()) {
				timeoutSeconds = (std::clamp)(args["timeout_seconds"].get<int>(), 1, 300);
			}
			if (args.contains("max_bytes") && args["max_bytes"].is_number_integer()) {
				const int value = args["max_bytes"].get<int>();
				maxBytes = (std::clamp)(value, 4096, 2097152);
			}
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return Utf8ToLocalText(r.dump());
		}

		const HttpFetchResult fetchResult = WebDocumentClient::FetchTextUrl(urlUtf8, timeoutSeconds, maxBytes);
		const ExtractedWebDocument document = WebDocumentExtractor::Extract(fetchResult);

		nlohmann::json links = nlohmann::json::array();
		for (const auto& link : document.links) {
			links.push_back({
				{"text", link.text},
				{"url", link.url}
			});
		}

		nlohmann::json r;
		r["ok"] = document.ok;
		r["url"] = document.url;
		r["http_status"] = document.httpStatus;
		r["content_type"] = document.contentType;
		r["title"] = document.title;
		r["plain_text"] = document.plainText;
		r["excerpt"] = document.excerpt;
		r["links"] = std::move(links);
		if (!document.error.empty()) {
			r["error"] = document.error;
		}
		outOk = document.ok;
		return Utf8ToLocalText(r.dump());
	}

	if (toolName == "get_current_page_code" ||
		toolName == "get_current_page_info" ||
		toolName == "list_imported_modules" ||
		toolName == "list_support_libraries" ||
		toolName == "get_support_library_info" ||
		toolName == "search_support_library_info" ||
		toolName == "get_module_public_info" ||
		toolName == "search_module_public_info" ||
		toolName == "list_program_items" ||
		toolName == "get_program_item_code" ||
		toolName == "switch_to_program_item_page" ||
		toolName == "search_project_keyword" ||
		toolName == "jump_to_search_result" ||
		toolName == "compile_with_output_path") {
		std::string resultJson;
		if (!RequestToolExecutionFromMainThread(toolName, argumentsJson, resultJson, outOk)) {
			return R"({"ok":false,"error":"main thread tool execution failed"})";
		}
		return resultJson;
	}

	nlohmann::json r;
	r["ok"] = false;
	r["error"] = "unknown tool: " + toolName;
	return Utf8ToLocalText(r.dump());
}

