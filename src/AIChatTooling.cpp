#include "AIChatTooling.h"
#include "AIChatToolingInternal.h"
#include "AIService.h"
#include "ConfigManager.h"
#include "Global.h"
#include "LocalMcpInstanceRegistry.h"
#include "LocalMcpServer.h"
#include "PowerShellToolRunner.h"
#include "TavilyClient.h"
#include "WebDocumentClient.h"
#include "WebDocumentExtractor.h"
#include "WinINetUtil.h"

#include <Windows.h>

#include <algorithm>
#include <chrono>
#include <cctype>
#include <condition_variable>
#include <format>
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

void ReplaceAllInPlace(std::string& text, const std::string& from, const std::string& to)
{
	if (from.empty()) {
		return;
	}

	size_t pos = 0;
	while ((pos = text.find(from, pos)) != std::string::npos) {
		text.replace(pos, from.size(), to);
		pos += to.size();
	}
}

std::string SanitizeSingleLineText(std::string text)
{
	ReplaceAllInPlace(text, "\\r\\n", " ");
	ReplaceAllInPlace(text, "\\n", " ");
	ReplaceAllInPlace(text, "\\r", " ");
	ReplaceAllInPlace(text, "\\t", " ");
	ReplaceAllInPlace(text, "\r\n", " ");
	ReplaceAllInPlace(text, "\n", " ");
	ReplaceAllInPlace(text, "\r", " ");
	ReplaceAllInPlace(text, "\t", " ");

	std::string collapsed;
	collapsed.reserve(text.size());
	bool previousWhitespace = false;
	for (unsigned char ch : text) {
		if (std::isspace(ch) != 0) {
			if (!previousWhitespace) {
				collapsed.push_back(' ');
				previousWhitespace = true;
			}
			continue;
		}
		collapsed.push_back(static_cast<char>(ch));
		previousWhitespace = false;
	}
	return TrimAsciiCopy(collapsed);
}

bool TryDecodeTextToWide(const std::string& text, std::wstring& outWide)
{
	outWide.clear();
	if (text.empty()) {
		return true;
	}

	int wideLen = MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0);
	UINT codePage = CP_UTF8;
	DWORD flags = MB_ERR_INVALID_CHARS;
	if (wideLen <= 0) {
		wideLen = MultiByteToWideChar(
			CP_ACP,
			0,
			text.data(),
			static_cast<int>(text.size()),
			nullptr,
			0);
		codePage = CP_ACP;
		flags = 0;
		if (wideLen <= 0) {
			return false;
		}
	}

	outWide.assign(static_cast<size_t>(wideLen), L'\0');
	if (MultiByteToWideChar(
		codePage,
		flags,
		text.data(),
		static_cast<int>(text.size()),
		outWide.data(),
		wideLen) <= 0) {
		outWide.clear();
		return false;
	}
	return true;
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

std::string ConvertUtf8ToGbkText(const std::string& text)
{
	if (text.empty()) {
		return std::string();
	}
	if (!IsValidUtf8Text(text)) {
		return text;
	}
	return ConvertCodePage(text, CP_UTF8, 936, MB_ERR_INVALID_CHARS);
}

std::string TruncateToolLogText(const std::string& text, size_t maxChars = 180)
{
	std::wstring wide;
	if (TryDecodeTextToWide(text, wide)) {
		if (wide.size() <= maxChars) {
			return text;
		}
		const std::string truncated = WideToUtf8(wide.substr(0, maxChars));
		if (!truncated.empty()) {
			return truncated + "...";
		}
	}

	if (text.size() <= maxChars) {
		return text;
	}
	return text.substr(0, maxChars) + "...";
}

std::string FormatToolLogText(const std::string& text)
{
	return TruncateToolLogText(SanitizeSingleLineText(text), 180);
}

std::string FormatToolLogJsonString(const std::string& jsonText)
{
	const std::string trimmed = TrimAsciiCopy(jsonText);
	if (trimmed.empty()) {
		return "null";
	}

	try {
		const nlohmann::json value = nlohmann::json::parse(trimmed);
		if (value.is_null() || (value.is_object() && value.empty())) {
			return "null";
		}
		return FormatToolLogText(value.dump());
	}
	catch (...) {
		return FormatToolLogText(jsonText);
	}
}

void LogInternalToolCallLine(const std::string& message)
{
	OutputStringToELog(ConvertUtf8ToGbkText("[Tool] " + message));
}

void LogInternalToolRequest(const std::string& toolName, const std::string& argumentsJson)
{
	LogInternalToolCallLine(">> " + toolName + "(" + FormatToolLogJsonString(argumentsJson) + ")");
}

void LogInternalToolResponse(const std::string& toolName, const std::string& resultJsonLocal, double elapsedMs)
{
	const std::string resultJsonUtf8 = LocalToUtf8Text(resultJsonLocal);
	LogInternalToolCallLine(std::format(
		"<< {} ({:.1f}ms) {}",
		toolName.empty() ? "unknown_tool" : toolName,
		elapsedMs,
		FormatToolLogJsonString(resultJsonUtf8)));
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

std::string BuildEndpointFromPortLocal(int port)
{
	if (port <= 0) {
		return std::string();
	}
	return std::format("http://127.0.0.1:{}/mcp", port);
}

bool TryResolveLocalInstanceTarget(
	const std::string& instanceId,
	int port,
	LocalMcpInstanceRegistry::InstanceRecord& outRecord,
	std::string& outError)
{
	outRecord = {};
	outError.clear();

	std::vector<LocalMcpInstanceRegistry::InstanceRecord> records;
	if (!LocalMcpInstanceRegistry::LoadInstances(records, &outError)) {
		if (outError.empty()) {
			outError = "load local mcp instances failed";
		}
		return false;
	}

	for (const auto& record : records) {
		if (!instanceId.empty() && record.instanceId == instanceId) {
			outRecord = record;
			return true;
		}
		if (instanceId.empty() && port > 0 && record.port == port) {
			outRecord = record;
			return true;
		}
	}

	outError = "target instance not found";
	return false;
}

std::string BuildForwardedToolResultJson(
	const LocalMcpInstanceRegistry::InstanceRecord& target,
	const std::string& toolName,
	const nlohmann::json& rpcResponse,
	bool& outForwardOk)
{
	outForwardOk = false;
	nlohmann::json result;
	result["ok"] = false;
	result["target_instance_id"] = target.instanceId;
	result["target_process_id"] = target.processId;
	result["target_port"] = target.port;
	result["target_endpoint"] = target.endpoint.empty() ? BuildEndpointFromPortLocal(target.port) : target.endpoint;
	result["tool_name"] = toolName;

	if (!rpcResponse.is_object()) {
		result["error"] = "invalid forwarded rpc response";
		return Utf8ToLocalText(result.dump());
	}

	if (rpcResponse.contains("error")) {
		result["error"] = rpcResponse["error"];
		if (rpcResponse.contains("id")) {
			result["rpc_id"] = rpcResponse["id"];
		}
		return Utf8ToLocalText(result.dump());
	}

	if (!rpcResponse.contains("result") || !rpcResponse["result"].is_object()) {
		result["error"] = "forwarded rpc result missing";
		return Utf8ToLocalText(result.dump());
	}

	const nlohmann::json rpcResult = rpcResponse["result"];
	const bool mcpIsError = rpcResult.value("isError", false);
	result["mcp_is_error"] = mcpIsError;

	if (rpcResult.contains("structuredContent")) {
		result["tool_result"] = rpcResult["structuredContent"];
	}
	else if (rpcResult.contains("content") && rpcResult["content"].is_array() && !rpcResult["content"].empty()) {
		const nlohmann::json& firstContent = rpcResult["content"].front();
		if (firstContent.is_object() &&
			firstContent.value("type", std::string()) == "text" &&
			firstContent.contains("text") &&
			firstContent["text"].is_string()) {
			const std::string text = firstContent["text"].get<std::string>();
			try {
				result["tool_result"] = nlohmann::json::parse(text);
			}
			catch (...) {
				result["tool_result_text"] = text;
			}
		}
	}

	if (rpcResult.contains("content")) {
		result["raw_content"] = rpcResult["content"];
	}

	outForwardOk = !mcpIsError;
	result["ok"] = outForwardOk;
	return Utf8ToLocalText(result.dump());
}

std::string ExecuteToolCallImpl(const std::string& toolName, const std::string& argumentsJson, bool& outOk)
{
	outOk = false;

	if (toolName == "list_local_mcp_instances") {
		std::vector<LocalMcpInstanceRegistry::InstanceRecord> records;
		std::string error;
		if (!LocalMcpInstanceRegistry::LoadInstances(records, &error)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = error.empty() ? "load local mcp instances failed" : error;
			r["registry_file"] = LocalMcpInstanceRegistry::GetRegistryFilePath();
			return Utf8ToLocalText(r.dump());
		}

		nlohmann::json rows = nlohmann::json::array();
		const std::string currentInstanceId = LocalMcpServer::GetInstanceId();
		for (const auto& record : records) {
			rows.push_back({
				{"instance_id", record.instanceId},
				{"is_current", !currentInstanceId.empty() && record.instanceId == currentInstanceId},
				{"process_id", record.processId},
				{"process_path", record.processPath},
				{"process_name", record.processName},
				{"port", record.port},
				{"endpoint", record.endpoint.empty() ? BuildEndpointFromPortLocal(record.port) : record.endpoint},
				{"source_file_path_hint", record.sourceFilePathHint},
				{"page_name_hint", record.pageNameHint},
				{"page_type_hint", record.pageTypeHint},
				{"last_seen_unix_ms", record.lastSeenUnixMs}
			});
		}

		nlohmann::json r;
		r["ok"] = true;
		r["registry_file"] = LocalMcpInstanceRegistry::GetRegistryFilePath();
		r["current_instance_id"] = currentInstanceId;
		r["count"] = records.size();
		r["instances"] = std::move(rows);
		outOk = true;
		return Utf8ToLocalText(r.dump());
	}

	if (toolName == "call_local_mcp_instance_tool") {
		nlohmann::json args;
		try {
			args = argumentsJson.empty() ? nlohmann::json::object() : nlohmann::json::parse(argumentsJson);
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("invalid arguments json: ") + ex.what();
			return Utf8ToLocalText(r.dump());
		}

		const std::string instanceId =
			args.contains("instance_id") && args["instance_id"].is_string()
			? args["instance_id"].get<std::string>()
			: std::string();
		const int port =
			args.contains("port") && args["port"].is_number_integer()
			? args["port"].get<int>()
			: 0;
		const std::string targetToolName =
			args.contains("tool_name") && args["tool_name"].is_string()
			? args["tool_name"].get<std::string>()
			: std::string();
		const int timeoutSeconds =
			args.contains("timeout_seconds") && args["timeout_seconds"].is_number_integer()
			? (std::clamp)(args["timeout_seconds"].get<int>(), 1, 120)
			: 30;
		const nlohmann::json targetArguments =
			args.contains("arguments") ? args["arguments"] : nlohmann::json::object();

		if (instanceId.empty() && port <= 0) {
			return R"({"ok":false,"error":"instance_id or port is required"})";
		}
		if (TrimAsciiCopy(targetToolName).empty()) {
			return R"({"ok":false,"error":"tool_name is required"})";
		}
		if (targetToolName == "call_local_mcp_instance_tool") {
			return R"({"ok":false,"error":"recursive forwarding of call_local_mcp_instance_tool is not allowed"})";
		}

		LocalMcpInstanceRegistry::InstanceRecord target;
		std::string resolveError;
		if (!TryResolveLocalInstanceTarget(instanceId, port, target, resolveError)) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = resolveError.empty() ? "target instance not found" : resolveError;
			r["requested_instance_id"] = instanceId;
			r["requested_port"] = port;
			r["registry_file"] = LocalMcpInstanceRegistry::GetRegistryFilePath();
			return Utf8ToLocalText(r.dump());
		}

		const std::string endpoint = target.endpoint.empty() ? BuildEndpointFromPortLocal(target.port) : target.endpoint;
		nlohmann::json rpcRequest = {
			{"jsonrpc", "2.0"},
			{"id", std::format("forward-{}-{}", GetCurrentProcessId(), GetTickCount64())},
			{"method", "tools/call"},
			{"params", {
				{"name", targetToolName},
				{"arguments", targetArguments}
			}}
		};

		const auto httpResult = PerformPostRequest(
			endpoint,
			rpcRequest.dump(),
			"Content-Type: application/json\r\n",
			timeoutSeconds * 1000,
			true,
			true);
		if (httpResult.second < 200 || httpResult.second >= 300) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = "forward http request failed";
			r["http_status"] = httpResult.second;
			r["target_instance_id"] = target.instanceId;
			r["target_port"] = target.port;
			r["target_endpoint"] = endpoint;
			r["response_text"] = LocalToUtf8Text(httpResult.first);
			return Utf8ToLocalText(r.dump());
		}

		try {
			const nlohmann::json rpcResponse = httpResult.first.empty()
				? nlohmann::json::object()
				: nlohmann::json::parse(httpResult.first);
			return BuildForwardedToolResultJson(target, targetToolName, rpcResponse, outOk);
		}
		catch (const std::exception& ex) {
			nlohmann::json r;
			r["ok"] = false;
			r["error"] = std::string("parse forwarded rpc response failed: ") + ex.what();
			r["target_instance_id"] = target.instanceId;
			r["target_port"] = target.port;
			r["target_endpoint"] = endpoint;
			r["response_text"] = LocalToUtf8Text(httpResult.first);
			return Utf8ToLocalText(r.dump());
		}
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
		toolName == "get_current_eide_info" ||
		toolName == "list_imported_modules" ||
		toolName == "list_support_libraries" ||
		toolName == "get_support_library_info" ||
		toolName == "search_support_library_info" ||
		toolName == "search_support_library_public_code" ||
		toolName == "read_support_library_public_code" ||
		toolName == "get_module_public_info" ||
		toolName == "search_module_public_info" ||
		toolName == "search_module_public_code" ||
		toolName == "read_module_public_code" ||
		toolName == "list_program_items" ||
		toolName == "get_program_item_real_code" ||
		toolName == "read_program_item_real_code" ||
		toolName == "edit_program_item_code" ||
		toolName == "multi_edit_program_item_code" ||
		toolName == "write_program_item_real_code" ||
		toolName == "diff_program_item_code" ||
		toolName == "restore_program_item_code_snapshot" ||
		toolName == "search_program_item_real_code" ||
		toolName == "list_program_item_symbols" ||
		toolName == "get_symbol_real_code" ||
		toolName == "edit_symbol_real_code" ||
		toolName == "insert_program_item_code_block" ||
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

} // namespace

std::string ExecuteToolCall(const std::string& toolName, const std::string& argumentsJson, bool& outOk, bool enableLog)
{
	if (!enableLog) {
		return ExecuteToolCallImpl(toolName, argumentsJson, outOk);
	}

	LogInternalToolRequest(toolName, argumentsJson);
	const auto startTime = std::chrono::steady_clock::now();
	const std::string result = ExecuteToolCallImpl(toolName, argumentsJson, outOk);
	const double elapsedMs = std::chrono::duration<double, std::milli>(
		std::chrono::steady_clock::now() - startTime).count();
	LogInternalToolResponse(toolName, result, elapsedMs);
	return result;
}

